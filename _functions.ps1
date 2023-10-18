Write-Host "Loading script '.\_functions.ps1'..." -NoNewline

#========================================================
#InfoLog output
#========================================================
function WriteInfoLog($message)
{   
   
    Write-Host $message -ForegroundColor White 
    Write-Log -Identity "MessageTrace" -Level Info -Message $message
}

#========================================================
#WarnLog output
#========================================================
function WriteWarnLog($message)
{  
    Write-Host $message -ForegroundColor Yellow
    Write-Log -Identity "MessageTrace" -Level Warning -Message $message
}

#========================================================
#ErrorLog output
#========================================================
function WriteErrorLog($message)
{   
    Write-Host $message -ForegroundColor Red
    Write-Log -Identity "MessageTrace" -Level Error -Message $message
}

#======================================================================
#Ensure the correct PS modules are installed to run the script correctly
#=======================================================================
function Install-RequiredModules {
    # Array of required modules
    $requiredModules = @('PnP.PowerShell','AzureAD','ExchangeOnlineManagement')

    foreach ($module in $requiredModules) {
        # Check if the module is installed
        if (Get-Module -ListAvailable -Name $module) {
            Write-Host "$module module is already installed." -ForegroundColor Green
        } else {
            # Prompt the user if they wish to install the module
            $install = Read-Host -Prompt "$module is not installed. Do you want to install it now? (yes/no)"

            if ($install -eq "yes") {
                try {
                    Install-Module -Name $module -AllowClobber -Scope CurrentUser -Force -Confirm:$false
                    Write-Host "$module module has been installed successfully!" -ForegroundColor Green
                } catch {
                    Write-Host "There was an error installing the $module module. Please run as Administrator or check internet connectivity." -ForegroundColor Red
                    exit
                }
            } else {
                Write-Host "The script cannot proceed without $module. Exiting..." -ForegroundColor Red
                exit
            }
        }
    }
}


#========================================================
#Generate the sitecollection.csv file to interrogate
#========================================================
function Get-SiteCollections {
    $sites = @();

    # Define the path to the CSV
    $csvPath = 'sitecollection.csv'

    # Check if the CSV file exists, and if it does, delete it
    if (Test-Path $csvPath) {
        Remove-Item $csvPath -Force
    }

    # Connect to SharePoint Online Admin Center
    $adminUrl = "https://$TenantD-admin.sharepoint.com"
    Connect-PnPOnline -Url $adminUrl -ClientId $appId -Thumbprint $thumbprint -Tenant $Tenant

    # Get site collections
    if ($scopeSiteCollectionsFilePath)
    {
        $sites=Import-Csv -Path $scopeSiteCollectionsFilePath 
        foreach ($site in $sites)
        {
            $sites += Get-PnPTenantSite -Identity $site
        }
    }
    else
    {
        $sites = Get-PnPTenantSite
    }

    # Create a generic list for results
    $results = New-Object 'System.Collections.Generic.List[PSCustomObject]'

    foreach ($site in $sites) {
        $result = [PSCustomObject]@{
            "Site Name"            = $site.Title
            "URL"                  = $site.Url
        }

        $results.Add($result)
    }

    # Export results to CSV
    $results | Export-Csv -Path $csvPath -NoTypeInformation

    # Disconnect from SharePoint
    # Disconnect-PnPOnline

    Write-Host "Done fetching site collection details."
}

function ScanList {
    param (
        [string]$listTitle,
        [string]$listRootFolder,
        [string]$listId
    )
    WriteInfoLog "Begin check the list $($listTitle)"
    $storage= Get-PnPFolderStorageMetric -FolderSiteRelativeUrl $listRootFolder
    if($null -ne $storage){
        if($storage.TotalFileCount -eq 0){
            WriteInfoLog "Finished check the empty list $($listTitle)."
            return [System.String]::Format('"{0}","{1}","{2}","{3}","{4}","{5}"',$listTitle,$storage.LastModified,$storage.TotalFileCount,"0","0","0")
        }
        [decimal]$totalsize=0;
        [decimal]::TryParse($storage.TotalSize,[ref]$totalsize);
        [decimal]$currentVersionsize=0;
        [decimal]::TryParse($storage.TotalFileStreamSize,[ref]$currentVersionsize);
        [decimal]$versionSize=$totalsize-$currentVersionsize;
        WriteInfoLog "Finished scan the list $($listTitle)"
        return [System.String]::Format('"{0}","{1}","{2}","{
    else {3}","{4}","{5}"',$listTitle,$storage.LastModified,$storage.TotalFileCount,$storage.TotalFileStreamSize/1024/1024,$storage.TotalSize/1024/1024,$versionSize/1024/1024)
    }
        WriteInfoLog "Finished check the list $($listTitle) with empty storage."
    }
    
    return "";
}

function ScanSiteCollection {
    param (
        [string]$sitecollection
    )
    try {
        # Connect to the specific site collection with MFA
        Connect-PnPOnline -Url $sitecollection -ClientId $appId -Thumbprint $thumbprint -Tenant $Tenant

        $rootWeb = Get-PnPWeb
        ScanSite $rootWeb.Id $sitecollection
        $webs = Get-PnPSubWeb -Recurse;
        foreach($web in $webs){
            try {
                ScanSite $web.Id $sitecollection
            }
 catch {
        WriteErrorLog "An error occurred while scanning the site collection $($sitecollection).Exception:$($_)"
            }
           
        }     
    }
    catch {
        WriteErrorLog "An error occurred while scaning the site collection $($sitecollection).Exception:$($_)" 
    }
    
}

function ScanSite {
   param (
       [string]$webId,
       [string]$sitecollection
    )
    $web = Get-PnPWeb -Identity $webId
    # Use our new connection function
    # Connect to the web
    Connect-PnPOnline -Url $web.Url -ClientId $appId -Thumbprint $thumbprint -Tenant $Tenant
    WriteInfoLog "Begin scan the web $($web.Url)"
    $lists =  Get-PnPList  | Where-Object {$_.BaseTemplate -eq 101 -and $_.Hidden -eq $false}
    foreach ($list in $lists){
        try {
            $folderName=$list.RootFolder.ServerRelativeUrl.Replace($web.ServerRelativeUrl,"")          
            if($folderName -eq "/Style Library"){
                continue;
            }
            if($folderName -eq "/FormServerTemplates"){
                continue;
            } 
            [array]$report=ScanList $list.Title $folderName $list.Id
            if($null -ne $report){          
                $reportLine=[System.String]::Format('"{0}","{1}",{2}',$sitecollection,$web.Url,$report[$report.Count-1])
                WriteReport $reportLine              
            }
            else {
                WriteWarnLog "The report line of the list $($list.Title) is null" 
            }
        }
        catch {
            WriteWarnLog "An error occurred while exporting the list $($list.Title).Exception:$($_)" 
        }
       
    }
    WriteInfoLog "Finished scan the web $($web.Url)"  

}

function WriteReport {
    param (
        [string]$report      
    )
    $report|Out-File -FilePath $outFilePath -Append -Encoding utf8
}

function Write-Log {
        [CmdletBinding()]
        param(
            [Parameter()]
            [ValidateNotNullOrEmpty()]
            [string]$Message,
 
            [Parameter()]
            [ValidateNotNullOrEmpty()]
            [ValidateSet('Info','Warning','Error')]
            [string]$Level='Info',
        
            [Parameter()]
            [string]$Identity
        )
 
        $Stamp=(Get-Date).ToString("yyyy/MM/dd HH:mm:ss:fff")
        $Line="$Stamp    $Identity    $Level    $Message"
        Add-Content $logfile -Value $Line -Encoding UTF8
    }

#========================================================
#Log module
#=======================================================

function Write-InfoWith-Log {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]$Message,
            
            [Parameter(Mandatory=$false)]
            [ValidateSet('Green','Blue','Yellow','Red')]
            [string]$ForegroundColor
        )
       if($ForegroundColor -ne $null){
            Write-Host $Message -ForegroundColor $ForegroundColor
       }
       else{
            Write-Host $Message
       }      
       Write-Log -Message $Message -Level Info -Identity "Output"       
    }

function Write-WarningWith-Log {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]$Message,
            
            [Parameter(Mandatory=$false)]
            [ValidateSet('Green','Blue','Yellow','Red')]
            [string]$ForegroundColor='Yellow'
        )
       if($ForegroundColor -ne $null){
            Write-Host $Message -ForegroundColor $ForegroundColor
       }
       else{
            Write-Host $Message
       }      
       Write-Log -Message $Message -Level Warning -Identity "Output"       
    }

function Write-ErrorWith-Log {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]$Message,
            
            [Parameter(Mandatory=$false)]
            [ValidateSet('Green','Blue','Yellow','Red')]
            [string]$ForegroundColor='Red'
        )
       if($ForegroundColor -ne $null){
            Write-Host $Message -ForegroundColor $ForegroundColor
       }
       else{
            Write-Host $Message
       }      
       Write-Log -Message $Message -Level Error -Identity "Output"       
    }
    
Write-Host 'Success' -ForegroundColor Green
