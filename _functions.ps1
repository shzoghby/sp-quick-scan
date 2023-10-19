Write-Host "Loading script '.\_functions.ps1'..." -NoNewline

function WriteReport {
    param (
        [string]$report      
    )
    $report | Out-File -FilePath $outFilePath -Append -Encoding utf8
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
 
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info',
        
        [Parameter()]
        [string]$Identity
    )
 
    $Stamp = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss:fff")
    $Line = "$Stamp    $Identity    $Level    $Message"
    Add-Content $logfile -Value $Line -Encoding UTF8
}

#========================================================
#Log module
#=======================================================

function Write-InfoWith-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
            
        [Parameter(Mandatory = $false)]
        [ValidateSet('Green', 'Blue', 'Yellow', 'Red')]
        [string]$ForegroundColor
    )
    if ($ForegroundColor -ne $null) {
        Write-Host $Message -ForegroundColor $ForegroundColor
    }
    else {
        Write-Host $Message
    }      
    Write-Log -Message $Message -Level Info -Identity "Output"       
}

function Write-WarningWith-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
            
        [Parameter(Mandatory = $false)]
        [ValidateSet('Green', 'Blue', 'Yellow', 'Red')]
        [string]$ForegroundColor = 'Yellow'
    )
    if ($ForegroundColor -ne $null) {
        Write-Host $Message -ForegroundColor $ForegroundColor
    }
    else {
        Write-Host $Message
    }      
    Write-Log -Message $Message -Level Warning -Identity "Output"       
}

function Write-ErrorWith-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
            
        [Parameter(Mandatory = $false)]
        [ValidateSet('Green', 'Blue', 'Yellow', 'Red')]
        [string]$ForegroundColor = 'Red'
    )
    if ($ForegroundColor -ne $null) {
        Write-Host $Message -ForegroundColor $ForegroundColor
    }
    else {
        Write-Host $Message
    }      
    Write-Log -Message $Message -Level Error -Identity "Output"       
}

#========================================================
#InfoLog output
#========================================================
function WriteInfoLog($message) {   
   
    Write-Host $message -ForegroundColor White 
    Write-Log -Identity "MessageTrace" -Level Info -Message $message
}

#========================================================
#WarnLog output
#========================================================
function WriteWarnLog($message) {  
    Write-Host $message -ForegroundColor Yellow
    Write-Log -Identity "MessageTrace" -Level Warning -Message $message
}

#========================================================
#ErrorLog output
#========================================================
function WriteErrorLog($message) {   
    Write-Host $message -ForegroundColor Red
    Write-Log -Identity "MessageTrace" -Level Error -Message $message
}

#======================================================================
#Ensure the correct PS modules are installed to run the script correctly
#=======================================================================
function Install-RequiredModules {
    # Array of required modules
    $requiredModules = @('PowerShellGet', 'PnP.PowerShell', 'AzureAD', 'ExchangeOnlineManagement')

    foreach ($module in $requiredModules) {
        # Check if the module is installed
        if (Get-Module -ListAvailable -Name $module) {
            Write-Host "$module module is already installed." -ForegroundColor Green
        }
        else {
            # Prompt the user if they wish to install the module
            $install = Read-Host -Prompt "$module is not installed. Do you want to install it now? (yes/no)"

            if ($install -eq "yes") {
                try {
                    Install-Module -Name $module -AllowClobber -Scope CurrentUser -Force -Confirm:$false
                    Save-Module $module -Repository PSGallery -Path "$PSScriptRoot\bin\Modules" -Force
                    Write-Host "$module module has been installed successfully!" -ForegroundColor Green
                }
                catch {
                    Write-Host "There was an error installing the $module module. Please run as Administrator or check internet connectivity." -ForegroundColor Red
                    exit
                }
            }
            else {
                Write-Host "The script cannot proceed without $module. Exiting..." -ForegroundColor Red
                exit
            }
        }
    }
}

#========================================================
#Get site collections from input file or all from tenant
#========================================================
function GetSitecollections {
    param (
        [string]$tenantFullName,
        [string]$clientId,
        [string]$thumbprint,
        [string]$inputFilePath
    )
    $tenantBase = $tenantFullName.SubString(0, $tenantFullName.IndexOf('.'));
    $returnSitecollections = @();   
    
    # Connect to SharePoint Online Admin Center
    $adminUrl = "https://$tenantBase-admin.sharepoint.com"
    Connect-PnPOnline -Url $adminUrl -ClientId $appId -Thumbprint $thumbprint -Tenant $tenantFullName

    # Get site collections
    if ($inputFilePath) {
        $sites = Import-Csv -Path $inputFilePath
        
        foreach ($site in $sites) {
            $siteCollection = Get-PnPTenantSite -Identity $site.url;
            $returnSitecollections += $siteCollection;
        }
    }
    else {
        $returnSitecollections = Get-PnPTenantSite
    }

    return $returnSitecollections;
    Write-Host "Done fetching site collection details."
}


function ScanLists {
    param (
        [string]$sitecollectionUrl,
        [string]$webUrl,
        [string]$listTitle,
        [string]$listRootFolder,
        [string]$listId
    )

    $listReturnObject = $null
    WriteInfoLog "Begin check the list $($listTitle)"
    $storage = Get-PnPFolderStorageMetric -FolderSiteRelativeUrl $listRootFolder
    if ($null -ne $storage) {
        if ($storage.TotalFileCount -eq 0) {
            WriteInfoLog "Finished check the empty list $($listTitle)."

            $listReturnObject = [PSCustomObject]@{
                SitecollectionUrl   = $sitecollectionUrl;
                WebUrl              = $webUrl;
                ListTitle           = $listTitle;
                LastModified        = $storage.LastModified;
                TotalFileCount      = 0;
                TotalFileStreamSize = 0;
                TotalSize           = 0;
                VersionSize         = 0;
            };

            #return [System.String]::Format('"{0}","{1}","{2}","{3}","{4}","{5}"', $listTitle, $storage.LastModified, $storage.TotalFileCount, "0", "0", "0")
        }
        else {
            [decimal]$totalsize = 0;
            [decimal]::TryParse($storage.TotalSize, [ref]$totalsize);
            [decimal]$currentVersionsize = 0;
            [decimal]::TryParse($storage.TotalFileStreamSize, [ref]$currentVersionsize);
            [decimal]$versionSize = $totalsize - $currentVersionsize;
            WriteInfoLog "Finished scan the list $($listTitle)"
            #return [System.String]::Format('"{0}","{1}","{2}","{3}","{4}","{5}"',$listTitle,$storage.LastModified,$storage.TotalFileCount,$storage.TotalFileStreamSize/1024/1024,$storage.TotalSize/1024/1024,$versionSize/1024/1024);

            $listReturnObject = [PSCustomObject]@{
                SitecollectionUrl   = $sitecollectionUrl;
                WebUrl              = $webUrl;
                ListTitle           = $listTitle;
                LastModified        = $storage.LastModified;
                TotalFileCount      = $($storage.TotalFileCount);
                TotalFileStreamSize = $($storage.TotalFileStreamSize / 1024 / 1024);
                TotalSize           = $($storage.TotalSize / 1024 / 1024);
                VersionSize         = $($versionSize / 1024 / 1024);
            };
        }
    }
    else {
        WriteInfoLog "Finished check the list $($listTitle) with empty storage."
    }

    return $listReturnObject;
}

function ScanWebs {
    param (
        [string]$tenantFullName,
        [string]$clientId,
        [string]$thumbprint,
        [string]$webId,
        [string]$webUrl,
        [string]$sitecollectionUrl
    )
    $listReportLines = New-Object 'System.Collections.Generic.List[PSCustomObject]'
    $web = Get-PnPWeb -Identity $webId
    # Use our new connection function
    # Connect to the web
    Connect-PnPOnline -Url $web.Url -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantFullName
    WriteInfoLog "Begin scan the web $($web.Url)"
    $lists = Get-PnPList  | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false }
    foreach ($list in $lists) {
        try {
            $folderName = $list.RootFolder.ServerRelativeUrl.Replace($web.ServerRelativeUrl, "")          
            if ($folderName -eq "/Style Library") {
                continue;
            }
            if ($folderName -eq "/FormServerTemplates") {
                continue;
            } 
            $listReturnObject = ScanLists -sitecollectionUrl $sitecollectionUrl -webUrl $webUrl -listTitle $list.Title -listRootFolder $folderName -listId $list.Id
            if ($null -ne $listReturnObject) {              
                $listReportLines.Add($listReturnObject[$listReturnObject.length-1]);
            }
            else {
                WriteWarnLog "The report line of the list $($list.Title) is null" 
            }
        }
        catch {
            WriteWarnLog "An error occurred while exporting the list $($list.Title).Exception:$($_)" 
        }
        
    }

    return $listReportLines;
    WriteInfoLog "Finished scan the web $($web.Url)"  
 
}

function ScanSiteCollection {
    param (
        [string]$tenantFullName,
        [string]$clientId,
        [string]$thumbprint,
        [string]$url
    )
    try {
        $webReportLines = New-Object 'System.Collections.Generic.List[PSCustomObject]'
        # Connect to the specific site collection with MFA
        Connect-PnPOnline -Url $url -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantFullName

        $rootWeb = Get-PnPWeb
        $webReportLines = ScanWebs -tenantFullName $tenantFullName -ClientId $clientId -Thumbprint $thumbprint -webId $rootWeb.Id -webUrl $rootWeb.Url -sitecollectionUrl $url
        $webs = Get-PnPSubWeb -Recurse;
        foreach ($web in $webs) {
            try {
                $webReportLine = ScanWebs -tenantFullName $tenantFullName -ClientId $clientId -Thumbprint $thumbprint -webId $web.Id -webUrl $web.Url -sitecollectionUrl $url
                $webReportLines.Add($webReportLine)
            }
            catch {
                WriteErrorLog "An error occurred while scanning the site collection $($url).Exception:$($_.Exception)"
            }
           
        }   
        return $webReportLines;  
    }
    catch {
        WriteErrorLog "An error occurred while scaning the site collection $($url).Exception:$($_.Exception)" 
    }
    
}
    
Write-Host 'Success' -ForegroundColor Green
