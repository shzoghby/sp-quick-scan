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
function WriteInfoLog {  
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
            
        [Parameter(Mandatory = $false)]
        [ValidateSet('Green', 'Blue', 'Yellow', 'Red')]
        [string] $ForegroundColor = 'White'
    )
   
    Write-Host $message -ForegroundColor $ForegroundColor 
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
    Write-Host "Done fetching site collection details." -ForegroundColor Green
}

function getDateTimeFilter {
    param (
        [string]$dayOrMonthOrYear,
        [int]$number
    )

    if($null -eq $number -or '' -eq $number)
    {
        $number = 30;
    }
    
    if ($dayOrMonthOrYear -eq 'day') {
        return "<View><Query><Where><Geq><FieldRef Name ='Modified'/><Value Type ='DateTime'><Today OffsetDays='-$number'/></Value></Geq></Where></Query></View>"
    }
    elseif ($dayOrMonthOrYear -eq 'month') {
        return "<View><Query><Where><Geq><FieldRef Name ='Modified'/><Value Type ='DateTime'><Today OffsetMonths='-$number'/></Value></Geq></Where></Query></View>"
    }
    elseif ($dayOrMonthOrYear -eq 'year') {
        return "<View><Query><Where><Geq><FieldRef Name ='Modified'/><Value Type ='DateTime'><Today OffsetYears='-$number'/></Value></Geq></Where></Query></View>"
    }
    else {
        return "<View><Query><Where><Geq><FieldRef Name ='Modified'/><Value Type ='DateTime'><Today OffsetDays='-$number'/></Value></Geq></Where></Query></View>"
    }
}

function GetFilesLastAccessedDate {
    param (
        [string]$sitecollectionUrl,
        [string]$webUrl,
        [string]$listTitle,
        [string]$listRootFolder,
        [string]$listId,
        [string]$dayOrMonthOrYear,
        [int]$number
    )
    {
        $libraryUrl = '/sites/YourSite/LibraryName/'
        $auditData = Get-PnPAuditLog -Query "<Query><Where><And><Eq><FieldRef Name='FileDirRef'/><Value Type='Text'>$libraryUrl</Value></Eq><Eq><FieldRef Name='Event' /><Value Type='String'>View</Value></Eq></And></Where></Query>"
        foreach ($entry in $auditData) 
        {
            $fileUrl = $entry.ItemUrl
            $lastAccessedDate = $entry.Occurred
            Write-Host "File URL: $fileUrl, Last Consultation Date: $lastAccessedDate"
        }
    }
}

function ScanFiles {
    param (
        [string]$sitecollectionUrl,
        [string]$webUrl,
        [string]$listTitle,
        [string]$listRootFolder,
        [string]$listId,
        [string]$dayOrMonthOrYear,
        [int]$number
    )

    $allFiles = @() # Result array to keep all file details
    WriteInfoLog "Begin check files in the list $($listTitle)"
    $query = getDateTimeFilter -dayOrMonthOrYear $dayOrMonthOrYear -number $number
    $ListItems = Get-PnPListItem -List $listTitle -Query $query
    #Enumerate all list items to get file details
    ForEach ($Item in $ListItems) {
        #Add file details to Result array
        $allFiles += New-Object PSObject -property $([ordered]@{
                SitecollectionUrl = $sitecollectionUrl;
                WebUrl            = $webUrl;
                ListTitle         = $listTitle;
                FileName          = $Item.FieldValues["FileLeafRef"]            
                FileID            = $Item.FieldValues["UniqueId"]
                FileType          = $Item.FieldValues["File_x0020_Type"]
                RelativeURL       = $Item.FieldValues["FileRef"]
                CreatedByEmail    = $Item.FieldValues["Author"].Email
                CreatedTime       = $Item.FieldValues["Created"]
                LastModifiedTime  = $Item.FieldValues["Modified"]
                ModifiedByEmail   = $Item.FieldValues["Editor"].Email
                FileSizeMB        = [Math]::Round(($Item.FieldValues["File_x0020_Size"] / 1024 / 1024), 2) #File size in MB
            })

    }
    
    WriteInfoLog "Finished scan files in the list $($listTitle)" -ForegroundColor Green
    return $allFiles;
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
            WriteInfoLog "Finished check the empty list $($listTitle)." -ForegroundColor Green

            $listReturnObject = New-Object PSObject -property $([ordered]@{
                    SitecollectionUrl     = $sitecollectionUrl;
                    WebUrl                = $webUrl;
                    ListTitle             = $listTitle;
                    LastModified          = $storage.LastModified;
                    TotalFileCount        = 0;
                    TotalFileStreamSizeMB = 0;
                    TotalSizeMB           = 0;
                    VersionSizeMB         = 0;
                });

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

            $listReturnObject = New-Object PSObject -property $([ordered]@{
                    SitecollectionUrl     = $sitecollectionUrl;
                    WebUrl                = $webUrl;
                    ListTitle             = $listTitle;
                    LastModified          = $storage.LastModified;
                    TotalFileCount        = $($storage.TotalFileCount);
                    TotalFileStreamSizeMB = $($storage.TotalFileStreamSize / 1024 / 1024);
                    TotalSizeMB           = $($storage.TotalSize / 1024 / 1024);
                    VersionSizeMB         = $($versionSize / 1024 / 1024);
                });
        }
    }
    else {
        WriteInfoLog "Finished check the list $($listTitle) with empty storage." -ForegroundColor Green
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
        [string]$sitecollectionUrl,
        [switch]$reportLevel,
        [string]$dayOrMonthOrYear,
        [int]$number
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
            if ($folderName -eq "/Style Library" -or $folderName -eq "/FormServerTemplates") {
                continue;
            }
            
            if ($reportLevel -eq "listLevel" -or $reportLevel -eq $null) {
                $returnObject = ScanLists -sitecollectionUrl $sitecollectionUrl -webUrl $webUrl -listTitle $list.Title -listRootFolder $folderName -listId $list.Id;
                if ($null -ne $returnObject) {              
                    $listReportLines.Add($returnObject[$returnObject.length - 1]);
                    #$listReportLines.Add($returnObject);
                }
                else {
                    WriteWarnLog "The report line of the list $($list.Title) is null" 
                }
            }
            else {
                $returnObject = ScanFiles -sitecollectionUrl $sitecollectionUrl -webUrl $webUrl -listTitle $list.Title -listRootFolder $folderName -listId $list.Id -dayOrMonthOrYear $dayOrMonthOrYear -number $number;
                if ($null -ne $returnObject) {              
                    $listReportLines.Add($returnObject);
                }
                else {
                    WriteWarnLog "The report line of the list $($list.Title) is null" 
                }
            } 
        }
        catch {
            WriteWarnLog "An error occurred while exporting the list $($list.Title).Exception:$($_)" 
        }
        
    }

    return $listReportLines;
    WriteInfoLog "Finished scan the web $($web.Url)" -ForegroundColor Green
 
}

function ScanSiteCollection {
    param (
        [string]$tenantFullName,
        [string]$clientId,
        [string]$thumbprint,
        [string]$url,
        [string]$reportLevel,
        [string]$dayOrMonthOrYear,
        [int]$number
    )
    try {
        $webReportLines = New-Object 'System.Collections.Generic.List[PSCustomObject]'
        # Connect to the specific site collection with MFA
        Connect-PnPOnline -Url $url -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantFullName

        $rootWeb = Get-PnPWeb
        $webReportLines = ScanWebs -tenantFullName $tenantFullName -ClientId $clientId -Thumbprint $thumbprint -webId $rootWeb.Id -webUrl $rootWeb.Url -sitecollectionUrl $url -reportLevel $reportLevel -dayOrMonthOrYear $dayOrMonthOrYear -number $number;
        $webs = Get-PnPSubWeb -Recurse;
        foreach ($web in $webs) {
            try {
                $webReportLine = ScanWebs -tenantFullName $tenantFullName -ClientId $clientId -Thumbprint $thumbprint -webId $web.Id -webUrl $web.Url -sitecollectionUrl $url -reportLevel $reportLevel -dayOrMonthOrYear $dayOrMonthOrYear -number $number;
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
