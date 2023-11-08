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
    $requiredModules = @('PowerShellGet', 'PnP.PowerShell', 'ExchangeOnlineManagement')

    foreach ($module in $requiredModules) {
        # Check if the module is installed
        if (Get-Module -ListAvailable -Name $module) {
            Write-Host "$module module is already installed." -ForegroundColor Green
            #Uninstall-Module $module -Force
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

function getLastModifiedDateQuery {
    param (
        [string]$lastModifiedDayOrMonthOrYear
    )

    $lastModifiedDayOrMonthOrYearArr = $lastModifiedDayOrMonthOrYear.Split(":");
    $number = 30;
    $dayOrMonthOrYear = 'day';

    if($null -ne $lastModifiedDayOrMonthOrYearArr -and $lastModifiedDayOrMonthOrYearArr.length -eq 2)
    {
        $dayOrMonthOrYear = $lastModifiedDayOrMonthOrYearArr[0];
    $number = $lastModifiedDayOrMonthOrYearArr[1];
    }

    if ($null -eq $number -or '' -eq $number) {
        $number = 30;
    }

    if ($null -eq $dayOrMonthOrYear -or '' -eq $dayOrMonthOrYear) {
        $dayOrMonthOrYear = 'day';
    }
    
    switch ($dayOrMonthOrYear) {
        'day' { return "<View><Query><Where><Leq><FieldRef Name ='Modified'/><Value Type ='DateTime'><Today OffsetDays='-$number'/></Value></Leq></Where></Query></View>" }
        'month' { return "<View><Query><Where><Leq><FieldRef Name ='Modified'/><Value Type ='DateTime'><Today OffsetMonths='-$number'/></Value></Leq></Where></Query></View>" }
        'year' { return "<View><Query><Where><Leq><FieldRef Name ='Modified'/><Value Type ='DateTime'><Today OffsetYears='-$number'/></Value></Leq></Where></Query></View>" }
        default { return "<View><Query><Where><Leq><FieldRef Name ='Modified'/><Value Type ='DateTime'><Today OffsetDays='-$number'/></Value></Leq></Where></Query></View>" }
    }
}

function getLastAccessedDateTimeFilter {
    param (
        [string]$lastAccessedDayOrMonthOrYear
    )
    $lastAccessedDayOrMonthOrYearArr = $lastAccessedDayOrMonthOrYear.Split(":");
    $number = 90;
    $dayOrMonthOrYear = 'day';

    if($null -ne $lastAccessedDayOrMonthOrYearArr -and $lastAccessedDayOrMonthOrYearArr.length -eq 2)
    {
        $dayOrMonthOrYear = $lastAccessedDayOrMonthOrYearArr[0];
        $number = $lastAccessedDayOrMonthOrYearArr[1];
    }

    if ($null -eq $number -or '' -eq $number) {
        $number = 90;
    }

    if ($null -eq $dayOrMonthOrYear -or '' -eq $dayOrMonthOrYear) {
        $dayOrMonthOrYear = 'day';
    }

    switch ($dayOrMonthOrYear) {
        'day' { return (Get-Date).AddDays(-$number) }
        'month' { return (Get-Date).AddMonths(-$number) }
        'year' { return (Get-Date).AddYears(-$number) }
        default { return (Get-Date).AddDays(-$number) }
    }
}

function GetFileLastAccessedDate {
    param (
        [string]$sitecollectionUrl,
        [string]$webUrl,
        [string] $fileRelativeUrl,
        [string]$lastAccessedDayOrMonthOrYear
    )
    $lastAccessedDate = '';
    $fileAbsoluteURL = "$($webUrl.SubString(0, $webUrl.IndexOf("/sites")))$fileRelativeUrl";
    WriteInfoLog "Begin get file last accessed date for file $($fileRelativeUrl)"

    #Set Dates
    $startDate = getLastAccessedDateTimeFilter -lastAccessedDayOrMonthOrYear $lastAccessedDayOrMonthOrYear
    $endDate = (Get-Date)
        
    #Search Unified Log
    $auditLog = Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -RecordType SharePointFileOperation -Operations FileAccessed -SessionId "WordDocs_SharepointViews"-SessionCommand ReturnLargeSet -ObjectIds $fileAbsoluteURL
    if ($null -ne $auditLog) {
        if ($auditLog.length -gt 1) {
            $auditSortedResults = $auditLog | Sort-Object -Property CreationDate | Select-Object -Last 1;
        }
        else {
            $auditSortedResults = $auditLog;
        }

        $auditLogResults = $auditSortedResults.AuditData | ConvertFrom-Json | Select-Object CreationTime, UserID, Operation, ClientIP, ObjectID;
        $lastAccessedDate = $auditLogResults.CreationTime;
    }

    WriteInfoLog "Finished get file last accessed date for file $($fileRelativeUrl)" -ForegroundColor Green;
    return $lastAccessedDate;
}

function ScanFiles {
    param (
        [string]$sitecollectionUrl,
        [string]$webUrl,
        [string]$listTitle,
        [string]$lastModifiedDayOrMonthOrYear,
        [string]$lastAccesseddayOrMonthOrYear,
        [string]$includeLastAccessed
    )

    $allFiles = @() # Result array to keep all file details
    WriteInfoLog "Begin check files in the list $($listTitle)"
    $query = getLastModifiedDateQuery -lastModifiedDayOrMonthOrYear $lastModifiedDayOrMonthOrYear
    $ListItems = Get-PnPListItem -List $listTitle -Query $query
    #Enumerate all list items to get file details
    ForEach ($Item in $ListItems) {
        if ($includeLastAccessed -eq "yes") {
            $lastAccessedDate = GetFileLastAccessedDate -sitecollectionUrl $sitecollectionUrl -webUrl $webUrl -fileRelativeUrl $($Item.FieldValues["FileRef"]) -lastAccesseddayOrMonthOrYear $lastAccesseddayOrMonthOrYear
        }

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
                LastAccessedDate  = $lastAccessedDate
                ModifiedByEmail   = $Item.FieldValues["Editor"].Email
                FileSizeMB        = [Math]::Ceiling(($Item.FieldValues["File_x0020_Size"] / 1024 / 1024)) #File size in MB
            });
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
                    TotalFileStreamSizeMB = [Math]::Ceiling($($storage.TotalFileStreamSize / 1024 / 1024));
                    TotalSizeMB           = [Math]::Ceiling($($storage.TotalSize / 1024 / 1024));
                    VersionSizeMB         = [Math]::Ceiling($($versionSize / 1024 / 1024));
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
        [string]$reportLevel,
        [string]$lastModifieddayOrMonthOrYear,
        [string]$includeLastAccessed,
        [string]$lastAccesseddayOrMonthOrYear
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
            
            switch ($reportLevel) {
                "listLevel" {
                    $returnObject = ScanLists -sitecollectionUrl $sitecollectionUrl -webUrl $webUrl -listTitle $list.Title -listRootFolder $folderName -listId $list.Id;
                    if ($null -ne $returnObject) {              
                        $listReportLines.Add($returnObject[$returnObject.length - 1]);
                        #$listReportLines.Add($returnObject);
                    }
                    else {
                        WriteWarnLog "The report line of the list $($list.Title) is null" 
                    }
                    break;
                }
                "fileLevel" {
                    $returnObject = ScanFiles -sitecollectionUrl $sitecollectionUrl -webUrl $webUrl -listTitle $list.Title -lastModifieddayOrMonthOrYear $lastModifieddayOrMonthOrYear -lastAccesseddayOrMonthOrYear $lastAccesseddayOrMonthOrYear -includeLastAccessed $includeLastAccessed

                    if ($null -ne $returnObject) {              
                        $listReportLines.Add($returnObject);
                    }
                    else {
                        WriteWarnLog "The report line of the list $($list.Title) is null" 
                    }

                    break;
                }
                Default {
                    $returnObject = ScanLists -sitecollectionUrl $sitecollectionUrl -webUrl $webUrl -listTitle $list.Title -listRootFolder $folderName -listId $list.Id;
                    if ($null -ne $returnObject) {              
                        $listReportLines.Add($returnObject[$returnObject.length - 1]);
                        #$listReportLines.Add($returnObject);
                    }
                    else {
                        WriteWarnLog "The report line of the list $($list.Title) is null" 
                    }
                    break;
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
        [string]$lastModifieddayOrMonthOrYear,
        [string]$includeLastAccessed,
        [string]$lastAccesseddayOrMonthOrYear
    )
    try {
        $webReportLines = New-Object 'System.Collections.Generic.List[PSCustomObject]'
        # Connect to the specific site collection with MFA
        Connect-PnPOnline -Url $url -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantFullName

        $rootWeb = Get-PnPWeb
        $webReportLines = ScanWebs -tenantFullName $tenantFullName -ClientId $clientId -Thumbprint $thumbprint -webId $rootWeb.Id -webUrl $rootWeb.Url -sitecollectionUrl $url -reportLevel $reportLevel -lastModifieddayOrMonthOrYear $lastModifieddayOrMonthOrYear -includeLastAccessed $includeLastAccessed -lastAccesseddayOrMonthOrYear $lastAccesseddayOrMonthOrYear
        $webs = Get-PnPSubWeb -Recurse;
        foreach ($web in $webs) {
            try {
                $webReportLine = ScanWebs -tenantFullName $tenantFullName -ClientId $clientId -Thumbprint $thumbprint -webId $web.Id -webUrl $web.Url -sitecollectionUrl $url -reportLevel $reportLevel -lastModifieddayOrMonthOrYear $lastModifieddayOrMonthOrYear -includeLastAccessed $includeLastAccessed -lastAccesseddayOrMonthOrYear $lastAccesseddayOrMonthOrYear
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

function GrantAppPermissions {
    param (
        [string]$servicePrincipalObjectId
    )
    # Retrieve the Service Principal for SharePoint Online
    $spOnline = Get-AzureADServicePrincipal -Filter "AppId eq '00000003-0000-0ff1-ce00-000000000000'"
    $exchangeOnine = Get-AzureADServicePrincipal -Filter "AppId eq '00000002-0000-0ff1-ce00-000000000000'"
 
    # Check if the permission already exists
    $existingSpOnlinePermission = Get-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipalObjectId | Where-Object { $_.ResourceId -eq $spOnline.ObjectId }
    $existingExchangeOnlinePermission = Get-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipalObjectId | Where-Object { $_.ResourceId -eq $exchangeOnine.ObjectId }
 
    # Grant Full Control permission to SharePoint Online if not already granted
    if (-not $existingSpOnlinePermission) {
        $fullControlPermission = $spOnline.AppRoles | Where-Object { $_.Value -eq "Sites.FullControl.All" }
        New-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipalObjectId -PrincipalId $servicePrincipalObjectId -Id $fullControlPermission.Id -ResourceId $spOnline.ObjectId
    }
 
    # GrantExchange.ManageAsApp permission to Office 365 Exchange Online if not already granted
    if (-not $existingExchangeOnlinePermission) {
        $exchangePermission = $exchangeOnine.AppRoles | Where-Object { $_.Value -eq "Exchange.ManageAsApp" }
        New-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipalObjectId -PrincipalId $servicePrincipalObjectId -Id $exchangePermission.Id -ResourceId $exchangeOnine.ObjectId
    } 
}
Write-Host 'Success' -ForegroundColor Green
