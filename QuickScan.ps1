#=====================================================================================================================
# Script Name:             QuickScan.ps1
# Description:             Quick Archiving Scan for AvePoint
# Author:                  (Frank Ren) - Modified by Bruce Berends
# Creation Date:           2023/10/13
# Last Modified By:        Bruce Berends
# Last Modified Date:      2023/10/13
#=====================================================================================================================

<#
$csvData = Import-Csv -Path ".\AppDetails.csv"
if (-not $csvData) {
    Write-Error "AppDetails.csv not found. Please ensure it exists in the script directory."
    exit
}
$appId = $csvData.AppId.Trim()
$thumbprint = $csvData.Thumbprint.Trim()
#>
#========================================================
#Variable Prompts
#========================================================
# Prompt for SharePoint credentials and Tenant information
$tenantFullName = Read-Host -Prompt "Please enter your tenant full name (e.g., yourdomain.onmicrosoft.com)";
$tenantBase = $tenantFullName.SubString(0,$tenantFullName.IndexOf('.'));


#========================================================
#Static Variables (no need to modify)
#========================================================
$pageSize=500
$execPath = Split-Path $MyInvocation.MyCommand.Path -parent
$jobId=[DateTime]::Now.ToString("yyyyMMddHHmmss")
$jobFolder=Join-Path $execPath $jobId

$inputFileName="input-sitecollections.csv";
$inputFilePath=Join-Path $execPath $inputFileName;

$outFileName = [System.String]::Format("Report_{0}.csv",[System.DateTime]::Now.ToString("yyyyMMddhhmmss"));
$outFilePath = Join-Path $jobFolder $outFileName;

#========================================================
#Functions module
#=======================================================
try
{
    . .\_functions.ps1
}
catch
{
    Write-Error "Could not load _functions.ps1 file. $_"
    exit
}

#========================================================
#Log directory
#=======================================================
try
{
    $parent=Split-Path -Parent $MyInvocation.MyCommand.Definition
    $logfile=[System.IO.Path]::Combine($parent,"Logs\Log.log")
    $logpath=Split-Path -Parent $logfile
    if (!(Test-Path $logpath)) {
        New-Item -ItemType Directory -Force -Path $logpath
    }
}
catch
{
    Write-Host "An error occured while writing errors to log. $_"
}

#========================================================
#Assemblies Install
#=======================================================
Install-RequiredModules
Import-Module PnP.PowerShell -ErrorAction SilentlyContinue # Ensure the PnP module is imported

#========================================================
#Retrieve Site Collections
#=======================================================
$siteCollections = GetSitecollections -tenantFullName $tenantFullName -tenantBase $-tenantBase -inputFilePath $inputFilePath

#========================================================
#Main Run
#=======================================================
try 
{
    $startTime=Get-Date
    WriteInfoLog "The job is running on $($startTime)."
    if([System.IO.Directory]::Exists($jobFolder) -eq $false){
        [System.IO.Directory]::CreateDirectory($jobFolder) | Out-Null
        WriteReport '"SiteCollection","Site","List","LastModified","TotalFileCount","TotalFileStreamSize","TotalSize","VersionSize"'
    }
    
    $siteCollectionReportLines  = @();
    foreach($sitecollection in $sitecollections){
        WriteInfoLog "Begin scan the site collection $($sitecollection.URL)"
        $siteCollectionReportLine = ScanSiteCollection tenantFullName $tenantFullName -clientId $appId -thumbprint $thumbprint -url $sitecollection.URL
        $siteCollectionReportLines += $siteCollectionReportLine;
        WriteInfoLog "Finished scan the site collection $($sitecollection.URL)"
    }
    $report|Out-File -FilePath $outFilePath -Append -Encoding utf8
    $siteCollectionReportLines | Export-Csv -Path $outFilePath
    WriteInfoLog "The job is finished on $(Get-Date)."
}
catch {
    WriteErrorLog "An error occurred while exporting the stroage of the excel.Exception:$($_.Exception)"   
}