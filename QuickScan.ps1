#=====================================================================================================================
# Script Name:             QuickScan.ps1
# Description:             Quick Archiving Scan for AvePoint
# Author:                  (Frank Ren) - Modified by Bruce Berends
# Creation Date:           2023/10/13
# Last Modified By:        Bruce Berends
# Last Modified Date:      2023/10/13
#=====================================================================================================================
param(
    [string] $tenantFullName = "engage2syddev.onmicrosoft.com",
    #$(Read-Host -Prompt "Please enter your tenant full name (e.g., yourdomain.onmicrosoft.com)"),
    #[string] $inputFileName = $(Read-Host -Prompt "Please enter name of site collections input file or press enter to scan the whole tenant (e.g., input-sitecollections.csv)")
    [string] $inputFileName = "input-sitecollections.csv"
)

$csvData = Import-Csv -Path ".\Appdetails.csv"
if (-not $csvData) {
    Write-Error "AppDetails.csv not found. Please ensure it exists in the script directory."
    exit
}
$appId = $csvData.AppId.Trim()
$thumbprint = $csvData.Thumbprint.Trim()

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

try
{
    Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking | out-null
    Import-Module PnP.PowerShell -DisableNameChecking | out-null
}
catch
{
    Write-Error "Could not load required assemblies"
    exit
}

#========================================================
#Retrieve Site Collections
#=======================================================
$siteCollections = GetSitecollections -tenantFullName $tenantFullName -clientId $appId -thumbprint $thumbprint -inputFilePath $inputFilePath

#========================================================
#Main Run
#=======================================================
try 
{
    $startTime=Get-Date
    WriteInfoLog "The job is running on $($startTime)."
    if([System.IO.Directory]::Exists($jobFolder) -eq $false){
        [System.IO.Directory]::CreateDirectory($jobFolder) | Out-Null
    }
    
    $siteCollectionReportLines  = @();
    foreach($sitecollection in $siteCollections){
        WriteInfoLog "Begin scan the site collection $($sitecollection.URL)"
        $siteCollectionReportLine = ScanSiteCollection -tenantFullName $tenantFullName -clientId $appId -thumbprint $thumbprint -url $sitecollection.URL
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