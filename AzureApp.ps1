#=====================================================================================================================
# Script Name:             AzureApp.ps1
# Description:             Azure app creation for Quick Archiving Scan for AvePoint
# Author:                  Bruce Berends
# Creation Date:           2023/10/13
# Last Modified By:        Bruce Berends
# Last Modified Date:      2023/10/13
#=====================================================================================================================

#========================================================
#Functions module
#=======================================================
try {
    . .\_functions.ps1
}
catch {
    Write-Error "Could not load _functions.ps1 file. $_"
    exit
}

$outputCsv = "appdetails.csv"

# 1. Install and import required modules
if ($PSVersionTable.PSEdition -eq "Desktop" -and (Get-Module -Name AzureAD -ListAvailable)) {
    Install-Module AzureAD -Force -AllowClobber -Scope CurrentUser
    Save-Module AzureAD -Repository AzureAD -Path "$PSScriptRoot\bin\Modules" -Force
    Import-Module AzureAD
}
else {
    Install-Module AzureAD.Standard.Preview -Force -AllowClobber -Scope CurrentUser
    Save-Module AzureAD.Standard.Preview -Repository PSGallery -Path "$PSScriptRoot\bin\Modules" -Force
    Import-Module AzureAD.Standard.Preview
}

# 2. Login to Azure AD
Connect-AzureAD

# 3. Check if the application already exists
$appName = "AvePointQuickScan TEST"
$app = Get-AzureADApplication -Filter "DisplayName eq '$appName'"

if (-not $app) {
    Write-Host "Creating new app..." -NoNewline
    # Create the Azure App without permissions first
    $app = New-AzureADApplication -DisplayName $appName
     
    # Generate a certificate for the app
    $certStartDate = (Get-Date).Date
    $certEndDate = $certStartDate.AddYears(1)
    $cert = New-SelfSignedCertificate -Subject "CN=$appName" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -NotAfter $certEndDate

    # Convert the certificate into Base64 format
    $base64Value = [System.Convert]::ToBase64String($cert.RawData)

    # Construct the key credential
    $certKeyCredential = New-Object Microsoft.Open.AzureAD.Model.KeyCredential
    $certKeyCredential.CustomKeyIdentifier = [System.Convert]::FromBase64String([System.Convert]::ToBase64String($cert.GetCertHash()))
    $certKeyCredential.EndDate = $certEndDate
    $certKeyCredential.Value = [System.Text.Encoding]::Default.GetBytes($base64Value)
    $certKeyCredential.StartDate = (Get-Date).AddMinutes(-10)
    $certKeyCredential.Type = "AsymmetricX509Cert"
    $certKeyCredential.Usage = "Verify"

    # Add the key credential to the Azure AD app
    Set-AzureADApplication -ObjectId $app.ObjectId -KeyCredentials @($certKeyCredential)
    $servicePrincipal = New-AzureADServicePrincipal -AppId $app.AppId

    # Output to Console
    Write-Host "App ID: $($app.AppId) - Certificate Thumbprint: $($cert.Thumbprint)" -ForegroundColor Green

    # Store the details for reference
    $outputObject = [PSCustomObject]@{
        'AppId'      = $app.AppId
        'Thumbprint' = $cert.Thumbprint
    }

    # Check if CSV exists and delete it if it does
    if (Test-Path $outputCsv) {
        Remove-Item $outputCsv -Force
    }

    Write-Host "Exporting CSV with new app details..." -NoNewline
    $outputObject | Export-Csv -Path $outputCsv -NoTypeInformation -Append
    Write-Host "Success" -ForegroundColor Green
}
else {
    Write-Warning "Azure AD application '$appName' already exists."
    $servicePrincipal = Get-AzureADServicePrincipal -Filter "DisplayName eq '$appName'"
}

if ($null -ne $app) {
    Write-Host "Granting App Permissions"
    if (-not $servicePrincipal) {
        Write-Error "Service Principal for $appName not found."
        exit
    }
    Start-Sleep -Seconds 10
    GrantAppPermissions -servicePrincipalObjectId $servicePrincipal.ObjectId
}
else {
    # Handle the error appropriately
    Write-Error "Azure AD application creation failed."
}