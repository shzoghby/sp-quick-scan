#=====================================================================================================================
# Script Name:             AzureApp.ps1
# Description:             Azure app creation for Quick Archiving Scan for AvePoint
# Author:                  Bruce Berends
# Creation Date:           2023/10/13
# Last Modified By:        Bruce Berends
# Last Modified Date:      2023/10/13
#=====================================================================================================================

$outputCsv = "appdetails.csv"

# Check if CSV exists and delete it if it does
if (Test-Path $outputCsv) {
    Remove-Item $outputCsv -Force
}

# 1. Install and import required modules
if ($PSVersionTable.PSEdition -eq "Desktop" -and (Get-Module -Name AzureAD -ListAvailable)) {
    Install-Module AzureAD -Force -AllowClobber -Scope CurrentUser
    Import-Module AzureAD
} else {
    Install-Module AzureAD.Standard.Preview -Force -AllowClobber -Scope CurrentUser
    Import-Module AzureAD.Standard.Preview
}

# 2. Login to Azure AD
Connect-AzureAD

# 3. Check if the application already exists
$appName = "AvePointQuickScan"
$app = Get-AzureADApplication -Filter "DisplayName eq '$appName'"

if (-not $app) {
# Create the Azure App without permissions first
$app = New-AzureADApplication -DisplayName $appName
$servicePrincipal = New-AzureADServicePrincipal -AppId $app.AppId

if ($app -ne $null) {

    $sp = Get-AzureADServicePrincipal -Filter "DisplayName eq '$appName'"
        if (-not $sp) {
            Write-Error "Service Principal for $appName not found."
            exit
         }

    Start-Sleep -Seconds 10
    # Retrieve the Service Principal for SharePoint Online
    $spOnline = Get-AzureADServicePrincipal -Filter "AppId eq '00000003-0000-0ff1-ce00-000000000000'"
    $exchangeOnine = Get-AzureADServicePrincipal -Filter "AppId eq '00000002-0000-0ff1-ce00-000000000000'"

    # Check if the permission already exists
    $existingSpOnlinePermission = Get-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipal.ObjectId | Where-Object {$_.ResourceId -eq $spOnline.ObjectId}
    $existingExchangeOnlinePermission = Get-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipal.ObjectId | Where-Object {$_.ResourceId -eq $exchangeOnine.ObjectId}

    
    # Grant Full Control permission to SharePoint Online if not already granted
    if (-not $existingSpOnlinePermission) {
    $fullControlPermission = $spOnline.AppRoles | Where-Object {$_.Value -eq "Sites.FullControl.All"}
    New-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipal.ObjectId -PrincipalId $servicePrincipal.ObjectId -Id $fullControlPermission.Id -ResourceId $spOnline.ObjectId
    }

    # GrantExchange.ManageAsApp permission to Office 365 Exchange Online if not already granted
    if (-not $existingExchangeOnlinePermission) {
    $exchangePermission = $spOnline.AppRoles | Where-Object {$_.Value -eq "Exchange.ManageAsApp"}
    New-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipal.ObjectId -PrincipalId $servicePrincipal.ObjectId -Id $exchangePermission.Id -ResourceId $exchangeOnine.ObjectId
    }

    # Add the required permissions to the Azure AD app
    #$servicePrincipal = Get-AzureADServicePrincipal -Filter "DisplayName eq '$appName'"
    #New-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipal.ObjectId -PrincipalId $servicePrincipal.ObjectId -ResourceId  (Get-AzureADServicePrincipal -Filter "AppId eq '00000003-0000-0ff1-ce00-000000000000'").ObjectId -Id (Get-AzureADServicePrincipal -Filter "AppId eq '00000003-0000-0ff1-ce00-000000000000'").AppRoles[0].Id
 
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

        # Output to Console
        Write-Output "App ID: $($app.AppId)"
        Write-Output "Certificate Thumbprint: $($cert.Thumbprint)"

        # Store the details for reference
        $outputObject = [PSCustomObject]@{
            'AppId'      = $app.AppId
            'Thumbprint' = $cert.Thumbprint
        }

        $outputObject | Export-Csv -Path $outputCsv -NoTypeInformation -Append
    } else {
        # Handle the error appropriately
        Write-Error "Azure AD application creation failed."
    }
} else {
    Write-Warning "Azure AD application '$appName' already exists."
}
