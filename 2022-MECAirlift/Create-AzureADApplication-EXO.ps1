<#
    .SYNOPSIS
    Creates an Azure AD Application for certificate-based access, adding the Exchange.ManageAsApp API permission and
    granting it the Exchange administrator AzureAD built-in role.

    .AUTHOR
    Michel de Rooij
    http://eightwone.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 1.0, August 3rd, 2020

    .DESCRIPTION

    .REQUIREMENTS
    - AzureAD or AzureADPreview Graph-based Azure Active Directory PowerShell modules.

    .PARAMETER CertificateFile
    Name of the certificate file to upload. This needs to be the public key (.cer) file.
    Make sure you install this certificate with the private key (.pfx) on the system where you 
    want to connect from.

    .PARAMETER Name
    Name of the Azure AD Application to register.

    .PARAMETER Workload
    Specify one or more workloads to enable managent for. Default is Exchange.
    
    Currently supported are:

    Workload        API Permission
    ------------------------------------------------
    Exchange        Exchange.ManageAsApp

    .EXAMPLE

#>
[cmdletbinding()]
Param(
   [Parameter(Mandatory=$true)]
   [string]$CertificateFile,

   [Parameter(Mandatory=$true)]
   [string]$Name,

   [Parameter(Mandatory=$false)]
   [ValidateSet('Exchange')]
   [string[]]$Workload='Exchange'
)
#Require -Module AzureADPreview

# Check if AzureAD/AzureADPreview module is installed
If( -not( Get-Module AzureAD -ListAvailable -ErrorAction SilentlyContinue) -and -not( Get-Module AzureADPreview -ListAvailable -ErrorAction SilentlyContinue)) {
    Throw( 'AzureAD Module does not seem to be installed; Use Install-Module AzureAD to install.')
}

If( Get-Module AzureAD -ListAvailable -ErrorAction SilentlyContinue) {
    Import-Module -Name AzureAD
}
Else {
    Import-Module -Name AzureADPreview
}

# Check if we are connected
Try { 
    $null= Get-AzureADApplication -ErrorAction SilentlyContinue 
}
Catch {
    Throw( 'Cannot run AzureAD cmdlets: {0}' -f $Error[0])
}

If( Test-Path -Path $CertificateFile) {

    Try {
        $ResolvedCertFile= Resolve-Path -Path $CertificateFile
        Write-Verbose ('Importing certificate from file {0}' -f $ResolvedCertFile)
        $certObj = New-Object -Type System.Security.Cryptography.X509Certificates.X509Certificate2( $ResolvedCertFile)
        $certRaw= $certObj.GetRawCertData()
        $certBase64Value = [System.Convert]::ToBase64String($certRaw)
        $certRaw = $certObj.GetCertHash()
        $certBase64Thumbprint = [System.Convert]::ToBase64String( $certRaw)
        $null = [System.Guid]::NewGuid().ToString() 
    }
    Catch {
        Throw( 'Problem importing certificate from {0}: {1}' -f $CertificateFile, $Error[0])
    }

    $existingApp= Get-AzureADApplication | Where-Object {$_.DisplayName -eq $Name}
    If( $existingApp.objectId) {
        if ($PSCmdlet.ShouldProcess($existingApp, 'Remove existing App registration')) {
            Remove-AzureADApplication -ObjectId $existingApp.ObjectId
        }
    }

    # Create new App registration if necessary
    $App= Get-AzureADApplication | Where-Object {$_.DisplayName -eq $Name}
    If(-not( $App.objectId)) {
        $App= New-AzureADApplication -DisplayName $Name
    }

    If( $App.objectId) {

        Write-Host ('Application ID {0}' -f $App.AppId)

        # Create App security principal
        $AppSP= Get-AzureADServicePrincipal -Filter ("AppId eq '{0}'" -f $App.AppId)
        If(-not( $AppSP.objectId)) {
            $AppSP= New-AzureADServicePrincipal -AppId $App.AppId -DisplayName $Name -AccountEnabled True 
        }

        $Permissions= New-Object System.Collections.ArrayList

        # Get the delegated API permission for Graph/User.Read
        $GraphSP= Get-AzureADServicePrincipal -All $true | Where-Object {$_.DisplayName -like 'Microsoft Graph'}
        $UserRead= $GraphSP.OAuth2Permissions | Where-Object {'User.Read' -contains $_.Value}
        $GraphPerms= New-Object -TypeName 'Microsoft.Open.AzureAD.Model.RequiredResourceAccess'
        $GraphPerms.ResourceAppId = $GraphSP.AppId
        $Permission1 = New-Object -TypeName 'Microsoft.Open.AzureAD.Model.ResourceAccess' -ArgumentList $UserRead.Id, 'Scope'
        $GraphPerms.ResourceAccess= $Permission1
        $null= $Permissions.Add( $GraphPerms)
        
        If( $Workload -icontains 'Exchange') {
            # Get the API permission for Exchange/ManageMyApps
            $EXOSP= Get-AzureADServicePrincipal -All $true | Where-Object {$_.DisplayName -like 'Office 365 Exchange Online'}
            $ManageMyApps= $EXOSP.AppRoles | Where-Object {'Exchange.ManageAsApp' -contains $_.Value}
            $ExPerms= New-Object -TypeName 'Microsoft.Open.AzureAD.Model.RequiredResourceAccess'
            $ExPerms.ResourceAppId = $EXOSP.AppId
            $Permission2 = New-Object -TypeName 'Microsoft.Open.AzureAD.Model.ResourceAccess' -ArgumentList $ManageMyApps.Id, 'Role'
            $ExPerms.ResourceAccess= $Permission2
            $null= $Permissions.Add( $ExPerms)
        }

        # Assign API permissions to Application
        Set-AzureADApplication -ObjectId $App.ObjectId -RequiredResourceAccess $Permissions

        # Assign certificate to Application
        $AAAKC= New-AzureADApplicationKeyCredential -ObjectId $App.objectId -CustomKeyIdentifier $certBase64Thumbprint  -Type AsymmetricX509Cert -Usage Verify -Value $certBase64Value -StartDate $certObj.NotBefore -EndDate $certObj.NotAfter


        $TenantId= (Get-AzureADCurrentSessionInfo).tenantId
        $AdminConsentURI= 'https://login.microsoftonline.com/{0}/adminconsent?client_id={1}' -f $TenantId, $App.AppId
#        Write-Host ('Use the following URL to grant admin consent to {0}:' -f $Name)
#        Write-Host ($AdminConsentURI)

        # 74658136-14ec-4630-ad9b-26e160ff0fc6 = main.iam.ad.ext.azure.com
#        $token = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate($context.Account, $context.Environment, $TenantId, $null, 'Never', $null, '74658136-14ec-4630-ad9b-26e160ff0fc6')
        $account= (Get-AzureADCurrentSessionInfo).Account
        $environment= (Get-AzureADCurrentSessionInfo).Environment
#        $token= [Microsoft.Open.Azure.AD.CommonLibrary.AzureSession]::AuthenticationFactory.Authenticate( $account , $environment, $tenantId, $null, 'Never', $null, 'AadGraphEndpointResourceId')
        $token= [Microsoft.Open.Azure.AD.CommonLibrary.AzureSession]::AuthenticationFactory.Authenticate( $account , $environment, $tenantId, $null, 'Never', $null, '74658136-14ec-4630-ad9b-26e160ff0fc6')
        #$token= [Microsoft.Open.Azure.AD.CommonLibrary.AzureSession]::AccessTokens['AccessToken']
        $correlationId= [Microsoft.Open.Azure.AD.CommonLibrary.AzureSession]::CorrelationId
        $headers = @{
            'Authorization' = 'Bearer {0}' -f $token.AccessToken
            'X-Requested-With'= 'XMLHttpRequest'
            'x-ms-client-request-id'= [guid]::NewGuid()
            'x-ms-correlation-id' = $correlationId.Guid
        }
        $url = 'https://main.iam.ad.ext.azure.com/api/RegisteredApplications/{0}/Consent?onBehalfOfAll=true' -f $App.AppId
        Invoke-RestMethod -Uri $url -Headers $headers -Method POST -ErrorAction Stop

}
    Else {
        Throw( 'Problem locating or creating Application')
    }
}
Else {
    Throw( 'File {0} does not exist or inaccessible.')
}