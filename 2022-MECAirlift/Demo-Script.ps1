<#
    .SYNOPSIS
    MECAIR212 Modernizing your Exchange scripts
    Script to support the demos for Microsoft Exchange Community Technical Airlift Demo Script

    .AUTHOR
    Michel de Rooij
    http://eightwone.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

    Version 1.0, September 14th, 2022
#>


#Requires -Version 5.0 –Modules AzureADPreview,Microsoft.Graph,ExchangeOnlineManagement,PSAzureMigrationAdvisor 
function Prompt { 'PS> '}
Clear-Host

Install-module ExchangeOnlineManagement -AllowPrerelease
# This will install the GA build
#Install-Module ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement
# Use prefixes to provide context
#Import-Module ExchangeOnlineManagement -Prefix Dev

# Clear previous sessions
#Get-Module tmp* | % { Remove-Module $_.Name }
#Get-PSSession | Where-Object {$_.ComputerName -like 'outlook.office365.com' } | Remove-PSSession

#####################################################
# Preperations Certificate-Based Authentication
#####################################################
# Create self-signed cert, comes with EOM module (admin role) or available at https://github.com/SharePoint/PnP-Partner-Pack/blob/master/scripts/Create-SelfSignedCertificate.ps1
$ModulePath= Join-Path (Split-Path (Split-Path -Path (get-module ExchangeOnlineManagement).Path -Parent) -Parent) 'netFramework'
$SSCScriptFile= Join-Path -Path $ModulePath -Child 'Create-SelfSignedCertificate.ps1'

# Create a (self-signed) certificate
. $SSCScriptFile -CommonName 'MECDemo' -StartDate 9/1/2022 -EndDate 9/30/2022 -Password (Read-Host -AsSecureString) -Force
Get-ChildItem Cert:\CurrentUser\My | Where {$_.Subject -eq 'CN=MECDemo'}

# Import certificate in local cert.store for usage
Import-PfxCertificate -CertStoreLocation Cert:\CurrentUser\My -FilePath .\MECDemo.pfx -Password (Read-Host -AsSecureString)

# Now we need to create app. registration (app-only) with Office 365 Exchange Online permissions, use this cert for 
# authentication, and assign Exchange Admin role to app (and other roles as required, eg other admin workloads). Grant consent.
# https://eightwone.com/2020/08/05/exchange-online-management-using-exov2-module/
# https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps

# Register app in Azure AD
# Note: Using PS7 to connect to AzureAD, use "Import-Module AzureAD -UseWindowsPowerShell" for compatibility reasons
#.\Create-AzureADApplication-EXO.ps1 -CertificateFile .\MECDemo.cer -Name MECDemo -Workload Exchange -Verbose

######################################################
### Connect using Certificate-Based Authentication
######################################################
# Set $TenantId, $OrgId and $AppId, and get Thumbprint of cert to use for authentication
#$TenantId= '12391a2c-c494-48d9-a2b5-32c9d88d38b2'
#$OrgId= 'contoso.onmicrosoft.com'
#$AppId= 'ad6edc04-249b-480c-6dbf-c77140c262b0'
. .\SetDemoVars.ps1
$CertThumbprint= (Get-ChildItem Cert:\CurrentUser\My | Where {$_.Subject -eq 'CN=MECDemo'}).Thumbprint
# Note: I'd recommend 1 certificate per app/service principal: Sign-In logs will show app, not who used it.

# Connect using CBA and REST mode
Connect-ExchangeOnline -Organization $OrgId -AppId $AppId -CertificateThumbprint $CertThumbprint -ShowBanner:$False
# Connect using CBA and RPS session
#Connect-ExchangeOnline -Prefix RPS -Organization $OrgId -AppId $AppId -CertificateThumbprint $CertThumbprint -UseRPSSession -ShowBanner:$False

# Alternative CBA connect methods - first we need password to read the pfx:
#$PfxPwd= Read-Host -AsSecureString
# Connect with certificate from file with password:
#Connect-ExchangeOnline -CertificateFilePath (Resolve-Path .\MECDemo.pfx).Path -CertificatePassword $PfxPwd -AppID $AppId -Organization $OrgId
# Connect with certificate object, created from file w/password:
#$CertObj= [System.Security.Cryptography.X509Certificates.X509Certificate2]::new( (Resolve-Path .\MECDemo.pfx).Path, [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $PfxPwd)))
#Connect-ExchangeOnline -Certificate $CertObj -AppID $AppId -Organization $OrgId

# Connect other workload shells using this certificate
# Workaround for PS7 compatibility issue with AzureAD module
Import-Module AzureADPreview -UseWindowsPowerShell
Connect-AzureAD -TenantId $OrgId -CertificateThumbprint $CertThumbprint -ApplicationId $AppId
#Connect-Graph -CertificateThumbprint $CertThumbprint -ClientId $AppId -TenantId $TenantId
#Connect-MicrosoftTeams -CertificateThumbprint $CertThumbprint -ApplicationId $AppId -TenantId $TenantId

#######################################################
# Analyzing EXO scripts for assessment
# Uses:
# - https://github.com/michelderooij/Analyze-ExoScript
#######################################################

# Analyse scripts and report on Exchange cmdlets
.\Analyze-ExoScript.ps1 -File '.\Demo\' -Organization $OrgId -CertificateThumbprint $CertThumbprint -AppId $AppId | ft -a

# Analyse scripts and report which scripts require RPS
.\Analyze-ExoScript.ps1 -File '.\Demo\' -Organization $OrgId -CertificateThumbprint $CertThumbprint -AppId $AppId | Where {$_.Type -eq 'RPS'} | ft -a

# Analyse scripts and report which scripts can perhaps benefit from REST-backed to REST-based mapping
.\Analyze-ExoScript.ps1 -File '.\Demo\' -Organization $OrgId -CertificateThumbprint $CertThumbprint -AppId $AppId | Where {$_.Alt} | ft -a

# Analyse scripts reporting all cmdlets (not only Exchange ones) and report which ones use PSSession (eg Remote PowerShell)
.\Analyze-ExoScript.ps1 -File '.\Demo\' -Organization $OrgId -CertificateThumbprint $CertThumbprint -AppId $AppId -ShowAll | Where {$_.Command -like '*-PSSession'} | ft -a

# Analyse scripts and report which scripts use Connect-ExchangeOnline using Credential (BA)
.\Analyze-ExoScript.ps1 -File '.\Demo\' -Organization $OrgId -CertificateThumbprint $CertThumbprint -AppId $AppId | Where {$_.Command -like 'Connect-ExchangeOnline' -and $_.Parameters -contains 'Credential'} | ft -a

#######################################################
# Analyzing scripts for MSOnline/AzureAD usage 
# Uses: 
# - https://merill.net/2022/04/graph-powershell-conversion-analyzer/
# - https://github.com/FriedrichWeinmann/PSAzureMigrationAdvisor
#######################################################

# Individual, interactive analysis 
# https://graphpowershell.merill.net

# Bulk analysis
#Install-Module PSAzureMigrationAdvisor -Scope AllUsers
Import-Module PSAzureMigrationAdvisor
Get-ChildItem .\Demo\*.ps1 | Read-AzScriptFile

# Be aware of feature discrepancies, eg New-MsolLicenseOptions doesn't map to MSGraph:
# $lic= New-MsolLicenseOptions -AccountSkuId Contoso:BPOS_STANDARD -DisabledPlans EXCHANGE_STANDARD
# Set-MsolUserLicense .. -LicenseOptions $lic
# See if there are other ways to accomplish this, in this case use Set-MgUserLicense and set license options:
# Note: When connecting using app, scopes determined by app-permissions, eg 'User.ReadWrite.All', 'Directory.ReadWrite.All'
Connect-MgGraph -ClientId $appid -CertificateThumbprint $CertThumbprint -TenantId $TenantId 
Select-MgProfile Beta
Get-MgUser -UserId olrik@myexchangelabs.com | Select -ExpandProperty AssignedLicenses
$LicenseOptions= @{
    SkuId= '6fd2c87f-b296-42f0-b197-1e91e994b900'		#ENTERPRISEPACK
    DisabledPlans= 'a23b959c-7ce8-4e57-9140-b90eb88a9e97'	#SWAY
}
Set-MgUserLicense -UserID olrik@myexchangelabs.com -AddLicenses $LicenseOptions -RemoveLicenses @()




#######################################################
# Extra
#######################################################
# With REST, no more PSSessions so this doesn't work:
# Invoke-Command -Session (Get-PSSession) -ScriptBlock { .. }
# Instead:
$ScriptBlock1 = {
    Get-Mailbox -ResultSize Unlimited | Get-MailboxStatistics
}
Invoke-Command -ScriptBlock $ScriptBlock1
# Multiple sessions: Use prefixes, eg Connect-ExchangeOnline .. -Prefix Prod
# but requires changing all related cmdlets to add these prefixes.

# When limiting ResultSet, order is undetermined and returned set may differ with same call.
# Get-EXOMailbox -ResultSize 1000 -PropertySet StatisticsSeed | Get-EXOMailboxStatistics

# Which EOM module cmdlets support PropertySets
(Get-Command -Module ExchangeOnlineManagement).Where{$_.Parameters.propertySets}

# What are PropertySets options
[System.Enum]::GetNames( [Microsoft.Exchange.Management.RestApiClient.GetExoMailbox+PropertySet])

# Some performance indicators
# Get-Mailbox vs Get-EXOMailbox retrieving Quota properties
1..10 | ForEach-Object { 
    Write-Host ('{0} ' -f $_) -NoNewLine
    (Measure-Command { Get-Mailbox -ResultSize 1000 | Select-Object Identity, *Quota }).TotalSeconds 
} | Measure-Object -Average

1..10 | ForEach-Object { 
    Write-Host ('{0} ' -f $_) -NoNewLine
    (Measure-Command { Get-EXOMailbox -ResultSize 1000 -PropertySets Quota }).TotalSeconds 
} | Measure-Object -Average 

# ForEach
1..10 | ForEach-Object {
    Write-Host ('{0} ' -f $_) -NoNewLine
    (Measure-Command {
        $MbxSet= Get-EXOMailbox -ResultSize 1000 -PropertySet StatisticsSeed 
        ForEach( $Mbx in $MbxSet) {
            Get-EXOMailboxStatistics -Identity $Mbx.Identity 
        }
    }).TotalSeconds
} | Measure-Object -Average

# Pipeline
1..10 | ForEach-Object {
    Write-Host ('{0} ' -f $_) -NoNewLine
    (Measure-Command {
        Get-EXOMailbox -ResultSize 1000 -PropertySet StatisticsSeed | Get-EXOMailboxStatistics
    }).TotalSeconds
} | Measure-Object -Average

