# TEC2023 Challenges Maintaining Scripts in the Modern Age
# Technical Experts Conference 2023 Demo Script
# https://github.com/michelderooij/Sessions
#
# Tip: In scripts, use Requires for required PS version, modules, and use -RunAsAdministrator if it requires elevation
#Requires -Version 5.0 –Modules AzureADPreview,Microsoft.Graph.Beta,Microsoft.Graph,ExchangeOnlineManagement,PSAzureMigrationAdvisor,Microsoft.Graph.Compatibility.AzureAD 

# Glorious PS theme & prompt by oh-my-posh, see https://ohmyposh.dev

function Prompt { 'PS> '}
Clear-Host

Install-module ExchangeOnlineManagement -AllowPrerelease
# This will install the GA build
#Install-Module ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement
# Tip: Use prefixes when accessing multiple tenants
#Import-Module ExchangeOnlineManagement -Prefix Dev

# Reporting/Updating Modules
# Helper functions from https://github.com/michelderooij/Connect-Office365Services
Report-Office365Modules -AllowPrerelease
# Update-Office365Modules -AllowPrerelease
# Script from https://github.com/12Knocksinna/Office365itpros/blob/master/UpdateOffice365PowerShellModules.PS1
.\Demo\UpdateOffice365PowerShellModules.PS1

######################################################
### Connect using Certificate-Based Authentication
######################################################
# Set $TenantId, $OrgId and $AppId, and get Thumbprint of cert to use for authentication
. .\SetDemoVars.ps1

$CertThumbprint= (Get-ChildItem Cert:\CurrentUser\My | Where {$_.Subject -eq 'CN=TEC2023DEMO'}).Thumbprint
# Note: Recommend 1 certificate per app/service principal: Sign-In logs will show app, not who used it.

# Connect using CBA and REST mode
Connect-ExchangeOnline -Organization $OrgId -AppId $AppId -CertificateThumbprint $CertThumbprint -ShowBanner:$False

# Alternative CBA connect methods - first we need password to read the pfx:
$PfxPwd= Read-Host -AsSecureString
# Connect with certificate from file with password:
Connect-ExchangeOnline -CertificateFilePath (Resolve-Path .\TEC2023DEMO.pfx).Path -CertificatePassword $PfxPwd -AppID $AppId -Organization $OrgId
# Connect with certificate object, created from file w/password:
$CertObj= [System.Security.Cryptography.X509Certificates.X509Certificate2]::new( (Resolve-Path .\TEC2023DEMO.pfx).Path, [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $PfxPwd)))
Connect-ExchangeOnline -Certificate $CertObj -AppID $AppId -Organization $OrgId

# Connect other workload shells using this certificate, provided you have the permission
# Workaround for PS7 compatibility issue with AzureAD module, this doesn't work:
Import-Module AzureADPreview -UseWindowsPowerShell
# Note: PS7 uses other module folder, add the 'old' folder to $env:PSModulePath or explicit import:
Import-Module -UseWindowsPowerShell 'C:\Program Files\PowerShell\Modules\AzureADPreview\2.0.2.183\AzureADPreview.psd1'
Connect-AzureAD -TenantId $OrgId -CertificateThumbprint $CertThumbprint -ApplicationId $AppId

# Connect to Graph using app & auth (certificate)
Connect-MgGraph -CertificateThumbprint $CertThumbprint -ClientId $AppId -TenantId $TenantId
# Note: When you see Connect-Graph or Connect-MgGraph, Connect-Graph is an alias for Connect-MgGraph

# Could use same principle for Teams, etc., provided permissions  same authentication method
#Connect-MicrosoftTeams -CertificateThumbprint $CertThumbprint -ApplicationId $AppId -TenantId $TenantId

#######################################################
# Analyzing EXO scripts for assessment
# https://github.com/michelderooij/Analyze-ExoScript
#######################################################

# Refresh cmdlet definitions from current EXOM version
#.\Analyze-ExoScript.ps1 -Organization $OrgId -CertificateThumbprint $CertThumbprint -AppId $AppId -Refresh

# Analyse scripts and report on Exchange cmdlets
.\Analyze-ExoScript.ps1 -File '.\Demo\' -Organization $OrgId -CertificateThumbprint $CertThumbprint -AppId $AppId | ft -a

# Analyse scripts and report which scripts can perhaps benefit from REST-backed to REST-based mapping
.\Analyze-ExoScript.ps1 -File '.\Demo\' -Organization $OrgId -CertificateThumbprint $CertThumbprint -AppId $AppId | Where {$_.Alt} | ft -a

# Analyse scripts reporting all cmdlets (not only Exchange ones) and report which ones use PSSession (eg Remote PowerShell)
.\Analyze-ExoScript.ps1 -File '.\Demo\' -Organization $OrgId -CertificateThumbprint $CertThumbprint -AppId $AppId -ShowAll | Where {$_.Command -like '*-PSSession'} | ft -a

# Analyse scripts and report which scripts use Connect-ExchangeOnline using Credential (BA)
.\Analyze-ExoScript.ps1 -File '.\Demo\' -Organization $OrgId -CertificateThumbprint $CertThumbprint -AppId $AppId | Where {$_.Command -like 'Connect-ExchangeOnline' -and $_.Parameters -contains 'Credential'} | ft -a

#######################################################
# Analyzing scripts for MSOnline/AzureAD usage 
# Uses: 
# https://merill.net/2022/04/graph-powershell-conversion-analyzer/
# https://github.com/FriedrichWeinmann/PSAzureMigrationAdvisor
#
# For individual, interactive analysis: 
# https://graphpowershell.merill.net
#######################################################

# Perform analysis
#Install-Module PSAzureMigrationAdvisor -Scope AllUsers
Import-Module PSAzureMigrationAdvisor
Get-ChildItem .\Demo\*.ps1 | Read-AzScriptFile
# Or limit to MSOnline, AzureAD or AzureADPreview cmdlets, e.g.
Get-ChildItem .\Demo\*.ps1 | Read-AzScriptFile -Type MSOnline

# Example of converted 'MSOnline' cmdlets:
# $lic= New-MsolLicenseOptions -AccountSkuId Contoso:BPOS_STANDARD -DisabledPlans EXCHANGE_STANDARD
# Set-MsolUserLicense .. -LicenseOptions $lic
# Note: When connecting using app, scopes determined by app-permissions, eg 'User.ReadWrite.All', 'Directory.ReadWrite.All'
# Important: Licensing beta endpoint only
Connect-MgGraph -ClientId $appid -CertificateThumbprint $CertThumbprint -TenantId $TenantId -NoWelcome

# Important: Licensing beta endpoint only!
Get-MgUser -UserId olrik@contoso.com | Select -ExpandProperty AssignedLicenses
Get-MgBetaUser -UserId olrik@contoso.com | Select -ExpandProperty AssignedLicenses

# Configure licensing sample:
$LicenseOptions= @{
    SkuId= '6fd2c87f-b296-42f0-b197-1e91e994b900'		#ENTERPRISEPACK
    DisabledPlans= 'a23b959c-7ce8-4e57-9140-b90eb88a9e97'	#SWAY
}
Set-MgBetaUserLicense -UserID olrik@contoso.com -AddLicenses $LicenseOptions -RemoveLicenses @()

# ConsistencyLevel
Get-MgUser -Filter 'accountEnabled eq false'
# This doesn't work - 'ne' requires advanced query
Get-MgUser -Filter 'accountEnabled ne false' 
# Use Beta endpoint, and adv. query 'mode'
Get-MgBetaUser -Filter 'accountEnabled ne false' -ConsistencyLevel eventual -Count 1




#######################################################
### TBD
#######################################################

# M365DevProxy
Start-Process -FilePath C:\M365DevProxy\m365proxy.exe -ArgumentList '-p','8080','--failure-rate','50'
# Configuring proxy for PowerShell Core
[System.Net.Http.HttpClient]::DefaultProxy = New-Object System.Net.WebProxy('0.0.0.0:8080')
# .. Do stuff. For monitoring EXO REST calls, add https://outlook.office.com/adminApi to UrlsToWatch in m365proxyrc.json
[System.Net.Http.HttpClient]::DefaultProxy = New-Object System.Net.WebProxy($null)
Stop-Process -Name m365proxy

#######################################################
# Extra
#######################################################

# When limiting ResultSet, order is undetermined on live systems, and results may differ with identical call.
Get-EXOMailbox -ResultSize 10 -PropertySet StatisticsSeed | Get-EXOMailboxStatistics

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

