###############################################################################################
# TEC2024 Workshop 
# Technical Experts Conference 2024 Workshop Demo Script
# Michel de Rooij & Jaap Wesselius
# https://github.com/michelderooij/Sessions
###############################################################################################

# Set execution policy to allow running scripts
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process -Force
function Prompt { "PS >"}
Clear-Host

######################################################
# Install required modules
######################################################
#Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
#Install-Module -Name Microsoft.Graph.Application -Scope CurrentUser -Force -AllowClobber
#Install-Module -Name Microsoft.Graph.Users -Scope CurrentUser -Force -AllowClobber
#Install-Module -Name Microsoft.Graph.Group -Scope CurrentUser -Force -AllowClobber
#Install-Module -Name Microsoft.Graph.DirectoryManagement -Scope CurrentUser -Force -AllowClobber
#Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber

# Beta
#Install-Module -Name Microsoft.Graph.Beta.Authentication -Scope CurrentUser -Force -AllowClobber
#Install-Module -Name Microsoft.Graph.Beta.Users -Scope CurrentUser -Force -AllowClobber
#Install-Module -Name Microsoft.Graph.Beta.Group -Scope CurrentUser -Force -AllowClobber

######################################################
# Connect to Graph 
######################################################
Import-Module -Name Microsoft.Graph.Authentication # Without Import, will be loaded as-needed by cmdlets
Import-Module -Name Microsoft.Graph.Users 

Find-Module Microsoft.Graph | Select-Object -ExpandProperty Dependencies | Format-Table -AutoSize
# Better digestable
Find-Module Microsoft.Graph | Select-Object -ExpandProperty Dependencies | Select-Object Name,RequiredVersion,MinimumVersion,MaximumVersion | Format-Table -AutoSize

# Interactive (delegated)
Connect-MgGraph -Scopes User.ReadWrite.All,Directory.ReadWrite.All,Group.ReadWrite.All

Get-MgContext

######################################################
# Connect to Exchange Online
######################################################
# This will install the GA build
Install-Module ExchangeOnlineManagement
#Install-module ExchangeOnlineManagement -AllowPrerelease
Import-Module ExchangeOnlineManagement
# Tip: Use prefixes when accessing multiple tenants within same session
#Import-Module ExchangeOnlineManagement -Prefix Dev

# Connect interactively
(Get-MgContext).Account
Connect-ExchangeOnline -UserPrincipalName (Get-MgContext).Account

######################################################
# Who Am I?
######################################################
Get-MgContext
[Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance.AuthContext

# Fetch my access token
$Response = Invoke-MgGraphRequest -Method Get -Uri /v1.0/me -OutputType HttpResponseMessage
$Response.RequestMessage.Headers.Authorization

# Decode token
$Response.Request`Message.Headers.Authorization.Parameter | Clip
Start-Process microsoft-edge:https://jwt.io

######################################################
# USERS
######################################################
#region Users
Get-MgUser -UserId philip@myexchangelabs.com
$User= Get-MgUser -UserId philip@myexchangelabs.com

# For Beta would prefix noun with Beta, eg.
Get-MgBetaUser -UserId philip@myexchangelabs.com

# Find info for a cmdlet, eg Update-MgUser, incl module & permissions
Find-MgGraphCommand -Command Update-MgUser | Format-List Command,Module,Permissions
Find-MgGraphCommand -Command Update-MgUser | Select -ExpandProperty Permissions

# Updating should work, we connected using User.ReadWrite.All
Update-MgUser -UserId michel@myexchangelabs.com -DisplayName 'Michel de Rooij'

# Help for a cmdlet
Get-Help Get-MgUser
Get-Help Get-MgUser -Examples
Get-Help Get-MgUser -ShowWindow

$User= Get-MgUser -UserId philip@myexchangelabs.com
$User | ConvertTo-Json
$User.ToJsonString()
# Tip: $User.ToJsonString() | code -

# Some properties are empty by default
$User.LastPasswordChangeDateTime
# Explicitly request these properties
$User= Get-MgUser -UserId philip@myexchangelabs.com -Property LastPasswordChangeDateTime
$User.LastPasswordChangeDateTime

$passwordProfile = @{
    forceChangePasswordNextSignIn = $true
    password = ([char[]]([char]33..[char]95) + ([char[]]([char]97..[char]126)) + 0..9 | sort {Get-Random})[0..20] -join ''
}
Update-MgUser -UserId francis@myexchangelabs.com -passwordProfile $passwordProfile

# Be aware of Eventual lag
Get-MgUser -UserId francis@myexchangelabs.com | Select-Object DisplayName
Update-MgUser -UserId francis@myexchangelabs.com -DisplayName 'Not Francis'
Get-MgUser -UserId francis@myexchangelabs.com | Select-Object DisplayName
Get-MgUser -Filter "userPrincipalName eq 'francis@myexchangelabs.com'" -ConsistencyLevel Eventual -CountVariable Count | Select-Object DisplayName
Update-MgUser -UserId francis@myexchangelabs.com -DisplayName 'Francis Blake'

# Get licensing positions
Get-MgSubscribedSku  | Select-Object SkuPartNumber,CapabilityStatus,ConsumedUnits,@{n='Enabled';e={$_.PrepaidUnits.Enabled}},@{n='SKU';e={ $_.SubscriptionIds -join ';'}} | Format-Table -AutoSize

# Get-SubscribedSku doesn't return lifecycle information, use MgDirectorySubscription
$skuData= Get-MgDirectorySubscription
#$skuData = Invoke-MgGraphRequest -Uri https://graph.microsoft.com/v1.0/directory/subscriptions -Method Get
$skuData | Select skuPartNumber,createdDateTime,isTrial,nextLifecycleDateTime,TotalLicenses | Format-Table -AutoSize 

$Now= Get-Date -Format 'o'  # ISO 8601 format
$skuData | Select-Object skuPartNumber, createdDateTime,isTrial,nextLifecycleDateTime,TotalLicenses,
    @{n='DaysToRenewal';e={ (New-Timespan -Start $Now -End (Get-Date $_.nextLifecycleDateTime -Format 'o')).Days}} | Format-Table -AutoSize

# Configure licensing sample:
$SKUToApply= Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq 'Office_365_w/o_Teams_Bundle_E5' }
# Which plans are in this SKU?
# what service plans are in SKU $SKUToApply?
$SKUToApply.ServicePlans | Select-Object ServicePlanId,ServicePlanName

# Disable SWAY
$DisablePlans= $SKUToApply.ServicePlans | Where-Object {$_.ServicePlanName -eq 'SWAY'}

$LicenseOptions= @{
    SkuId= $SKUToApply.SkuId
    DisabledPlans= $DisablePlans.ServicePlanId
}
Set-MgUserLicense -UserID olrik@myexchangelabs.com -AddLicenses $LicenseOptions -RemoveLicenses @()
Get-MgUserLicenseDetail -UserID olrik@myexchangelabs.com | Select-Object SkuId,ServicePlanId,ServicePlanName

# Remove a license
$SKUtoRemove= Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq 'FLOW_FREE' }
# Need to specify addLicenses 
Set-MgUserlicense -UserId olrik@myexchangelabs.com -AddLicenses @() -RemoveLicenses $SKUtoRemove.SkuId 

# Get one page of users
Get-MgUser
# Get all users, module uses paging if necessary
Get-MgUser -All
# Get all users using page size of 1 - use debug to see all the individual calls
Get-MgUser -All -PageSize 1 -Debug

# Count no. of users
Get-MgUser -CountVariable Records -ConsistencyLevel eventual -All | Out-Null
$Records

# Pay attention to case-sensitivity in oData filters
Get-MgUser -Filter 'AccountEnabled eq True' -All
Get-MgUser -Filter 'AccountEnabled eq true' -All
# Pay attention to supported operators
Get-MgUser -Filter 'AccountEnabled ne true' -All
# And pay attention to advanced queries
Get-MgUser -Filter 'not(AccountEnabled eq true)' -All
Get-MgUser -Filter 'not(AccountEnabled eq true)' -All -ConsistencyLevel eventual -CountVariable Records
$Records

# Which users have EXCHANGE_S_ENTERPRISE (Plan 2)
Get-MgSubscribedSku -Property SkuPartNumber, ServicePlans | Select-Object SkuPartNumber, @{n='ServicePlans';e={$_.ServicePlans.ServicePlanId}}
Get-MgUserLicenseDetail -UserId philip@myexchangelabs.com

# Get all groups someone is a member of - Get-MgUserMemberOf requires GroupMember.Read.All
Connect-MgGraph -Scopes GroupMember.Read.All
$Groups= Get-MgUserMemberOf -UserId philip@myexchangelabs.com
ForEach( $Group in $Groups) {
    $Object= [PsCustomObject]@{
        Id= $Group.Id
        DisplayName= $Group.AdditionalProperties['displayName']
        Type= $Group.AdditionalProperties['@odata.type']
    }
    $Object
}

# Find users not member of a group
# Won't work -Filter on Expanded property:
Get-MgUser -ExpandProperty MemberOf –Filter 'memberOf/$count eq 0' 
# Workaround
Get-MgUser -ExpandProperty MemberOf | Where-Object {$_.memberOf.Count -eq 0}  

# For Sign-In activity, need AuditLog.Read.All
Connect-MgGraph -Scopes AuditLog.Read.All
# SignInActivity not default property, need to use -Property parameter
# Also, UserId with Property only works when specifying Guid, so we use Filter
$User= Get-MgUser -Filter "userPrincipalName eq 'michel@myexchangelabs.com'" -Property SignInActivity
$User | Select-Object Id, DisplayName, @{n='LastSignIn';e={$_.SignInActivity.LastSignInDateTime}}

# Get all users not signed in after January 1st, 2024
Get-MgUser -Filter 'signInActivity/lastSignInDateTime le 2024-01-01T00:00:00Z'
# Get all users not signed in last 90 days
$NinetyDaysAgo= [DateTime]::Today.AddDays(-90)
$FormattedDate= $NinetyDaysAgo.ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ssZ")
Get-MgUser -Filter ('signInActivity/lastSignInDateTime le {0}' -f $FormattedDate)

# Get licensed users 
Get-MgUser -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -Count Records -All
# Get licensed users with password age > 90 days
$Users= Get-MgUser -Filter 'assignedLicenses/$count ne 0' -Property id,displayName,lastPasswordChangeDateTime -ConsistencyLevel eventual -Count Records -All
ForEach( $User in $Users) {
    $User | Select-Object Id,DisplayName,lastPasswordChangeDateTime 
}

# Get users with specific license
Get-MgUser -Filter "assignedPlans/any(x:x/servicePlanId eq efb87545-963c-4e0d-99df-69c6916d9eb0)" -All -ConsistencyLevel eventual -CountVariable Records  

# Get groups used for group-based license assignment(s)
Get-MgGroup -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -Count Records -All


# Get Users using Search to filter. Advanced query mode.
Get-MgUser -Search "userPrincipalName:philip@myexchangelabs.com" -ConsistencyLevel eventual -CountVariable Records
# Combine Search & Filter
Get-MgUser -Search "userPrincipalName:michel" -Filter "accountEnabled eq true" -ConsistencyLevel eventual -CountVariable Records
#endregion

# Sign In Activity
Import-Module Microsoft.Graph.Reports
Get-MgAuditLogSignIn -Top 10 | Select createdDateTime,AppId,appDisplayName,userDisplayName

# Working with Extension Attributes (similar principle for custom 'security' attributes)
Update-MgUser –UserId newtest@myexchangelabs.com -onPremisesExtensionAttributes @{ extensionAttribute15= 'THX1138'}
Get-MgUser -Filter "onPremisesExtensionAttributes/ExtensionAttribute15 eq 'THX1138'" -ConsistencyLevel eventual -CountVariable count

# Clearing values: this won't work:
Update-MgUser –UserId newtest@myexchangelabs.com -CompanyName $null
# Workaround .. use direct patch (=update) queries
$User= Get-MgUser -UserId newtest@myexchangelabs.com
Invoke-MgGraphRequest -Method Patch -Uri ('https://graph.microsoft.com/v1.0/Users/{0}' -f $User.Id) -Body @{ CompanyName= $null } -Debug

# Working with Administrative Units
Get-MgDirectoryAdministrativeUnit
$AU= Get-MgDirectoryAdministrativeUnit -Filter "displayName eq 'Demo AU'"
$User= Get-MgUser -UserId newtest@myexchangelabs.com
$odataId = 'https://graph.microsoft.com/v1.0/users/{0}' –f $User.Id
New-MgDirectoryAdministrativeUnitMemberByRef -AdministrativeUnitId $AU.Id -OdataId $odataId
# Remove call differs, as AUs can contain users, groups, devices etc.
Remove-MgDirectoryAdministrativeUnitMemberByRef -AdministrativeUnitId $AU.Id –DirectoryObjectId $User.Id

######################################################
# GROUPS
######################################################
#region groups
Find-MgGraphCommand -Command New-MgGroup | Format-List Command,Module,Permissions

# Trying to create a DG (or MESG) in Graph won't work:
New-MgGroup -UniqueName G-SGDemo -MailNickname G-SGDemo -DisplayName 'G-SGDemo' -MailEnabled -SecurityEnabled:$false
# Create "plain old security group"
New-MgGroup -UniqueName G-SGDemo -MailNickname G-SGDemo -DisplayName 'G-SGDemo' -MailEnabled:$false -SecurityEnabled

Get-MgGroup -Filter "MailNickname eq 'G-SGDemo'" 

$Group= Get-MgGroup -Filter "MailNickname eq 'G-SGDemo'" 
Remove-MgGroup -GroupId $Group.Id

# Create security group with dynamic membership using splatting
$Params= @{
    UniqueName= 'G-DSGDemo'
    MailNickname= 'G-DSGDemo'
    DisplayName= 'G-DSGDemo'
    MailEnabled= $false
    SecurityEnabled= $true
    GroupTypes= 'DynamicMembership'
    MembershipRuleProcessingState= 'On'
    MembershipRule= 'user.department -eq "Sales"'
}
New-MgGroup @Params

Get-MgGroup -Filter "MailNickname eq 'G-DSGDemo'" 
$Group= Get-MgGroup -Filter "MailNickname eq 'G-DSGDemo'" 
# Preview dynamic members
Get-MgGroupMember -GroupId $Group.id | Select-Object *
# Use additionalProperties to get other default properties, eg DisplayName
# NOTE: oData references are case-sensitive, see below examples:
Get-MgGroupMember -GroupId $Group.id | Select-Object Id, @{n='DisplayName';e={$_.additionalProperties.DisplayName}}
Get-MgGroupMember -GroupId $Group.id | Select-Object Id, @{n='DisplayName';e={$_.additionalProperties.displayName}}

Remove-MgGroup -GroupId $Group.Id

# Distribution Groups (&  Mail-Enabled Security Groups, use Security for Type )
New-DistributionGroup -Name G-DGDemo -primarySmtpAddress g-dgdemo@myexchangelabs.com –Type Distribution
Add-DistributionGroupMember -Identity G-DGDemo -Member philip@myexchangelabs.com
Remove-DistributionGroup -Identity G-DGDemo -Confirm:$false

# Dynamic Distribution Groups
New-DynamicDistributionGroup -Name G-DDGDemo -primarySmtpAddress g-ddgdemo@myexchangelabs.com -ConditionalDepartment 'Sales'

$Group= Get-DynamicDistributionGroup -Identity G-DDGDemo
Get-Recipient -RecipientPreviewFilter $Group.RecipientFilter

Remove-DistributionGroup -Identity G-DDGDemo -Confirm:$false

# M365 Groups
New-MgGroup –DisplayName 'Marketing via Graph' –GroupTypes 'Unified' -MailEnabled –SecurityEnabled:$False -MailNickname M365G-Marketing-Graph
Get-MgGroup -Filter "MailNickname eq 'M365G-Marketing-Graph'"
$Group= Get-MgGroup -Filter "MailNickname eq 'M365G-Marketing-Graph'"
Update-MgGroup -GroupId $Group.Id -HideFromAddressLists:$true
# We can check EXO using Group Id (will locate it using externalDirectoryObjectId in EXO)
Get-UnifiedGroup -Identity $Group.Id | Select-Object Name,Guid,externalDirectoryObjectId,HiddenFromAddressListsEnabled
Add-RecipientPermission -Identity $Group.Id -Trustee philip@myexchangelabs.com -AccessRights SendAs -Confirm:$false
Set-UnifiedGroup -Identity $Group.Id -GrantSendOnBehalfTo philip@myexchangelabs.com

Remove-MgGroup -GroupId $Group.Id

# Same, via EXO
New-UnifiedGroup –DisplayName 'Marketing via EXO' –Alias M365G-Marketing-EXO 
Set-UnifiedGroup -Identity M365G-Marketing-EXO 

# Manage members via EXO
Add-UnifiedGroupLinks -Identity M365G-Marketing-EXO -LinkType Members -Links philip@myexchangelabs.com
Add-RecipientPermission -Identity M365G-Marketing-EXO -Trustee philip@myexchangelabs.com -AccessRights SendAs -Confirm:$false
Set-UnifiedGroup -Identity M365G-Marketing-EXO -GrantSendOnBehalfTo philip@myexchangelabs.com
Remove-UnifiedGroup -Identity M365G-Marketing-EXO -Confirm:$false

# Group Settings
Get-MgGroupSettingTemplateGroupSettingTemplate | Select-Object Id,displayName

# Create new group setting (when not yet exists)
$Template= Get-MgGroupSettingTemplateGroupSettingTemplate | Where-Object {$_.DisplayName -eq 'Group.Unified'}
$Template.Values | Select-Object Name,DefaultValue,Type,Description
New-MgGroupSetting -TemplateId $Template.Id -Values $Values

# Get current group setting and change an element
$Settings= Get-MgGroupSetting | Where-Object {$_.DisplayName -eq 'Group.Unified'}
$params = @{
    templateId = $Settings.Id
    values = @(
        @{
            name = "EnableGroupCreation"
            value = "false"
        }
    )
}
# Write new settings back
Update-MgBetaDirectorySetting –DirectorySettingId $Settings.Id –Body $params
# Verify our change
Get-MgGroupSetting | Where-Object {$_.DisplayName -eq 'Group.Unified'} | Select-Object -ExpandProperty Values

#endregion
######################################################
# MAILBOXES
######################################################
#region Mailboxes

#endregion
######################################################
# MISC
######################################################
#region Misc

# Create self-signed cert. For older PS versions, use script https://github.com/SharePoint/PnP-Partner-Pack/blob/master/scripts/Create-SelfSignedCertificate.ps1
# Clean up old cert: Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Subject -eq 'CN=TEC2024' } | Remove-Item
$Certificate=New-SelfSignedCertificate –Subject 'TEC2024' -CertStoreLocation 'Cert:\CurrentUser\My' -NotBefore (Get-Date).AddDays(-1) -NotAfter (Get-Date).AddDays(30)

# Export certificate to .pfx and .cer with password
Export-Certificate -Cert $Certificate -FilePath .\TEC2024.cer -Type CERT

#  Sets CBA variables $TenantId and $AppId
. .\Set-DemoVars.ps1

$Certificate= Get-ChildItem -path Cert:\CurrentUser\My | Where-Object { $_.Subject -eq 'CN=TEC2024' } | Sort-Object -Property NotBefore -Descending | Select-Object -First 1
$Certificate | Format-List Subject, Thumbprint, NotBefore, NotAfter
$CertThumbprint= $Certificate.Thumbprint

# Connecting Graph PowerShell using CBA
# Note: Sign-In logs will show app & cert, not who used it.
Connect-MgGraph -CertificateThumbprint $CertThumbprint -ClientId $AppId -TenantId $TenantId

# Connect to EXO using CBA
# Note that Exchange needs domain for organization instead of tenant id, but we're already 
# connected to graph, so we can fetch this as initial domain from the org's verified domains:
$OrgId= (Get-MgOrganization).VerifiedDomains | Where-Object { $_.IsInitial -eq $true } | Select-Object -ExpandProperty Name
Connect-ExchangeOnline -Organization $OrgId -AppId $AppId -CertificateThumbprint $CertThumbprint

# Tip:Get TenantId from domain (public info)
$Domain='myexchangelabs.onmicrosoft.com'
(Invoke-RestMethod -Uri ('https://login.microsoftonline.com/{0}/v2.0/.well-known/openid-configuration') -f $Domain).jwks_uri.split('/')[3]

# Could use same principle for Teams a.o.
#Connect-MicrosoftTeams -CertificateThumbprint $CertThumbprint -ApplicationId $AppId -TenantId $TenantId

# Paging
$Iteration= 1
$AllUsers= [System.Collections.ArrayList]::new()
# Query first page with users (pagesize 10 for demonstration)
$Response= Invoke-MgGraphRequest -Method Get -Uri 'https://graph.microsoft.com/v1.0/users?$top=10' -OutputType PSObject
$null= $Response.value.ForEach( { $AllUsers.Add( $_ ) } )
While( $null -ne $Response.'@odata.nextLink') {
    $Iteration++
    # Fetch next page until nextLink is empty (no more data)
    $Response= Invoke-MgGraphRequest -Method Get -Uri $Response.'@odata.nextLink'
    $null= $Response.value.ForEach( { $AllUsers.Add( $_ ) } )
}
Write-Host ('Total users {0} fetched in {1} iterations' -f $AllUsers.Count, $Iteration)

# MS Commerce
Install-Module MSCommerce
Import-Module MSCommerce
Connect-MSCommerce
#Connect-MSCommerce -CertificateThumbprint $CertThumbprint -ClientId $AppId -TenantId $TenantId

Get-MSCommercePolicies
Get-MSCommercePolicy -PolicyId AllowSelfServicePurchase
# Get products assigned to policy
Get-MSCommerceProductPolicies -PolicyId AllowSelfServicePurchase
# Disable all self-service purchase options
Get-MSCommerceProductPolicies -PolicyId AllowSelfServicePurchase | ForEach-Object { 
    Update-MSCommerceProductPolicy -PolicyId $_.PolicyId -ProductId $_.ProductId -Enabled $false 
}

# Generate simple report, using ArrayList 
$Report= [System.Collections.ArrayList]::new()
Get-MgUser -All | ForEach-Object {
    $LicenseDetails= Get-MgUserLicenseDetail -UserId $_.Id
    $Obj= [PSCustomObject][ordered]@{
        User= $_.UserPrincipalName
        Licenses= $LicenseDetails.SkuPartNumber -Join ';'
    }
    $null= $Report.Add( $Obj)
}
$Report

# ArrayList instead of Array
# Add 10.000 numbers to array
$Array= @()
Measure-Command {
  1..10000 | ForEach{ $Array+= $_ }
} | Select TotalMilliseconds

# Add 10.000 numbers to arraylist
$ArrayList= [System.Collections.ArrayList]::new()
Measure-Command {
  1..10000 | ForEach{ $null= $ArrayList.Add( $_)}
} | Select TotalMilliseconds

#endregion
################################################################################
# THE END
################################################################################
