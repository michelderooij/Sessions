# TEC2025 Session
# https://github.com/michelderooij/Sessions

# Tip: Leverage Requires for required PS version, modules, and -RunAsAdministrator as required
#Requires -Version 5.1 –Modules Microsoft.Graph.Authentication

function Prompt { "PS >" }
Clear-Host

# Set Guids for Demo
. .\Set-DemoVars.ps1

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Groups

# Connect to MS Graph using registered app with certificate
Connect-MgGraph -tenantId $TenantId -appId $AppId -CertificateThumbprint $CertThumbprint

###############################################################################
# Invoke-MgGraphRequest
###############################################################################
Get-MgUser -UserId michel@contoso.com

Invoke-MgGraphRequest -Method GET -Uri https://graph.microsoft.com/v1.0/users/michel@contoso.com
Invoke-MgGraphRequest -Method GET -Uri https://graph.microsoft.com/v1.0/users/a9978aa6-528c-4eca-9235-71c587d9fdc1 -ContentType PSObject

Set-MgUser -UserId a9978aa6-528c-4eca-9235-71c587d9fdc1 -CompanyName 'Contoso'
# JSON payload
$body= @'
{
    "companyName": "Contoso"
}
'@
Invoke-MgGraphRequest -Method PATCH -Uri https://graph.microsoft.com/v1.0/users/michel@contoso.com -Body $body
# Hash tables are also accepted
$props= @{
    companyName= 'Contoso'
}
$JSONbody= $props | ConvertTo-Json
Invoke-MgGraphRequest -Method PATCH -Uri https://graph.microsoft.com/v1.0/users/michel@contoso.com -Body $JSONbody

###############################################################################
# Select, Filter, OrderBy, Expand
# Property support https://learn.microsoft.com/en-us/graph/aad-advanced-queries
###############################################################################

# Select
Get-MgUser -Select Id,country | Where-Object { $_.country -eq 'NL'} | Select Id,Country
Get-MgUser -Filter "country eq 'NL'"

Get-MgUser -Select Id,DisplayName,accountEnabled | Where-Object { $_.accountEnabled}
Get-MgUser -Select Id,DisplayName -Filter 'accountEnabled eq true'

Get-MgUser -All -OrderBy DisplayName

# Expand
$User= Get-MgUser -UserId michel@contoso.com -Select Id -ExpandProperty MemberOf | Select-Object Id,memberOf
$User.MemberOf.AdditionalProperties

###############################################################################
# Pagination
###############################################################################
$Results= [System.Collections.Generic.List[PSObject]]::new()
$Data= Invoke-MgGraphRequest –Method Get –Uri 'https://graph.microsoft.com/v1.0/users?$top=10' -OutputType PSObject
$Data.Value | ForEach-Object { $Results.Add( $_ ) }
While( $Data.'@odata.nextLink') {
    $Data= Invoke-MgGraphRequest –Method Get –Uri $Data.'@odata.nextLink' -OutputType PSObject
    $Data.Value | ForEach-Object { $Results.Add( $_ ) }
}
$Results.Count

# Use debug to see what is happening when using Cmdlets
Get-MgUser -PageSize 5 -Top 10 -Debug

###############################################################################
# Parallelism
###############################################################################
$ResultsBag= [System.Collections.Concurrent.ConcurrentBag[System.Object]]::new()
Get-MgUser -All | ForEach-Object -Parallel {
    $UserGroups= Get-MgUserMemberOf -UserId $_.Id -All
    If( $UserGroups) {
        $Obj= [PSCustomObject]@{
            User= $_.UserPrincipalName
            GroupCount= $UserGroups.Count
        }
        ($using:ResultsBag).Add( $Obj)
    }
}
$ResultsBag | Sort-Object GroupCount -Descending

###############################################################################
# Batching
###############################################################################
# We use a stack collection to easy push/pop user ids and check if we exhausted the list
$UserStack= [System.Collections.Stack]::new()
Get-MgUser -Filter "startswith(DisplayName,'Demo')" -All -Select Id | ForEach-Object { $UserStack.Push( $_.Id) }

$BatchSize= 20
While( $UserStack.Count -gt 0) {
    Write-Host ('Creating batch with {1} users left to process' -f ( $Batches.Count + 1), $UserStack.Count)
    $BatchRequest= [System.Collections.Generic.List[Object]]::new()
    While( $UserStack.Count -gt 0 -and $BatchRequest.Count -lt $BatchSize) {
        $UserId= $UserStack.Pop()
        $Request= @{
            Id= $BatchRequest.Count + 1
            Method= 'GET'
            Url= '/users/{0}?$select=id,companyname,displayname,city,country' -f $UserId
        }
        $BatchRequest.Add( $Request)
    }
    $Body= @{ 'requests'= $BatchRequest }
    $Response= Invoke-MgGraphRequest -Method POST 'https://graph.microsoft.com/v1.0/$batch' -Body $Body
    $Response.responses.Body | Select-Object id,companyname,displayName,city,country
}

###############################################################################
# Parallel & Expand
###############################################################################
$ResultsBag= [System.Collections.Concurrent.ConcurrentBag[System.Object]]::new()
Get-MgUser -All -Expand MemberOf  | ForEach-Object -Parallel {
    If( $_.MemberOf) {
        $Obj= [PSCustomObject]@{
            User= $_.UserPrincipalName
            GroupCount= ($_.memberOf).Count
        }
        ($using:ResultsBag).Add( $Obj)
    }
}
$ResultsBag | Sort-Object GroupCount -Descending

###############################################################################
# PowerShell Practices
###############################################################################

# Where-Object vs .Where()
$coll= 1..1000000
measure-command { $coll | Where-Object {$_ -ge 5000 }  }
measure-command { $coll.Where( {$_ -ge 5000 }) }

# ToJsonString
$user= get-Mguser -userid a9978aa6-528c-4eca-9235-71c587d9fdc1
$user.ToJsonString()

###############################################################################
# Tools
###############################################################################

# Profiler
Install-Module -Name PSProfiler
# Just time a script block
Invoke-Script { .\Clean-DemoUsers.ps1 }
# Trace a script block and profile its execution
$Trace= Trace-Script { .\Create-DemoUsers1.ps1 }

$Trace.Top50SelfDuration



###############################################################################
# Regular, Parallel, Batching, Batching & Parallelism comparison
###############################################################################

$trace1=@{}; $trace2=@{}
1..3 | % {
    $trace1[$_]= Trace-Script -ScriptBlock { & .\Create-DemoUsers1.ps1 }
    .\Clean-DemoUsers.ps1
    $trace2[$_]= Trace-Script -ScriptBlock { & .\Create-DemoUsers2.ps1 }
    .\Clean-DemoUsers.ps1
    $trace3[$_]= Trace-Script -ScriptBlock { & .\Create-DemoUsers3.ps1 }
    .\Clean-DemoUsers.ps1
    $trace4[$_]= Trace-Script -ScriptBlock { & .\Create-DemoUsers4.ps1 }
    .\Clean-DemoUsers.ps1
}

Write-Host ('Script1 avg duration: {0}s' -f ($Trace1 | measure -Property { $_.TotalDuration.TotalSeconds} -Average))
Write-Host ('Script2 avg duration: {0}s' -f ($Trace2 | measure -Property { $_.TotalDuration.TotalSeconds} -Average))
Write-Host ('Script3 avg duration: {0}s' -f ($Trace3 | measure -Property { $_.TotalDuration.TotalSeconds} -Average))
Write-Host ('Script4 avg duration: {0}s' -f ($Trace4 | measure -Property { $_.TotalDuration.TotalSeconds} -Average))















