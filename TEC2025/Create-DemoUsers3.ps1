# Create-DemoUsers3.ps1 (Batching)

$RequestList= [System.Collections.Generic.List[PSCustomObject]]::new()
$BatchSize= 20
Import-Csv .\Users.csv | ForEach-Object {
    $Request= @{
        Id= $RequestList.Count + 1
        Method= 'POST'
        Url= '/users'
        Headers= @{ 'Content-Type' = 'application/json' }
        Body= @{
            userType = 'Member'
            companyName= $_.Company
            displayName = $_.DisplayName
            mail = $_.Mail
            mailNickname= $_.MailNickName
            userPrincipalName = $_.UserPrincipalName
            passwordProfile= @{
                password = ([char[]]([char]33..[char]95) + ([char[]]([char]97..[char]126)) + 0..9 | sort {Get-Random})[0..20] -join ''
                forceChangePasswordNextSignIn = $True
            }
            accountEnabled = $false
        }
    }
    $RequestList.Add( $Request )

    If( $RequestList.Count -ge $BatchSize) {
        # Batch payload full, send it and reinitialize request stack
        $Body= @{ 'requests'= $RequestList } | ConvertTo-Json -Depth 4
        $Response= Invoke-MgGraphRequest -Method POST 'https://graph.microsoft.com/v1.0/$batch' -Body $Body
        $RequestList.Clear()
    }
}
# Send batch with remaining requests
If( $RequestList.Count -gt 0) {
    $Body= @{ 'requests'= $RequestList } | ConvertTo-Json -Depth 4
    $Response= Invoke-MgGraphRequest -Method POST 'https://graph.microsoft.com/v1.0/$batch' -Body $Body
    $RequestList.Clear()
}

