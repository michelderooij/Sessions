# Create-DemoUsers2.ps1 (Parallelism)

Import-Csv .\Users.csv | ForEach-Object -ThrottleLimit 10 -Parallel {
    $NewProps = @{
        UserType = 'Member'
        CompanyName= $_.Company
        DisplayName = $_.DisplayName
        Mail = $_.Mail
        MailNickname= $_.MailNickName
        UserPrincipalName = $_.UserPrincipalName
        PasswordProfile= @{
            Password = ([char[]]([char]33..[char]95) + ([char[]]([char]97..[char]126)) + 0..9 | sort {Get-Random})[0..20] -join ''
            ForceChangePasswordNextSignIn = $True
        }
        AccountEnabled = $false
    }
    New-MgUser @NewProps
}

