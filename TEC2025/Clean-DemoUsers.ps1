# Remove-DemoUsers.ps1

Get-MgUser -Filter "startswith(userPrincipalName,'demo-')" -All | ForEach -Parallel { Remove-MgUser -UserId $_.id } -ThrottleLimit 10
