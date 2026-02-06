Connect-ExchangeOnline

$MailboxSecurityGroup = "[INPUT_HERE]"  # The Mail-Enabled Security Group with your shared mailbox in
$ManagedIdentityAppId = "[INPUT_HERE]" # The App/Client ID of your Managed Identity
$PolicyDescription = "[INPUT_HERE]"

New-ApplicationAccessPolicy -AppId $ManagedIdentityAppId -PolicyScopeGroupId $MailboxSecurityGroup -AccessRight RestrictAccess -Description $PolicyDescription

$TestShared = Test-ApplicationAccessPolicy -AppId $ManagedIdentityAppId -Identity $MailboxSecurityGroup
$TestUser = Test-ApplicationAccessPolicy -AppId $ManagedIdentityAppId -Identity "[INPUT_HERE]"

Write-Host "Access to Shared Mailbox: $($TestShared.AccessCheckResult)" -ForegroundColor Green
Write-Host "Access to User Mailbox:   $($TestUser.AccessCheckResult)" -ForegroundColor Yellow
