Connect-ExchangeOnline

$MailboxEmail = "[INPUT_HERE]"  # The Shared Mailbox
$ManagedIdentityAppId = "[INPUT_HERE]" # The App/Client ID of your Managed Identity
$PolicyDescription = "[INPUT_HERE]"

New-ApplicationAccessPolicy -AppId $ManagedIdentityAppId -PolicyScopeGroupId $MailboxEmail -AccessRight RestrictAccess -Description $PolicyDescription

$TestShared = Test-ApplicationAccessPolicy -AppId $ManagedIdentityAppId -Identity $MailboxEmail
$TestUser = Test-ApplicationAccessPolicy -AppId $ManagedIdentityAppId -Identity "[INPUT_HERE]"

Write-Host "Access to Shared Mailbox: $($TestShared.AccessCheckResult)" -ForegroundColor Green
Write-Host "Access to User Mailbox:   $($TestUser.AccessCheckResult)" -ForegroundColor Yellow
