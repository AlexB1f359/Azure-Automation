Connect-MgGraph -Identity

$TenantName = Get-AutomationVariable -Name trigram
$SharedMailbox = Get-AutomationVariable -Name mailbox
$to = Get-AutomationVariable -Name to

Function Send-AutomatedEmail {
    param(
        [Parameter (Mandatory = $false)]
        [string]$From,
        [Parameter (Mandatory = $true)]
        [string]$Subject,
        [Parameter (Mandatory = $true)]
        $To,
        [Parameter (Mandatory = $true)]
        [string]$Body
    )
    if ([string]::IsNullOrEmpty( $From )) {
        $From = $SharedMailbox
    }
    $ParamTable =   @{
        Subject =   $Subject
        From    =   $SharedMailbox
        To      =   $To
        Type    =   "html"
        Body    =   $body
    }
    $ToRecipients = [System.Collections.Generic.List[Hashtable]]::new()
    $ParamTable.To | ForEach-Object {
        [void]$ToRecipients.Add(@{
                emailAddress = @{
                    address = $_
                }
            })
    }
    $params =   @{
        Message =   @{
            Subject =   $ParamTable.Subject
            Body    =   @{
                ContentType =   $ParamTable.Type
                Content     =   $ParamTable.Body
            }
            ToRecipients    =   $ToRecipients
        }
        SaveToSentItems =   "false"
    }
    try {
        Send-MgUserMail -UserId $ParamTable.From -BodyParameter $params -ErrorAction Stop
        Write-Output "Email sent to:"
        $ParamTable.To
    }
    catch {
        Write-Error $Error[0]
    }
}

### Push Cert
$30days = ((get-date).AddDays(30)).ToString("yyyy-MM-dd")
$pushuri = "https://graph.microsoft.com/beta/deviceManagement/applePushNotificationCertificate"
$pushcert = Invoke-MgGraphRequest -Uri $pushuri -Method Get -OutputType PSObject
$pushexpiryplaintext = $pushcert.expirationDateTime
$pushexpiry = ($pushcert.expirationDateTime).ToString("yyyy-MM-dd")
if ($pushexpiry -lt $30days) {
    write-host "Cert Expiring" -ForegroundColor Red

    $PUSHSendMailSplat = @{
        Subject        = "$TenantName - MDM Push Cert Report - $Date"
        Body           = "Your Apple Push Certificate is due to expire on <br>
                        $pushexpiryplaintext <br>
                        Please Renew before this date
                        "
    }

    Send-AutomatedEmail @PUSHSendMailSplat
}
else {
write-host "All fine" -ForegroundColor Green
}

#VPP
$30days = ((get-date).AddDays(30)).ToString("yyyy-MM-dd")
$vppuri = "https://graph.microsoft.com/beta/deviceAppManagement/vppTokens"
$vppcert = Invoke-MgGraphRequest -Uri $vppuri -Method Get -OutputType PSObject
$vppexpiryvalue = $vppcert.value
$vppexpiryplaintext = $vppexpiryvalue.expirationDateTime
$vppexpiry = ($vppexpiryvalue.expirationDateTime).ToString("yyyy-MM-dd")
if ($vppexpiry -lt $30days) {
    write-host "Cert Expiring" -ForegroundColor Red
    #Send Mail
    $VPPSendMailSplat = @{
        Subject        = "$TenantName - VPP Token Report - $Date"
        Body           = "Your Apple VPP Token is due to expire on <br>
                        $vppexpiryplaintext <br>
                        Please Renew before this date
                        "
    }

    Send-AutomatedEmail @VPPSendMailSplat
}
else {
write-host "All fine" -ForegroundColor Green
}

#DEP
$30days = ((get-date).AddDays(30)).ToString("yyyy-MM-dd")
$depuri = "https://graph.microsoft.com/beta/deviceManagement/depOnboardingSettings"
$depcert = Invoke-MgGraphRequest -Uri $depuri -Method Get -OutputType PSObject
$depexpiryvalue = $depcert.value
$depexpiryplaintext = $depexpiryvalue.tokenexpirationDateTime

$depexpiry = ($depexpiryvalue.tokenExpirationDateTime).ToString("yyyy-MM-dd")
if ($depexpiry -lt $30days) {
    write-host "Cert Expiring" -ForegroundColor Red
    #Send Mail
    $DEPSendMailSplat = @{
        Subject        = "$TenantName - DEP Expiry Report - $Date"
        Body           = "Your Apple DEP Token is due to expire on <br>
                        $depexpiryplaintext <br>
                        Please Renew before this date
                        "
    }
    Send-AutomatedEmail @DEPSendMailSplat
}
else {
write-host "All fine" -ForegroundColor Green
}
