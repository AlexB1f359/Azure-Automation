try {
    Connect-MgGraph -Identity -ErrorAction Stop
}
catch {
    throw "Critical Error: Failed to connect to Microsoft Graph via Managed Identity. $_"
}

$TenantName = Get-AutomationVariable -Name trigram
$SharedMailbox = Get-AutomationVariable -Name mailbox
$recipients = Get-AutomationVariable -Name to
$homeTenantID = Get-AutomationVariable -Name tenantid
$recipients = ($recipients -split ",").Trim()
$ExpiryWindowDays = 90

Function Set-CellColor {   
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory,Position=0)]
        [string]$Property,
        [Parameter(Mandatory,Position=1)]
        [string]$Color,
        [Parameter(Mandatory,ValueFromPipeline)]
        [Object[]]$InputObject,
        [Parameter(Mandatory)]
        [string]$Filter,
        [switch]$Row
    )
    Begin {
        Write-Verbose "Function Set-CellColor begins"
        $Index = $null
        If ($Filter) {   
            $FilterString = $Filter.ToUpper().Replace($Property.ToUpper(),"`$_")
            Try {
                [scriptblock]$FilterScript = [scriptblock]::Create($FilterString)
            }
            Catch {
                Write-Warning "$(Get-Date): ""$FilterString"" caused an error, stopping script!"
                Write-Warning $Error[0]
                Exit
            }
        }
        Else {
            Write-Warning "No Filter was provided, which is required."
            Exit
        }
    }
    Process {
        ForEach ($Line in $InputObject) {   
            If (($null -eq $Index) -and ($Line -match "<th")) {   
                Write-Verbose "Processing headers..."
                $Search = $Line | Select-String -Pattern '<th.*?>(.*?)<\/th>' -AllMatches
                $Index = 0 

                ForEach ($Match in $Search.Matches)
                {   If ($Match.Groups[1].Value -eq $Property)
                    {   Break 
                    }
                    $Index ++
                }

                If ($Index -ge $Search.Matches.Count) 
                {   
                    Write-Warning "$(Get-Date): Unable to locate property: $Property in table header."
                    $Index = $null 
                }
                else {
                    Write-Verbose "$Property column found at index: $Index"
                }
            }
            ElseIf (($Line -match "<td") -and ($null -ne $Index)) {   
                $Search = $Line | Select-String -Pattern '<td.*?>(.*?)<\/td>' -AllMatches
                if ($Search.Matches.Count -gt $Index) {
                    $StringValue = $Search.Matches[$Index].Groups[1].Value
                    $Value = $StringValue -as [double]
                    if ($null -eq $value) {
                        $Value = $StringValue
                    }
                    If ($Value | Where-Object -FilterScript $FilterScript) {
                        If ($Row) {  
                            Write-Verbose "Criteria met!  Changing row to $Color..."
                            If ($Line -match "<tr style=""background-color:(.+?)"">") {
                                $Line = $Line -replace "<tr style=""background-color:$($Matches[1])","<tr style=""background-color:$Color"
                            }
                            Else {
                            $Line = $Line.Replace("<tr>","<tr style=""background-color:$Color"">")
                            }
                        }
                        Else {   
                            Write-Verbose "Criteria met!  Changing cell to $Color..."
                            $Line = $Line.Replace($Search.Matches[$Index].Value,"<td style=""background-color:$Color"">$StringValue</td>")
                        }
                    }
                }
                else {
                    Write-Warning "$(Get-Date): A data row was found with fewer columns than the header. Skipping cell coloring for this row."
                }
            }
            Write-Output $Line
        }
    }
    End {
        Write-Verbose "Function Set-CellColor completed"
    }
}
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

Function Get-EntraCredentialExpiry {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [int]$ExpiryThresholdInDays
    )
    Write-Verbose "Starting credential scan..."
    $rawFindings = [System.Collections.Generic.List[PSObject]]::new()
    $today = (Get-Date).ToUniversalTime()
    $ProcessList = {
        Param($List, $Type)
        foreach ($item in $List) {
            if ($item.DisplayName -eq "P2P Server") { continue }
            $allCreds = @()
            if ($item.PasswordCredentials) { $allCreds += $item.PasswordCredentials }
            if ($item.KeyCredentials) { $allCreds += $item.KeyCredentials }
            $activeCreds = $allCreds | Where-Object { $_.EndDateTime -gt $today }
            $expiredCreds = $allCreds | Where-Object { $_.EndDateTime -le $today }
            if ($activeCreds) {
                foreach ($cred in $activeCreds) {
                    $endDate = $cred.EndDateTime
                    $daysToExpire = (New-TimeSpan -Start $today -End $endDate).Days
                    if ($daysToExpire -le $ExpiryThresholdInDays) {
                        $cType = if ($cred.Type) { "Certificate" } else { "Secret" }
                        [void]$rawFindings.Add([PSCustomObject]@{
                            DisplayName    = $item.DisplayName
                            AppID          = $item.AppId
                            ObjectType     = $Type
                            CredentialType = $cType
                            EndDate        = $endDate.ToLocalTime().ToString("yyyy-MM-dd")
                            DaysToExpire   = $daysToExpire
                            Status         = "Expiring Soon"
                        })
                    }
                }
            }
            elseif ($expiredCreds) {
                $latestExpired = $expiredCreds | Sort-Object EndDateTime -Descending | Select-Object -First 1
                $endDate = $latestExpired.EndDateTime
                $daysToExpire = (New-TimeSpan -Start $today -End $endDate).Days
                $cType = if ($latestExpired.Type) { "Certificate" } else { "Secret" }
                [void]$rawFindings.Add([PSCustomObject]@{
                    DisplayName    = $item.DisplayName
                    AppID          = $item.AppId
                    ObjectType     = $Type
                    CredentialType = $cType
                    EndDate        = $endDate.ToLocalTime().ToString("yyyy-MM-dd")
                    DaysToExpire   = $daysToExpire
                    Status         = "Fully Inactive"
                })
            }
        }
    }
    Write-Verbose "Checking App Registrations..."
    try {
        $applications = Get-MgApplication -All -Select "DisplayName,AppId,PasswordCredentials,KeyCredentials" -ErrorAction Stop
        & $ProcessList -List $applications -Type "App Registration"
    } catch { Write-Error "Failed to get MgApplication: $_" }
    Write-Verbose "Checking Enterprise Applications..."
    try {$servicePrincipals = Get-MgServicePrincipal -All -Select "DisplayName,AppId,PasswordCredentials,KeyCredentials,ServicePrincipalType,AppOwnerOrganizationId" -ErrorAction Stop | Where-Object {($_.ServicePrincipalType -ne 'ManagedIdentity') -and ($_.AppOwnerOrganizationId -eq $homeTenantID)}
        
        & $ProcessList -List $servicePrincipals -Type "Enterprise Application"
    } catch { Write-Error "Failed to get MgServicePrincipal: $_" }
    Write-Verbose "Deduplicating results by AppID..."
    $finalReport = [System.Collections.Generic.List[PSObject]]::new()
    $grouped = $rawFindings | Group-Object AppID
    foreach ($group in $grouped) {
        $hasActive = $group.Group | Where-Object { $_.Status -eq "Expiring Soon" }
        if ($hasActive) {
            foreach ($row in $hasActive) { [void]$finalReport.Add($row) }
        }
        else {
            $winner = $group.Group | Sort-Object EndDate -Descending | Select-Object -First 1
            [void]$finalReport.Add($winner)
        }
    }
    Write-Verbose "Scan complete. Found $($finalReport.Count) unique items."
    return $finalReport
}
Write-Host "- Capturing Entra ID Expiration Information." -ForegroundColor Yellow
Write-Host " "
$Date = Get-Date -Format dd-MM-yyyy
$expiringCredentials = Get-EntraCredentialExpiry -ExpiryThresholdInDays $ExpiryWindowDays
$Title = @"
<title>Azure Entra ID Credential Report</title>
"@
$Header = @"
<style>
BODY {font-family:verdana;}
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; padding: 5px; background-color: #d1c3cd;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black; padding: 5px}
</style>
"@
$ReportHeader = @"
<h1>Azure Entra ID Credential Report</h1>
<p>The following report was run on $Date for tenant '$TenantName'.</p>
<p>Scope:</p>
<ul>
<li><b>Expiring Soon:</b> Active credentials expiring within <b>$ExpiryWindowDays days</b>.</li>
<li><b>Fully Inactive:</b> Applications with <b>zero</b> active credentials (abandoned).</li>
</ul>
"@
if ($expiringCredentials.Count) {
    $ReportRequired = $true
    $AZSPInfoHTML = $expiringCredentials | 
        Sort-Object DaysToExpire | 
        ConvertTo-Html -Fragment -Property DisplayName, AppID, ObjectType, CredentialType, EndDate, DaysToExpire, Status |
        Set-CellColor DaysToExpire '#820101' -Filter "DaysToExpire -lt 0" | 
        Set-CellColor DaysToExpire '#FF7F7F' -Filter "DaysToExpire -ge 0 -and DaysToExpire -lt 30" |
        Set-CellColor DaysToExpire '#e8f800' -Filter "DaysToExpire -ge 30 -and DaysToExpire -le 90"
    $AZSPHTML = $ReportHeader + $AZSPInfoHTML
} 
else {
    $ReportRequired = $false
    $AZSPHTML = $ReportHeader + "<p><b>No expiring or abandoned credentials found.</b></p>"
}
$FinalHTML = $Title + $Header + $AZSPHTML + "</body></html>"
if ($ReportRequired -eq $true) {
    $SendMailSplat = @{
        Subject        = "$TenantName - ACTION REQUIRED: Azure Credential Report - $Date"
        Body           = $FinalHTML
        To             = $recipients
    }
} 
else {
    $SendMailSplat = @{
        Subject        = "$TenantName - Azure Credential Report - $Date"
        Body           = $FinalHTML
        To             = $recipients
    }
}
Send-AutomatedEmail @SendMailSplat
Write-Host "$TenantName Credential Report Emailed." -ForegroundColor Yellow