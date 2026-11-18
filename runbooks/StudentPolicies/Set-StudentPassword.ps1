<#
.SYNOPSIS
  Azure Automation Runbook (PowerShell 7.6) für das SuS-Benutzermanagement.
.DESCRIPTION
  Das Runbook lädt alle Entra-Benutzer und grenzt den Verarbeitungsumfang mit
  konfigurierbaren Include-Regeln, einer Ignore-Liste und optionalen Test-UPNs ein.

  Für Benutzer mit einer erkannten Schuleinheit werden bei Bedarf Konten aktiviert,
  Passwort-Policies sichergestellt, bei Steig-Benutzern die ageGroup korrigiert und
  für neue Benutzer mit leerem Verarbeitungsattribut ein zufälliges Passwort gesetzt.
  Neue Zugangsdaten und reaktivierte Konten werden pro Schuleinheit in einer sortierten Sammelmail versendet. Bei reaktivierten Konten steht in der Passwortspalte "bestehend gem. Pupil".
  Erst nach erfolgreichem Mailversand wird der Verarbeitungsstempel gesetzt.

  Benutzer, deren konfiguriertes Zuordnungsattribut "extern" enthält, werden deaktiviert.
  Optional können auch Benutzer mit leerem Zuordnungsattribut deaktiviert werden.

  DryRun verhindert alle schreibenden Aktionen und fachlichen Passwortmails.
  Debug verhindert keine Änderungen. Debug steuert eine technische Änderungsmail und
  setzt bei fachlichen Mails optional das technische Supportpostfach in CC.

  Passwörter erscheinen ausschließlich in den fachlichen Passwortmails. Sie werden
  weder in Logs noch in technischen Supportmails ausgegeben.
  Regulär nicht verarbeitete Benutzer werden nach Ausschlussgrund zusammengefasst.
  Lokal synchronisierte Benutzer und unbekannte Schuleinheiten werden gesammelt als
  je eine Warnung mit separaten Detailzeilen im technischen Log ausgegeben.
.VORAUSSETZUNGEN
  Azure Automation Runtime: PowerShell 7.6
  Vorab importiertes Modul:
    - Microsoft.Graph.Authentication

  Microsoft Graph Application Permissions der Managed Identity:
    - User.ReadWrite.All
    - Mail.Send
    - Sites.Selected (nur bei aktiviertem SharePoint-Logging)

  Das Runbook ist als UTF-8 zu speichern. PowerShell 7.6 verarbeitet UTF-8 nativ.
#>
#requires -Version 7.6
#requires -Modules Microsoft.Graph.Authentication

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
Set-StrictMode -Version 3.0

# -----------------------------------------------------------------------------
# KONFIGURATION
# -----------------------------------------------------------------------------

# OR-Verknuepfung: Mindestens eine Regel muss passen.
# Eine leere Liste inkludiert absichtlich niemanden.
# Property unterstützt direkte und verschachtelte Pfade.
$IncludeUserPropertyRules = @(
    [ordered]@{ Property = 'userPrincipalName'; Pattern = '@sus.srb\.sg$' }
    # [ordered]@{ Property = 'department'; Pattern = '^Schule$' }
)

# Ist die Liste nicht leer, werden schreibende Aktionen und fachliche Mails nur für
# diese UPNs ausgeführt. Eine leere Liste erlaubt alle inkludierten Benutzer.
$TestUserPrincipalNames = @()

# Diese Benutzer sind vollständig geschuetzt.
$IgnoredUserPrincipalNames = [ordered]@{
    # 'servicekonto@schule-rorschacherberg.ch' = 'Technisches Konto'
}

# Dieses Attribut enthält Schuleinheit, Klasse und optional den Begriff "extern".
# Verschachtelte Pfade werden unterstützt.
$SchoolAssignmentProperty = 'Department'

# Benutzer mit leerem Zuordnungsattribut optional deaktivieren.
$DisableUsersWithEmptySchoolAssignment = $true

# Erkennung der Schuleinheiten im Zuordnungsattribut.
$SchoolUnitPatterns = [ordered]@{
    Klosterguet = '(?i)klosterguet'
    Steig       = '(?i)steig'
    Wildenstein = '(?i)wildenstein'
}
$ExternalUserPattern = '(?i)extern'

# Absenderkonto für alle durch das Runbook versendeten Nachrichten.
$MailSenderUserPrincipalName = 'informatik@schule-rorschacherberg.ch'

# Fachliche Empfänger je Schuleinheit. Pro Schuleinheit wird höchstens eine
# Sammelmail versendet. Sie enthält neue Konten mit generiertem Passwort und
# reaktivierte Konten mit "bestehend gem. Pupil" in der Passwortspalte.
$SchoolSupportRecipients = [ordered]@{
    Klosterguet = 'Maria.Gabathuler@schule-rorschacherberg.ch'
    Steig       = 'daniel.zoller@schule-rorschacherberg.ch'
    Wildenstein = 'Renato.Locher@schule-rorschacherberg.ch'
}

# Fachlicher Empfänger für Meldungen zu deaktivierten Konten.
$DisabledAccountNotificationRecipient = 'daniel.zoller@schule-rorschacherberg.ch'

# Empfänger technischer Statusmeldungen und Laufprotokolle.
$TechnicalSupportRecipient = 'daniel.zoller@schule-rorschacherberg.ch'

# Fachliche Sammelmail je Schuleinheit für neue und reaktivierte Konten.
$SendSchoolAccountMails = $true

# Separate fachliche Sammelmail zu deaktivierten Konten.
$SendDisabledAccountMails = $true

# Bereits deaktivierte Konten nicht bei jedem Lauf erneut melden.
$IncludeAlreadyDisabledUsersInDisabledAccountMail = $false

# Technische Statusmail mit Änderungen, Warnungen, Fehlern und Log-Anhang.
$EnableTechnicalSupportMail = $true
$SendTechnicalMailOnWarnings = $true

# Debug verhindert keine Änderungen. Bei Änderungen wird eine technische
# Statusmail versendet. Fachliche Mails bleiben an die normalen Empfänger
# adressiert; optional erhält der technische Support zusätzlich CC.
$Debug = $true
$CcTechnicalSupportInDebug = $true

# DryRun verhindert alle schreibenden Aktionen und fachlichen Mails.
$DryRun = $false

# Speicherung der Nachrichten im Ordner "Gesendete Elemente" des Absenders.
$SaveSchoolAccountMailsToSentItems = $true
$SaveTechnicalMailsToSentItems = $true

# Passwort und Benutzer-Sollwerte.
$PasswordPattern = 'ULLDDDDLLU'
$ForceChangePasswordNextSignIn = $false
$RequiredPasswordPolicy = 'DisablePasswordExpiration'
$ProcessedAttribute = 'extensionAttribute15'
$SteigAgeGroup = 'notAdult'

# SharePoint-Logging. Drive-ID leer lassen, um den Upload zu deaktivieren.
$SharePointDriveId = 'b!fx3JtQSXbkK3O1uu9sq1-c4Dqce5o5dPiIzsCi2qbYr3pomhXM6NTbYNHSepNYt6'
$SharePointLogFolder = 'Runbook-Logs/SuS-Accounts'
$NoActionLogSubfolder = 'NoAction'

# Optional für eine benutzerseitig zugewiesene Managed Identity.
$ManagedIdentityClientId = '1b646a0d-5a88-4588-9b29-847a25a7e103'

# -----------------------------------------------------------------------------
# LAUFSTATUS
# -----------------------------------------------------------------------------

$RunId = [guid]::NewGuid().ToString('N').Substring(0, 8)
$StartedAt = Get-Date
$LogLines = [System.Collections.Generic.List[string]]::new()
$ChangeLines = [System.Collections.Generic.List[string]]::new()
$TechnicalMailChangesByUser = @{}
$PasswordDeliveryFailures = [System.Collections.Generic.List[object]]::new()
$LocallySyncedExcludedUsers = [System.Collections.Generic.List[object]]::new()
$UnknownSchoolUnitUsers = [System.Collections.Generic.List[object]]::new()
$WarningCount = 0
$ErrorCount = 0
$FatalError = $false
$GraphConnected = $false
$ResolvedLogPath = ''

$Counters = [ordered]@{
    EntraUsersLoaded             = 0
    IncludedUsers                = 0
    IgnoredUsers                 = 0
    IncludeExcluded             = 0
    GuestsExcluded              = 0
    TestScopeExcluded           = 0
    LocallySyncedExcluded       = 0
    EmptyAssignment             = 0
    UnknownSchoolUnit           = 0
    UnchangedUsers               = 0
    AccountsActivated           = 0
    ExternalAlreadyDisabled     = 0
    ExternalDisabled            = 0
    EmptyAssignmentDisabled     = 0
    AgeGroupsCorrected          = 0
    PasswordPoliciesCorrected   = 0
    PasswordsSet                = 0
    SchoolAccountMailsSent           = 0
    DisabledAccountMailsSent     = 0
    ProcessedStampsSet          = 0
    ProcessedStampErrors        = 0
    PasswordDeliveryFailures    = 0
    ChangesPlanned              = 0
}

# -----------------------------------------------------------------------------
# ZEIT, LOGGING UND ÄNDERUNGEN
# -----------------------------------------------------------------------------

function Get-LocalTime {
    try { $zone = [System.TimeZoneInfo]::FindSystemTimeZoneById('W. Europe Standard Time') }
    catch { $zone = [System.TimeZoneInfo]::FindSystemTimeZoneById('Europe/Zurich') }
    [System.TimeZoneInfo]::ConvertTime([System.DateTimeOffset]::UtcNow, $zone)
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )
    $line = '[{0}] [{1}] [{2}] {3}' -f (Get-LocalTime).ToString('yyyy-MM-dd HH:mm:ss zzz'), $RunId, $Level, $Message
    $LogLines.Add($line)
    switch ($Level) {
        'WARN'  { $script:WarningCount++; Write-Warning $line }
        'ERROR' { $script:ErrorCount++; Write-Error $line -ErrorAction Continue }
        default { Write-Output $line }
    }
}

function Write-Section {
    param([Parameter(Mandatory)][string]$Title)
    foreach ($line in @('', ('=' * 96), "  $Title", ('=' * 96))) {
        $LogLines.Add($line)
        Write-Output $line
    }
}

function Write-UserSection {
    param(
        [Parameter(Mandatory)]$User,
        [Parameter(Mandatory)][string]$Reason
    )
    $heading = '  BENUTZER: {0} <{1}> | {2}' -f $User.displayName, $User.userPrincipalName, $Reason
    foreach ($line in @('', ('-' * 96), $heading, ('-' * 96))) {
        $LogLines.Add($line)
        Write-Output $line
    }
}

function Add-Change {
    param(
        [Parameter(Mandatory)][string]$Type,
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][string]$Description
    )
    $line = '[{0}] {1}: {2}' -f $Type, $Identity, $Description
    $ChangeLines.Add($line)
    if (-not $script:TechnicalMailChangesByUser.ContainsKey($Identity)) {
        $script:TechnicalMailChangesByUser[$Identity] = [System.Collections.Generic.List[object]]::new()
    }
    $script:TechnicalMailChangesByUser[$Identity].Add([pscustomobject]@{ Type = $Type; Description = $Description })
    $script:Counters.ChangesPlanned = $ChangeLines.Count
    Write-Log "ÄNDERUNG: $line$(if ($DryRun) { ' [DRY RUN]' })"
}

# Fasst regulär nicht verarbeitete Benutzer zusammen. Erwartbare Gruppen werden
# bewusst nur als Anzahl ausgegeben. Relevante Auffälligkeiten erzeugen pro Gruppe
# genau eine Warnung; Benutzerdetails folgen als INFO-Zeilen.
function Write-UserProcessingSummary {
    Write-Log 'Zusammenfassung der nicht verarbeiteten Benutzer:'
    Write-Log ('  Ignore-Liste                     : {0}' -f $Counters.IgnoredUsers)
    Write-Log ('  Gastbenutzer                    : {0}' -f $Counters.GuestsExcluded)
    Write-Log ('  Ausserhalb Include-Regeln       : {0}' -f $Counters.IncludeExcluded)
    Write-Log ('  Ausserhalb Testbegrenzung       : {0}' -f $Counters.TestScopeExcluded)
    Write-Log ('  Lokal synchronisiert            : {0}' -f $Counters.LocallySyncedExcluded)
    Write-Log ('  Leere Zuordnung ohne Aktion     : {0}' -f $Counters.EmptyAssignment)
    Write-Log ('  Unbekannte Schuleinheit         : {0}' -f $Counters.UnknownSchoolUnit)
    Write-Log ('  Bereits deaktiviert (extern)    : {0}' -f $Counters.ExternalAlreadyDisabled)
    Write-Log ('  Keine Änderung erforderlich     : {0}' -f $Counters.UnchangedUsers)

    if ($LocallySyncedExcludedUsers.Count -gt 0) {
        Write-Log (('{0} lokal synchronisierte Benutzer wurden aus Sicherheitsgründen nicht verarbeitet. ' +
            'Details folgen im Log.') -f $LocallySyncedExcludedUsers.Count) -Level WARN
        foreach ($item in @($LocallySyncedExcludedUsers | Sort-Object UserPrincipalName)) {
            Write-Log ('  - {0}' -f $item.UserPrincipalName)
        }
    }
    if ($UnknownSchoolUnitUsers.Count -gt 0) {
        Write-Log (('Bei {0} Benutzern wurde keine bekannte Schuleinheit erkannt. ' +
            'Details folgen im Log.') -f $UnknownSchoolUnitUsers.Count) -Level WARN
        foreach ($item in @($UnknownSchoolUnitUsers | Sort-Object UserPrincipalName)) {
            $assignment = if ([string]::IsNullOrWhiteSpace([string]$item.AssignmentValue)) { '<leer>' } else { [string]$item.AssignmentValue }
            Write-Log ('  - {0} | {1}: {2}' -f $item.UserPrincipalName, $SchoolAssignmentProperty, $assignment)
        }
    }
}

# -----------------------------------------------------------------------------
# PROPERTY-, SCOPE- UND KONFIGURATIONSHILFEN
# -----------------------------------------------------------------------------

function Get-NestedPropertyValue {
    param($InputObject, [Parameter(Mandatory)][string]$Path)
    $current = $InputObject
    foreach ($segment in ($Path -split '\.')) {
        if ($null -eq $current) { return $null }
        if ($current -is [System.Collections.IDictionary]) {
            if (-not $current.Contains($segment)) { return $null }
            $current = $current[$segment]
        }
        else {
            $property = $current.PSObject.Properties[$segment]
            if ($null -eq $property) { return $null }
            $current = $property.Value
        }
    }
    return $current
}

function Test-IncludeConfiguration {
    if (@($IncludeUserPropertyRules).Count -eq 0) {
        throw 'IncludeUserPropertyRules ist leer. Aus Sicherheitsgründen wird niemand verarbeitet.'
    }
    foreach ($rule in @($IncludeUserPropertyRules)) {
        if ($null -eq $rule) { throw 'Include-Regel darf nicht null sein.' }
        $property = [string]$rule.Property
        $pattern = [string]$rule.Pattern
        if ([string]::IsNullOrWhiteSpace($property)) { throw 'Include-Regel ohne Property gefunden.' }
        if ([string]::IsNullOrWhiteSpace($pattern)) { throw "Include-Regel für '$property' hat kein Pattern." }
        try { [void][regex]::new($pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase) }
        catch { throw "Ungültige Include-Regex für '$property': $pattern. $($_.Exception.Message)" }
    }
}

function Test-Configuration {
    Test-IncludeConfiguration
    if ([string]::IsNullOrWhiteSpace($SchoolAssignmentProperty)) { throw 'SchoolAssignmentProperty ist leer.' }
    if ([string]::IsNullOrWhiteSpace($ProcessedAttribute)) { throw 'ProcessedAttribute ist leer.' }
    if ($ProcessedAttribute -notmatch '^extensionAttribute(?:[1-9]|1[0-5])$') {
        throw "ProcessedAttribute muss extensionAttribute1 bis extensionAttribute15 sein: '$ProcessedAttribute'."
    }
    if ([string]::IsNullOrWhiteSpace($MailSenderUserPrincipalName)) { throw 'MailSender ist leer.' }
    if ([string]::IsNullOrWhiteSpace($PasswordPattern)) { throw 'PasswordPattern ist leer.' }
    if ($PasswordPattern -notmatch '^[ULD]+$') { throw "PasswordPattern enthält ungültige Zeichen: '$PasswordPattern'." }
    try { [void][regex]::new([string]$ExternalUserPattern) }
    catch { throw "Ungültige Regex für externe Benutzer: $($_.Exception.Message)" }
    foreach ($unit in $SchoolUnitPatterns.Keys) {
        if (-not $SchoolSupportRecipients.Contains($unit) -or [string]::IsNullOrWhiteSpace([string]$SchoolSupportRecipients[$unit])) {
            throw "Keine Supportadresse für Schuleinheit '$unit' konfiguriert."
        }
        try { [void][regex]::new([string]$SchoolUnitPatterns[$unit]) }
        catch { throw "Ungültige Regex für Schuleinheit '$unit': $($_.Exception.Message)" }
    }
    if ($SendDisabledAccountMails -and [string]::IsNullOrWhiteSpace($DisabledAccountNotificationRecipient)) {
        throw 'DisabledAccountNotificationRecipient ist leer.'
    }
    if ($Debug -and $CcTechnicalSupportInDebug -and [string]::IsNullOrWhiteSpace($TechnicalSupportRecipient)) {
        throw 'Debug-CC ist aktiviert, aber LogSupportMail ist leer.'
    }
}

function Test-UserIncluded {
    param([Parameter(Mandatory)]$User)
    foreach ($rule in @($IncludeUserPropertyRules)) {
        $value = Get-NestedPropertyValue -InputObject $User -Path ([string]$rule.Property)
        foreach ($candidate in @($value)) {
            if ($null -ne $candidate -and ([string]$candidate -match [string]$rule.Pattern)) { return $true }
        }
    }
    return $false
}

function Test-TestScopeUser {
    param([string]$Upn)
    if (@($TestUserPrincipalNames).Count -eq 0) { return $true }
    return $TestUserPrincipalNames -contains $Upn
}

function Get-GraphSelectProperties {
    $base = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($name in @(
        'id', 'displayName', 'userPrincipalName', 'userType', 'accountEnabled', 'ageGroup',
        'passwordPolicies', 'onPremisesSyncEnabled', 'onPremisesExtensionAttributes'
    )) { [void]$base.Add($name) }
    foreach ($rule in @($IncludeUserPropertyRules)) {
        $root = ([string]$rule.Property -split '\.')[0]
        if ($root) { [void]$base.Add($root) }
    }
    $assignmentRoot = ($SchoolAssignmentProperty -split '\.')[0]
    if ($assignmentRoot) { [void]$base.Add($assignmentRoot) }
    return (@($base) | Sort-Object) -join ','
}

function Get-SchoolUnit {
    param([AllowNull()][string]$AssignmentValue)
    if ([string]::IsNullOrWhiteSpace($AssignmentValue)) { return $null }
    foreach ($unit in $SchoolUnitPatterns.Keys) {
        if ($AssignmentValue -match [string]$SchoolUnitPatterns[$unit]) { return [string]$unit }
    }
    return $null
}

# -----------------------------------------------------------------------------
# GRAPH
# -----------------------------------------------------------------------------

function Invoke-Graph {
    param(
        [ValidateSet('GET', 'POST', 'PATCH')][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [object]$Body
    )
    $parameters = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
    if ($PSBoundParameters.ContainsKey('Body')) {
        $parameters.Body = $Body | ConvertTo-Json -Depth 15 -Compress
        $parameters.ContentType = 'application/json; charset=utf-8'
    }
    Invoke-MgGraphRequest @parameters
}

function Get-GraphCollection {
    param([Parameter(Mandatory)][string]$Uri)
    $all = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while (-not [string]::IsNullOrWhiteSpace([string]$next)) {
        $page = Invoke-Graph -Method GET -Uri $next
        if ($null -eq $page -or -not ($page.PSObject.Properties.Name -contains 'value')) {
            throw "Graph-Antwort für '$next' enthält keine value-Eigenschaft."
        }
        foreach ($item in @($page.value)) { if ($null -ne $item) { $all.Add($item) } }
        $nextProperty = $page.PSObject.Properties['@odata.nextLink']
        $next = if ($null -ne $nextProperty) { [string]$nextProperty.Value } else { $null }
    }
    return $all
}

function Update-EntraUser {
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][hashtable]$Changes
    )
    Invoke-Graph -Method PATCH -Uri "v1.0/users/$UserId" -Body $Changes | Out-Null
}

# -----------------------------------------------------------------------------
# PASSWORT UND VERARBEITUNGSSTEMPEL
# -----------------------------------------------------------------------------

function New-SchoolPassword {
    $sets = @{
        U = 'ABCDEFGHJKMNPQRSTUVWXYZ'
        L = 'abcdefghijkmnpqrstuvwxyz'
        D = '23456789'
    }
    return -join $PasswordPattern.ToCharArray().ForEach({
        $set = $sets[[string]$_]
        if (-not $set) { throw "Ungültiges Zeichen '$_' im Passwortmuster." }
        $set[[System.Security.Cryptography.RandomNumberGenerator]::GetInt32($set.Length)]
    })
}

function Get-RequiredPasswordPolicies {
    param([AllowNull()][string]$CurrentPolicies)
    $policies = [System.Collections.Generic.List[string]]::new()
    foreach ($policy in @($CurrentPolicies -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })) {
        if (-not $policies.Contains($policy)) { $policies.Add($policy) }
    }
    if (-not $policies.Contains($RequiredPasswordPolicy)) { $policies.Add($RequiredPasswordPolicy) }
    return $policies -join ','
}
function Test-PasswordPolicyPresent {
    param([AllowNull()][string]$CurrentPolicies)
    if ([string]::IsNullOrWhiteSpace($CurrentPolicies)) { return $false }
    return @($CurrentPolicies -split ',' | ForEach-Object { $_.Trim() }) -contains $RequiredPasswordPolicy
}

function Set-ProcessedStamp {
    param([Parameter(Mandatory)]$User)
    if ([bool]$User.onPremisesSyncEnabled) {
        throw "$ProcessedAttribute ist bei lokal synchronisierten Benutzern über Graph schreibgeschützt."
    }
    $stamp = 'PROCESSED|{0}|{1}' -f (Get-LocalTime).ToString('yyyy-MM-ddTHH:mm:sszzz'), $RunId
    Update-EntraUser -UserId ([string]$User.Id) -Changes @{
        onPremisesExtensionAttributes = @{ $ProcessedAttribute = $stamp }
    }
}

# -----------------------------------------------------------------------------
# MAIL-HILFSFUNKTIONEN
# -----------------------------------------------------------------------------

function ConvertTo-HtmlEncoded {
    param([AllowNull()][string]$Text)
    [System.Net.WebUtility]::HtmlEncode([string]$Text)
}

function Get-DebugCcRecipients {
    param([AllowNull()][string]$PrimaryRecipient)
    if ($Debug -and
        $CcTechnicalSupportInDebug -and
        -not [string]::IsNullOrWhiteSpace($TechnicalSupportRecipient) -and
        -not $TechnicalSupportRecipient.Equals($PrimaryRecipient, [StringComparison]::OrdinalIgnoreCase)) {
        return @(@{ emailAddress = @{ address = $TechnicalSupportRecipient } })
    }
    return @()
}

function Send-SchoolAccountMail {
    param(
        [Parameter(Mandatory)][string]$SchoolUnit,
        [Parameter(Mandatory)][object[]]$Users
    )
    $rows = foreach ($user in @($Users | Sort-Object Class, UserPrincipalName)) {
        $displayName = ConvertTo-HtmlEncoded $user.DisplayName
        $upn = ConvertTo-HtmlEncoded $user.UserPrincipalName
        $class = ConvertTo-HtmlEncoded $user.Class
        $password = ConvertTo-HtmlEncoded $user.Password
        @"
<tr>
  <td style="border:1px solid #d7dce2;padding:9px 10px;">$displayName</td>
  <td style="border:1px solid #d7dce2;padding:9px 10px;">$class</td>
  <td style="border:1px solid #d7dce2;padding:9px 10px;font-family:Consolas,monospace;font-weight:700;">$password</td>
  <td style="border:1px solid #d7dce2;padding:9px 10px;white-space:nowrap;">$upn</td>
</tr>
"@
    }
    $html = @"
<div style="font-family:Segoe UI,Arial,sans-serif;max-width:900px;color:#222;line-height:1.45;">
  <h2 style="color:#0f6cbd;margin-bottom:6px;">SuS-Accounts: $(ConvertTo-HtmlEncoded $SchoolUnit)</h2>
  <p style="margin-top:0;">Folgende neuen oder reaktivierten Benutzerkonten wurden am $((Get-LocalTime).ToString('dd.MM.yyyy HH:mm')) verarbeitet. Bei reaktivierten Konten gilt das bereits in Pupil hinterlegte Passwort.</p>
  <table style="border-collapse:collapse;width:100%;background:#fff;">
    <thead>
      <tr>
        <th style="border:1px solid #d7dce2;padding:9px 10px;background:#f3f6f9;text-align:left;">Name</th>
        <th style="border:1px solid #d7dce2;padding:9px 10px;background:#f3f6f9;text-align:left;">Klasse</th>
        <th style="border:1px solid #d7dce2;padding:9px 10px;background:#f3f6f9;text-align:left;">Passwort</th>
        <th style="border:1px solid #d7dce2;padding:9px 10px;background:#f3f6f9;text-align:left;">Username</th>
      </tr>
    </thead>
    <tbody>$($rows -join "`n")</tbody>
  </table>
  <p style="padding:10px 12px;background:#f3f7fb;border-left:4px solid #0f6cbd;border-radius:4px;">
    Bitte das Passwort beim entsprechenden SuS-Account in Pupil notieren.
  </p>
  <p style="padding:10px 12px;background:#fff4ce;border-left:4px solid #f0b429;border-radius:4px;">
    <strong>Vertraulich:</strong> Passwörter sicher behandeln und diese Mail nach der Verteilung löschen.
  </p>
  <div style="font-size:11px;color:#666;border-top:1px solid #ddd;margin-top:22px;padding-top:10px;">
    Run-ID: $RunId · Zeitpunkt: $((Get-LocalTime).ToString('dd.MM.yyyy HH:mm:ss')) · Debug: $Debug
  </div>
</div>
"@
    $message = @{
        subject = "$SchoolUnit - neue oder reaktivierte SuS-Accounts"
        body = @{ contentType = 'HTML'; content = $html }
        toRecipients = @(@{ emailAddress = @{ address = $SchoolSupportRecipients[$SchoolUnit] } })
    }
    $cc = @(Get-DebugCcRecipients -PrimaryRecipient ([string]$SchoolSupportRecipients[$SchoolUnit]))
    if ($cc.Count -gt 0) { $message.ccRecipients = $cc }
    Invoke-Graph -Method POST -Uri "v1.0/users/$([uri]::EscapeDataString($MailSenderUserPrincipalName))/sendMail" -Body @{
        message = $message
        saveToSentItems = $SaveSchoolAccountMailsToSentItems
    } | Out-Null
}

function Send-DisabledAccountMail {
    param([Parameter(Mandatory)][object[]]$Users)
    $rows = foreach ($user in @($Users | Sort-Object AssignmentValue, UserPrincipalName)) {
        $assignment = if ([string]::IsNullOrWhiteSpace([string]$user.AssignmentValue)) {
            'keine Zuordnung'
        }
        else {
            [string]$user.AssignmentValue
        }
        @"
<tr>
  <td style="border:1px solid #d7dce2;padding:9px 10px;">$(ConvertTo-HtmlEncoded $user.DisplayName)</td>
  <td style="border:1px solid #d7dce2;padding:9px 10px;">$(ConvertTo-HtmlEncoded $assignment)</td>
  <td style="border:1px solid #d7dce2;padding:9px 10px;">$(ConvertTo-HtmlEncoded $user.UserPrincipalName)</td>
  <td style="border:1px solid #d7dce2;padding:9px 10px;">$(ConvertTo-HtmlEncoded $user.Status)</td>
</tr>
"@
    }
    $html = @"
<div style="font-family:Segoe UI,Arial,sans-serif;max-width:900px;color:#222;line-height:1.45;">
  <h2 style="color:#0f6cbd;margin-bottom:6px;">Deaktivierte oder erkannte Benutzerkonten</h2>
  <p style="margin-top:0;">Folgende Konten wurden im aktuellen Lauf als extern oder ohne Zuordnung erkannt.</p>
  <table style="border-collapse:collapse;width:100%;background:#fff;">
    <thead><tr>
      <th style="border:1px solid #d7dce2;padding:9px 10px;background:#f3f6f9;text-align:left;">Name</th>
      <th style="border:1px solid #d7dce2;padding:9px 10px;background:#f3f6f9;text-align:left;">Zuordnung</th>
      <th style="border:1px solid #d7dce2;padding:9px 10px;background:#f3f6f9;text-align:left;">Username</th>
      <th style="border:1px solid #d7dce2;padding:9px 10px;background:#f3f6f9;text-align:left;">Status</th>
    </tr></thead>
    <tbody>$($rows -join "`n")</tbody>
  </table>
  <div style="font-size:11px;color:#666;border-top:1px solid #ddd;margin-top:22px;padding-top:10px;">
    Run-ID: $RunId · Zeitpunkt: $((Get-LocalTime).ToString('dd.MM.yyyy HH:mm:ss')) · Dry Run: $DryRun · Debug: $Debug
  </div>
</div>
"@
    $message = @{
        subject = 'SRB SuS-Usermanagement: deaktivierte Konten'
        body = @{ contentType = 'HTML'; content = $html }
        toRecipients = @(@{ emailAddress = @{ address = $DisabledAccountNotificationRecipient } })
    }
    $cc = @(Get-DebugCcRecipients -PrimaryRecipient ([string]$DisabledAccountNotificationRecipient))
    if ($cc.Count -gt 0) { $message.ccRecipients = $cc }
    Invoke-Graph -Method POST -Uri "v1.0/users/$([uri]::EscapeDataString($MailSenderUserPrincipalName))/sendMail" -Body @{
        message = $message
        saveToSentItems = $SaveTechnicalMailsToSentItems
    } | Out-Null
}

function Get-TechnicalMailIssueSummaryHtml {
    $issueLines = @($LogLines | Where-Object { $_ -match '\[(ERROR|WARN)\]' })
    if ($issueLines.Count -eq 0) { return '' }
    $items = foreach ($line in $issueLines) {
        $message = [regex]::Replace([string]$line, '^\[[^\]]+\] \[[^\]]+\] \[(?:ERROR|WARN)\]\s*', '')
        '<li style="margin:5px 0;">' + (ConvertTo-HtmlEncoded $message) + '</li>'
    }
    return '<h3 style="margin-bottom:6px;">Fehler und Hinweise</h3><div style="border:1px solid #e6a700;border-radius:8px;padding:10px 14px;margin:8px 0 18px;background:#fff8e1;"><ul style="margin:0;padding-left:20px;">' + ($items -join '') + '</ul></div>'
}

function Get-TechnicalMailChangesHtml {
    if ($TechnicalMailChangesByUser.Count -eq 0) { return '<p style="color:#666;">Keine Änderungen.</p>' }
    $blocks = foreach ($identity in @($TechnicalMailChangesByUser.Keys | Sort-Object)) {
        $details = foreach ($item in @($TechnicalMailChangesByUser[$identity])) {
            '<li style="margin:3px 0;"><strong>' + (ConvertTo-HtmlEncoded $item.Type) + ':</strong> ' + (ConvertTo-HtmlEncoded $item.Description) + '</li>'
        }
        '<div style="border:1px solid #e1e5ea;border-radius:8px;padding:12px 14px;margin:8px 0;background:#fff;"><div style="font-weight:600;color:#172b4d;">' + (ConvertTo-HtmlEncoded $identity) + '</div><ul style="margin:6px 0 0;padding-left:20px;">' + ($details -join '') + '</ul></div>'
    }
    return $blocks -join ''
}

function Get-TechnicalMailExclusionSummaryHtml {
    $entries = @(
        [pscustomobject]@{ Label = 'Ignore-Liste'; Count = $Counters.IgnoredUsers }
        [pscustomobject]@{ Label = 'Gastbenutzer'; Count = $Counters.GuestsExcluded }
        [pscustomobject]@{ Label = 'Ausserhalb Include-Regeln'; Count = $Counters.IncludeExcluded }
        [pscustomobject]@{ Label = 'Ausserhalb Testbegrenzung'; Count = $Counters.TestScopeExcluded }
        [pscustomobject]@{ Label = 'Lokal synchronisiert'; Count = $Counters.LocallySyncedExcluded }
        [pscustomobject]@{ Label = 'Leere Zuordnung ohne Aktion'; Count = $Counters.EmptyAssignment }
        [pscustomobject]@{ Label = 'Unbekannte Schuleinheit'; Count = $Counters.UnknownSchoolUnit }
        [pscustomobject]@{ Label = 'Bereits deaktiviert (extern)'; Count = $Counters.ExternalAlreadyDisabled }
        [pscustomobject]@{ Label = 'Keine Änderung erforderlich'; Count = $Counters.UnchangedUsers }
    )
    $visibleEntries = @($entries | Where-Object { $_.Count -gt 0 })
    if ($visibleEntries.Count -eq 0) { return '' }
    $items = foreach ($entry in $visibleEntries) {
        '<li style="margin:3px 0;">' + (ConvertTo-HtmlEncoded ('{0}: {1}' -f $entry.Label, $entry.Count)) + '</li>'
    }
    return '<h3 style="margin-bottom:6px;">Nicht verarbeitete Benutzer</h3>' +
        '<div style="border:1px solid #e1e5ea;border-radius:8px;padding:10px 14px;margin:8px 0 18px;background:#f8f9fa;">' +
        '<ul style="margin:0;padding-left:20px;">' + ($items -join '') + '</ul></div>'
}

function Get-PasswordDeliveryFailureHtml {
    if ($PasswordDeliveryFailures.Count -eq 0) { return '' }
    $items = foreach ($failure in @($PasswordDeliveryFailures | Sort-Object SchoolUnit, UserPrincipalName)) {
        '<li style="margin:5px 0;">' + (ConvertTo-HtmlEncoded "$($failure.UserPrincipalName) ($($failure.SchoolUnit))") + '</li>'
    }
    return @"
<h3 style="color:#a4262c;margin-bottom:6px;">Kritischer Zustellfehler</h3>
<div style="border:1px solid #d13438;border-radius:8px;padding:10px 14px;margin:8px 0 18px;background:#fde7e9;">
  <p style="margin-top:0;"><strong>Für diese Benutzer wurde bereits ein neues Passwort gesetzt, die Passwortmail konnte jedoch nicht versendet werden.</strong></p>
  <p>Es wurde kein Verarbeitungsstempel gesetzt. Beim nächsten Lauf wird ein neues Passwort erzeugt. Passwörter werden in dieser technischen Meldung nicht angezeigt.</p>
  <ul style="margin-bottom:0;padding-left:20px;">$($items -join '')</ul>
</div>
"@
}

function Send-TechnicalSupportMail {
    param([Parameter(Mandatory)][string]$Status)
    $content = $LogLines -join "`r`n"
    $bytes = [Text.Encoding]::UTF8.GetBytes($content)
    $name = 'SuS_{0}_{1}.log' -f (Get-LocalTime).ToString('yyyy-MM-dd_HHmmss'), $RunId
    $modeHtml = if ($DryRun) {
        '<p style="padding:10px 12px;background:#f3f7fb;border-left:4px solid #0f6cbd;border-radius:4px;">Vorschau: Es wurden keine Änderungen ausgeführt und keine fachlichen Mails versendet. Geplante Passwortsetzungen enthalten bewusst keine Passwörter.</p>'
    } else { '' }
    $intro = if ($Status -eq 'FEHLER') { 'Der Lauf wurde mit Fehlern beendet.' }
             elseif ($Status -eq 'WARNUNG') { 'Der Lauf wurde mit Hinweisen beendet.' }
             elseif ($Status -eq 'VORSCHAU') { 'Der geplante Lauf wurde als Vorschau abgeschlossen.' }
             elseif ($ChangeLines.Count -gt 0) { 'Das SuS-Usermanagement wurde erfolgreich abgeschlossen.' }
             else { 'Es waren keine Änderungen erforderlich.' }
    $html = @"
<div style="font-family:Segoe UI,Arial,sans-serif;max-width:800px;color:#222;line-height:1.45;">
  <h2 style="color:#0f6cbd;margin-bottom:6px;">SuS-Usermanagement: $Status</h2>
  <p style="margin-top:0;">$intro</p>
  $modeHtml
  $(Get-PasswordDeliveryFailureHtml)
  $(Get-TechnicalMailIssueSummaryHtml)
  $(Get-TechnicalMailExclusionSummaryHtml)
  <h3 style="margin-bottom:6px;">Änderungen</h3>
  $(Get-TechnicalMailChangesHtml)
  <div style="font-size:11px;color:#666;border-top:1px solid #ddd;margin-top:24px;padding-top:10px;">
    <strong>Technische Informationen</strong><br>
    Run-ID: $RunId · Zeitpunkt: $((Get-LocalTime).ToString('dd.MM.yyyy HH:mm:ss')) · Dry Run: $DryRun · Debug: $Debug<br>
    Einzelaktionen: $($ChangeLines.Count) · Warnungen: $WarningCount · Fehler: $ErrorCount<br>
    Das vollständige technische Log befindet sich im Anhang. Passwörter sind weder im Log noch in dieser Mail enthalten.
  </div>
</div>
"@
    Invoke-Graph -Method POST -Uri "v1.0/users/$([uri]::EscapeDataString($MailSenderUserPrincipalName))/sendMail" -Body @{
        message = @{
            subject = "SRB SuS-Usermanagement: $Status"
            body = @{ contentType = 'HTML'; content = $html }
            toRecipients = @(@{ emailAddress = @{ address = $TechnicalSupportRecipient } })
            attachments = @(@{
                '@odata.type' = '#microsoft.graph.fileAttachment'
                name = $name
                contentType = 'text/plain; charset=utf-8'
                contentBytes = [Convert]::ToBase64String($bytes)
            })
        }
        saveToSentItems = $SaveTechnicalMailsToSentItems
    } | Out-Null
}

# -----------------------------------------------------------------------------
# SHAREPOINT-LOGGING
# -----------------------------------------------------------------------------

function Save-LogToSharePoint {
    if ([string]::IsNullOrWhiteSpace($SharePointDriveId) -or $LogLines.Count -eq 0) { return }
    $name = 'SuS_{0}_{1}.log' -f (Get-LocalTime).ToString('yyyy-MM-dd_HHmmss'), $RunId
    $folder = $SharePointLogFolder.Trim('/')
    if ($ChangeLines.Count -eq 0) {
        $folder = (@($folder, $NoActionLogSubfolder.Trim('/')) | Where-Object { $_ }) -join '/'
    }
    $path = if ($folder) { "$folder/$name" } else { $name }
    $script:ResolvedLogPath = $path
    $bytes = [Text.Encoding]::UTF8.GetBytes(($LogLines -join "`r`n"))
    Invoke-MgGraphRequest -Method PUT -Uri "v1.0/drives/$SharePointDriveId/root:/$path`:/content" -Body $bytes -ContentType 'text/plain; charset=utf-8' -ErrorAction Stop | Out-Null
}

# -----------------------------------------------------------------------------
# HAUPTABLAUF
# -----------------------------------------------------------------------------

$UsersPendingSchoolAccountMailByUnit = @{}
foreach ($unit in $SchoolUnitPatterns.Keys) {
    $UsersPendingSchoolAccountMailByUnit[$unit] = [System.Collections.Generic.List[object]]::new()
}
$UsersPendingDisabledAccountMail = [System.Collections.Generic.List[object]]::new()

try {
    Write-Section '1. START UND KONFIGURATION'
    Test-Configuration
    Write-Log "Runbook gestartet. DryRun=$DryRun; Debug=$Debug; IncludeRegeln=$(@($IncludeUserPropertyRules).Count); Testbenutzer=$(@($TestUserPrincipalNames).Count); Ignore=$($IgnoredUserPrincipalNames.Count)"
    Write-Log "Zuordnungsattribut: $SchoolAssignmentProperty; Leere Zuordnung deaktivieren: $DisableUsersWithEmptySchoolAssignment; Bereits deaktivierte Benutzer mailen: $IncludeAlreadyDisabledUsersInDisabledAccountMail"
    $module = Get-Module Microsoft.Graph.Authentication -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    Write-Log "Modul Microsoft.Graph.Authentication: $($module.Version)"

    Write-Section '2. GRAPH-VERBINDUNG UND BENUTZER LADEN'
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($ManagedIdentityClientId)) {
        Connect-MgGraph -Identity -NoWelcome | Out-Null
    } else {
        Connect-MgGraph -Identity -ClientId $ManagedIdentityClientId -NoWelcome | Out-Null
    }
    $GraphConnected = $true
    Write-Log 'Mit Microsoft Graph verbunden.' -Level OK

    $select = Get-GraphSelectProperties
    $users = @(Get-GraphCollection -Uri "v1.0/users?`$select=$select&`$top=999")
    $Counters.EntraUsersLoaded = $users.Count
    Write-Log "$($users.Count) Entra-Benutzer geladen." -Level OK

    Write-Section '3. BENUTZER VALIDIEREN UND VERARBEITEN'
    foreach ($user in $users) {
        $upn = [string]$user.userPrincipalName
        try {
            if ($IgnoredUserPrincipalNames.Contains($upn)) { $Counters.IgnoredUsers++; continue }
            if ([string]$user.userType -eq 'Guest') { $Counters.GuestsExcluded++; continue }
            if (-not (Test-UserIncluded $user)) { $Counters.IncludeExcluded++; continue }
            $Counters.IncludedUsers++

            $inTestScope = Test-TestScopeUser $upn
            if (-not $inTestScope) { $Counters.TestScopeExcluded++; continue }

            $assignmentValue = [string](Get-NestedPropertyValue -InputObject $user -Path $SchoolAssignmentProperty)
            $isEmptyAssignment = [string]::IsNullOrWhiteSpace($assignmentValue)
            $isExternal = -not $isEmptyAssignment -and $assignmentValue -match $ExternalUserPattern

            if ([bool]$user.onPremisesSyncEnabled) {
                $Counters.LocallySyncedExcluded++
                $LocallySyncedExcludedUsers.Add([pscustomobject]@{ UserPrincipalName = $upn })
                continue
            }

            if ($isExternal -or ($isEmptyAssignment -and $DisableUsersWithEmptySchoolAssignment)) {
                $reason = if ($isExternal) { 'Extern' } else { 'Leere Zuordnung' }
                if (-not [bool]$user.accountEnabled) {
                    $status = 'Bereits deaktiviert'
                    if ($isExternal) { $Counters.ExternalAlreadyDisabled++ }
                    if (-not $IncludeAlreadyDisabledUsersInDisabledAccountMail) {
                        continue
                    }
                }
                else {
                    Write-UserSection -User $user -Reason $reason.ToUpperInvariant()
                    Add-Change -Type 'Konto deaktivieren' -Identity $upn -Description $reason
                    if ($DryRun) { $status = 'VORSCHAU: würde deaktiviert' }
                    else {
                        Update-EntraUser -UserId ([string]$user.id) -Changes @{ accountEnabled = $false }
                        $status = 'Deaktiviert'
                        if ($isExternal) { $Counters.ExternalDisabled++ } else { $Counters.EmptyAssignmentDisabled++ }
                        Write-Log "Konto deaktiviert: $upn" -Level OK
                    }
                }

                $UsersPendingDisabledAccountMail.Add([pscustomobject]@{
                    DisplayName = [string]$user.displayName
                    UserPrincipalName = $upn
                    AssignmentValue = $assignmentValue
                    Reason = $reason
                    Status = $status
                })
                continue
            }

            if ($isEmptyAssignment) {
                $Counters.EmptyAssignment++
                continue
            }

            $unit = Get-SchoolUnit -AssignmentValue $assignmentValue
            if (-not $unit) {
                $Counters.UnknownSchoolUnit++
                $UnknownSchoolUnitUsers.Add([pscustomobject]@{
                    UserPrincipalName = $upn
                    AssignmentValue = $assignmentValue
                })
                continue
            }

            $processedValue = [string]$user.onPremisesExtensionAttributes.$ProcessedAttribute
            $needsPassword = [string]::IsNullOrWhiteSpace($processedValue)
            $needsActivation = -not [bool]$user.accountEnabled
            $needsAgeGroup = $unit -eq 'Steig' -and [string]$user.ageGroup -ne $SteigAgeGroup
            $needsPasswordPolicy = -not (Test-PasswordPolicyPresent -CurrentPolicies ([string]$user.passwordPolicies))

            if (-not $needsPassword -and -not $needsActivation -and -not $needsAgeGroup -and -not $needsPasswordPolicy) {
                $Counters.UnchangedUsers++
                continue
            }

            Write-UserSection -User $user -Reason 'VERARBEITUNG'

            if ($needsActivation) {
                Add-Change -Type 'Konto aktivieren' -Identity $upn -Description "Erkannte Schuleinheit: $unit"
                if (-not $DryRun) {
                    Update-EntraUser -UserId ([string]$user.id) -Changes @{ accountEnabled = $true }
                    $Counters.AccountsActivated++
                    Write-Log "Konto aktiviert: $upn" -Level OK
                }
            }

            if ($needsAgeGroup) {
                Add-Change -Type 'AgeGroup' -Identity $upn -Description "$($user.ageGroup) -> $SteigAgeGroup"
                if (-not $DryRun) {
                    Update-EntraUser -UserId ([string]$user.id) -Changes @{ ageGroup = $SteigAgeGroup }
                    $Counters.AgeGroupsCorrected++
                    Write-Log "AgeGroup aktualisiert: $upn" -Level OK
                }
            }

            if ($needsPasswordPolicy -and -not $needsPassword) {
                Add-Change -Type 'Password Policy' -Identity $upn -Description "Sicherstellen: $RequiredPasswordPolicy"
                if (-not $DryRun) {
                    Update-EntraUser -UserId ([string]$user.id) -Changes @{ passwordPolicies = (Get-RequiredPasswordPolicies -CurrentPolicies ([string]$user.passwordPolicies)) }
                    $Counters.PasswordPoliciesCorrected++
                    Write-Log "Password-Policy aktualisiert: $upn" -Level OK
                }
            }

            # Bestehende, reaktivierte Konten werden ohne Passwortänderung in
            # dieselbe fachliche Sammelmail wie neue Konten aufgenommen.
            if (-not $needsPassword) {
                if ($needsActivation -and -not $DryRun) {
                    $UsersPendingSchoolAccountMailByUnit[$unit].Add([pscustomobject]@{
                        Id = [string]$user.id
                        DisplayName = [string]$user.displayName
                        UserPrincipalName = $upn
                        Class = $assignmentValue
                        AccountAction = 'Konto reaktiviert'
                        Password = 'bestehend gem. Pupil'
                        RequiresProcessedStamp = $false
                        onPremisesSyncEnabled = [bool]$user.onPremisesSyncEnabled
                    })
                    Write-Log "Reaktivierter Benutzer für Sammelmail vorgemerkt: $upn" -Level OK
                }
                continue
            }

            Add-Change -Type 'Passwort setzen' -Identity $upn -Description "Neuer Benutzer; Schuleinheit=$unit; Klasse='$assignmentValue'; Passwort wird nicht protokolliert"
            if ($DryRun) { continue }

            $password = New-SchoolPassword
            Update-EntraUser -UserId ([string]$user.id) -Changes @{
                passwordProfile = @{
                    password = $password
                    forceChangePasswordNextSignIn = $ForceChangePasswordNextSignIn
                }
                passwordPolicies = (Get-RequiredPasswordPolicies -CurrentPolicies ([string]$user.passwordPolicies))
            }
            $Counters.PasswordsSet++
            if ($needsPasswordPolicy) { $Counters.PasswordPoliciesCorrected++ }
            $UsersPendingSchoolAccountMailByUnit[$unit].Add([pscustomobject]@{
                Id = [string]$user.id
                DisplayName = [string]$user.displayName
                UserPrincipalName = $upn
                Class = $assignmentValue
                AccountAction = if ($needsActivation) { 'Neu eingerichtet und reaktiviert' } else { 'Neu eingerichtet' }
                Password = $password
                RequiresProcessedStamp = $true
                onPremisesSyncEnabled = [bool]$user.onPremisesSyncEnabled
            })
            Write-Log "Neues Passwort gesetzt und Benutzer für Sammelmail vorgemerkt: $upn" -Level OK
        }
        catch {
            Write-Log "${upn}: Verarbeitung fehlgeschlagen: $($_.Exception.Message)" -Level ERROR
        }
    }

    Write-UserProcessingSummary

    Write-Section '4. FACHLICHE SAMMELMAILS UND VERARBEITUNGSSTEMPEL'
    foreach ($unit in $SchoolUnitPatterns.Keys) {
        $schoolMailUsers = @($UsersPendingSchoolAccountMailByUnit[$unit])
        if ($schoolMailUsers.Count -eq 0) { continue }
        if ($DryRun) { continue }
        if (-not $SendSchoolAccountMails) {
            Write-Log "Sammelmail für $unit ist deaktiviert. Neue Benutzer werden nicht gestempelt; Reaktivierungen wurden bereits ausgeführt und nicht fachlich gemeldet." -Level WARN
            foreach ($schoolMailUser in @($schoolMailUsers | Where-Object { $_.RequiresProcessedStamp })) {
                $PasswordDeliveryFailures.Add([pscustomobject]@{ SchoolUnit = $unit; UserPrincipalName = $schoolMailUser.UserPrincipalName })
            }
            $Counters.PasswordDeliveryFailures += @($schoolMailUsers | Where-Object { $_.RequiresProcessedStamp }).Count
            continue
        }
        try {
            Send-SchoolAccountMail -SchoolUnit $unit -Users $schoolMailUsers
            $Counters.SchoolAccountMailsSent++
            Write-Log "Sammelmail für $unit an $($SchoolSupportRecipients[$unit]) versendet." -Level OK
            foreach ($schoolMailUser in @($schoolMailUsers | Where-Object { $_.RequiresProcessedStamp })) {
                try {
                    Set-ProcessedStamp -User $schoolMailUser
                    $Counters.ProcessedStampsSet++
                    Write-Log "Verarbeitungsstempel gesetzt: $($schoolMailUser.UserPrincipalName)" -Level OK
                }
                catch {
                    $Counters.ProcessedStampErrors++
                    Write-Log "$($schoolMailUser.UserPrincipalName): Mail versendet, Stempel fehlgeschlagen: $($_.Exception.Message)" -Level ERROR
                }
            }
        }
        catch {
            foreach ($schoolMailUser in @($schoolMailUsers | Where-Object { $_.RequiresProcessedStamp })) {
                $PasswordDeliveryFailures.Add([pscustomobject]@{ SchoolUnit = $unit; UserPrincipalName = $schoolMailUser.UserPrincipalName })
            }
            $Counters.PasswordDeliveryFailures += @($schoolMailUsers | Where-Object { $_.RequiresProcessedStamp }).Count
            Write-Log "${unit}: Sammelmail fehlgeschlagen. $($schoolMailUsers.Count) Benutzer wurden nicht gestempelt: $($_.Exception.Message)" -Level ERROR
        }
    }

    if ($UsersPendingDisabledAccountMail.Count -gt 0) {
        if ($DryRun) {
            Write-Log "VORSCHAU: Kein Mail zu deaktivierten Konten versendet. Erkannte Benutzer: $($UsersPendingDisabledAccountMail.Count)"
        }
        elseif (-not $SendDisabledAccountMails) {
            Write-Log 'Mail zu deaktivierten Konten ist deaktiviert.' -Level WARN
        }
        else {
            try {
                Send-DisabledAccountMail -Users $UsersPendingDisabledAccountMail.ToArray()
                $Counters.DisabledAccountMailsSent++
                Write-Log "Statusmail für $($UsersPendingDisabledAccountMail.Count) Konten an $($DisabledAccountNotificationRecipient) versendet." -Level OK
            }
            catch { Write-Log "Statusmail für deaktivierte Konten fehlgeschlagen: $($_.Exception.Message)" -Level ERROR }
        }
    }

    Write-Section '5. SCHLUSSZUSAMMENFASSUNG'
    foreach ($key in $Counters.Keys) { Write-Log ('{0,-34}: {1}' -f $key, $Counters[$key]) }
    Write-Log ('Warnungen                         : {0}' -f $WarningCount)
    Write-Log ('Fehler                            : {0}' -f $ErrorCount)
    Write-Log ('DryRun                            : {0}' -f $DryRun)
    Write-Log ('Debug                             : {0}' -f $Debug)
    Write-Log ('Dauer                             : {0:n1} Sekunden' -f ((Get-Date) - $StartedAt).TotalSeconds)
    Write-Log ('Vorgesehener Logbereich           : {0}' -f $(if ($ChangeLines.Count) { $SharePointLogFolder } else { "$SharePointLogFolder/$NoActionLogSubfolder" }))
    if ($ErrorCount -eq 0 -and $WarningCount -eq 0) { Write-Log 'Runbook erfolgreich beendet.' -Level OK }
    else { Write-Log 'Runbook mit Hinweisen beendet.' }
}
catch {
    $FatalError = $true
    $lineInfo = if ($_.InvocationInfo.ScriptLineNumber) { " (Skriptzeile $($_.InvocationInfo.ScriptLineNumber))" } else { '' }
    Write-Log "FATALER ABBRUCH${lineInfo}: $($_.Exception.Message)" -Level ERROR
    if ($null -ne $_.ErrorDetails -and -not [string]::IsNullOrWhiteSpace([string]$_.ErrorDetails.Message)) {
        Write-Log "Details: $($_.ErrorDetails.Message)" -Level ERROR
    }
    throw
}
finally {
    try {
        Save-LogToSharePoint
        if ($ResolvedLogPath) { Write-Output "SharePoint-Log gespeichert: $ResolvedLogPath. Run-ID: $RunId" }
    }
    catch { Write-Warning "SharePoint-Log konnte nicht gespeichert werden: $($_.Exception.Message)" }

    $mailForIssues = $EnableTechnicalSupportMail -and ($ErrorCount -gt 0 -or ($SendTechnicalMailOnWarnings -and $WarningCount -gt 0) -or $FatalError)
    $mailForDebugChanges = $EnableTechnicalSupportMail -and $Debug -and $ChangeLines.Count -gt 0
    $needMail = -not [string]::IsNullOrWhiteSpace($TechnicalSupportRecipient) -and ($mailForIssues -or $mailForDebugChanges)
    if ($needMail -and $GraphConnected) {
        try {
            $status = if ($FatalError -or $ErrorCount -gt 0) { 'FEHLER' }
                      elseif ($WarningCount -gt 0) { 'WARNUNG' }
                      elseif ($DryRun -and $ChangeLines.Count -gt 0) { 'VORSCHAU' }
                      elseif ($ChangeLines.Count -gt 0) { 'ÄNDERUNGEN' }
                      else { 'INFO' }
            Send-TechnicalSupportMail -Status $status
            Write-Output "Supportmail an $TechnicalSupportRecipient versendet. Run-ID: $RunId"
        }
        catch { Write-Warning "Supportmail konnte nicht versendet werden: $($_.Exception.Message)" }
    }

    if ($GraphConnected) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
}
