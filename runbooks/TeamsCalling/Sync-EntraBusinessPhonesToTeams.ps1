<#
.SYNOPSIS
  Azure Automation Runbook (PowerShell 7.6) zur Synchronisation von Entra businessPhones
  mit Microsoft Teams Direct Routing.

.DESCRIPTION
  Das Runbook gleicht die in Microsoft Entra ID hinterlegten businessPhones mit den
  Direct-Routing-Zuweisungen in Microsoft Teams ab. Entra businessPhones ist dabei die
  fuehrende Sollquelle. Verwaltet werden ausschliesslich Telefonnummern, die einem der
  konfigurierten erlaubten Nummernbereiche entsprechen. Reservierte Nummern und Ziele
  ausserhalb des definierten Benutzerumfangs werden nicht automatisch veraendert.

  Der Lauf erfolgt in klar getrennten Phasen:
    1. Konfiguration pruefen und Verbindungen zu Microsoft Graph und Microsoft Teams
       herstellen.
    2. Entra-Benutzer und bestehende Teams-Telefonnummernzuweisungen laden.
    3. Benutzer und Sollnummern validieren sowie einen Aktionsplan erstellen. Dabei werden
       unter anderem Gastkonten, deaktivierte oder nicht lizenzierte Benutzer, Benutzer
       gemaess Exclude-Regeln, ungueltige Nummern und doppelte Sollnummern ausgeschlossen.
    4. Verwaltete Teams-Nummern freigeben, wenn businessPhones leer ist.
    5. Nummern neu zuweisen, verschieben oder korrigieren und die erforderlichen Calling-,
       Voice-Routing- und UsageLocation-Einstellungen sicherstellen.
    6. Ergebnisse protokollieren, das Log in SharePoint speichern und je nach Ergebnis
       Support- oder Konfliktmails versenden.

  Vor jeder schreibenden Aktion wird die geplante Aenderung protokolliert. Nach einer
  Telefonnummernzuweisung oder -entfernung prueft das Runbook den sichtbaren Teams-Status.
  Schlaegt eine Verschiebung nach bereits erfolgter Entfernung fehl, wird soweit moeglich
  die vorherige Zuweisung wiederhergestellt.

  DryRun steuert, ob geplante Aenderungen tatsaechlich ausgefuehrt werden. Debug steuert
  ausschliesslich den Versand einer Aenderungsmail und verhindert keine Aenderungen.

  WICHTIG: Fuer den ersten Einsatz DryRun=$true belassen.

.VORAUSSETZUNGEN
  Azure Automation Runtime: PowerShell 7.6
  Vorab importierte Module (nicht im Runbook installieren):
    - Microsoft.Graph.Authentication
    - MicrosoftTeams

  Microsoft Graph Application Permissions der Managed Identity:
    - User.ReadWrite.All                  (usageLocation)
    - Mail.Send                           (Supportmail)
    - Sites.Selected                     (SharePoint-Log; auf der Site 'Accounts' wurde
                                          der Managed Identity die Rolle 'write' erteilt)

  Microsoft Teams:
    - Eine passende Entra-Rolle, z. B. Teams Telephony Administrator oder eine nach
      aktueller Microsoft-Dokumentation geeignete Rolle fuer die verwendeten Cmdlets.
    - App-basierte Teams-PowerShell-Authentifizierung muss fuer die Managed Identity
      bzw. den zugehoerigen Service Principal eingerichtet sein.

  HINWEIS ZUR TEAMS-ANMELDUNG:
  MicrosoftTeams 7.9.0 stellt in Azure Automation den Parametersatz -Identity bereit.
  Das Runbook verwendet daher direkt Connect-MicrosoftTeams -Identity.

  Codierung: UTF-8.
#>
#requires -Version 7.6

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
Set-StrictMode -Version 3.0

# -----------------------------------------------------------------------------
# KONFIGURATION
# -----------------------------------------------------------------------------
# Sollwerte fuer Teams-Richtlinien und den Entra-Nutzungsstandort.
$TeamsCallingPolicy = 'AllowCalling'
$VoiceRoutingPolicy = 'Peoplefone'
$UsageLocation = 'CH'
# Sicherheitsgrenze: Nur Nummern, die mindestens einem Muster entsprechen, werden verwaltet.
$AllowedPhoneNumberPatterns = @(
    '^\+415810082\d{2}$'
    '^\+415810081\d{2}$'
)
# Einzelne Nummern innerhalb der erlaubten Bereiche bleiben von Zuweisungen ausgeschlossen.
$BlockedPhoneNumbers = [ordered]@{
    '+41581008166' = 'Reserviert fuer Evakutations-Anlage der Schuleinheit Klosterguet'
    '+41581008167' = 'Reserviert fuer Evakutations-Anlage der Schuleinheit Wildenstein'
    '+41581008168' = 'Reserviert fuer Evakutations-Anlage der Schuleinheit Steig (Realschulhaus)'
    '+41581008169' = 'Reserviert fuer Evakutations-Anlage der Schuleinheit Steig (Sekundarschulhaus)'
}
# Mindestens einer dieser Serviceplaene muss beim Benutzer aktiv sein.
$TeamsPhoneServicePlanIds = @(
    '4828c8ec-dc2e-4779-b502-87ac9ce28ab7'
)

# DryRun protokolliert den Aktionsplan, fuehrt aber keine schreibenden Aktionen aus.
$DryRun = $false
# Debug steuert nur den Versand einer Aenderungsmail. Es verhindert keine Aenderung.
$Debug = $true

# Ist die Liste nicht leer, werden Veraenderungen nur fuer diese UPNs ausgefuehrt.
$TestUserPrincipalNames = @()

# Diese Benutzer und ihre bestehenden Zuweisungen sind ausdruecklich geschuetzt.
$IgnoredUserPrincipalNames = [ordered]@{
    'seb.wst@schule-rorschacherberg.ch' = 'Telefonnummer ist der CallQueue zugewiesen.'
    'Anja.Thunemann@schule-rorschacherberg.ch' = 'Standort Bildungskommission mit externer geschäftlicher Telefonnummer in Pupil'
}

# OR-Verknuepfung: Mindestens eine Regel muss passen.
# Eine leere Liste inkludiert absichtlich niemanden.
# Property unterstuetzt direkte und verschachtelte Pfade.
$IncludeUserPropertyRules = @(
    # [ordered]@{ Property = 'OfficeLocation';  Pattern = 'Steig' }
    # [ordered]@{ Property = 'OfficeLocation';  Pattern = 'Klosterguet' }   
    # [ordered]@{ Property = 'OfficeLocation';  Pattern = 'Wildenstein' }   
    # [ordered]@{ Property = 'OfficeLocation';  Pattern = 'Wartegg' }
    # [ordered]@{ Property = 'OfficeLocation';  Pattern = 'Wiesental' }
    # [ordered]@{ Property = 'OfficeLocation';  Pattern = 'Schulverwaltung' }    
    # [ordered]@{ Property = 'OfficeLocation';  Pattern = 'Tageshort' }   
      [ordered]@{ Property = 'userPrincipalName'; Pattern = '@schule-rorschacherberg\.ch$' }
    # [ordered]@{ Property = 'jobTitle';             Pattern = '\S' }
    # [ordered]@{ Property = 'onPremisesExtensionAttributes.extensionAttribute1'; Pattern = '^Telefonie$' }
)

# OR-Verknuepfung: Sobald eine Regel passt, wird der Benutzer ausgeschlossen.
# Eine leere Liste schliesst niemanden aus.
# Property unterstuetzt direkte und verschachtelte Pfade.
$ExcludeUserPropertyRules = @(
      [ordered]@{ Property = 'OfficeLocation';  Pattern = 'Tageshort' }
      [ordered]@{ Property = 'OfficeLocation';  Pattern = 'Mehrzweckhalle' }
      [ordered]@{ Property = 'OfficeLocation';  Pattern = 'Schulbus' }
    # Bildungskommission muss included sein, bis Dominique Baumgartner über Gemeinde-Phone kommuniziert; voraussichtlich mindestens 1.1.27
    # [ordered]@{ Property = 'OfficeLocation';  Pattern = 'Bildungskommission' }
      [ordered]@{ Property = 'businessPhones';  Pattern  = '71[\s.-]*858' } # Schulverwaltung         
    # [ordered]@{ Property = 'userPrincipalName'; Pattern = '^service-' }
    # [ordered]@{ Property = 'jobTitle';          Pattern = '^Schulleitung$' }
    # [ordered]@{ Property = 'onPremisesExtensionAttributes.extensionAttribute1'; Pattern = '^KeineTelefonie$' }
)

# Laufzeit-, Verifikations- und Wiederherstellungsparameter.
$TeamsPageSize = 1000
$MaximumRestoreAttempts = 1
$VerificationAttempts = 4
$VerificationDelaySeconds = 10
# Empfaenger und Schalter fuer Ergebnis-, Warnungs- und Konfliktmeldungen.
$SendSupportMail = $true
$SendMailOnWarnings = $true
$MailSender = 'informatik@schule-rorschacherberg.ch'
$LogSupportMail = 'daniel.zoller@schule-rorschacherberg.ch'
# Separate Meldung fuer doppelte und reservierte Sollnummern in Entra ID.
$SendPhoneConflictMail = $true
$PhoneConflictMail = 'schulverwaltung@rorschacherberg.ch'
# Ziel fuer die dauerhafte Ablage des technischen Laufprotokolls.
$SharePointDriveId = 'b!fx3JtQSXbkK3O1uu9sq1-c4Dqce5o5dPiIzsCi2qbYr3pomhXM6NTbYNHSepNYt6'
$SharePointLogFolder = 'Runbook-Logs/TeamsCalling'
$NoActionLogSubfolder = 'NoAction'
$ManagedIdentityClientId = ''

# -----------------------------------------------------------------------------
# LAUFSTATUS
# -----------------------------------------------------------------------------
# Pro Lauf getrennte Protokoll-, Mail-, Fehler- und Zaehlerdaten.
$RunId = [guid]::NewGuid().ToString('N').Substring(0, 8)
$StartedAt = Get-Date
$LogLines = [System.Collections.Generic.List[string]]::new()
$ChangeLines = [System.Collections.Generic.List[string]]::new()
$MailChangesByUser = @{}
$DuplicatePhoneConflicts = [System.Collections.Generic.List[object]]::new()
$BlockedPhoneConflicts = [System.Collections.Generic.List[object]]::new()
$DuplicateConflictErrorCount = 0
$BlockedConflictErrorCount = 0
$WarningCount = 0
$ErrorCount = 0
$FatalError = $false
$GraphConnected = $false
$TeamsConnected = $false
$ResolvedLogPath = ''
$Counters = [ordered]@{
    EntraUsersLoaded = 0; IgnoredUsers = 0; IncludeExcluded = 0
    ExcludeExcluded = 0; GuestsExcluded = 0; DisabledExcluded = 0; UnlicensedExcluded = 0
    EmptyBusinessPhones = 0; NumbersRemoved = 0
    NumbersAssigned = 0; NumbersChanged = 0; AssignmentsAlreadyCorrect = 0
    CallingPoliciesCorrected = 0; VoiceRoutingPoliciesCorrected = 0
    UsageLocationsCorrected = 0; InvalidNumbers = 0; BlockedNumbers = 0
    DuplicateDesiredNumbers = 0; LicenseErrors = 0; Conflicts = 0
    ProtectedConflicts = 0; RestoresSucceeded = 0; RestoresFailed = 0
    ChangesPlanned = 0
}

# Liefert einen Zeitstempel in der lokalen Schweizer Zeitzone fuer Logs und Mails.
function Get-LocalTime {
    try { $zone = [System.TimeZoneInfo]::FindSystemTimeZoneById('W. Europe Standard Time') }
    catch { $zone = [System.TimeZoneInfo]::FindSystemTimeZoneById('Europe/Zurich') }
    [System.TimeZoneInfo]::ConvertTime([System.DateTimeOffset]::UtcNow, $zone)
}
function Write-Log {
    param([Parameter(Mandatory)][string]$Message,
          [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO')
    $line = '[{0}] [{1}] [{2}] {3}' -f (Get-LocalTime).ToString('yyyy-MM-dd HH:mm:ss zzz'),$RunId,$Level,$Message
    $LogLines.Add($line)
    switch ($Level) {
        'WARN'  { $script:WarningCount++; Write-Warning $line }
        'ERROR' { $script:ErrorCount++; Write-Error $line -ErrorAction Continue }
        default { Write-Output $line }
    }
}
function Write-Section {
    param([Parameter(Mandatory)][string]$Title)
    foreach ($line in @('',('=' * 96),"  $Title",('=' * 96))) { $LogLines.Add($line); Write-Output $line }
}
function Write-UserSection {
    param([Parameter(Mandatory)]$User,[Parameter(Mandatory)][string]$Reason)
    foreach ($line in @('',('-' * 96),('  BENUTZER: {0} <{1}> | {2}' -f $User.displayName,$User.userPrincipalName,$Reason),('-' * 96))) {
        $LogLines.Add($line); Write-Output $line
    }
}
# Registriert eine geplante Aktion zentral fuer Log, Zaehler und Supportmail.
function Add-Change {
    param([Parameter(Mandatory)][string]$Type,[Parameter(Mandatory)][string]$Identity,
          [Parameter(Mandatory)][string]$Description)
    $line = '[{0}] {1}: {2}' -f $Type,$Identity,$Description
    $ChangeLines.Add($line)
    if (-not $script:MailChangesByUser.ContainsKey($Identity)) {
        $script:MailChangesByUser[$Identity] = [System.Collections.Generic.List[object]]::new()
    }
    $script:MailChangesByUser[$Identity].Add([pscustomobject]@{ Type=$Type; Description=$Description })
    $script:Counters.ChangesPlanned = $ChangeLines.Count
    Write-Log "AENDERUNG: $line$(if($DryRun){' [DRY RUN]'})"
}

# -----------------------------------------------------------------------------
# INCLUDE- UND PROPERTY-HILFSFUNKTIONEN
# -----------------------------------------------------------------------------
# Liest direkte oder verschachtelte Eigenschaften anhand eines Punktpfads.
function Get-NestedPropertyValue {
    param($InputObject,[Parameter(Mandatory)][string]$Path)
    $current = $InputObject
    foreach ($segment in ($Path -split '\.')) {
        if ($null -eq $current) { return $null }
        if ($current -is [System.Collections.IDictionary]) {
            if (-not $current.Contains($segment)) { return $null }
            $current = $current[$segment]
        } else {
            $property = $current.PSObject.Properties[$segment]
            if ($null -eq $property) { return $null }
            $current = $property.Value
        }
    }
    return $current
}
# Validiert die Include-Regeln frueh, damit fehlerhafte Regex-Ausdruecke den Lauf stoppen.
function Test-IncludeConfiguration {
    foreach ($rule in @($IncludeUserPropertyRules)) {
        if ($null -eq $rule) { throw 'Include-Regel darf nicht null sein.' }
        $property = [string]$rule.Property
        $pattern = [string]$rule.Pattern
        if ([string]::IsNullOrWhiteSpace($property)) { throw 'Include-Regel ohne Property gefunden.' }
        if ([string]::IsNullOrWhiteSpace($pattern)) { throw "Include-Regel fuer '$property' hat kein Pattern." }
        try { [void][regex]::new($pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase) }
        catch { throw "Ungueltige Include-Regex fuer '$property': $pattern. $($_.Exception.Message)" }
    }
}
# Validiert die Exclude-Regeln frueh, damit fehlerhafte Regex-Ausdruecke den Lauf stoppen.
function Test-ExcludeConfiguration {
    foreach ($rule in @($ExcludeUserPropertyRules)) {
        if ($null -eq $rule) { throw 'Exclude-Regel darf nicht null sein.' }
        $property = [string]$rule.Property
        $pattern = [string]$rule.Pattern
        if ([string]::IsNullOrWhiteSpace($property)) { throw 'Exclude-Regel ohne Property gefunden.' }
        if ([string]::IsNullOrWhiteSpace($pattern)) { throw "Exclude-Regel fuer '$property' hat kein Pattern." }
        try { [void][regex]::new($pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase) }
        catch { throw "Ungueltige Exclude-Regex fuer '$property': $pattern. $($_.Exception.Message)" }
    }
}
# OR-Logik: Ein Benutzer ist enthalten, sobald eine konfigurierte Regel passt.
function Test-UserIncluded {
    param([Parameter(Mandatory)]$User)
    if (@($IncludeUserPropertyRules).Count -eq 0) { return $false }
    foreach ($rule in @($IncludeUserPropertyRules)) {
        $value = Get-NestedPropertyValue -InputObject $User -Path ([string]$rule.Property)
        foreach ($candidate in @($value)) {
            if ($null -ne $candidate -and ([string]$candidate -match [string]$rule.Pattern)) { return $true }
        }
    }
    return $false
}
# OR-Logik: Ein Benutzer ist ausgeschlossen, sobald eine konfigurierte Regel passt.
function Test-UserExcluded {
    param([Parameter(Mandatory)]$User)
    if (@($ExcludeUserPropertyRules).Count -eq 0) { return $false }
    foreach ($rule in @($ExcludeUserPropertyRules)) {
        $value = Get-NestedPropertyValue -InputObject $User -Path ([string]$rule.Property)
        foreach ($candidate in @($value)) {
            if ($null -ne $candidate -and ([string]$candidate -match [string]$rule.Pattern)) { return $true }
        }
    }
    return $false
}
function Test-TestScopeUser {
    param([string]$Upn)
    if ($TestUserPrincipalNames.Count -eq 0) { return $true }
    return $TestUserPrincipalNames -contains $Upn
}
# Baut die minimale Graph-$select-Liste inklusive der Wurzeleigenschaften aus Include- und Exclude-Regeln.
function Get-GraphSelectProperties {
    $base = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($name in @('id','displayName','userPrincipalName','userType','accountEnabled','businessPhones','usageLocation','assignedLicenses','assignedPlans')) { [void]$base.Add($name) }
    foreach ($rule in @($IncludeUserPropertyRules) + @($ExcludeUserPropertyRules)) {
        $root = ([string]$rule.Property -split '\.')[0]
        if (-not [string]::IsNullOrWhiteSpace($root)) { [void]$base.Add($root) }
    }
    return (@($base) | Sort-Object) -join ','
}

# -----------------------------------------------------------------------------
# TELEFONNUMMERN
# -----------------------------------------------------------------------------
function Test-AllowedPhoneNumber {
    param([string]$PhoneNumber)
    foreach ($pattern in $AllowedPhoneNumberPatterns) { if ($PhoneNumber -match $pattern) { return $true } }
    return $false
}
# Normalisiert gaengige Schweizer Schreibweisen und erzwingt ein eindeutiges E.164-Format.
function ConvertTo-SwissE164 {
    param([Parameter(Mandatory)][string]$Value)
    $v = $Value.Trim()
    $v = $v -replace '^tel\s*:', ''
    $v = $v -replace '[\s\.\-\(\)/]', ''
    if ($v -match '^0041') { $v = '+' + $v.Substring(2) }
    elseif ($v -match '^0\d+$') { $v = '+41' + $v.Substring(1) }
    if ($v -notmatch '^\+[1-9]\d{7,14}$') { throw "Kein gueltiges E.164-Format: '$Value'" }
    if ($v -notmatch '^\+41\d{9}$') { throw "Keine eindeutige Schweizer Nummer mit 9 nationalen Ziffern: '$Value'" }
    return $v
}
# Klassifiziert businessPhones als leer, mehrfach, ungueltig, blockiert oder gueltig.
function Get-BusinessPhoneState {
    param($BusinessPhones)
    $values = @($BusinessPhones | ForEach-Object { if ($null -ne $_ -and -not [string]::IsNullOrWhiteSpace([string]$_)) { ([string]$_).Trim() } })
    if ($values.Count -eq 0) { return [pscustomobject]@{ State='Empty'; Values=@(); Number=$null; Error=$null } }
    if ($values.Count -gt 1) { return [pscustomobject]@{ State='Multiple'; Values=$values; Number=$null; Error=$null } }
    try {
        $number = ConvertTo-SwissE164 -Value $values[0]
        if (-not (Test-AllowedPhoneNumber $number)) { return [pscustomobject]@{ State='Invalid'; Values=$values; Number=$number; Error='Nummer liegt ausserhalb der erlaubten Bereiche.' } }
        if ($BlockedPhoneNumbers.Contains($number)) { return [pscustomobject]@{ State='Blocked'; Values=$values; Number=$number; Error=$BlockedPhoneNumbers[$number] } }
        return [pscustomobject]@{ State='Valid'; Values=$values; Number=$number; Error=$null }
    } catch { return [pscustomobject]@{ State='Invalid'; Values=$values; Number=$null; Error=$_.Exception.Message } }
}
function Test-TeamsPhonePlan {
    param($User)
    foreach ($plan in @($User.assignedPlans)) {
        if (($TeamsPhoneServicePlanIds -contains [string]$plan.servicePlanId) -and ([string]$plan.capabilityStatus -eq 'Enabled')) { return $true }
    }
    return $false
}

# -----------------------------------------------------------------------------
# GRAPH
# -----------------------------------------------------------------------------
# Einheitlicher Wrapper fuer authentifizierte Microsoft-Graph-Aufrufe.
function Invoke-Graph {
    param([ValidateSet('GET','POST','PATCH')][string]$Method,[Parameter(Mandatory)][string]$Uri,[object]$Body)
    $p = @{ Method=$Method; Uri=$Uri; OutputType='PSObject'; ErrorAction='Stop' }
    if ($PSBoundParameters.ContainsKey('Body')) { $p.Body=$Body | ConvertTo-Json -Depth 15 -Compress; $p.ContentType='application/json; charset=utf-8' }
    Invoke-MgGraphRequest @p
}
# Laedt eine vollstaendige Graph-Sammlung und folgt automatisch @odata.nextLink.
function Get-GraphCollection {
    param([Parameter(Mandatory)][string]$Uri)
    $all = [System.Collections.Generic.List[object]]::new(); $next = $Uri
    while (-not [string]::IsNullOrWhiteSpace([string]$next)) {
        $page = Invoke-Graph -Method GET -Uri $next
        if ($null -eq $page -or -not ($page.PSObject.Properties.Name -contains 'value')) { throw "Graph-Antwort fuer '$next' enthaelt keine value-Eigenschaft." }
        foreach ($item in @($page.value)) { if ($null -ne $item) { $all.Add($item) } }
        $property = $page.PSObject.Properties['@odata.nextLink']
        $next = if ($null -ne $property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) { [string]$property.Value } else { $null }
    }
    return $all
}
function Update-EntraUsageLocation {
    param($User)
    if ([string]$User.usageLocation -eq $UsageLocation) { return }
    Add-Change -Type 'UsageLocation' -Identity $User.userPrincipalName -Description "$($User.usageLocation) -> $UsageLocation"
    if (-not $DryRun) { Invoke-Graph -Method PATCH -Uri "v1.0/users/$($User.id)" -Body @{ usageLocation=$UsageLocation } | Out-Null }
    $script:Counters.UsageLocationsCorrected++
}

# -----------------------------------------------------------------------------
# TEAMS UND LAUFZEIT-INDIZES
# -----------------------------------------------------------------------------
# Liest alle Teams-Telefonnummernzuweisungen seitenweise ein.
function Get-AllTeamsPhoneAssignments {
    $all=[System.Collections.Generic.List[object]]::new(); $skip=0
    while ($true) {
        $page=@(Get-CsPhoneNumberAssignment -Top $TeamsPageSize -Skip $skip -ErrorAction Stop)
        foreach($item in $page){$all.Add($item)}
        if($page.Count -lt $TeamsPageSize){break}; $skip += $TeamsPageSize
    }
    return $all
}
# Kompatibilitaetshilfe fuer unterschiedliche Eigenschaftsnamen zwischen Modulversionen.
function Get-ObjectPropertyValue {
    param($InputObject,[Parameter(Mandatory)][string[]]$PropertyNames,$DefaultValue=$null)
    if ($null -eq $InputObject) { return $DefaultValue }
    $item=@($InputObject)|Select-Object -First 1
    foreach($name in $PropertyNames){$p=$item.PSObject.Properties[$name];if($null-ne$p){return $p.Value}}
    return $DefaultValue
}
function Get-AssignmentNumber { param($Assignment) $v=Get-ObjectPropertyValue $Assignment @('TelephoneNumber','PhoneNumber');if($null-eq$v){return $null};[string]$v }
function Get-AssignmentOwnerId { param($Assignment) $v=Get-ObjectPropertyValue $Assignment @('AssignedPstnTargetId','TargetId','AssignedPstnTarget');if($null-eq$v){return $null};[string]$v }
function Get-AssignmentNumberType { param($Assignment) [string](Get-ObjectPropertyValue $Assignment @('NumberType','Type') '') }
function Get-AssignmentTargetType { param($Assignment) [string](Get-ObjectPropertyValue $Assignment @('TargetType','AssignedPstnTargetType') '') }
function Get-UserPhoneAssignment {
    param([string]$Identity)
    @(Get-CsPhoneNumberAssignment -AssignedPstnTargetId $Identity -ErrorAction Stop) | Where-Object { (Get-AssignmentNumberType $_) -eq 'DirectRouting' } | Select-Object -First 1
}
# Prueft mit Wiederholungen, ob eine Teams-Zuweisung den erwarteten Zustand erreicht hat.
function Test-PhoneAssignment {
    param([string]$Identity,[AllowNull()][string]$ExpectedNumber)
    for($i=1;$i-le$VerificationAttempts;$i++){
        $actual=Get-AssignmentNumber (Get-UserPhoneAssignment $Identity)
        if(([string]::IsNullOrWhiteSpace($ExpectedNumber)-and[string]::IsNullOrWhiteSpace($actual))-or$actual-eq$ExpectedNumber){return $true}
        if($i-lt$VerificationAttempts){Start-Sleep -Seconds $VerificationDelaySeconds}
    }
    return $false
}
# Aktualisiert nur die lokalen Laufzeitindizes nach einer erfolgreichen oder simulierten Aktion.
function Set-RuntimeAssignment {
    param([string]$OwnerId,[string]$OwnerIdentity,[AllowNull()][string]$OldNumber,[AllowNull()][string]$NewNumber)
    if ($OldNumber -and $script:AssignmentsByNumber.ContainsKey($OldNumber)) { $script:AssignmentsByNumber.Remove($OldNumber) }
    if ($OwnerId) { $script:CurrentNumberByOwner[$OwnerId] = $NewNumber }
    if ($NewNumber) {
        $script:AssignmentsByNumber[$NewNumber] = [pscustomobject]@{
            TelephoneNumber=$NewNumber; AssignedPstnTargetId=$OwnerId
            NumberType='DirectRouting'; TargetType='User'; RuntimeIdentity=$OwnerIdentity
        }
    }
}
# Entfernt ausschliesslich eine Nummer aus dem verwalteten Bereich und verifiziert das Ergebnis.
function Remove-ManagedPhone {
    param([string]$Identity,[string]$OwnerId,[string]$Number,[switch]$ResetPolicies,[string]$Reason='Entfernen')
    if(-not(Test-AllowedPhoneNumber $Number)){throw "Sicherheitsstopp: $Number liegt ausserhalb der verwalteten Bereiche."}
    Add-Change -Type 'Nummer entfernen' -Identity $Identity -Description "$Number ($Reason)"
    if(-not$DryRun){
        Remove-CsPhoneNumberAssignment -Identity $Identity -TelephoneNumber $Number -NumberType DirectRouting -ErrorAction Stop | Out-Null
        if(-not(Test-PhoneAssignment $Identity $null)){throw "Verifikation fehlgeschlagen: $Number ist bei $Identity weiterhin sichtbar."}
    }
    Set-RuntimeAssignment -OwnerId $OwnerId -OwnerIdentity $Identity -OldNumber $Number -NewNumber $null
    if($ResetPolicies){
        Add-Change -Type 'Policies vererben' -Identity $Identity -Description 'Calling- und Voice-Routing-Policy auf Vererbung zuruecksetzen'
        if(-not$DryRun){Grant-CsTeamsCallingPolicy -Identity $Identity -PolicyName $null -ErrorAction Stop;Grant-CsOnlineVoiceRoutingPolicy -Identity $Identity -PolicyName $null -ErrorAction Stop}
    }
    $script:Counters.NumbersRemoved++
}
# Weist eine verwaltete Direct-Routing-Nummer zu und verifiziert das Ergebnis.
function Set-ManagedPhone {
    param([string]$Identity,[string]$OwnerId,[string]$Number,[string]$Reason='Zuweisen')
    if(-not(Test-AllowedPhoneNumber $Number)){throw "Sicherheitsstopp: $Number liegt ausserhalb der verwalteten Bereiche."}
    $old=if($script:CurrentNumberByOwner.ContainsKey($OwnerId)){$script:CurrentNumberByOwner[$OwnerId]}else{$null}
    Add-Change -Type 'Nummer zuweisen' -Identity $Identity -Description "$Number ($Reason)"
    if(-not$DryRun){
        Set-CsPhoneNumberAssignment -Identity $Identity -TelephoneNumber $Number -NumberType DirectRouting -ErrorAction Stop | Out-Null
        if(-not(Test-PhoneAssignment $Identity $Number)){throw "Verifikation fehlgeschlagen: $Number ist bei $Identity nicht sichtbar."}
    }
    Set-RuntimeAssignment -OwnerId $OwnerId -OwnerIdentity $Identity -OldNumber $old -NewNumber $Number
}
# Ermittelt die aktuell wirksamen Calling- und Voice-Routing-Policy-Namen robust.
function Get-TeamsUserPolicyState {
    param(
        [Parameter(Mandatory)]
        [string]$Identity
    )

    $u = Get-CsOnlineUser -Identity $Identity -ErrorAction Stop

    $callingValue = Get-ObjectPropertyValue `
        -InputObject $u `
        -PropertyNames @('TeamsCallingPolicy') `
        -DefaultValue $null

    $routingValue = Get-ObjectPropertyValue `
        -InputObject $u `
        -PropertyNames @('OnlineVoiceRoutingPolicy') `
        -DefaultValue $null

    $calling = if ($callingValue -is [string]) {
        [string]$callingValue
    }
    elseif ($null -ne $callingValue) {
        Get-ObjectPropertyValue `
            -InputObject $callingValue `
            -PropertyNames @('Name', 'PolicyName', 'Identity')`
            -DefaultValue ''
    }
    else {
        ''
    }

    $routing = if ($routingValue -is [string]) {
        [string]$routingValue
    }
    elseif ($null -ne $routingValue) {
        Get-ObjectPropertyValue `
            -InputObject $routingValue `
            -PropertyNames @('Name', 'PolicyName', 'Identity') `
            -DefaultValue ''
    }
    else {
        ''
    }

    [pscustomobject]@{
        Calling = $calling
        Routing = $routing
    }
}

function Set-RequiredPolicies {
    param([string]$Identity)
    $state=Get-TeamsUserPolicyState $Identity
    if($state.Calling-ne$TeamsCallingPolicy){Add-Change 'Calling Policy' $Identity "$($state.Calling) -> $TeamsCallingPolicy";if(-not$DryRun){Grant-CsTeamsCallingPolicy -Identity $Identity -PolicyName $TeamsCallingPolicy -ErrorAction Stop};$script:Counters.CallingPoliciesCorrected++}
    if($state.Routing-ne$VoiceRoutingPolicy){Add-Change 'Voice Routing Policy' $Identity "$($state.Routing) -> $VoiceRoutingPolicy";if(-not$DryRun){Grant-CsOnlineVoiceRoutingPolicy -Identity $Identity -PolicyName $VoiceRoutingPolicy -ErrorAction Stop};$script:Counters.VoiceRoutingPoliciesCorrected++}
}
# Verhindert automatische Verschiebungen von geschuetzten oder nicht eindeutig verwaltbaren Zielen.
function Test-ProtectedAssignmentHolder {
    param($Assignment,[string]$HolderId)
    $targetType=Get-AssignmentTargetType $Assignment
    if($targetType -and $targetType -notmatch '^(User|Unknown)$'){return "Geschuetzter Teams-Zieltyp '$targetType'"}
    if(-not$script:UsersById.ContainsKey($HolderId)){return 'Unbekanntes oder nicht als Entra-Benutzer aufloesbares Teams-Ziel'}
    $u=$script:UsersById[$HolderId];$upn=[string]$u.userPrincipalName
    if($IgnoredUserPrincipalNames.Contains($upn)){return 'Benutzer steht in der Ignore-Liste'}
    if(-not(Test-UserIncluded $u)){return 'Benutzer ist durch keine Include-Regel erfasst'}
    if(Test-UserExcluded $u){return 'Benutzer ist durch eine Exclude-Regel ausgeschlossen'}
    if(-not(Test-TestScopeUser $upn)){return 'Benutzer liegt ausserhalb der Testbenutzerbegrenzung'}
    return $null
}
# Best-Effort-Rollback fuer Nummern, die waehrend einer fehlgeschlagenen Verschiebung entfernt wurden.
function Restore-Number {
    param([string]$Identity,[string]$OwnerId,[string]$Number,[string]$Context)
    if([string]::IsNullOrWhiteSpace($Number)){return $true}
    for($r=1;$r-le$MaximumRestoreAttempts;$r++){
        try{Set-ManagedPhone -Identity $Identity -OwnerId $OwnerId -Number $Number -Reason "Rollback: $Context";$script:Counters.RestoresSucceeded++;return $true}
        catch{Write-Log "Rollback bei $Identity fuer $Number fehlgeschlagen: $($_.Exception.Message)" -Level ERROR}
    }
    $script:Counters.RestoresFailed++;return $false
}

# -----------------------------------------------------------------------------
# MAIL UND SHAREPOINT
# -----------------------------------------------------------------------------
function ConvertTo-HtmlEncoded { param([string]$Text) [Net.WebUtility]::HtmlEncode($Text) }
function Format-PhoneForMail {
    param([AllowNull()][string]$Number)
    if($Number -match '^\+41(\d{2})(\d{3})(\d{2})(\d{2})$'){return "+41 $($Matches[1]) $($Matches[2]) $($Matches[3]) $($Matches[4])"}
    return $Number
}
# Verdichtet technische Einzelaktionen pro Benutzer fuer eine lesbare Supportmail.
function Get-MailChangeGroups {
    $newAssignments=[System.Collections.Generic.List[object]]::new()
    $changes=[System.Collections.Generic.List[object]]::new()
    foreach($identity in @($MailChangesByUser.Keys | Sort-Object)){
        $items=@($MailChangesByUser[$identity])
        $types=@($items.Type)
        $assigned=@($items | Where-Object {$_.Type -eq 'Nummer zuweisen'} | Select-Object -Last 1)
        $removed=@($items | Where-Object {$_.Type -eq 'Nummer entfernen'} | Select-Object -First 1)
        $policyTypes=@($items | Where-Object {$_.Type -in @('Calling Policy','Voice Routing Policy','UsageLocation','Policies vererben')} | ForEach-Object {$_.Type})
        $newNumber=if($assigned.Count){([regex]::Match([string]$assigned[0].Description,'\+\d+')).Value}else{''}
        $oldNumber=if($removed.Count){([regex]::Match([string]$removed[0].Description,'\+\d+')).Value}else{''}
        $details=[System.Collections.Generic.List[string]]::new()
        if($oldNumber -and $newNumber){$details.Add("Telefonnummer: $(Format-PhoneForMail $oldNumber) → $(Format-PhoneForMail $newNumber)")}
        elseif($newNumber){$details.Add("Telefonnummer: $(Format-PhoneForMail $newNumber)")}
        elseif($oldNumber){$details.Add("Telefonnummer gelöscht: $(Format-PhoneForMail $oldNumber)")}
        $corrected=[System.Collections.Generic.List[string]]::new()
        if($policyTypes -contains 'Calling Policy'){$corrected.Add('Calling Policy')}
        if($policyTypes -contains 'Voice Routing Policy'){$corrected.Add('Voice Routing Policy')}
        if($policyTypes -contains 'UsageLocation'){$corrected.Add('Nutzungsstandort')}
        if($policyTypes -contains 'Policies vererben'){$corrected.Add('Richtlinien auf Vererbung zurückgesetzt')}
        if($corrected.Count){$details.Add('Zusätzlich: '+($corrected -join ', '))}
        $entry=[pscustomobject]@{Identity=$identity;Details=@($details)}
        if($newNumber -and -not $oldNumber){$newAssignments.Add($entry)}else{$changes.Add($entry)}
    }
    [pscustomobject]@{NewAssignments=@($newAssignments);Changes=@($changes)}
}
function Convert-MailEntriesToHtml {
    param([object[]]$Entries)
    if(@($Entries).Count-eq0){return '<p style="color:#666;margin:6px 0 18px 0;">Keine Einträge.</p>'}
    $rows=foreach($entry in @($Entries)){
        $detailHtml=(@($entry.Details) | ForEach-Object {'<div style="color:#444;margin-top:3px;">'+(ConvertTo-HtmlEncoded $_)+'</div>'}) -join ''
        '<div style="border:1px solid #e1e5ea;border-radius:8px;padding:12px 14px;margin:8px 0;background:#fff;"><div style="font-weight:600;color:#172b4d;">'+(ConvertTo-HtmlEncoded $entry.Identity)+'</div>'+$detailHtml+'</div>'
    }
    return ($rows -join '')
}
# Bereitet Fehler und Warnungen aus dem technischen Log fuer die Supportmail auf.
# Falls eine Supportmail versendet wird, werden auch doppelte Sollnummern aufgefuehrt.
function Get-MailIssueSummaryHtml {
    $issueLines = @(
        $LogLines | Where-Object {
            $_ -match '\[(ERROR|WARN)\]'
        }
    )

    if ($issueLines.Count -eq 0) { return '' }

    $items = foreach ($line in $issueLines) {
        $message = [regex]::Replace(
            [string]$line,
            '^\[[^\]]+\] \[[^\]]+\] \[(?:ERROR|WARN)\]\s*',
            ''
        )
        '<li style="margin:5px 0;">' + (ConvertTo-HtmlEncoded $message) + '</li>'
    }

    return @"
  <h3 style="margin-bottom:6px;">Fehler und Hinweise</h3>
  <div style="border:1px solid #e6a700;border-radius:8px;padding:10px 14px;margin:8px 0 18px 0;background:#fff8e1;">
    <ul style="margin:0;padding-left:20px;">$($items -join '')</ul>
  </div>
"@
}
function Send-SupportLogMail {
    param(
        [Parameter(Mandatory)]
        [string]$Status
    )

    # Technisches Log fuer den Anhang vorbereiten
    $content = $LogLines -join "`r`n"
    $bytes = [Text.Encoding]::UTF8.GetBytes($content)

    $name = 'TeamsCalling_{0}_{1}.log' -f `
        (Get-LocalTime).ToString('yyyy-MM-dd_HHmmss'), `
        $RunId

    # Aenderungen fuer die lesbare Darstellung gruppieren
    $groups = Get-MailChangeGroups

    # Hinweis nur bei einem Dry Run anzeigen
    $modeText = if ($DryRun) {
        'Vorschau: Es wurden keine Änderungen ausgeführt.'
    }
    else {
        ''
    }

    # Vollstaendigen HTML-Block nur erzeugen, wenn ein Hinweis vorhanden ist
    $modeHtml = if (-not [string]::IsNullOrWhiteSpace($modeText)) {
        @"
<p style="padding:10px 12px;background:#f3f7fb;border-left:4px solid #0f6cbd;border-radius:4px;">
  $(ConvertTo-HtmlEncoded $modeText)
</p>
"@
    }
    else {
        ''
    }

    # Einleitung abhaengig vom Ergebnis des Runbook-Laufs
    $intro = if ($Status -eq 'FEHLER') {
        'Der Lauf wurde mit Fehlern beendet.'
    }
    elseif ($Status -eq 'WARNUNG') {
        'Der Lauf wurde mit Hinweisen beendet.'
    }
    elseif ($ChangeLines.Count -gt 0) {
        'Die Telefonsynchronisation wurde erfolgreich abgeschlossen.'
    }
    else {
        'Es waren keine Änderungen erforderlich.'
    }

    # Fehler und Warnungen als kurze, direkt sichtbare Zusammenfassung aufbereiten
    $issueSummaryHtml = Get-MailIssueSummaryHtml

    # HTML-Inhalt der Mail
    $html = @"
<div style="font-family:Segoe UI,Arial,sans-serif;max-width:760px;color:#222;line-height:1.45;">
  <h2 style="color:#0f6cbd;margin-bottom:6px;">Teams-Telefonie: $Status</h2>
  <p style="margin-top:0;">$intro</p>

  $modeHtml

  $issueSummaryHtml

  <h3 style="margin-bottom:6px;">Neue Nummernvergaben</h3>
  $(Convert-MailEntriesToHtml $groups.NewAssignments)

  <h3 style="margin-bottom:6px;">Änderungen und Löschungen</h3>
  $(Convert-MailEntriesToHtml $groups.Changes)

  <div style="font-size:11px;color:#666;border-top:1px solid #ddd;margin-top:24px;padding-top:10px;">
    <strong>Technische Informationen</strong><br>
    Run-ID: $RunId · Zeitpunkt: $((Get-LocalTime).ToString('dd.MM.yyyy HH:mm:ss')) · Dry Run: $DryRun · Debug: $Debug<br>
    Einzelaktionen: $($ChangeLines.Count) · Warnungen: $WarningCount · Fehler ohne Telefonnummernkonflikte: $([Math]::Max(0, $ErrorCount - $DuplicateConflictErrorCount))<br>
    Das vollständige technische Log befindet sich im Anhang.
  </div>
</div>
"@

    # Graph-Payload erstellen
    $body = @{
        message = @{
            subject = "SRB Teams-Telefonie: $Status"
            body = @{
                contentType = 'HTML'
                content     = $html
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $LogSupportMail
                    }
                }
            )
            attachments = @(
                @{
                    '@odata.type' = '#microsoft.graph.fileAttachment'
                    name          = $name
                    contentType   = 'text/plain; charset=utf-8'
                    contentBytes  = [Convert]::ToBase64String($bytes)
                }
            )
        }
        saveToSentItems = $false
    }

    # Mail ueber Microsoft Graph versenden
    Invoke-Graph `
        -Method POST `
        -Uri "v1.0/users/$([uri]::EscapeDataString($MailSender))/sendMail" `
        -Body $body |
        Out-Null
}
function Send-BlockedPhoneConflictMail {
    if($BlockedPhoneConflicts.Count-eq0){return}
    $conflictHtml=foreach($conflict in @($BlockedPhoneConflicts | Sort-Object Number,UserPrincipalName)){
        '<div style="border:1px solid #d83b01;border-radius:8px;padding:12px 14px;margin:10px 0;background:#fff4f0;">'+
        '<div style="font-size:17px;font-weight:700;">'+(ConvertTo-HtmlEncoded (Format-PhoneForMail $conflict.Number))+'</div>'+
        '<div style="margin-top:6px;"><strong>'+(ConvertTo-HtmlEncoded $conflict.DisplayName)+'</strong><br>'+
        '<span style="color:#555;">'+(ConvertTo-HtmlEncoded $conflict.UserPrincipalName)+'</span></div>'+
        '<div style="color:#7a2e0e;margin-top:8px;">'+(ConvertTo-HtmlEncoded $conflict.Reason)+'</div></div>'
    }
    $html=@"
<div style="font-family:Segoe UI,Arial,sans-serif;max-width:760px;color:#222;line-height:1.45;">
  <h2 style="color:#d83b01;margin-bottom:6px;">Reservierte Telefonnummer in Entra ID</h2>
  <p>Bei der Telefonsynchronisation wurde mindestens eine Telefonnummer gefunden, die in der Konfiguration als reserviert oder gesperrt hinterlegt ist. Für die betroffenen Benutzer wurden keine Änderungen an der Teams-Telefonie vorgenommen.</p>
  <p style="padding:10px 12px;background:#fff4ce;border-left:4px solid #d83b01;border-radius:4px;"><strong>Bitte in Pupil korrigieren:</strong> Die Telefonnummern werden aus Pupil in das Feld <strong>Telefon Geschäftlich</strong> von Entra ID übernommen. Korrigiere deshalb die betroffenen Einträge direkt in Pupil. Beim nächsten Lauf wird die Zuordnung erneut geprüft.</p>
  $($conflictHtml -join '')
  <div style="font-size:11px;color:#666;border-top:1px solid #ddd;margin-top:24px;padding-top:10px;">
    <strong>Technische Informationen</strong><br>
    Run-ID: $RunId · Zeitpunkt: $((Get-LocalTime).ToString('dd.MM.yyyy HH:mm:ss')) · Gefundene reservierte Nummern: $($BlockedPhoneConflicts.Count)
  </div>
</div>
"@
    $to=@(@{emailAddress=@{address=$PhoneConflictMail}})
    $cc=@()
    if(-not[string]::IsNullOrWhiteSpace($LogSupportMail)-and$LogSupportMail-ne$PhoneConflictMail){$cc=@(@{emailAddress=@{address=$LogSupportMail}})}
    $message=@{subject="SRB Teams-Telefonie: Reservierte Telefonnummer in Entra ID";body=@{contentType='HTML';content=$html};toRecipients=$to}
    if($cc.Count){$message.ccRecipients=$cc}
    $body=@{message=$message;saveToSentItems=$false}
    Invoke-Graph -Method POST -Uri "v1.0/users/$([uri]::EscapeDataString($MailSender))/sendMail" -Body $body | Out-Null
}
function Send-DuplicatePhoneConflictMail {
    if($DuplicatePhoneConflicts.Count-eq0){return}
    $conflictHtml=foreach($conflict in @($DuplicatePhoneConflicts | Sort-Object Number)){
        $usersHtml=(@($conflict.Users) | Sort-Object UserPrincipalName | ForEach-Object {
            '<li><strong>'+(ConvertTo-HtmlEncoded $_.DisplayName)+'</strong><br><span style="color:#555;">'+(ConvertTo-HtmlEncoded $_.UserPrincipalName)+'</span></li>'
        }) -join ''
        '<div style="border:1px solid #f0b429;border-radius:8px;padding:12px 14px;margin:10px 0;background:#fffaf0;"><div style="font-size:17px;font-weight:700;">'+(ConvertTo-HtmlEncoded (Format-PhoneForMail $conflict.Number))+'</div><ul style="margin-bottom:0;">'+$usersHtml+'</ul></div>'
    }
    $html=@"
<div style="font-family:Segoe UI,Arial,sans-serif;max-width:760px;color:#222;line-height:1.45;">
  <h2 style="color:#a15c00;margin-bottom:6px;">Doppelte Telefonnummern in Entra ID</h2>
  <p>Bei der Telefonsynchronisation wurden Telefonnummern gefunden, die mehreren Benutzern zugewiesen sind. Für diese Benutzer wurden keine Änderungen an der Teams-Telefonie vorgenommen.</p>
  <p style="padding:10px 12px;background:#fff4ce;border-left:4px solid #f0b429;border-radius:4px;"><strong>Bitte in Pupil korrigieren:</strong> Die Telefonnummern werden aus Pupil in das Feld <strong>Telefon Geschäftlich</strong> von Entra ID übernommen. Korrigiere deshalb die betroffenen Einträge direkt in Pupil. Beim nächsten Lauf wird die korrigierte Zuordnung erneut geprüft.</p>
  $($conflictHtml -join '')
  <div style="font-size:11px;color:#666;border-top:1px solid #ddd;margin-top:24px;padding-top:10px;">
    <strong>Technische Informationen</strong><br>
    Run-ID: $RunId · Zeitpunkt: $((Get-LocalTime).ToString('dd.MM.yyyy HH:mm:ss')) · Gefundene Nummernkonflikte: $($DuplicatePhoneConflicts.Count)
  </div>
</div>
"@
    $to=@(@{emailAddress=@{address=$PhoneConflictMail}})
    $cc=@()
    if(-not[string]::IsNullOrWhiteSpace($LogSupportMail)-and$LogSupportMail-ne$PhoneConflictMail){$cc=@(@{emailAddress=@{address=$LogSupportMail}})}
    $message=@{subject="SRB Teams-Telefonie: Doppelte Telefonnummern in Entra ID";body=@{contentType='HTML';content=$html};toRecipients=$to}
    if($cc.Count){$message.ccRecipients=$cc}
    $body=@{message=$message;saveToSentItems=$false}
    Invoke-Graph -Method POST -Uri "v1.0/users/$([uri]::EscapeDataString($MailSender))/sendMail" -Body $body | Out-Null
}
# Speichert das Laufprotokoll; Laeufe ohne Aktionen landen im separaten NoAction-Unterordner.
function Save-LogToSharePoint {
    if([string]::IsNullOrWhiteSpace($SharePointDriveId)-or$LogLines.Count-eq0){return}
    $name='TeamsCalling_{0}_{1}.log' -f (Get-LocalTime).ToString('yyyy-MM-dd_HHmmss'),$RunId
    $folder=$SharePointLogFolder.Trim('/')
    if($ChangeLines.Count-eq0){$folder=(@($folder,$NoActionLogSubfolder.Trim('/'))|Where-Object{$_})-join'/'}
    $path=if($folder){"$folder/$name"}else{$name};$script:ResolvedLogPath=$path
    $bytes=[Text.Encoding]::UTF8.GetBytes(($LogLines-join"`r`n"))
    Invoke-MgGraphRequest -Method PUT -Uri "v1.0/drives/$SharePointDriveId/root:/$path`:/content" -Body $bytes -ContentType 'text/plain; charset=utf-8' -ErrorAction Stop | Out-Null
}

# -----------------------------------------------------------------------------
# HAUPTABLAUF
# -----------------------------------------------------------------------------
# Der Hauptablauf arbeitet zuerst mit einem vollstaendigen Plan und fuehrt ihn danach phasenweise aus.
$Plans=[System.Collections.Generic.List[object]]::new()
$UsersById=@{};$UsersByUpn=@{};$AssignmentsByNumber=@{};$CurrentNumberByOwner=@{}
try {
    Write-Section '1. START UND KONFIGURATION'
    Test-IncludeConfiguration
    Test-ExcludeConfiguration
    Write-Log "Runbook gestartet. DryRun=$DryRun; Debug=$Debug; IncludeRegeln=$(@($IncludeUserPropertyRules).Count); ExcludeRegeln=$(@($ExcludeUserPropertyRules).Count); Testbenutzer=$($TestUserPrincipalNames.Count); Ignore=$($IgnoredUserPrincipalNames.Count)"
    foreach($m in @('Microsoft.Graph.Authentication','MicrosoftTeams')){$v=(Get-Module $m -ListAvailable|Sort-Object Version -Descending|Select-Object -First 1).Version;Write-Log "Modul ${m}: $v"}

    Write-Section '2. VERBINDUNGEN'
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    if([string]::IsNullOrWhiteSpace($ManagedIdentityClientId)){Connect-MgGraph -Identity -NoWelcome|Out-Null}else{Connect-MgGraph -Identity -ClientId $ManagedIdentityClientId -NoWelcome|Out-Null}
    $GraphConnected=$true;Write-Log 'Mit Microsoft Graph verbunden.' -Level OK
    Import-Module MicrosoftTeams -ErrorAction Stop
    if(-not[string]::IsNullOrWhiteSpace($ManagedIdentityClientId)){throw 'Teams-Verbindung ist aktuell auf die systemseitig zugewiesene Managed Identity ausgelegt.'}
    Connect-MicrosoftTeams -Identity -ErrorAction Stop|Out-Null;$TeamsConnected=$true

    Write-Section '3. DATEN LADEN'
    $select=Get-GraphSelectProperties
    $users=@(Get-GraphCollection -Uri "v1.0/users?`$select=$select&`$top=999")
    $Counters.EntraUsersLoaded=$users.Count
    foreach($u in $users){$UsersById[[string]$u.id]=$u;$UsersByUpn[[string]$u.userPrincipalName]=$u}
    $assignments=@(Get-AllTeamsPhoneAssignments)
    foreach($a in $assignments){$n=Get-AssignmentNumber $a;$o=Get-AssignmentOwnerId $a;if($n){$AssignmentsByNumber[$n]=$a};if($o-and(Get-AssignmentNumberType $a)-eq'DirectRouting'){$CurrentNumberByOwner[$o]=$n}}
    Write-Log "$($users.Count) Benutzer und $($assignments.Count) Teams-Zuweisungen geladen." -Level OK

    Write-Section '4. VALIDIERUNG UND AKTIONSPLAN'
    # Nummer -> Liste der Benutzer mit dieser Sollnummer; Grundlage fuer die Duplikaterkennung.
    $desired=@{}
    foreach($u in $users){
        $upn=[string]$u.userPrincipalName
        if($IgnoredUserPrincipalNames.Contains($upn)){$Counters.IgnoredUsers++;continue}
        if([string]$u.userType-eq'Guest'){$Counters.GuestsExcluded++;continue}
        if(-not[bool]$u.accountEnabled){$Counters.DisabledExcluded++;continue}
        if(-not(Test-UserIncluded $u)){$Counters.IncludeExcluded++;continue}
        if(Test-UserExcluded $u){$Counters.ExcludeExcluded++;continue}
        if(@($u.assignedLicenses).Count-eq0){$Counters.UnlicensedExcluded++;continue}
        $state=Get-BusinessPhoneState $u.businessPhones
        $currentNumber=if($CurrentNumberByOwner.ContainsKey([string]$u.id)){$CurrentNumberByOwner[[string]$u.id]}else{$null}
        $p=[pscustomobject]@{User=$u;Upn=$upn;State=$state.State;DesiredNumber=$state.Number;CurrentNumber=$currentNumber;Eligible=$true;InScope=(Test-TestScopeUser $upn);Duplicate=$false;Blocked=$false;Action='None';Error=$state.Error}
        switch($state.State){
            'Empty'{$Counters.EmptyBusinessPhones++;if($currentNumber-and(Test-AllowedPhoneNumber $currentNumber)){$p.Action='Remove'}}
            'Multiple'{$p.Eligible=$false;$p.Action='Error';$p.Error="Mehrere businessPhones-Werte: $($state.Values-join' | ')"}
            'Invalid'{$Counters.InvalidNumbers++;$p.Eligible=$false;$p.Action='Error'}
            'Blocked'{
                $Counters.BlockedNumbers++
                $p.Eligible=$false
                $p.Blocked=$true
                $p.Action='Error'
                $p.Error="Die Telefonnummer $($state.Number) darf nicht zugewiesen werden: $($state.Error)"
                $BlockedPhoneConflicts.Add([pscustomobject]@{
                    Number=[string]$state.Number
                    Reason=[string]$state.Error
                    DisplayName=[string]$u.displayName
                    UserPrincipalName=[string]$upn
                })
            }
            'Valid'{if(-not(Test-TeamsPhonePlan $u)){$Counters.LicenseErrors++;$p.Eligible=$false;$p.Action='Error';$p.Error='Kein aktiver Teams-Phone-Serviceplan vorhanden.'}else{if(-not$desired.ContainsKey($state.Number)){$desired[$state.Number]=[System.Collections.Generic.List[object]]::new()};$desired[$state.Number].Add($p);$p.Action=if($currentNumber-eq$state.Number){'Ensure'}else{'Assign'}}}
        }
        $Plans.Add($p)
    }
    # Doppelte Sollnummern werden fuer alle betroffenen Benutzer blockiert und separat gemeldet.
    foreach($number in @($desired.Keys)){
        if($desired[$number].Count-gt1){
            $Counters.DuplicateDesiredNumbers++
            $conflictUsers=@($desired[$number] | ForEach-Object {
                [pscustomobject]@{ DisplayName=[string]$_.User.displayName; UserPrincipalName=[string]$_.Upn }
            })
            $DuplicatePhoneConflicts.Add([pscustomobject]@{ Number=[string]$number; Users=$conflictUsers })
            foreach($p in $desired[$number]){
                $p.Duplicate=$true;$p.Eligible=$false;$p.Action='Error'
                $p.Error="Doppelte Sollnummer $number bei: $($desired[$number].Upn-join', ')"
            }
        }
    }
    foreach($p in $Plans|Where-Object{$_.Action-eq'Error'}){
        Write-UserSection $p.User 'VALIDIERUNGSFEHLER'
        Write-Log "$($p.Upn): $($p.Error)" -Level ERROR
        if($p.Duplicate){$script:DuplicateConflictErrorCount++}
        if($p.Blocked){$script:BlockedConflictErrorCount++}
    }

    Write-Section '5. LEERE SOLLNUMMERN FREIGEBEN'
    # Freigaben erfolgen vor Neuzuweisungen, damit Nummern innerhalb desselben Laufs verfuegbar werden.
    foreach($p in $Plans|Where-Object{$_.Action-eq'Remove'}){
        if(-not$p.InScope){Write-Log "Testbegrenzung: Entfernung bei $($p.Upn) blockiert." -Level WARN;continue}
        try{Write-UserSection $p.User 'ENTFERNUNG';Remove-ManagedPhone -Identity $p.Upn -OwnerId $p.User.id -Number $p.CurrentNumber -ResetPolicies -Reason 'businessPhones ist leer'}catch{Write-Log "$($p.Upn): Entfernung fehlgeschlagen: $($_.Exception.Message)" -Level ERROR}
    }

Write-Section '6. ZUWEISUNGEN UND KONFLIKTE'

# Zaehlerstand zu Beginn der Phase merken, damit die Anzahl der in dieser
# Phase geplanten oder versuchten Korrekturen ausgewiesen werden kann.
$phase6ChangeCountBefore = $ChangeLines.Count

# Pro Benutzer: Basiswerte korrigieren, Konflikte pruefen, alte Nummer
# entfernen und Sollnummer setzen.
#
# Reine ENSURE-Pruefungen werden nicht mehr als eigener Benutzerblock
# protokolliert. Notwendige Korrekturen bleiben ueber Add-Change sichtbar.
foreach($p in $Plans | Where-Object {
    $_.Eligible -and
    $_.InScope -and
    $_.Action -in @('Assign', 'Ensure')
}){
    $removedHolder = $null
    $removedTargetOld = $null

    try {
        # Benutzerblock nur anzeigen, wenn eine Telefonnummer zugewiesen
        # oder verschoben werden soll.
        if ($p.Action -eq 'Assign') {
            Write-UserSection $p.User 'ASSIGN'
        }

        Update-EntraUsageLocation $p.User
        Set-RequiredPolicies $p.Upn

        # Die Telefonnummer stimmt bereits. Allfaellige Policy- oder
        # UsageLocation-Korrekturen wurden oben durch Add-Change geloggt.
        if ($p.Action -eq 'Ensure') {
            $Counters.AssignmentsAlreadyCorrect++
            continue
        }
            $holder=if($AssignmentsByNumber.ContainsKey($p.DesiredNumber)){$AssignmentsByNumber[$p.DesiredNumber]}else{$null}
            $holderId=Get-AssignmentOwnerId $holder
            if($holderId-and$holderId-ne[string]$p.User.id){
                $Counters.Conflicts++;$protection=Test-ProtectedAssignmentHolder $holder $holderId
                if ($protection) {
                $Counters.ProtectedConflicts++

                $holderUser = if ($UsersById.ContainsKey([string]$holderId)) {
                    $UsersById[[string]$holderId]
                } else {
                    $null
                }

                $holderIdentity = if ($holderUser) {
                    [string]$holderUser.userPrincipalName
                } else {
                    [string]$holderId
                }

                throw (
                    "Konflikt: Die Nummer {0} ist aktuell {1} zugewiesen. " +
                    "Die Verschiebung an {2} wurde blockiert: {3}"
                ) -f (
                    $p.DesiredNumber,
                    $holderIdentity,
                    $p.Upn,
                    $protection
                )
            }
                if((Get-AssignmentNumberType $holder)-ne'DirectRouting'){throw "Konflikt mit nicht verwaltbarem Nummerntyp '$(Get-AssignmentNumberType $holder)'."}
                $holderUser=$UsersById[[string]$holderId];$holderIdentity=[string]$holderUser.userPrincipalName
                Remove-ManagedPhone -Identity $holderIdentity -OwnerId $holderId -Number $p.DesiredNumber -Reason "Verschiebung an $($p.Upn)"
                $removedHolder=[pscustomobject]@{Identity=$holderIdentity;OwnerId=$holderId;Number=$p.DesiredNumber}
            }
            if($p.CurrentNumber-and$p.CurrentNumber-ne$p.DesiredNumber){
                if(-not(Test-AllowedPhoneNumber $p.CurrentNumber)){throw "Bestehende Nummer $($p.CurrentNumber) darf nicht ersetzt werden."}
                Remove-ManagedPhone -Identity $p.Upn -OwnerId $p.User.id -Number $p.CurrentNumber -Reason "Wechsel auf $($p.DesiredNumber)"
                $removedTargetOld=$p.CurrentNumber
            }
            Set-ManagedPhone -Identity $p.Upn -OwnerId $p.User.id -Number $p.DesiredNumber
            if($p.CurrentNumber){$Counters.NumbersChanged++}else{$Counters.NumbersAssigned++}
            Write-Log "$($p.DesiredNumber) erfolgreich an $($p.Upn) zugewiesen." -Level OK
        } catch {
            Write-Log "$($p.Upn): Verarbeitung fehlgeschlagen: $($_.Exception.Message)" -Level ERROR
            if($removedTargetOld-and-not$DryRun){[void](Restore-Number $p.Upn $p.User.id $removedTargetOld 'urspruengliche Zielbenutzernummer')}
            if($removedHolder-and-not$DryRun){[void](Restore-Number $removedHolder.Identity $removedHolder.OwnerId $removedHolder.Number 'urspruenglicher Nummerninhaber')}
        }
    }

    $phase6Changes = $ChangeLines.Count - $phase6ChangeCountBefore

    Write-Log (
        'Telefonnummer bereits korrekt bei {0} Benutzern.' -f
        $Counters.AssignmentsAlreadyCorrect
    ) -Level OK

    if ($phase6Changes -gt 0) {
        Write-Log (
            'In dieser Phase geplante oder versuchte Korrekturen: {0}.' -f
            $phase6Changes
        )
    }
    else {
        Write-Log 'In dieser Phase waren keine Korrekturen erforderlich.' -Level OK
    }

    Write-Section '7. SCHLUSSZUSAMMENFASSUNG'
    foreach($key in $Counters.Keys){Write-Log ('{0,-34}: {1}'-f$key,$Counters[$key])}
    Write-Log ('Warnungen                         : {0}'-f$WarningCount)
    Write-Log ('Fehler                            : {0}'-f$ErrorCount)
    Write-Log ('DryRun                            : {0}'-f$DryRun)
    Write-Log ('Debug                             : {0}'-f$Debug)
    Write-Log ('Aenderungen geplant/versucht      : {0}'-f$ChangeLines.Count)
    Write-Log ('Vorgesehener Logbereich           : {0}'-f$(if($ChangeLines.Count){$SharePointLogFolder}else{"$SharePointLogFolder/$NoActionLogSubfolder"}))
    if($ErrorCount-eq0-and$WarningCount-eq0){Write-Log 'Runbook erfolgreich beendet.' -Level OK}else{Write-Log 'Runbook mit Hinweisen beendet.'}
}
catch{
    $FatalError=$true;$fatalLine=if($_.InvocationInfo.ScriptLineNumber){" (Skriptzeile $($_.InvocationInfo.ScriptLineNumber))"}else{''}
    Write-Log "FATALER ABBRUCH${fatalLine}: $($_.Exception.Message)" -Level ERROR
    if($null-ne$_.ErrorDetails-and-not[string]::IsNullOrWhiteSpace([string]$_.ErrorDetails.Message)){Write-Log "Details: $($_.ErrorDetails.Message)" -Level ERROR}
    throw
}
# Abschlussarbeiten laufen auch nach Fehlern: Log/Mails schreiben und alle Sitzungen trennen.
finally{
    try{Save-LogToSharePoint;if($ResolvedLogPath){Write-Output "SharePoint-Log gespeichert: $ResolvedLogPath. Run-ID: $RunId"}}catch{Write-Warning "SharePoint-Log konnte nicht gespeichert werden: $($_.Exception.Message)"}
    # Doppelte Entra-Sollnummern werden immer separat gemeldet, unabhaengig von Debug und DryRun.
    if($SendPhoneConflictMail-and$DuplicatePhoneConflicts.Count-gt0-and-not[string]::IsNullOrWhiteSpace($PhoneConflictMail)-and$GraphConnected){
        try{Send-DuplicatePhoneConflictMail;Write-Output "Telefonnummernkonflikte an $PhoneConflictMail mit CC an $LogSupportMail versendet. Run-ID: $RunId"}catch{Write-Warning "Konfliktmail konnte nicht versendet werden: $($_.Exception.Message)"}
    }
    # Reservierte Telefonnummern werden an dieselbe Konfliktadresse gemeldet.
    if($SendPhoneConflictMail-and$BlockedPhoneConflicts.Count-gt0-and-not[string]::IsNullOrWhiteSpace($PhoneConflictMail)-and$GraphConnected){
        try{Send-BlockedPhoneConflictMail;Write-Output "Reservierte Telefonnummern an $PhoneConflictMail mit CC an $LogSupportMail versendet. Run-ID: $RunId"}catch{Write-Warning "Mail fuer reservierte Telefonnummern konnte nicht versendet werden: $($_.Exception.Message)"}
    }
    # Doppelte Sollnummern loesen bei aktiviertem Debug zusaetzlich zur
    # Konfliktmail eine Supportmail an LogSupportMail aus. Reservierte Nummern
    # bleiben unabhaengig von Debug normale Validierungsfehler.
    $nonDuplicateErrors=[Math]::Max(0,$ErrorCount-$DuplicateConflictErrorCount)
    $mailForBlockedNumbers=$SendSupportMail-and$Counters.BlockedNumbers-gt0
    $mailForIssues=$SendSupportMail-and($nonDuplicateErrors-gt0-or$mailForBlockedNumbers-or($SendMailOnWarnings-and$WarningCount-gt0)-or$FatalError)
    $mailForDebugChanges=$SendSupportMail-and$Debug-and$ChangeLines.Count-gt0
    $mailForDebugDuplicates=$SendSupportMail-and$Debug-and$DuplicatePhoneConflicts.Count-gt0
    $needMail=-not[string]::IsNullOrWhiteSpace($LogSupportMail)-and($mailForIssues-or$mailForDebugChanges-or$mailForDebugDuplicates)

    Write-Output ("Supportmail-Pruefung: SendSupportMail={0}; GraphConnected={1}; ErrorCount={2}; DuplicateConflictErrorCount={3}; BlockedConflictErrorCount={4}; NonDuplicateErrors={5}; BlockedNumbers={6}; DuplicatePhoneConflicts={7}; WarningCount={8}; FatalError={9}; Debug={10}; ChangeLines={11}; NeedMail={12}" -f $SendSupportMail,$GraphConnected,$ErrorCount,$DuplicateConflictErrorCount,$BlockedConflictErrorCount,$nonDuplicateErrors,$Counters.BlockedNumbers,$DuplicatePhoneConflicts.Count,$WarningCount,$FatalError,$Debug,$ChangeLines.Count,$needMail)

    if($needMail){
        if(-not$GraphConnected){
            Write-Warning "Supportmail ist erforderlich, kann aber nicht versendet werden, weil keine Microsoft-Graph-Verbindung besteht."
        }else{
            try{
                $status=if($FatalError-or$nonDuplicateErrors-gt0-or$Counters.BlockedNumbers-gt0-or($Debug-and$DuplicatePhoneConflicts.Count-gt0)){'FEHLER'}elseif($WarningCount-gt0){'WARNUNG'}elseif($ChangeLines.Count-gt0){if($DryRun){'VORSCHAU'}else{'ÄNDERUNGEN'}}else{'INFO'}
                Send-SupportLogMail $status
                Write-Output "Supportmail an $LogSupportMail versendet. Run-ID: $RunId"
            }catch{
                Write-Warning "Supportmail konnte nicht versendet werden: $($_.Exception.Message)"
                if($null-ne$_.ErrorDetails-and-not[string]::IsNullOrWhiteSpace([string]$_.ErrorDetails.Message)){
                    Write-Warning "Microsoft-Graph-Fehlerdetails: $($_.ErrorDetails.Message)"
                }
            }
        }
    }
    if($TeamsConnected){Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue|Out-Null}
    if($GraphConnected){Disconnect-MgGraph -ErrorAction SilentlyContinue|Out-Null}
}
