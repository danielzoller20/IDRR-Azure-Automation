# CLAUDE.md – IDRR Azure Automation

> Ergänzt die gemeinsamen Regeln aus `~/schulverbund/CLAUDE.md` (Handbuch: `claude/CLAUDE.shared.md`).

## Was dieses Tool tut

Azure-Automation-Runbooks (PowerShell 7.6), die nach dem PUPIL-Sync SuS-Konten (`Set-StudentPassword`) und die Teams-Telefonie (`Sync-EntraBusinessPhonesToTeams`) in Microsoft 365 nachführen.

## Technik

- Sprache: PowerShell 7.6 (UTF-8)
- Hosting: Azure Automation, Ressourcengruppe `UserAutomation`, Accounts `StudentPolicies` und `TeamsCalling`
- Abhängigkeiten: `Microsoft.Graph.Authentication`, `MicrosoftTeams`; Anmeldung per Managed Identity

## Befehle

```powershell
# Syntaxprüfung (lokal, ohne Ausführung)
Get-ChildItem runbooks -Recurse -Filter *.ps1 | ForEach-Object {
  $e = $null; $null = [System.Management.Automation.Language.Parser]::ParseFile($_.FullName, [ref]$null, [ref]$e); if ($e) { $_.FullName; $e }
}
# Lint (falls installiert)
Invoke-ScriptAnalyzer -Path runbooks -Recurse
```

Tests: keine. Die Runbooks werden in Azure mit `$DryRun = $true` getestet.

## Struktur

```
runbooks/
  StudentPolicies/Set-StudentPassword.ps1
  TeamsCalling/Sync-EntraBusinessPhonesToTeams.ps1
docs/betrieb.md        Konfiguration, Mitteilungen, Logging, Berechtigungen, Fehlerbehebung
offene-punkte.md       Was noch überarbeitet werden muss (Nummern A1 …)
```

## Regeln für dieses Repo

- Runbooks nie gegen den produktiven Tenant ausführen. Änderungen nur im Code, Ausführung macht ein Mensch.
- Ändert sich ein Wert im Konfigurationsblock, `docs/betrieb.md` und ggf. `README.md` im selben Commit nachführen.
- Graph-Berechtigungen und Entra-Rollen nicht erweitern ohne Rückfrage.
- `offene-punkte.md` vor grösseren Änderungen lesen; erledigte Punkte im selben Commit entfernen, neue Befunde eintragen.
- Passwörter dürfen nie in Logs oder technische Mails gelangen.
- Neue Einträge in Ignore-Listen oder Mailempfänger nur auf ausdrückliche Anweisung.
- Betrifft eine Änderung Übergreifendes (PUPIL-Attribute, dynamische Gruppen, Zuständigkeiten), Handbuch `idrr-handbuch` (`uebersicht/pupil-m365.md`) anpassen.
