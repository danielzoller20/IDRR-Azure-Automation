# IDRR Azure Automation

Azure-Automation-Runbooks (PowerShell 7.6), die nach dem täglichen PUPIL-Sync SuS-Konten und die Teams-Telefonie in Microsoft 365 nachführen.

> Übergreifende Dokumentation (Datenfluss PUPIL → Microsoft 365, Attribute, dynamische Gruppen, Zuständigkeiten) steht im Handbuch: [`idrr-handbuch` → uebersicht/pupil-m365.md](https://github.com/danielzoller20/idrr-handbuch/blob/main/uebersicht/pupil-m365.md).

## Überblick

| | |
|---|---|
| Zuständig | 2nd- oder 3rd-Level-Support (Person: TODO) |
| Stellvertretung | TODO |
| Status | aktiv |
| Genutzt von | Schule Rorschacherberg, Schuleinheiten Klosterguet, Steig, Wildenstein (siehe [Handbuch: schulen.md](https://github.com/danielzoller20/idrr-handbuch/blob/main/uebersicht/schulen.md)) |
| Hosting | Azure Automation, Ressourcengruppe `UserAutomation` (Subscription, Region: TODO) |

## Runbooks

| Runbook | Automation Account | Zeitplan | Zweck | Quelle |
|---|---|---|---|---|
| `Sync-EntraBusinessPhonesToTeams` | `TeamsCalling` | täglich 06:10 | Gleicht Entra `businessPhones` mit Teams Direct Routing ab, setzt Calling-/Voice-Routing-Policy und `usageLocation` | [runbooks/TeamsCalling/Sync-EntraBusinessPhonesToTeams.ps1](runbooks/TeamsCalling/Sync-EntraBusinessPhonesToTeams.ps1) |
| `Set-StudentPassword` | `StudentPolicies` | täglich 06:20 | Setzt Passwort und Passwort-Policy für neue SuS-Konten, aktiviert/deaktiviert Konten gemäss Klassenzuordnung, meldet an Support | [runbooks/StudentPolicies/Set-StudentPassword.ps1](runbooks/StudentPolicies/Set-StudentPassword.ps1) |

Beide Runbooks laufen nach dem PUPIL-Sync (ca. 05:00). Korrekturen an Klasse, Standort und Telefonnummer erfolgen deshalb in PUPIL, nicht in Entra.

## Funktionsweise

Kurzfassung. Details, Konfigurationswerte, Mitteilungen und Logging: [docs/betrieb.md](docs/betrieb.md).

**Set-StudentPassword**

1. Alle Entra-Benutzer laden, eingrenzen über Include-Regeln (UPN der SuS-Domain), Ignore-Liste, Testbegrenzung. Gäste und lokal synchronisierte Konten werden ausgelassen.
2. `department` enthält «extern» oder ist leer → Konto deaktivieren, Sammelmail «deaktivierte Konten».
3. Schuleinheit aus `department` erkennen → Konto aktivieren, `DisablePasswordExpiration` sicherstellen, bei Steig `ageGroup = notAdult`.
4. `extensionAttribute15` leer → zufälliges Passwort setzen, Sammelmail je Schuleinheit, erst danach Stempel `PROCESSED|Zeitpunkt|Run-ID`.

**Sync-EntraBusinessPhonesToTeams**

1. Entra-Benutzer und alle Teams-Nummernzuweisungen laden.
2. Validieren (Gäste, deaktivierte, unlizenzierte, ausgeschlossene Benutzer, ungültige, reservierte und doppelte Nummern) und Aktionsplan erstellen.
3. Leere `businessPhones` → verwaltete Nummer entfernen, Policies auf Vererbung.
4. Nummern zuweisen oder verschieben, Policies und `usageLocation` sicherstellen, Ergebnis in Teams verifizieren, bei Fehlern Rollback versuchen.

Beide: `DryRun` verhindert schreibende Aktionen; `Debug` verhindert nichts, sondern steuert nur zusätzliche Mails.

## Daten und Datenschutz

- Bearbeitete Personendaten: Name, UPN, Klasse (`department`), Standort (`officeLocation`), geschäftliche Telefonnummer, Kontostatus. SuS-Initialpasswörter (nur in den fachlichen Sammelmails, nie im Log).
- Herkunft: PUPIL → Microsoft 365 (Entra ID). Ziel: Entra ID, Microsoft Teams, Mails an Support und Schulverwaltung.
- Speicherort der Logs: SharePoint-Team «Accounts», Ordner `Runbook-Logs/…`. Aufbewahrung: TODO.

## Konfiguration

Die Konfiguration steht als Variablenblock am Anfang jedes Runbooks (Abschnitt `KONFIGURATION`). Übersicht und aktuelle Werte: [docs/betrieb.md](docs/betrieb.md).

Secrets: keine. Beide Runbooks melden sich per Managed Identity an (`StudentPolicies`: benutzerseitig zugewiesen, `TeamsCalling`: systemseitig zugewiesen). SharePoint-Drive-ID und Client-ID der Managed Identity sind Kennungen, keine Secrets.

## Lokale Entwicklung

Die Runbooks sind für Azure Automation (Managed Identity) geschrieben und laufen lokal nicht ohne Anpassung der Anmeldung.

Syntaxprüfung lokal (PowerShell 7):

```powershell
$null = [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path .\runbooks\StudentPolicies\Set-StudentPassword.ps1), [ref]$null, [ref]$errors); $errors
```

## Deployment

Aktuell manuell: Skript im Azure-Portal ins Runbook kopieren (Runtime PowerShell 7.6, UTF-8), zuerst mit `$DryRun = $true` testen, dann veröffentlichen.

TODO: Deployment per GitHub Actions (OIDC) prüfen, damit das Repo die einzige Quelle bleibt.

Benötigte Module im Automation Account (vorab importiert):

| Account | Module |
|---|---|
| `StudentPolicies` | `Microsoft.Graph.Authentication` |
| `TeamsCalling` | `Microsoft.Graph.Authentication`, `MicrosoftTeams` (getestet mit 7.9.0) |

## Betrieb

- Logs: SharePoint-Team «Accounts», `Runbook-Logs/SuS-Accounts` und `Runbook-Logs/TeamsCalling` (Läufe ohne Änderung im Unterordner `NoAction`).
- Häufige Fehler und Lösungen: [docs/betrieb.md#fehlerbehebung](docs/betrieb.md#fehlerbehebung)
- Notfall: siehe Handbuch [`ablaeufe/notfall.md`](https://github.com/danielzoller20/idrr-handbuch/blob/main/ablaeufe/notfall.md)

## Changelog

Siehe Git-Historie und GitHub Releases. TODO: `CHANGELOG.md` anlegen, sobald versioniert wird.
