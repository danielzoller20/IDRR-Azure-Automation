# Betriebsdokumentation Runbooks

Konfiguration, Mitteilungen, Logging, Berechtigungen und Fehlerbehebung der beiden Runbooks.
Datenfluss PUPIL → Microsoft 365, Attribute und dynamische Gruppen: Handbuch [`uebersicht/pupil-m365.md`](https://github.com/danielzoller20/idrr-handbuch/blob/main/uebersicht/pupil-m365.md).

| Version | Datum | Änderung | Autor/Freigabe |
|---|---|---|---|
| 1.0 | 15.10.2026 | Erster dokumentierter Stand (Word) | Daniel Zoller |
| 1.1 | 18.11.2026 | Nach Markdown übernommen, mit Skriptstand abgeglichen | TODO Freigabe |

> Massgebend sind die Werte im Skript. Ändert sich ein Wert im Konfigurationsblock, wird diese Seite im selben Pull Request nachgeführt.

## Azure-Struktur

| Typ | Name | Aufgabe |
|---|---|---|
| Resource Group | `UserAutomation` | Gemeinsame Ressourcengruppe der Automationen |
| Automation Account | `TeamsCalling` | Teams-Direct-Routing-Zuweisungen, Mailversand, Logging |
| Runbook | `Sync-EntraBusinessPhonesToTeams` | Setzt oder entfernt Direct-Routing-Zuweisungen, meldet Konflikte, loggt |
| Automation Account | `StudentPolicies` | Passwort- und Kontorichtlinien SuS, Mailversand, Logging |
| Runbook | `Set-StudentPassword` | Setzt Passwort und Policies, meldet per Mail, loggt |

Subscription und Region: TODO.

## Set-StudentPassword (Automation StudentPolicies)

Läuft täglich um 06:20. Verarbeitet Konten der SuS-Domain: stellt Konto- und Passwortvorgaben sicher, aktiviert oder deaktiviert Konten gemäss PUPIL-Zuordnung und informiert den Support.

### Verarbeitungslogik

| Situation | Aktion |
|---|---|
| UPN in Ignore-Liste, Gastkonto, keine Include-Regel passt, ausserhalb Testbegrenzung | nichts (nur gezählt) |
| Konto lokal synchronisiert (`onPremisesSyncEnabled`) | nichts, gesammelte Warnung im Log |
| `department` enthält «extern» | Konto deaktivieren, Mail «deaktivierte Konten» |
| `department` leer und `$DisableUsersWithEmptySchoolAssignment = $true` | Konto deaktivieren, Mail «deaktivierte Konten» |
| `department` ohne bekannte Schuleinheit | nichts, gesammelte Warnung im Log |
| Schuleinheit erkannt, Konto deaktiviert | Konto aktivieren; ohne neues Passwort in Sammelmail mit «bestehend gem. Pupil» |
| Schuleinheit erkannt, Password-Policy fehlt | `DisablePasswordExpiration` ergänzen |
| Schuleinheit Steig, `ageGroup` ≠ `notAdult` | `ageGroup` setzen |
| Schuleinheit erkannt, `extensionAttribute15` leer | Passwort setzen, Sammelmail, nach Versand Stempel setzen |

Aktivierung, Policy und `ageGroup` werden bei **jedem** Lauf geprüft, nicht nur bei leerem `extensionAttribute15`. Nur das Passwort hängt am Stempel.

### Wesentliche Konfiguration

| Variable | Aktueller Wert | Wirkung |
|---|---|---|
| `$IncludeUserPropertyRules` | UPN endet auf `@sus.srb.sg` | Mehrere Regeln OR-verknüpft. Leere Liste: Lauf bricht ab, niemand wird verarbeitet. |
| `$TestUserPrincipalNames` | leer | Optional auf einzelne UPNs einschränken. |
| `$IgnoredUserPrincipalNames` | leer | Schützt definierte Konten vollständig. |
| `$SchoolAssignmentProperty` | `Department` | Attribut mit der Klassenbezeichnung aus PUPIL. |
| `$SchoolUnitPatterns` | enthält `klosterguet`, `steig`, `wildenstein` | Bestimmt Mailempfänger und `ageGroup`. Erste passende Einheit gewinnt. |
| `$ExternalUserPattern` | enthält `extern` | SuS mit «extern» in der Klasse werden deaktiviert (Austritte an weiterführende oder externe Schulen). |
| `$DisableUsersWithEmptySchoolAssignment` | `$true` | Konten ohne Klassenzuordnung werden deaktiviert. Wird das Semester in PUPIL vor dem 31.8. umgestellt, sind ehemalige SuS der 3. Oberstufe ohne Klasse. |
| `$PasswordPattern` | `ULLDDDDLLU` | U = Grossbuchstabe, L = Kleinbuchstabe, D = Ziffer (ohne verwechselbare Zeichen). Zehn Zeichen. |
| `$ForceChangePasswordNextSignIn` | `$false` | SuS müssen das Passwort nicht ändern. |
| `$RequiredPasswordPolicy` | `DisablePasswordExpiration` | Passwort läuft nicht ab. |
| `$SteigAgeGroup` | `notAdult` | `ageGroup` für SuS der Schuleinheit Steig. |
| `$ProcessedAttribute` | `extensionAttribute15` | Stempel `PROCESSED\|Zeitpunkt\|Run-ID`, gesetzt nach erfolgreichem Versand der Passwortmail. |
| `$IncludeAlreadyDisabledUsersInDisabledAccountMail` | `$false` | Bereits deaktivierte Konten werden nicht erneut gemeldet. |
| `$SaveSchoolAccountMailsToSentItems` | `$true` | Sammelmails (mit Passwörtern) bleiben in «Gesendete Elemente» des Absenders. |
| `$ManagedIdentityClientId` | gesetzt | Benutzerseitig zugewiesene Managed Identity. |
| `$DryRun` | `$false` | `$true` verhindert alle schreibenden Aktionen und fachlichen Mails. |
| `$Debug` | `$true` | Verhindert nichts. Technische Änderungsmail bei Änderungen; mit `$CcTechnicalSupportInDebug` geht jede fachliche Mail in CC an den technischen Support. |

### Mitteilungen

| Mitteilung | Empfänger (Variable) | Inhalt / Auslöser |
|---|---|---|
| Neue oder reaktivierte SuS-Accounts | 1st-Level-Support der Schuleinheit (`$SchoolSupportRecipients`) | Eine Sammelmail je Schuleinheit: Name, Klasse, Benutzername, bei neuen Konten das Passwort, bei reaktivierten «bestehend gem. Pupil». Bei `Debug` CC an technischen Support. |
| Deaktivierte Konten | 2nd-Level-Support oder Schulverwaltung (`$DisabledAccountNotificationRecipient`) | Sammelmail für Konten, die als extern oder ohne Zuordnung deaktiviert wurden. |
| Technische Statusmail | 2nd-Level-Support (`$TechnicalSupportRecipient`) | Bei Fehlern, optional bei Warnungen, bei `Debug` und Änderungen. Zusammenfassung, Änderungen, Log als Anhang. |
| Kritischer Zustellfehler | 2nd-Level-Support (`$TechnicalSupportRecipient`) | Passwort gesetzt, fachliche Mail nicht zugestellt. Kein Stempel; nächster Lauf erzeugt ein neues Passwort. |

**Passwortschutz:** Passwörter stehen nur in den fachlichen Sammelmails (inkl. Debug-CC und Kopie in «Gesendete Elemente» des Absenders), nie im Log oder in technischen Mails.

### Logging

- SharePoint-Team «Accounts» (`$SharePointDriveId`), Ordner `Runbook-Logs/SuS-Accounts`.
- Läufe ohne Änderung im Unterordner `NoAction`.
- Dateiname `SuS_YYYY-MM-DD_HHMMSS_RunID.log`.
- Jede Zeile: lokaler Zeitstempel, Run-ID, Level `INFO`, `OK`, `WARN` oder `ERROR`.
- Inhalt: Konfiguration, Verbindungen, Zähler, Änderungen, Ausschlussgründe, Warnungen, Fehler, Laufdauer.

## Sync-EntraBusinessPhonesToTeams (Automation TeamsCalling)

Läuft täglich um 06:10. `businessPhones` in Entra ist die führende Quelle. Gleicht mit Teams Direct Routing ab und stellt Teams-Richtlinien und `usageLocation` sicher.

### Wesentliche Konfiguration

| Variable | Aktueller Wert | Wirkung |
|---|---|---|
| `$IncludeUserPropertyRules` | UPN endet auf `@schule-rorschacherberg.ch` | Mehrere Regeln OR-verknüpft. Leere Liste: niemand. |
| `$ExcludeUserPropertyRules` | `OfficeLocation` enthält Tageshort, Mehrzweckhalle oder Schulbus; `businessPhones` enthält `71 858` (Schulverwaltung) | Mehrere Regeln OR-verknüpft. Bildungskommission bleibt bewusst eingeschlossen (voraussichtlich bis mindestens 1.1.2027). |
| `$IgnoredUserPrincipalNames` | 2 Einträge: Konto der Call Queue; Konto Standort Bildungskommission mit externer Nummer | Schützt Konten und ihre bestehenden Nummernzuweisungen. |
| `$AllowedPhoneNumberPatterns` | `+41 58 100 81xx` und `+41 58 100 82xx` | Nur diese Nummern werden verwaltet. Andere Nummern in `businessPhones` erzeugen bei jedem Lauf einen Fehler. |
| `$BlockedPhoneNumbers` | `+41 58 100 81 66` bis `…69` | Reserviert für die Evakuationsanlagen (Klosterguet, Wildenstein, Steig Real, Steig Sek). |
| `$TeamsPhoneServicePlanIds` | `4828c8ec-dc2e-4779-b502-87ac9ce28ab7` | Aktiver Teams-Phone-Serviceplan nötig. |
| `$TeamsCallingPolicy` | `AllowCalling` | |
| `$VoiceRoutingPolicy` | `Peoplefone` | |
| `$UsageLocation` | `CH` | |
| `$ManagedIdentityClientId` | leer | Systemseitig zugewiesene Managed Identity (Teams-Anmeldung unterstützt nur diese). |
| `$DryRun` | `$false` | `$true` protokolliert nur den Aktionsplan. |
| `$Debug` | `$true` | Verhindert nichts; steuert die Änderungsmail. |

### Verhalten in Sonderfällen

- `businessPhones` leer → bestehende, verwaltete Direct-Routing-Nummer wird entfernt, Calling- und Voice-Routing-Policy auf Vererbung zurückgesetzt.
- Deaktivierte oder unlizenzierte Benutzer werden übersprungen; eine bestehende Nummer bleibt bis dahin zugewiesen. Konten austretender Mitarbeitender werden gelöscht, damit wird die Nummer frei.
- Nummer ist bei einem anderen Benutzer zugewiesen → Verschiebung, ausser der bisherige Inhaber ist geschützt (Ignore-Liste, Exclude-Regel, keine Include-Regel, kein Benutzer, z. B. Resource Account). Dann Fehler, keine Änderung.
- Nach jeder Zuweisung oder Entfernung wird der Teams-Status bis zu 4 × im Abstand von 10 s geprüft. Schlägt eine Verschiebung fehl, wird die vorherige Zuweisung wenn möglich wiederhergestellt.

### Mitteilungen

| Mitteilung | Empfänger (Variable) | Inhalt / Auslöser |
|---|---|---|
| Support- und Änderungsmail | 2nd-Level-Support (`$LogSupportMail`) | Bei Fehlern, reservierten Nummern, optional bei Warnungen, bei `Debug` und Änderungen. Neue Vergaben, Änderungen, Löschungen, Hinweise; Log im Anhang. |
| Doppelte Telefonnummer | Schulverwaltung (`$PhoneConflictMail`, eigene Domain `rorschacherberg.ch`, gewollt), CC `$LogSupportMail` | Separate Konfliktmail. Keine Änderung für betroffene Benutzer. Korrektur in PUPIL. |
| Reservierte Telefonnummer | Schulverwaltung (`$PhoneConflictMail`), CC `$LogSupportMail` | Separate Konfliktmail. Keine Änderung für betroffene Benutzer. Korrektur in PUPIL. |

### Logging

- SharePoint-Team «Accounts» (`$SharePointDriveId`), Ordner `Runbook-Logs/TeamsCalling`, ohne Änderungen im Unterordner `NoAction`.
- Dateiname `TeamsCalling_YYYY-MM-DD_HHMMSS_RunID.log`.
- Inhalt: Validierung, geplante Aktionen, ausgeführte Änderungen, Verifikation, Konflikte, Rollback-Versuche, Schlusszähler.

## Berechtigungen und Rollen

Soll-Stand gemäss Runbooks. Im Tenant verifiziert am 14.10.2026.

| Automation Account | Graph (Application) | Entra-Rolle | Weitere |
|---|---|---|---|
| `StudentPolicies` | `User.ReadWrite.All`, `Mail.Send`, `Sites.Selected` | User Administrator | SharePoint: Schreibrecht auf die Site «Accounts» |
| `TeamsCalling` | `User.ReadWrite.All` (`usageLocation`), `Mail.Send`, `Sites.Selected` | Teams Telephony Administrator | SharePoint: Schreibrecht auf die Site «Accounts»; App-basierte Teams-Anmeldung (`Connect-MicrosoftTeams -Identity`) |

## Fehlerbehebung

| Beobachtung | Prüfung / Massnahme |
|---|---|
| SuS fehlt in Gruppen | PUPIL-Klasse und Standort prüfen, dann `department` in Entra und die dynamische Regel. Datum der Klassenzuordnung prüfen: PUPIL schreibt erst ab Eintrittsdatum. |
| Neues SuS-Konto erhält kein Passwort | `extensionAttribute15`, `department`, Kontoart, lokale Synchronisation, Mailzustellung und technisches Log prüfen. |
| Mitarbeitende fehlen in Standortgruppe | Standort in PUPIL und `officeLocation` / `physicalDeliveryOfficeName` in Entra prüfen. |
| Teams-Nummer wird nicht gesetzt | `businessPhones`, erlaubten Nummernbereich, Exclude-Regeln, Kontostatus, Lizenz und Teams-Phone-Serviceplan prüfen. |
| Doppelte oder reservierte Nummer | In PUPIL korrigieren. Der nächste Lauf prüft erneut. |
| Runbook meldet keine Änderung | `NoAction`-Log prüfen. Bei korrektem Zielzustand normal. |
| SharePoint-Upload schlägt fehl | `Sites.Selected`, Site-Berechtigung, Drive-ID und Zielpfad prüfen. |
| Graph- oder Teams-Anmeldung schlägt fehl | Managed Identity, Modulversion, Graph-App-Rechte und Entra-Rollenzuweisung prüfen. |

## Betriebskontrolle

- Nach Änderungen an PUPIL-Attributmapping, Membership Rules oder Runbook-Konfiguration zuerst mit `$DryRun = $true` testen.
- Konflikt- und Fehlermails zeitnah bearbeiten; die Quelldaten werden meist in PUPIL korrigiert.
- SharePoint-Logs regelmässig auf Warnungen und wiederkehrende Fehler prüfen.
- Berechtigungen der Managed Identities mindestens jährlich und nach Funktionsänderungen nach Least Privilege prüfen.
- Änderungen an Nummernbereichen, reservierten Nummern, Supportrollen und Schuleinheiten in Skript und Dokumentation gemeinsam nachführen.
