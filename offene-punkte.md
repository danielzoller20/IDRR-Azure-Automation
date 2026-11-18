# Offene Punkte – idrr-azure-automation

Was in diesem Repo noch überarbeitet, ergänzt oder entschieden werden muss.

Priorität: **Hoch** = Risiko für Betrieb, Sicherheit oder Datenschutz · **Mittel** = Fehlverhalten oder Lücke im Betrieb · **Niedrig** = Kosmetik, Konsistenz, Doku.

Regel und Format: Handbuch [`konventionen/offene-punkte.md`](https://github.com/danielzoller20/idrr-handbuch/blob/main/konventionen/offene-punkte.md) (Entscheid 0004). Erledigte Punkte werden im selben Pull Request entfernt.

Stand: 18.11.2026

Details der Runbooks: [`docs/betrieb.md`](docs/betrieb.md)

Quelle der Befunde: Review vom 18.11.2026 (statische Prüfung, nichts ausgeführt). Die Skripte sind unverändert übernommen; keiner der Punkte ist umgesetzt.

## Sicherheit und Berechtigungen

| Nr. | Prio | Was | Betrifft | Art der Änderung |
|---|---|---|---|---|
| A1 | Hoch | `Mail.Send` (Application) erlaubt das Senden aus jedem Postfach des Tenants. In Exchange Online per RBAC for Applications bzw. Application Access Policy auf das Absenderpostfach `informatik@` beschränken. | beide Managed Identities | Konfiguration Tenant |
| A2 | Hoch | Entra-Rolle User Administrator gilt tenantweit (Passwort-Reset für alle Nicht-Admins). Nur auf eine Administrative Unit mit den SuS-Konten vergeben. | `StudentPolicies` | Konfiguration Tenant |
| A3 | Hoch | SuS-Passwörter bleiben dauerhaft in Postfächern: Kopie in «Gesendete Elemente» (`$SaveSchoolAccountMailsToSentItems = $true`) und CC an den technischen Support (`$Debug` + `$CcTechnicalSupportInDebug`). Werte anpassen oder Aufbewahrungsrichtlinie für diese Mails festlegen. | `Set-StudentPassword` | Konfiguration Runbook |

## Runbook-Logik

| Nr. | Prio | Was | Betrifft | Art der Änderung |
|---|---|---|---|---|
| A4 | Hoch | Keine Notbremse bei Massenänderungen. Liefert PUPIL leere Werte, werden alle SuS ohne Klasse deaktiviert bzw. allen Mitarbeitenden die Nummer entfernt. Schwellwert einbauen (z. B. > 10 Aktionen oder > 5 % des Bestands → Abbruch, nur Supportmail). | beide Runbooks | Code |
| A6 | Mittel | Schlägt der Verarbeitungsstempel nach dem Mailversand fehl, erzeugt der nächste Lauf ein zweites Passwort; das gemailte erste gilt nicht mehr. | `Set-StudentPassword` | Code oder Doku |
| A7 | Mittel | Schlägt die Sammelmail fehl, werden reaktivierte Konten beim nächsten Lauf nicht erneut gemeldet. | `Set-StudentPassword` | Code |
| A8 | Mittel | Nummern ausserhalb des erlaubten Bereichs erzeugen bei jedem Lauf einen ERROR und eine FEHLER-Mail (Daueralarm). Als Hinweis behandeln oder per Exclude-Regel ausschliessen. | `Sync-EntraBusinessPhonesToTeams` | Code oder Konfiguration |
| A9 | Mittel | Kein Retry bei Graph-Throttling (429/503); ein Fehler beim Laden der Benutzer bricht den ganzen Lauf ab. | beide Runbooks | Code |
| A10 | Niedrig | Include-Regex `@sus.srb\.sg$`: erster Punkt nicht maskiert, korrekt `@sus\.srb\.sg$`. | `Set-StudentPassword` | Konfiguration Runbook |
| A11 | Niedrig | Fehlermeldung in `Test-Configuration` nennt `LogSupportMail` statt `$TechnicalSupportRecipient`. | `Set-StudentPassword` | Code |
| A12 | Niedrig | Bereits deaktivierte Konten mit leerer Zuordnung werden in keinem Zähler erfasst. | `Set-StudentPassword` | Code |
| A13 | Niedrig | Mail «deaktivierte Konten» nutzt `$SaveTechnicalMailsToSentItems` statt einer eigenen Einstellung. | `Set-StudentPassword` | Code |
| A14 | Niedrig | Zähler zählen im DryRun hoch (im SuS-Runbook nicht). Vereinheitlichen. | `Sync-EntraBusinessPhonesToTeams` | Code |
| A15 | Niedrig | Kein `#requires -Modules`, MicrosoftTeams-Version nicht festgehalten (Doku nennt 7.9.0). `saveToSentItems` fest auf `$false` statt konfigurierbar. | `Sync-EntraBusinessPhonesToTeams` | Code |

## Klären

| Nr. | Prio | Was |
|---|---|---|
| A16 | Mittel | Fachliche Mailempfänger sind persönliche Postfächer. Bei Abwesenheit oder Austritt gehen Passwörter und Fehlermeldungen verloren. Funktionspostfächer oder Verteilergruppen einsetzen? |
| A18 | Mittel | Beide Runbooks laufen produktiv mit `$Debug = $true`. Gewollter Normalzustand? |
| A19 | Niedrig | Initialpasswort (Muster `ULLDDDDLLU`, rund 39 Bit) gilt dauerhaft (`ForceChangePasswordNextSignIn = $false`, kein Ablauf). Bewusst so lassen? Entscheid festhalten. |
| A20 | Niedrig | Kommentar in `Sync-EntraBusinessPhonesToTeams` nennt eine Person mit Namen. Durch Funktion ersetzen. |

## Dokumentation und Betrieb

| Nr. | Prio | Was | Wo |
|---|---|---|---|
| A21 | Hoch | Aufbewahrung der Logs in SharePoint festlegen (enthalten Namen, UPNs, Klassen von SuS). | README, [datenschutz.md](https://github.com/danielzoller20/idrr-handbuch/blob/main/konventionen/datenschutz.md) |
| A22 | Mittel | Deployment: Runbooks werden manuell ins Portal kopiert. Repo als einzige Quelle sichern (GitHub Actions mit OIDC oder fester Ablauf). | README |
| A23 | Mittel | Zuständige Person und Stellvertretung eintragen. | README, [tools.md](https://github.com/danielzoller20/idrr-handbuch/blob/main/uebersicht/tools.md) |
| A24 | Niedrig | Subscription und Region der Automation Accounts ergänzen. | README, `docs/betrieb.md`, [architektur.md](https://github.com/danielzoller20/idrr-handbuch/blob/main/uebersicht/architektur.md) |
| A25 | Niedrig | Freigabe der Betriebsdoku 1.1 eintragen. | `docs/betrieb.md` |
| A26 | Niedrig | `CHANGELOG.md` anlegen, sobald versioniert wird. | Repo |
