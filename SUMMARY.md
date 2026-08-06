# Code-Zusammenfassung: Urlaubsplaner

> Stand der Analyse: 6. August 2026, Branch `main`
>
> Repository: `MoinsenWerner/urlaubsplaner`

## 1. Zweck der Anwendung

Die Anwendung ist ein webbasierter Urlaubs-, Abwesenheits- und Desksharing-Planer. Sie verwaltet Benutzer mit mehreren Rollen, zeigt Jahresmatrizen für Abwesenheiten und Desksharing an, importiert bestehende Excel-/Word-Daten, exportiert neue Excel-Übersichten und versendet Initialzugänge per selbst zugestellter, DKIM-signierter E-Mail.

Die gesamte Backend-Logik befindet sich derzeit in der einzelnen Flask-Datei `app.py`. Die Benutzeroberfläche besteht aus Jinja2-Templates sowie drei statischen Dateien für Gestaltung, Zellenauswahl und Zeilensortierung.

## 2. Technischer Aufbau

- **Backend:** Python, Flask, Flask-Login
- **Datenbank:** SQLite über direktes `sqlite3`, ohne ORM
- **Frontend:** serverseitige Jinja2-Templates, CSS und Vanilla JavaScript
- **Excel:** `openpyxl`
- **Word:** `python-docx`; für alte `.doc`-Dateien optional das Systemprogramm `antiword`
- **Feiertage:** Paket `holidays`, Bundesland Bayern
- **E-Mail:** direkte SMTP-Zustellung an die MX-Server des Empfängers, DNS-Abfragen, STARTTLS und DKIM
- **Verschlüsselung:** Fernet-Schlüssel wird deterministisch aus `SECRET_KEY` abgeleitet
- **Tests:** `pytest`, aktuell zentral in `tests/test_app.py`

## 3. Start und Laufzeit

Direktstart:

```bash
python3 app.py
```

Die Anwendung startet dann auf:

```text
0.0.0.0:4010
```

Am Dateiende ist `debug=True` gesetzt. Das ist für einen öffentlich erreichbaren Produktivbetrieb ungeeignet und muss dort deaktiviert werden.

Beim Direktstart werden ausgeführt:

1. `init_db()`
2. `start_mail_server()`
3. Flask-Entwicklungsserver

Wichtig: Der Mail-Worker wird nur im `if __name__ == "__main__"`-Block gestartet. Wird die Flask-App über Gunicorn, uWSGI oder einen anderen WSGI-Server importiert, läuft die Mailwarteschlange nicht automatisch. Dann muss der Worker separat oder über eine Application-Factory gestartet werden.

Vor jedem Request wird `init_db()` erneut aufgerufen. Die Funktion ist überwiegend idempotent, verursacht aber wiederholte Schema- und Seed-Prüfungen und ist kein Ersatz für ein richtiges Migrationssystem.

## 4. Konfiguration und Laufzeitdateien

Die wichtigsten Umgebungsvariablen stehen in `.env.example`:

| Variable | Bedeutung |
|---|---|
| `SECRET_KEY` | Flask-Sitzungen und Ableitung des Fernet-Schlüssels |
| `URLAUBSPLANER_DB` | Optionaler Pfad zur SQLite-Datenbank; Standard `instance/urlaubsplaner.sqlite3` |
| `PUBLIC_BASE_URL` | Basis-URL für Initialpasswort-Links |
| `INITIAL_ADMIN_PASSWORD` | Optionales Passwort für den ersten Admin |
| `MAIL_FROM` | Absenderadresse |
| `MAIL_HOSTNAME` | SMTP-HELO/EHLO-Hostname |
| `DKIM_SELECTOR` | DKIM-Selector |
| `DKIM_PRIVATE_KEY` | Pfad zum privaten DKIM-Schlüssel |
| `MAIL_WORKER_INTERVAL` | Wartezeit des Mail-Workers in Sekunden |
| `MAIL_MAX_ATTEMPTS` | Maximale Zustellversuche |
| `MAIL_IPV4_ONLY` | Erzwingt IPv4 für die direkte SMTP-Zustellung |

Folgende Verzeichnisse werden zur Laufzeit erzeugt und sind über `.gitignore` ausgeschlossen:

- `instance/`
- `instance/uploads/`
- `instance/profile-images/`
- standardmäßig `instance/urlaubsplaner.sqlite3`
- standardmäßig `instance/dkim_private.pem`

`SECRET_KEY` muss dauerhaft stabil sein. Eine Änderung macht bestehende Sitzungen ungültig und verhindert außerdem das Entschlüsseln noch nicht bestätigter Initialpasswörter und bereits erzeugter Initial-Links.

## 5. Datenmodell

### `users`

Enthält Benutzerkonto und Profildaten:

- Benutzername, Vorname, Nachname
- Passwort-Hash
- Pflicht zum Ändern des Initialpassworts
- verschlüsseltes Initialpasswort
- Bestätigungsstatus der Initialdatenanzeige
- E-Mail-Adresse
- Geburtsdatum
- Dateiname des Profilbilds

Ein Teil dieser Spalten wird durch manuelle `ALTER TABLE`-Migrationen in `init_db()` ergänzt.

### `user_roles`

Mehrfachrollen pro Benutzer. Mögliche Rollen in Prioritätsreihenfolge:

1. `admin`
2. `ausbilder`
3. `putzchef`
4. `azubi`
5. `normal`
6. `desksharing`

Hat ein Benutzer keine gültige ausgewählte Rolle, wird automatisch `normal` verwendet.

### `entries`

Urlaubs- und Abwesenheitseinträge:

- Benutzer
- ISO-Datum
- internes Matrixkürzel
- Ersteller

Die Eindeutigkeit gilt für `(user_id, entry_date, code)`. Technisch können dadurch mehrere unterschiedliche Codes am selben Tag existieren. Die Bulk-Bearbeitung löscht dagegen vor dem Einfügen alle Einträge der Zelle. Einzel- und Bulk-Bearbeitung haben damit nicht exakt dieselbe Semantik.

### `desksharing_entries`

Ein Desksharing-Status pro Benutzer und Datum. Durch den Unique-Constraint `(user_id, entry_date)` wird ein vorhandener Status überschrieben.

### `user_matrix_order`

Speichert die vom Admin festgelegte Zeilenreihenfolge getrennt für:

- `vacation`
- `desksharing`

Beim Speichern muss immer die vollständige, duplikatfreie Liste aller aktuell sichtbaren Benutzer übergeben werden.

### `entry_mappings`

Zentrale Definition jedes Abwesenheitstyps:

- Importkürzel
- Matrixkürzel und Matrixfarbe
- Schaltflächenbezeichnung
- Exportkürzel, Exportfarbe und Exportbeschreibung

### `entry_mapping_roles`

Rollenabhängige Sichtbarkeit und optionale Überschreibungen für:

- Matrixkürzel
- Matrixfarbe
- Exportkürzel
- Exportfarbe

Die Priorität richtet sich nach `ROLE_ORDER`. Bei mehreren Rollen gewinnt die erste Rolle mit einem passenden Override.

### `app_metadata`

Speichert derzeit insbesondere, ob die Standard-Mappings bereits initial angelegt wurden.

### `mail_outbox`

Persistente E-Mail-Warteschlange mit:

- Empfänger
- vollständiger signierter Nachricht
- Status `pending`, `sending`, `delivered` oder `failed`
- Anzahl der Versuche
- nächster Versuch
- letzter Fehler
- Erstellungszeit

## 6. Rollen und Berechtigungen

### Admin

- vollständige Bearbeitung aller Benutzer und Kalendereinträge
- Erstellen und Löschen von Benutzern
- Rollen und E-Mail-Adressen aller Benutzer bearbeiten
- Initialpasswörter zurücksetzen und anzeigen
- Import von Excel-/Word-Dateien
- Zuordnungen/Mappings verwalten
- Matrixreihenfolgen speichern
- Desksharing-Einträge bearbeiten

### Ausbilder

- Mitgliederseite öffnen
- Initialpasswörter ausschließlich für Benutzer mit Rolle `azubi` zurücksetzen/anzeigen
- erweiterter Zeitraum in der Urlaubsübersicht
- eigene Einträge gemäß Mapping-Berechtigung

### Putzchef, Azubi, Normal, Desksharing

Diese Rollen erhalten unterschiedliche, über die Mapping-Tabelle konfigurierbare Eintragstypen. Standardmäßig gelten unter anderem:

- `normal`/`desksharing`: geplanter und genehmigter Urlaub
- `azubi`: zusätzlich Berufsschule
- `putzchef`: zusätzlich Weihnachtsputz
- `ausbilder`: zusätzliche Ausbildungsarten

Berechtigungen sollten nicht nur in der UI geändert werden. Entscheidend sind `entry_mapping_roles`, `allowed_codes()` und `may_edit()` im Backend.

## 7. Anmeldung und Initialzugänge

Beim ersten Start ohne Benutzer wird ein Admin erstellt. Das Passwort stammt entweder aus `INITIAL_ADMIN_PASSWORD` oder wird zufällig erzeugt und im Serverlog ausgegeben.

Neue Benutzer erhalten:

1. ein zufälliges achtstelliges Initialpasswort,
2. einen Passwort-Hash in der Datenbank,
3. zusätzlich eine Fernet-verschlüsselte Version zur einmaligen Anzeige,
4. `must_change_password = 1`,
5. einen verschlüsselten Link zur Seite `/initial-login`,
6. eine E-Mail in der persistenten Mailwarteschlange.

Bis der Administrator/Ausbilder die Anzeige der neu erstellten Initialdaten bestätigt, wird dessen Sitzung auf der Mitgliederseite festgehalten. Der neue Benutzer muss beim ersten Login ein Passwort mit mindestens acht Zeichen setzen.

Die Initial-Links besitzen keine eigene Ablaufzeit. Ihre Gültigkeit endet praktisch erst, wenn das Initialpasswort geändert oder zurückgesetzt wurde. Die Fehlermeldung bezeichnet ungültige Links dennoch als „ungültig oder abgelaufen“.

## 8. Urlaubs- und Abwesenheitsmatrix

Die Hauptseite `/` zeigt ausschließlich Montag bis Freitag.

Zeitraum:

- Admin und Ausbilder: 2. Juli des Vorjahres bis 2. Juli des Folgejahres
- andere Benutzer: 1. Januar bis 31. Dezember des gewählten Jahres

Zusätzliche Markierungen:

- gesetzliche Feiertage in Bayern über das Paket `holidays`
- bayerische Schulferien aus einer fest im Code hinterlegten Tabelle

Die Schulferien sind nur für 2025, 2026 und 2027 hinterlegt. Ab 2028 liefert die Funktion keine Ferienmarkierungen mehr, bis die Tabelle erweitert oder durch eine externe/algorithmische Datenquelle ersetzt wird.

### Krank-Einträge (`KR`)

Die tatsächliche Krankheitsfarbe ist nur sichtbar für:

- Admins,
- den betroffenen Benutzer selbst,
- Ausbilder bei betroffenen Azubis.

Andere Benutzer sehen den Eintrag neutral als „Abwesend“. Änderungen an Farben oder Codes müssen diese Sonderlogik berücksichtigen.

### Einzel- und Mehrfachauswahl

- `/entry` fügt oder löscht einen bestimmten Code.
- `/bulk-entry` bearbeitet eine rechteckige Mehrfachauswahl.
- Beim Bulk-Einfügen werden vorher sämtliche Einträge der Zelle gelöscht.
- Beim Bulk-Löschen werden nur Einträge entfernt, die der aktuelle Benutzer entfernen darf.

`static/matrix-selection.js` implementiert Rechteckauswahl per Maus und Touch. `static/matrix-order.js` implementiert den Admin-Sortiermodus.

## 9. Dynamische Eintragstypen

Die Seite `/entry-mappings` erlaubt Admins, Eintragstypen zu erstellen und zu bearbeiten. Ein Mapping verbindet Import, Matrixdarstellung, Rollen und Export.

Wichtige Regeln:

- Importkürzel und Matrixkürzel müssen jeweils eindeutig sein.
- Farben müssen sechsstellige RGB-Hexwerte sein.
- Beim Umbenennen eines Matrixkürzels werden vorhandene `entries.code`-Werte mit geändert.
- Beim Löschen eines Mappings bleiben vorhandene Einträge mit dem alten Code bestehen.
- Solche verwaisten Einträge werden anschließend weder normal dargestellt noch exportiert, solange kein neues Mapping mit diesem Code existiert.

Vor dem Löschen eines produktiv verwendeten Mappings muss daher entschieden werden, ob die vorhandenen Einträge migriert oder ebenfalls gelöscht werden sollen.

## 10. Desksharing

Die Seite `/desksharing` zeigt ein Kalenderjahr mit den Statuswerten:

- `Anwesend`
- `Homeoffice`
- `Abwesend`

Nur Benutzer mit Rolle `desksharing` erscheinen in dieser Matrix. Bearbeiten darf derzeit ausschließlich ein Admin. Die Zeilenreihenfolge wird getrennt von der Urlaubsmatrix gespeichert.

## 11. Excel-Import

Administratoren können `.xlsx` und `.xlsm` hochladen.

Der Import:

- sucht automatisch die Zeile mit den meisten erkennbaren Datumswerten,
- kann Datumswerte zusätzlich aus den speziellen Zeilen `Monat` und `Ferienzeit` rekonstruieren,
- liest Kürzel direkt aus Zellen,
- kann leere Zellen anhand ihrer Füllfarbe und einer Legende zuordnen,
- ignoriert bekannte Überschriften und Legendenbezeichnungen,
- erzeugt unbekannte Benutzer automatisch,
- erzeugt für automatisch angelegte Benutzer Initialzugänge und Mailwarteschlangen-Einträge,
- importiert nur Codes, für die ein Mapping existiert.

Benutzernamen werden aus Vor- und Nachname als kleingeschriebene Punktnotation gebildet. Änderungen an Namenslogik oder Importformat können Duplikate oder neue Benutzer erzeugen.

`.xlsm` wird mit `data_only=True` gelesen. VBA-Makros werden nicht ausgeführt und beim späteren Export nicht mitgegeben.

## 12. Excel-Export

`POST /download` erstellt eine neue `.xlsx`-Datei für einen gewählten Jahresbereich.

Enthalten sind:

- Werktage,
- Ferien- und Feiertagszeile,
- Benutzer mit mindestens einem Eintrag im gewählten Zeitraum,
- rollenabhängige Exportkürzel und Farben,
- Legende.

Nicht enthalten sind:

- Profilbilder,
- Makros/VBA,
- Benutzer ohne Eintrag,
- unbekannte oder verwaiste Codes ohne Mapping.

Wenn mehrere Codes für dieselbe Zelle existieren, wird nur der zuerst gefundene exportiert. Die Datenbankabfrage besitzt dafür keine explizite Sortierung; die Ausgabe kann bei solchen Mehrfacheinträgen daher unerwartet sein.

## 13. Word-Import für Desksharing

Unterstützt werden:

- `.docx` über `python-docx`,
- alte `.doc`-Dateien bevorzugt über `antiword`,
- als letzter Fallback eine einfache UTF-8-/UTF-16-Textauswertung.

Erkannt werden sowohl Tabellen als auch Textformen wie Kalenderwoche, Wochentag, Status und Namenslisten.

Die Zuordnung erfolgt ausschließlich über den Vornamen. Existiert derselbe Vorname bei mehreren Desksharing-Benutzern, wird er absichtlich nicht eindeutig zugeordnet und somit nicht importiert.

## 14. E-Mail-Zustellung

Die Anwendung verwendet keinen konfigurierten SMTP-Relay. Sie ermittelt die MX-Server der Empfängerdomain und versucht die Nachricht direkt über Port 25 zuzustellen.

Ablauf:

1. MIME-Nachricht mit Text- und HTML-Teil erzeugen.
2. Persistenten 2048-Bit-RSA-DKIM-Schlüssel erzeugen oder laden.
3. Nachricht mit DKIM signieren.
4. Nachricht in `mail_outbox` speichern.
5. Worker ermittelt MX- und optional A-Records.
6. STARTTLS wird verwendet, falls der Zielserver es anbietet.
7. Temporäre Fehler erhalten exponentielles Backoff.
8. Permanente 5xx-Ablehnungen werden nicht erneut versucht.

Für zuverlässige Zustellung müssen insbesondere stimmen:

- öffentlicher PTR-/Reverse-DNS-Eintrag der Absender-IP,
- SPF,
- veröffentlichter DKIM-TXT-Record,
- erreichbarer Port 25,
- korrekter HELO-Hostname,
- keine dynamische oder geblockte Absender-IP.

Der private DKIM-Schlüssel und die SQLite-Datenbank müssen gemeinsam gesichert werden.

## 15. Templates und statische Dateien

### Templates

- `base.html`: gemeinsames Layout und Navigation
- `login.html`: normale Anmeldung
- `initial-login.html`: Initialpasswort ändern
- `profile.html`: E-Mail, Geburtsdatum und Profilbild
- `change-password.html`: Passwortänderung
- `index.html`: Urlaubsmatrix
- `desksharing.html`: Desksharing-Matrix
- `members.html`: Benutzer-, Rollen- und Initialdatenverwaltung
- `entry-mappings.html`: Mapping-Verwaltung
- `upload.html`: Excel-/Word-Upload
- `years.html`: verfügbare Export-/Jahresansicht

### Statische Dateien

- `static/style.css`: Layout, responsive Matrix, Sticky-Bereiche und mobile Darstellung
- `static/matrix-selection.js`: rechteckige Zellenauswahl
- `static/matrix-order.js`: Drag-/Sortiermodus und Speichern der Reihenfolge

## 16. Tests

`tests/test_app.py` deckt unter anderem ab:

- Admin-Initialisierung und Login
- Mehrfachrollen
- rollenbasierte Bearbeitung
- Bulk-Einträge
- Excel-/XLSM-Import
- Excel-Export ohne VBA und Bilder
- Initial-Links und Passwort-Reset
- E-Mail-Adressen und Profilbilder
- DKIM-Signatur und direkte SMTP-Zustellung
- permanente SMTP-Fehler
- Word-/DOCX-Desksharing-Import
- Matrixreihenfolgen
- dynamische Mapping-Erstellung, Overrides, Umbenennung und Löschung
- Import der mitgelieferten Beispiel-Excel-Datei

Ausführen:

```bash
python3 -m pip install -r requirements.txt
pytest -q
```

Bei Änderungen an Rollen, Mapping, Import, Export, Mail oder Datenbank müssen die vorhandenen Tests angepasst und erweitert werden. In dieser Analyse wurden die Tests nicht ausgeführt; die Zusammenfassung basiert auf dem aktuellen Repository-Code.

## 17. Besonders wichtige Risiken und Seiteneffekte

1. **Produktivbetrieb mit `debug=True`:** deaktivieren.
2. **Unsicherer Standard-Secret-Key:** ohne gesetzte Variable wird `dev-change-me` verwendet.
3. **Kein CSRF-Schutz:** zustandsändernde HTML-Formulare und JSON-Endpunkte besitzen keine CSRF-Tokens.
4. **Mail-Worker bei WSGI:** startet beim Import der App nicht automatisch.
5. **`init_db()` bei jedem Request:** Schemaänderungen gehören in versionierte Migrationen.
6. **Schulferien enden 2027:** Datenquelle rechtzeitig aktualisieren.
7. **Mapping-Löschung erzeugt verwaiste Einträge.**
8. **Mehrere Codes pro Tag möglich:** Einzel-, Bulk-, Anzeige- und Exportlogik sind dabei nicht vollständig einheitlich.
9. **Initialpasswörter entschlüsselbar gespeichert:** notwendig für die aktuelle Anzeige, aber besonders schützenswert.
10. **Dateiupload nur über Endung geprüft:** keine Inhalts-, Größen- oder Malwareprüfung.
11. **Direkte SMTP-Zustellung:** betrieblich anspruchsvoll und von DNS-/Providerkonfiguration abhängig.
12. **Hardcodierte Firmenadresse:** `default_email()` verwendet immer `@m-a-i.de`.

## 18. Vorgehen bei zukünftigen Änderungen

### Vor jeder Änderung

1. SQLite-Datenbank sichern.
2. `instance/profile-images/`, Uploads und DKIM-Schlüssel sichern.
3. Aktuellen Branch und Commit dokumentieren.
4. Tests ausführen und Ausgangszustand festhalten.

### Bei Datenbankänderungen

- Keine destruktiven Änderungen direkt in `init_db()` einbauen.
- Versionierte Migrationen einführen, zum Beispiel Alembic oder ein eigenes Schema-Versionssystem.
- Fremdschlüssel in SQLite explizit mit `PRAGMA foreign_keys = ON` aktivieren; die Deklarationen allein erzwingen sonst nicht zuverlässig alle Löschregeln.
- Migration zuerst an einer Kopie der Produktivdatenbank testen.

### Bei Rollen- oder Mappingänderungen

Mindestens prüfen:

- Backend-Berechtigung über `allowed_codes()` und `may_edit()`
- sichtbare Buttons
- Rollen-Overrides
- bestehende Einträge
- Importkürzel
- Exportkürzel und Farben
- Krankheits-Datenschutz
- Tests für Benutzer mit mehreren Rollen

### Bei Import-/Exportänderungen

- Beispiel-Excel-Datei beibehalten und als Regressionstest nutzen.
- leere farbige Zellen, echte Datumszellen und rekonstruierte Datumsüberschriften testen.
- keine Makroerhaltung versprechen, solange `openpyxl` nicht mit einem geeigneten `keep_vba`-Workflow verwendet wird.
- Mehrfacheinträge pro Zelle eindeutig behandeln.

### Bei Deploymentänderungen

- Flask-Debugserver durch einen WSGI-Server ersetzen.
- Mail-Worker als separaten Prozess oder eindeutig genau einmal pro Deployment starten.
- Reverse Proxy, HTTPS, Uploadlimits und sichere Session-Cookies konfigurieren.
- regelmäßige Datenbank- und DKIM-Backups einrichten.
- Healthcheck und strukturierte Logs ergänzen.

## 19. Sinnvolle zukünftige Aufteilung

Die aktuelle Monolith-Datei sollte schrittweise getrennt werden:

```text
app/
├── __init__.py              # Application-Factory
├── config.py
├── db.py                    # Verbindung und Migrationen
├── auth.py                  # Login, Initialzugänge, Profil
├── members.py               # Benutzer- und Rollenverwaltung
├── vacation.py              # Urlaubsmatrix und Einträge
├── desksharing.py
├── mappings.py
├── import_export/
│   ├── excel.py
│   └── word.py
├── mail/
│   ├── messages.py
│   ├── dkim.py
│   └── worker.py
├── templates/
└── static/
```

Die Aufteilung sollte funktionserhaltend und mit Tests nach jedem Schritt erfolgen.

## 20. Kurzreferenz der wichtigsten Routen

| Route | Funktion | Zugriff |
|---|---|---|
| `/login` | Anmeldung | öffentlich |
| `/initial-login` | Initialpasswort setzen | öffentlich mit Zugangsdaten/Token |
| `/` | Urlaubsmatrix | angemeldet |
| `/entry` | Einzeleintrag ändern | angemeldet, rollenabhängig |
| `/bulk-entry` | Zellbereich ändern | angemeldet, rollenabhängig |
| `/profile` | Profil verwalten | angemeldet |
| `/members` | Mitgliederverwaltung | Admin/Ausbilder, Aktionen teilweise nur Admin |
| `/entry-mappings` | Eintragstypen verwalten | Admin |
| `/desksharing` | Desksharing anzeigen | angemeldet |
| `/desksharing/bulk-entry` | Desksharing bearbeiten | Admin |
| `/matrix-order/<matrix_name>` | Zeilenreihenfolge speichern | Admin |
| `/upload` | Excel/Word importieren | Admin |
| `/download` | Excel exportieren | angemeldet |

---

Diese Datei beschreibt den aktuellen Ist-Zustand. Sie sollte bei jeder Änderung an Architektur, Datenbank, Rollen, Routen, Import-/Exportformaten oder Deployment aktualisiert werden.
