# calcal

## Lokale Web-App starten

1. Im Projektordner `npm start` ausführen.
2. Im Browser `http://localhost:4173` öffnen.
3. Die App speichert Termine lokal im Browser und funktioniert offline nach dem ersten Laden weiter.

## Funktionen

1. Kalender 2026 von Montag bis Freitag.
2. Zeitslots von 08:00 bis 12:00 in 10-Minuten-Schritten.
3. Terminformular mit Name, KG, EZ, Telefonnummer und Anliegen.
4. Tagesliste für das Sekretariat mit Datum, Uhrzeit und Name.
5. Archiv mit Name, Telefonnummer und vergangener Terminzeit.
6. Pastell-Ampelfarben pro Tag ohne Blau.
7. Druckansicht für die Tagesliste im Hochformat über den Browser.

## Dateien

1. `index.html` enthält die Oberfläche.
2. `styles.css` enthält das Layout und die Pastellfarben.
3. `app.js` enthält Kalenderlogik, lokale Speicherung und Druckansicht.

## Windows Desktop (Electron)

### Varianten

1. Portable EXE (ohne Installation)
2. Per-User Installer (ohne Admin, wenn Richtlinien es erlauben)

### Lokaler Build auf Windows

1. `npm install`
2. `build-windows-release.cmd` starten
3. Fertige Dateien liegen im Ordner `dist/`

### GitHub Build (Download ohne lokalen Build)

1. In GitHub den Workflow `Build Windows Release` starten
2. Nach Abschluss das Artifact `calcal-windows-build` herunterladen
3. Darin sind die EXE-Dateien:
	- `CalCal-Portable-...exe`
	- `CalCal-Setup-...exe`