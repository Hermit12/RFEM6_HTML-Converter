# RFEM HTML zu DOCX Konverter

## Übersicht
Dieses Programm konvertiert RFEM HTML-Dateien in DOCX-Format (Microsoft Word). Es kann sowohl neue Word-Dokumente erstellen als auch HTML-Inhalte in bestehende Word-Dokumente ab einer bestimmten Seite einfügen.


## Funktionen
- Konvertierung von HTML-Dateien in DOCX-Format
- Einfügen von HTML-Inhalten in bestehende Word-Dokumente
- Beibehaltung der Formatierung des vorhandenen Dokuments
- Anpassbare Schriftarten und -größen
- Beibehaltung von Tabellenfarben
- Einbindung von Bildern
- Automatische Anpassung der Tabellenbreite
- Öffnen des erstellten Dokuments nach der Konvertierung

## Installationsanleitung für Windows 11

### Voraussetzungen
- Windows 11
- Internetverbindung

### Schritt 1: Python installieren
1. Besuchen Sie [python.org](https://www.python.org/downloads/)
2. Klicken Sie auf die große Schaltfläche "Download Python 3.x.x" (neueste Version)
3. Führen Sie die heruntergeladene Datei aus
4. **WICHTIG:** Aktivieren Sie das Kontrollkästchen "Add Python to PATH" während der Installation
5. Klicken Sie auf "Install Now"

### Schritt 2: Konverter-Programm herunterladen
1. Laden Sie diese Dateien herunter:
   - `rfem6_html_converter.py` (Hauptprogramm)
   - `start_converter.bat` (Startdatei für einfache Benutzung)
2. Speichern Sie beide Dateien im selben Ordner auf Ihrem Computer (z.B. auf dem Desktop)

### Schritt 3: Das Programm starten

#### Empfohlene Methode: Startdatei verwenden
1. Doppelklicken Sie auf die Datei `start_converter.bat`
2. Die Startdatei wird automatisch:
   - Prüfen, ob Python installiert ist
   - Alle erforderlichen Pakete installieren
   - Das Konverter-Programm starten

#### Alternative Methode 1: Direkt starten
1. Navigieren Sie zum Speicherort der `rfem6_html_converter.py`
2. Doppelklicken Sie auf die Datei
3. Wenn Windows fragt, mit welchem Programm Sie die Datei öffnen möchten, wählen Sie Python

#### Alternative Methode 2: Über die Kommandozeile
1. Öffnen Sie die Windows-Eingabeaufforderung (drücken Sie `Win + R`, tippen Sie `cmd` ein und drücken Sie Enter)
2. Navigieren Sie zum Verzeichnis, in dem Sie die Dateien gespeichert haben
   ```
   cd C:\Users\IhrName\Desktop
   ```
3. Installieren Sie die erforderlichen Pakete
   ```
   pip install beautifulsoup4 requests pillow cairosvg python-docx
   ```
4. Starten Sie das Programm
   ```
   python rfem6_html_converter.py
   ```

## Programmbedienung

### HTML-Datei konvertieren
1. Klicken Sie auf "Durchsuchen..." neben "HTML-Datei" und wählen Sie Ihre RFEM-HTML-Datei aus
2. Standardmäßig wird das Ausgabeverzeichnis auf den Ordner der HTML-Datei gesetzt
3. Passen Sie die Formatierungsoptionen an (Schriftart, Schriftgröße, etc.)
4. Klicken Sie auf "Konvertieren"
5. Nach der Konvertierung werden Sie gefragt, ob Sie das erstellte Dokument öffnen möchten

### In bestehendes Word-Dokument einfügen
1. Aktivieren Sie das Kontrollkästchen "In bestehendes Word-Dokument einfügen"
2. Klicken Sie auf "Durchsuchen..." neben "DOCX-Datei" und wählen Sie Ihr bestehendes Word-Dokument aus
3. Geben Sie die Seitennummer ein, ab der der HTML-Inhalt eingefügt werden soll
4. Die vorhandene Formatierung im Dokument bleibt erhalten
5. Klicken Sie auf "Konvertieren"

### Formatierungsoptionen
- **Schriftart**: Wählen Sie die Schriftart für den Text (Standard: Calibri)
- **Schriftgröße**: Wählen Sie die Schriftgröße für normalen Text (Standard: 11)
- **Tabellenfarben beibehalten**: Behält die Hintergrundfarben der Tabellen bei
- **Bilder einbinden**: Fügt Bilder aus der HTML-Datei in das Word-Dokument ein
- **Tabellenbreite anpassen**: Passt die Breite der Tabellen automatisch an
- **Schriftgröße Tabellen**: Legt die Schriftgröße für Text in Tabellen fest
- **Max. PNG-Bildbreite**: Maximale Breite für PNG-Bilder in Zentimetern
- **Breite 2. Spalte**: Bestimmt die Breite der zweiten Spalte in Tabellen in Zentimetern

## Tipps zur Fehlerbehebung

### Programm startet nicht
- Stellen Sie sicher, dass Python korrekt installiert ist
- Öffnen Sie die Eingabeaufforderung und geben Sie `python --version` ein, um zu prüfen, ob Python erkannt wird
- Versuchen Sie, das Programm mit der `start_converter.bat` zu starten, die eine detailliertere Fehlermeldung anzeigen wird
- Achten Sie auf Fehlermeldungen im Kommandozeilenfenster, das von der .bat-Datei geöffnet wird

### Fehlermeldung beim Installieren der Pakete
- Versuchen Sie, die Pakete einzeln zu installieren:
  ```
  pip install beautifulsoup4
  pip install requests
  pip install pillow
  pip install cairosvg
  pip install python-docx
  ```
- Wenn bei `cairosvg` ein Fehler auftritt, können Sie versuchen, es über conda zu installieren oder diese Funktionalität zu deaktivieren (SVG-Bilder werden dann nicht konvertiert)

### Startdatei (.bat) funktioniert nicht
- Rechtsklicken Sie auf die `start_converter.bat` und wählen Sie "Ausführen als Administrator"
- Stellen Sie sicher, dass sich die .bat-Datei im selben Verzeichnis wie die `rfem6_html_converter.py` befindet
- Überprüfen Sie, ob Windows-Skripte auf Ihrem System blockiert sind:
  1. Öffnen Sie die Windows-Sicherheitseinstellungen
  2. Gehen Sie zu "App & Browsersteuerung" > "SmartScreen für Microsoft Store-Apps"
  3. Deaktivieren Sie temporär die Überprüfung oder fügen Sie eine Ausnahme hinzu

### Probleme bei der Konvertierung
- Stellen Sie sicher, dass die HTML-Datei ein RFEM-Format hat
- Prüfen Sie, ob Sie Schreibrechte im Ausgabeverzeichnis haben
- Bei Problemen mit großen Dokumenten, versuchen Sie, kleinere Teile zu konvertieren

### Bilder werden nicht angezeigt
- Stellen Sie sicher, dass die Bilder im gleichen Verzeichnis wie die HTML-Datei oder in einem Unterverzeichnis liegen
- Prüfen Sie, ob die Bilder in der HTML-Datei korrekt referenziert sind

## Unterstützung
Bei Fragen oder Problemen wenden Sie sich bitte an die IT-Abteilung oder den Entwickler des Programms.
