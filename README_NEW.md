# DocConverter - Universal Office zu PDF Konverter

Ein umfassendes Java-Tool zur Konvertierung von Office-Dokumenten in PDF-Format.

## Features

### DOCX zu PDF Konvertierung
- **✅ Rich-Text-Formatierung**: Fett, kursiv, Farben, Unterstreichung
- **✅ Tabellen mit Farben**: Zellhintergründe und Textformatierung
- **✅ Bilder und Grafiken**: Inline-Bilder mit intelligentem Layout
- **✅ Headers und Footers**: Kopf- und Fußzeilen
- **✅ Automatische Seitennummerierung**: "Seite X von Y" Format
- **✅ Seitenumbrüche**: Intelligente Layout-Optimierung
- **✅ A4-Layout**: Professionelle Seitenränder

### XLSX zu PDF Konvertierung
- **✅ Multi-Sheet-Support**: Alle Arbeitsblätter in einem PDF
- **✅ Tabellenformatierung**: Farben und Zellformatierung
- **✅ Zahlenformatierung**: Währungen, Dezimalstellen
- **✅ Header-Hervorhebung**: Automatische Kopfzeilen-Formatierung
- **✅ Querformat**: Optimiert für Excel-Darstellung

## Unterstützte Formate

| Eingabe | Ausgabe | Status |
|---------|---------|--------|
| .docx   | .pdf    | ✅ Vollständig unterstützt |
| .xlsx   | .pdf    | ✅ Vollständig unterstützt |
| .pptx   | .pdf    | 🔄 Geplant |
| .doc    | .pdf    | 🔄 Geplant |
| .xls    | .pdf    | 🔄 Geplant |

## Technologie-Stack

- **Java 21** mit Maven
- **Apache POI 5.2.4** für Office-Dokument-Verarbeitung
- **iText 7.2.5** für PDF-Generierung
- **JUnit 5** für umfassende Tests

## Installation & Build

```bash
# Repository klonen
git clone <repository-url>
cd docconverter

# Projekt kompilieren
mvn clean compile

# Tests ausführen (6 Tests)
mvn test

# JAR-File erstellen
mvn package
```

## Verwendung

### Kommandozeile

```bash
# DOCX zu PDF
java -jar target/docconverter-1.0.0.jar dokument.docx ausgabe.pdf

# XLSX zu PDF  
java -jar target/docconverter-1.0.0.jar tabelle.xlsx ausgabe.pdf
```

### Als Maven-Exec

```bash
# DOCX konvertieren
mvn exec:java -Dexec.mainClass="ch.std.doc.converter.DocConverter" -Dexec.args="dokument.docx ausgabe.pdf"

# XLSX konvertieren
mvn exec:java -Dexec.mainClass="ch.std.doc.converter.DocConverter" -Dexec.args="tabelle.xlsx ausgabe.pdf"
```

### Programmatische Verwendung

```java
DocConverter converter = new DocConverter();

// DOCX konvertieren
converter.convertDocxToPdf("eingabe.docx", "ausgabe.pdf");

// XLSX konvertieren  
converter.convertXlsxToPdf("tabelle.xlsx", "ausgabe.pdf");
```

## Test-Dateien erstellen

```bash
# Einfache DOCX-Datei erstellen
mvn exec:java -Dexec.mainClass="ch.std.doc.converter.DocxCreator"

# Komplexe DOCX-Datei mit allen Features erstellen
mvn exec:java -Dexec.mainClass="ch.std.doc.converter.ExtendedDocxCreator"

# Excel-Datei mit Verkaufsdaten erstellen
mvn exec:java -Dexec.mainClass="ch.std.doc.converter.ExcelCreator"
```

## Architektur

```
src/main/java/ch/std/doc/converter/
├── DocConverter.java          # Hauptkonverter-Klasse
├── DocxCreator.java          # Einfache DOCX-Erstellung
├── ExtendedDocxCreator.java  # Komplexe DOCX-Erstellung
└── ExcelCreator.java         # Excel-Datei-Erstellung

src/test/java/ch/std/doc/converter/
└── DocConverterTest.java     # Umfassende Tests (6 Tests)
```

## Features im Detail

### DOCX-Konvertierung
- **Zweistufiger PDF-Prozess** für korrekte Seitenzahlen
- **Inline-Bild-Verarbeitung** mit kompaktem Layout
- **Intelligente Seitenumbrüche** für bessere Lesbarkeit
- **Farbtreue Konvertierung** von Text und Hintergründen
- **Rich-Text-Support** mit allen Formatierungen

### XLSX-Konvertierung
- **Multi-Sheet-Support** - alle Arbeitsblätter in einem PDF
- **Automatische Spaltenbreiten** basierend auf Inhalt
- **Datentyp-Erkennung** (Text, Zahlen, Datum, Formeln)
- **Währungsformatierung** für Finanzberichte
- **Header-Hervorhebung** für bessere Übersicht

## Beispiel-Ausgaben

Das Tool generiert professionelle PDFs mit:
- Seitennummerierung im Format "Seite X von Y"
- Erhaltung der Original-Formatierung
- Optimierte Layouts für verschiedene Inhaltstypen
- Querformat für Excel-Dateien für bessere Lesbarkeit

## Roadmap

- [ ] PowerPoint (.pptx) Unterstützung
- [ ] Legacy-Formate (.doc, .xls) Unterstützung  
- [ ] Batch-Konvertierung mehrerer Dateien
- [ ] Konfigurierbare Output-Optionen
- [ ] GUI-Interface
- [ ] Cloud-Integration

## Tests

Das Projekt enthält 6 umfassende Tests:
- Einfache DOCX-Konvertierung
- Komplexe DOCX-Konvertierung mit allen Features
- Excel-Konvertierung mit Tabellen und Formatierung
- Legacy-Methoden-Tests
- Unsupported-Format-Handling

```bash
mvn test
# Tests run: 6, Failures: 0, Errors: 0, Skipped: 0
```

## Support

Bei Fragen oder Problemen erstellen Sie bitte ein Issue im Repository.
