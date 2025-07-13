# DocConverter - DOCX zu PDF Konverter

Ein einfaches Java-Programm zur Konvertierung von Microsoft Word-Dokumenten (DOCX) in PDF-Dateien.

## Features

- **DOCX zu PDF Konvertierung**: Konvertiert Word-Dokumente in PDF-Format
- **✅ Vollständige Tabellen-Unterstützung**: Alle Tabellen mit Struktur und Inhalt
- **✅ Tabellen-Farben**: Zellhintergründe und Textfarben werden übertragen
- **✅ Professionelles Layout**: A4-Format mit korrekten Seitenrändern
- **✅ Intelligentes Paging**: Automatische Seitenumbrüche bei Kapiteln
- **✅ Bilder und Grafiken**: Eingebettete PNG/JPG-Bilder werden übertragen  
- **✅ Header und Footer**: Kopf- und Fußzeilen werden konvertiert
- **✅ Textformatierung**: Überschriften, Fettdruck und Ausrichtung
- **✅ Strukturierte Darstellung**: Hierarchische Gliederung und Abstände
- **Einfache Bedienung**: Kommandozeilen-Interface für einfache Verwendung
- **Standalone JAR**: Alle Abhängigkeiten sind im JAR enthalten

## Technologie-Stack

- **Java 21**: Moderne Java-Version
- **Apache POI 5.2.4**: Für das Lesen von DOCX-Dateien
- **iText 7.2.5**: Für die PDF-Generierung
- **Maven**: Build-Management
- **JUnit 5**: Testing Framework

## Installation und Build

### Voraussetzungen
- Java 21 oder höher
- Maven 3.6 oder höher

### Projekt kompilieren
```bash
mvn clean package
```

Dies erstellt eine ausführbare JAR-Datei: `target/docconverter-1.0.0.jar`

## Verwendung

### Kommandozeile
```bash
java -jar target/docconverter-1.0.0.jar <input.docx> <output.pdf>
```

### Beispiel
```bash
# Beispiel-DOCX erstellen
java -cp target/docconverter-1.0.0.jar ch.std.doc.converter.DocxCreator beispiel.docx

# DOCX zu PDF konvertieren
java -jar target/docconverter-1.0.0.jar beispiel.docx beispiel.pdf
```

### Programmatische Verwendung
```java
DocConverter converter = new DocConverter();
converter.convertDocxToPdf("input.docx", "output.pdf");
```

## Projektstruktur

```
docconverter/
├── src/
│   ├── main/
│   │   └── java/
│   │       └── ch/std/doc/converter/
│   │           ├── DocConverter.java      # Hauptklasse
│   │           └── DocxCreator.java       # Utility zum Erstellen von Test-DOCX
│   └── test/
│       └── java/
│           └── ch/std/doc/converter/
│               └── DocConverterTest.java  # Unit Tests
├── pom.xml                               # Maven-Konfiguration
├── README.md                             # Diese Datei
└── USAGE.md                              # Detaillierte Verwendungsanweisungen
```

## Abhängigkeiten

- **commons-io**: File-Utilities
- **Apache POI**: 
  - poi: Core-Bibliothek
  - poi-ooxml: OOXML-Unterstützung für DOCX
  - poi-scratchpad: Erweiterte Features
- **iText 7**:
  - kernel: PDF-Core-Funktionalität
  - layout: PDF-Layout-Engine

## Tests ausführen

```bash
mvn test
```

## Erweiterte Verwendung

### Eigene DOCX-Dateien erstellen
Das Projekt enthält eine `DocxCreator`-Klasse zum Erstellen von Beispiel-DOCX-Dateien:

```bash
java -cp target/docconverter-1.0.0.jar ch.std.doc.converter.DocxCreator mein-dokument.docx
```

### Formatierung
Der Konverter erkennt und konvertiert:
- **✅ Fette Texte** → werden als Überschriften behandelt (größere Schrift)
- **✅ Normale Paragraphen** → werden mit Standardschrift konvertiert
- **✅ Tabellen** → komplette Tabellenstruktur mit allen Zellen
- **✅ Tabellen-Farben** → Zellhintergründe und Textfarben (RGB-Hex-Codes)
- **✅ Bilder** → PNG, JPG und andere Formate werden eingebettet
- **✅ Header/Footer** → Kopf- und Fußzeilen werden übertragen
- **✅ Zeilenumbrüche** → werden beibehalten
- **✅ Textausrichtung** → Links-, Rechts- und Zentrierung

## Limitierungen

- Unterstützt derzeit nur DOCX-Format (nicht DOC)
- ~~Grundlegende Formatierung~~ **→ Jetzt erweiterte Formatierung!**
- ~~Keine Bilder oder komplexe Layouts~~ **→ Bilder und Tabellen werden unterstützt!**
- Hyperlinks werden als normaler Text dargestellt
- Eingebettete Objekte (Word-Charts, Excel-Tabellen) werden nicht konvertiert

## Verbesserungen in Version 1.0

### ✅ **Neu hinzugefügt:**
- **Vollständige Tabellen-Unterstützung** - alle Tabellen werden mit korrekter Struktur konvertiert
- **🎨 Tabellen-Farben** - Zellhintergründe und Textfarben werden korrekt übertragen
- **📄 Professionelles Layout** - A4-Format mit korrekten Seitenrändern (2cm/1cm)
- **📑 Intelligentes Paging** - automatische Seitenumbrüche bei Hauptkapiteln
- **🏗️ Strukturierte Darstellung** - hierarchische Überschriften mit passenden Größen
- **📏 Optimierte Abstände** - professionelle Margins und Padding
- **🖼️ Bilder-Integration** - PNG, JPG mit Rahmen und Bildunterschriften
- **📊 Erweiterte Tabellen** - Borders, Padding und numerische Ausrichtung
- **Header/Footer-Verarbeitung** - Kopf- und Fußzeilen werden übertragen  
- **RGB-Hex-Farbunterstützung** - alle gängigen Farbformate (#4472C4, F2F2F2, etc.)
- **Umfassende Tests** - automatisierte Tests für komplexe Dokumente

## Entwicklung

### Code-Qualität
- Java 21 moderne Syntax
- Umfassende Javadoc-Dokumentation
- Unit Tests mit JUnit 5
- Maven für Dependency Management

### Logging
Das Projekt verwendet SLF4J für Logging. Die Log-Meldungen sind derzeit minimal gehalten.

## Lizenz

Dieses Projekt steht unter der MIT-Lizenz.

## Mitwirkende

Entwickelt mit GitHub Copilot als Coding-Assistent.