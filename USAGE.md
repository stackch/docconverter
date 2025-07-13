# DocConverter - Verwendungsanleitung

## Grundlegende Verwendung

### 1. DOCX zu PDF konvertieren
```bash
java -jar target/docconverter-1.0.0.jar input.docx output.pdf
```

### 2. Einfaches Test-DOCX erstellen
```bash
java -cp target/docconverter-1.0.0.jar ch.std.doc.converter.DocxCreator beispiel.docx
```

### 3. Komplexes Test-DOCX erstellen
```bash
java -cp target/docconverter-1.0.0.jar ch.std.doc.converter.ExtendedDocxCreator komplexes-dokument.docx
```

## Erweiterte Features

### ExtendedDocxCreator - Komplexe Dokumente

Der `ExtendedDocxCreator` erstellt umfangreiche Testdokumente mit:

#### 📄 **Dokumentstruktur (4+ Seiten)**
- **Seite 1**: Titelseite mit Inhaltsverzeichnis
- **Seite 2**: Verschiedene Tabellenformate
- **Seite 3**: Eingebettete Bilder und Grafiken
- **Seite 4**: Zusätzlicher Inhalt und Zusammenfassung

#### 📊 **Tabellen-Features**
- **Mitarbeitertabelle**: 
  - Farbige Header (blau mit weißer Schrift)
  - Abwechselnde Zeilenfärbung
  - Rechtsbündige Zahlenspalten
  
- **Verkaufstabelle**:
  - Grüne Header-Gestaltung
  - Hervorgehobene Summenzeile
  - Quartalsdaten mit Formatierung

#### 🖼️ **Bilder und Grafiken**
- **Automatisch generierte Balkendiagramme**
  - Quartalsumsätze mit farbigen Balken
  - Achsenbeschriftung und Titel
  
- **Geometrische Formen**
  - Kreis, Rechteck und Dreieck
  - Transparente Farbgebung
  - Verschiedene Farben (Blau, Rot, Grün)

#### 📋 **Header und Footer**
- **Header**: Zentrierter Dokumenttitel auf allen Seiten
- **Footer**: Automatische Seitennummerierung ("Seite X von Y")

#### 🎨 **Formatierung**
- Verschiedene Schriftgrößen (10pt - 24pt)
- Fette, kursive und farbige Texte
- Überschriften mit Farbkodierung
- Professionelle Abstände und Ausrichtung

## Beispiele

### Vollständiger Workflow
```bash
# 1. Projekt kompilieren
mvn clean package

# 2. Komplexes DOCX erstellen
java -cp target/docconverter-1.0.0.jar ch.std.doc.converter.ExtendedDocxCreator mein-dokument.docx

# 3. Zu PDF konvertieren
java -jar target/docconverter-1.0.0.jar mein-dokument.docx mein-dokument.pdf

# 4. Ergebnis prüfen
ls -la mein-dokument.*
```

### Batch-Konvertierung
```bash
# Mehrere Dateien konvertieren
for file in *.docx; do
    java -jar target/docconverter-1.0.0.jar "$file" "${file%.docx}.pdf"
done
```

## Unterstützte Features bei der PDF-Konvertierung

### ✅ Vollständig unterstützt
- **Texte**: Alle Schriftgrößen und -stile
- **Tabellen**: Struktur, Zelleninhalt und grundlegende Formatierung
- **Bilder**: PNG, JPG und andere gängige Formate
- **Seitenlayout**: Umbrüche und Seitennummerierung
- **Header/Footer**: Werden in PDF übertragen

### ⚠️ Eingeschränkt unterstützt
- **Tabellen-Formatierung**: Zellfarben werden als Graustufen konvertiert
- **Farbige Texte**: Werden als Fettdruck interpretiert
- **Komplexe Layouts**: Werden vereinfacht dargestellt

### ❌ Nicht unterstützt
- **Hyperlinks**: Werden als normaler Text dargestellt
- **Eingebettete Objekte**: Word-Charts, Excel-Tabellen etc.
- **Animationen**: Nur statischer Inhalt

## Technische Details

### Dateiformate
- **Eingabe**: .docx (Office Open XML)
- **Ausgabe**: .pdf (Portable Document Format)

### Bibliotheken
- **Apache POI**: DOCX-Verarbeitung
- **iText 7**: PDF-Generierung
- **Java AWT**: Grafik-Erstellung

### Performance
- **Kleine Dateien** (< 1MB): < 1 Sekunde
- **Mittlere Dateien** (1-10MB): 1-5 Sekunden  
- **Große Dateien** (> 10MB): 5+ Sekunden

### Speicherverbrauch
- Etwa 2-3x die Größe der Eingabedatei
- Empfohlen: Mindestens 512MB Heap für große Dokumente

## Fehlerbehebung

### Häufige Probleme

#### OutOfMemoryError
```bash
# Heap-Speicher erhöhen
java -Xmx2g -jar target/docconverter-1.0.0.jar input.docx output.pdf
```

#### Datei nicht gefunden
```bash
# Absolute Pfade verwenden
java -jar target/docconverter-1.0.0.jar /vollständiger/pfad/zu/input.docx output.pdf
```

#### Korrupte DOCX-Datei
```bash
# Datei-Integrität prüfen
file input.docx
# Sollte "Microsoft Word 2007+" anzeigen
```

### Logging aktivieren
```bash
# Detaillierte Ausgabe
java -Dlog4j.configuration=file:log4j.properties -jar target/docconverter-1.0.0.jar input.docx output.pdf
```

## API-Verwendung

### Programmatische Nutzung
```java
import ch.std.doc.converter.DocConverter;
import ch.std.doc.converter.ExtendedDocxCreator;

public class MeinConverter {
    public static void main(String[] args) throws IOException {
        // Komplexes DOCX erstellen
        ExtendedDocxCreator.createComplexDocx("test.docx");
        
        // Zu PDF konvertieren
        DocConverter converter = new DocConverter();
        converter.convertDocxToPdf("test.docx", "test.pdf");
        
        System.out.println("Konvertierung abgeschlossen!");
    }
}
```

## Lizenz und Credits

- **Apache POI**: Apache License 2.0
- **iText**: AGPL/Commercial License
- **DocConverter**: MIT License

Entwickelt mit GitHub Copilot als KI-Assistent.
```java
DocConverter converter = new DocConverter();
converter.convertDocxToPdf("input.docx", "output.pdf");
```

## Features:
- Konvertierung von DOCX zu PDF
- Erhaltung der Grundformatierung
- Erkennung von Überschriften (fett gedruckte Texte)
- Fehlerbehandlung

## Abhängigkeiten:
- Apache POI für DOCX-Verarbeitung
- iText 7 für PDF-Generierung
