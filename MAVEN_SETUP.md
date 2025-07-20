# Maven Setup und Verwendung

## ✅ Maven Wrapper erfolgreich installiert!

Du hast jetzt einen **Maven Wrapper** in deinem Projekt, der Maven automatisch herunterlädt und verwendet.

## Verwendung

### Option 1: Direkter Maven Wrapper (erfordert JAVA_HOME)
```cmd
# JAVA_HOME setzen und Maven verwenden
$env:JAVA_HOME = "C:\Program Files\Eclipse Adoptium\jdk-21.0.7.6-hotspot"
.\mvnw.cmd clean compile
.\mvnw.cmd test
.\mvnw.cmd package
```

### Option 2: Convenience Script (empfohlen)
```cmd
# Einfacher mit dem bereitgestellten Script
.\maven-run.cmd clean compile
.\maven-run.cmd test
.\maven-run.cmd package
```

## Häufige Maven-Kommandos

```cmd
# Projekt kompilieren
.\maven-run.cmd clean compile

# Tests kompilieren (ohne ausführen)
.\maven-run.cmd clean compile test-compile

# Alle Tests ausführen
.\maven-run.cmd test

# Nur spezifische Tests
.\maven-run.cmd test -Dtest=DocxToPdfConverterTest

# Nur Inhaltstests
.\maven-run.cmd test -Dtest="*Test#*Content*"

# JAR-Datei erstellen
.\maven-run.cmd clean package

# Dependencies herunterladen
.\maven-run.cmd dependency:resolve

# Maven-Version anzeigen
.\maven-run.cmd --version
```

## Projektstruktur nach Maven-Setup

```
docconverter/
├── .mvn/
│   └── wrapper/
│       ├── maven-wrapper.properties
│       └── maven-wrapper.jar (wird automatisch heruntergeladen)
├── mvnw.cmd                    # Maven Wrapper für Windows
├── maven-run.cmd              # Convenience Script
├── pom.xml                    # Maven-Konfiguration
└── src/
    ├── main/java/             # Produktiver Code
    └── test/java/             # Test-Code
```

## Vorteile des Maven Wrapper

✅ **Keine globale Maven-Installation nötig**  
✅ **Automatisches Download der korrekten Maven-Version**  
✅ **Konsistente Build-Umgebung für alle Entwickler**  
✅ **Funktioniert ohne Admin-Rechte**  

## Nächste Schritte

1. **Projekt kompilieren:**
   ```cmd
   .\maven-run.cmd clean compile
   ```

2. **Tests ausführen:**
   ```cmd
   .\maven-run.cmd test
   ```

3. **JAR erstellen:**
   ```cmd
   .\maven-run.cmd clean package
   ```

Der Maven Wrapper lädt beim ersten Aufruf automatisch Maven 3.9.6 herunter und verwendet es für alle nachfolgenden Builds.
