package ch.std.doc.converter.app;

import ch.std.doc.converter.core.DocumentConverter;
import ch.std.doc.converter.core.DocumentConverterFactory;

import java.io.File;

/**
 * Spezielle Anwendung zur Konvertierung der Generalversammlung.doc
 */
public class GeneralversammlungApp {
    
    public static void main(String[] args) {
        String inputPath = "Generalversammlung.doc";
        String outputPath = "Generalversammlung_output.pdf";
        
        try {
            System.out.println("🔄 GENERALVERSAMMLUNG.DOC KONVERTIERUNG");
            System.out.println("═".repeat(50));
            
            // Eingabe validieren
            File inputFile = new File(inputPath);
            if (!inputFile.exists()) {
                System.err.println("❌ Eingabedatei nicht gefunden: " + inputPath);
                System.err.println("   Aktuelles Verzeichnis: " + System.getProperty("user.dir"));
                return;
            }
            
            System.out.println("📂 Eingabe: " + inputFile.getName() + " (" + inputFile.length() + " Bytes)");
            
            // Konverter erstellen
            DocumentConverter converter = DocumentConverterFactory.createConverter("test.doc");
            System.out.println("🔧 Konverter: " + converter.getConverterName());
            
            // Konvertierung durchführen
            System.out.println("⚙️  Konvertiere DOC zu PDF...");
            long startTime = System.currentTimeMillis();
            
            converter.convertToPdf(inputPath, outputPath);
            
            long duration = System.currentTimeMillis() - startTime;
            System.out.println("✅ Konvertierung abgeschlossen in " + duration + "ms");
            
            // Ausgabe validieren
            File outputFile = new File(outputPath);
            if (outputFile.exists()) {
                System.out.println("📄 Ausgabe: " + outputFile.getName() + " (" + outputFile.length() + " Bytes)");
                System.out.println("✅ PDF erfolgreich erstellt!");
                System.out.println("📁 Gespeichert unter: " + outputFile.getAbsolutePath());
            } else {
                System.err.println("❌ PDF-Datei wurde nicht erstellt!");
            }
            
        } catch (Exception e) {
            System.err.println("❌ Fehler bei der Konvertierung: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
