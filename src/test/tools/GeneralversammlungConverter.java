package tools;

import ch.std.doc.converter.core.DocumentConverter;
import ch.std.doc.converter.core.DocumentConverterFactory;
import ch.std.doc.converter.utils.PdfContentValidator;

import java.io.File;

/**
 * Tool zur manuellen Konvertierung und Analyse der Generalversammlung.doc
 */
public class GeneralversammlungConverter {
    
    public static void main(String[] args) {
        String inputPath = "s:\\develop\\java\\docconverter\\Generalversammlung.doc";
        String outputPath = "s:\\develop\\java\\docconverter\\Generalversammlung.pdf";
        
        try {
            System.out.println("🔄 GENERALVERSAMMLUNG.DOC KONVERTIERUNG");
            System.out.println("═".repeat(50));
            
            // Eingabe validieren
            File inputFile = new File(inputPath);
            if (!inputFile.exists()) {
                System.err.println("❌ Eingabedatei nicht gefunden: " + inputPath);
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
                
                // Inhalt analysieren
                System.out.println("\n📊 INHALTSANALYSE");
                System.out.println("─".repeat(30));
                
                String pdfText = PdfContentValidator.extractTextFromPdf(outputFile);
                int pageCount = PdfContentValidator.getPageCount(outputFile);
                var lines = PdfContentValidator.extractLines(outputFile);
                
                System.out.println("Seiten: " + pageCount);
                System.out.println("Zeichen: " + pdfText.length());
                System.out.println("Zeilen: " + lines.size());
                
                // Erste paar Zeilen anzeigen
                System.out.println("\n📝 TEXTVORSCHAU");
                System.out.println("─".repeat(30));
                for (int i = 0; i < Math.min(10, lines.size()); i++) {
                    String line = lines.get(i);
                    if (line.length() > 60) {
                        line = line.substring(0, 57) + "...";
                    }
                    System.out.println(String.format("%2d: %s", i + 1, line));
                }
                
                // Suche nach typischen Versammlungsthemen
                System.out.println("\n🔍 THEMENSUCHE");
                System.out.println("─".repeat(30));
                String[] keywords = {
                    "versammlung", "protokoll", "tagesordnung", "beschluss", 
                    "antrag", "vorstand", "mitglieder", "abstimmung",
                    "geschäftsbericht", "finanzen", "budget", "wahl"
                };
                
                for (String keyword : keywords) {
                    boolean found = PdfContentValidator.containsText(outputFile, keyword);
                    if (found) {
                        System.out.println("✓ " + keyword);
                    }
                }
                
                System.out.println("\n✅ Analyse abgeschlossen!");
                System.out.println("📁 PDF gespeichert unter: " + outputPath);
                
            } else {
                System.err.println("❌ PDF-Datei wurde nicht erstellt!");
            }
            
        } catch (Exception e) {
            System.err.println("❌ Fehler bei der Konvertierung: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
