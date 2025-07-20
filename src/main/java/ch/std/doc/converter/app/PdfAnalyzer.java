package ch.std.doc.converter.app;

import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.canvas.parser.PdfTextExtractor;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * Analysiert den Inhalt der konvertierten Generalversammlung.pdf
 */
public class PdfAnalyzer {
    
    public static void main(String[] args) {
        String pdfPath = "Generalversammlung_output.pdf";
        
        try {
            File pdfFile = new File(pdfPath);
            if (!pdfFile.exists()) {
                System.err.println("❌ PDF nicht gefunden: " + pdfPath);
                return;
            }
            
            System.out.println("📊 PDF-INHALTSANALYSE");
            System.out.println("═".repeat(60));
            System.out.println("📄 Datei: " + pdfFile.getName() + " (" + pdfFile.length() + " Bytes)");
            
            // PDF öffnen und analysieren
            try (PdfDocument pdfDoc = new PdfDocument(new PdfReader(pdfPath))) {
                int pageCount = pdfDoc.getNumberOfPages();
                System.out.println("📚 Seiten: " + pageCount);
                
                // Text von allen Seiten extrahieren
                StringBuilder fullText = new StringBuilder();
                List<String> lines = new ArrayList<>();
                
                for (int i = 1; i <= pageCount; i++) {
                    String pageText = PdfTextExtractor.getTextFromPage(pdfDoc.getPage(i));
                    fullText.append(pageText);
                    if (i < pageCount) {
                        fullText.append("\n");
                    }
                    
                    System.out.println("\n📖 SEITE " + i + ":");
                    System.out.println("─".repeat(40));
                    
                    // Zeilen der Seite
                    String[] pageLines = pageText.split("\n");
                    for (String line : pageLines) {
                        String trimmed = line.trim();
                        if (!trimmed.isEmpty()) {
                            lines.add(trimmed);
                            System.out.println("   " + trimmed);
                        }
                    }
                }
                
                // Gesamtstatistiken
                String totalText = fullText.toString();
                System.out.println("\n📊 GESAMTSTATISTIKEN");
                System.out.println("─".repeat(40));
                System.out.println("Gesamtzeichen: " + totalText.length());
                System.out.println("Textzeilen: " + lines.size());
                System.out.println("Wörter (geschätzt): " + totalText.split("\\s+").length);
                
                // Themensuche
                System.out.println("\n🔍 THEMENSUCHE");
                System.out.println("─".repeat(40));
                String[] keywords = {
                    "Generalversammlung", "Versammlung", "Protokoll", "Tagesordnung", 
                    "Beschluss", "Antrag", "Vorstand", "Mitglieder", "Abstimmung",
                    "Geschäftsbericht", "Finanzen", "Budget", "Wahl", "Verein",
                    "Satzung", "Jahresbericht", "Kassenprüfer", "Entlastung"
                };
                
                int foundKeywords = 0;
                for (String keyword : keywords) {
                    boolean found = totalText.toLowerCase().contains(keyword.toLowerCase());
                    if (found) {
                        System.out.println("✓ " + keyword);
                        foundKeywords++;
                    }
                }
                
                if (foundKeywords == 0) {
                    System.out.println("⚠️  Keine typischen Versammlungs-Begriffe gefunden");
                } else {
                    System.out.println("\n📈 " + foundKeywords + " von " + keywords.length + " Schlüsselwörtern gefunden");
                }
                
                // Volltext anzeigen (falls kurz)
                if (totalText.length() < 1000) {
                    System.out.println("\n📝 VOLLTEXT");
                    System.out.println("─".repeat(40));
                    System.out.println(totalText);
                }
                
                // Qualitätsbewertung
                System.out.println("\n🎯 QUALITÄTSBEWERTUNG");
                System.out.println("─".repeat(40));
                
                boolean hasText = totalText.length() > 10;
                boolean hasGermanChars = totalText.matches(".*[äöüÄÖÜß].*");
                boolean hasStructure = lines.size() > 1;
                boolean hasSubstance = totalText.length() > 100;
                
                System.out.println("Text extrahiert: " + (hasText ? "✅" : "❌"));
                System.out.println("Deutsche Umlaute: " + (hasGermanChars ? "✅" : "❌"));
                System.out.println("Struktur erkannt: " + (hasStructure ? "✅" : "❌"));
                System.out.println("Substanzieller Inhalt: " + (hasSubstance ? "✅" : "❌"));
                
                int score = (hasText ? 1 : 0) + (hasGermanChars ? 1 : 0) + 
                           (hasStructure ? 1 : 0) + (hasSubstance ? 1 : 0);
                
                System.out.println("\n🏆 GESAMTBEWERTUNG: " + score + "/4 " + getScoreEmoji(score));
                
            }
            
        } catch (Exception e) {
            System.err.println("❌ Fehler bei der PDF-Analyse: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    private static String getScoreEmoji(int score) {
        return switch (score) {
            case 4 -> "🥇 Excellent";
            case 3 -> "🥈 Gut";
            case 2 -> "🥉 Akzeptabel";
            case 1 -> "⚠️ Problematisch";
            default -> "❌ Fehlgeschlagen";
        };
    }
}
