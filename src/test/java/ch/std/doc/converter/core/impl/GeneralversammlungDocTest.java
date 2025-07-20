package ch.std.doc.converter.core.impl;

import ch.std.doc.converter.core.DocumentConverter;
import ch.std.doc.converter.core.DocumentConverterFactory;
import ch.std.doc.converter.utils.PdfContentValidator;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.io.TempDir;
import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.nio.file.Path;

/**
 * Spezifischer Test für die Konvertierung der Generalversammlung.doc Datei
 */
@DisplayName("Generalversammlung.doc Content Tests")
public class GeneralversammlungDocTest {
    
    @TempDir
    Path tempDir;
    
    private DocumentConverter converter;
    private File inputFile;
    private File outputFile;
    
    @BeforeEach
    void setUp() {
        converter = DocumentConverterFactory.createConverter("test.doc");
        assertNotNull(converter);
        
        // Verwende die echte Generalversammlung.doc Datei
        inputFile = new File("s:\\develop\\java\\docconverter\\Generalversammlung.doc");
        outputFile = tempDir.resolve("Generalversammlung.pdf").toFile();
    }
    
    @Test
    @DisplayName("Generalversammlung.doc existiert und ist zugänglich")
    public void testInputFileExists() {
        assertTrue(inputFile.exists(), "Generalversammlung.doc sollte existieren");
        assertTrue(inputFile.canRead(), "Generalversammlung.doc sollte lesbar sein");
        assertTrue(inputFile.length() > 0, "Generalversammlung.doc sollte nicht leer sein");
        
        System.out.println("✓ Eingabedatei validiert: " + inputFile.getName() + 
                          " (" + inputFile.length() + " Bytes)");
    }
    
    @Test
    @DisplayName("DOC zu PDF Konvertierung funktioniert")
    public void testBasicConversion() throws Exception {
        if (!inputFile.exists()) {
            System.out.println("Warnung: Generalversammlung.doc nicht gefunden");
            return;
        }
        
        // Konvertierung durchführen
        assertDoesNotThrow(() -> {
            converter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
        }, "DOC-Konvertierung sollte ohne Exception erfolgreich sein");
        
        // Basis-Validierung
        assertTrue(outputFile.exists(), "PDF-Datei sollte erstellt werden");
        assertTrue(outputFile.length() > 0, "PDF-Datei sollte nicht leer sein");
        
        System.out.println("✓ PDF erfolgreich erstellt: " + outputFile.getName() + 
                          " (" + outputFile.length() + " Bytes)");
    }
    
    @Test
    @DisplayName("PDF-Inhalt wird korrekt extrahiert")
    public void testPdfContentExtraction() throws Exception {
        if (!inputFile.exists()) {
            System.out.println("Warnung: Generalversammlung.doc nicht gefunden");
            return;
        }
        
        // Konvertierung durchführen
        converter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
        
        // Text aus PDF extrahieren
        String pdfText = PdfContentValidator.extractTextFromPdf(outputFile);
        
        assertNotNull(pdfText, "PDF-Text sollte extrahiert werden können");
        assertFalse(pdfText.trim().isEmpty(), "PDF sollte Text enthalten");
        
        // Zeige ersten Teil des extrahierten Textes
        String preview = pdfText.substring(0, Math.min(300, pdfText.length()));
        System.out.println("✓ PDF-Text extrahiert (" + pdfText.length() + " Zeichen)");
        System.out.println("Textvorschau:");
        System.out.println("─".repeat(50));
        System.out.println(preview + "...");
        System.out.println("─".repeat(50));
    }
    
    @Test
    @DisplayName("PDF-Struktur und Qualität validieren")
    public void testPdfStructureAndQuality() throws Exception {
        if (!inputFile.exists()) {
            System.out.println("Warnung: Generalversammlung.doc nicht gefunden");
            return;
        }
        
        converter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
        
        // Seitenanzahl
        int pageCount = PdfContentValidator.getPageCount(outputFile);
        assertTrue(pageCount > 0, "PDF sollte mindestens eine Seite haben");
        
        // Text-Qualität
        String pdfText = PdfContentValidator.extractTextFromPdf(outputFile);
        assertTrue(pdfText.length() > 50, "PDF sollte substantiellen Text enthalten");
        
        // Deutsche Texte erwarten (Generalversammlung ist wahrscheinlich deutsch)
        boolean hasGermanText = pdfText.toLowerCase().contains("versammlung") || 
                               pdfText.toLowerCase().contains("protokoll") ||
                               pdfText.toLowerCase().contains("tagesordnung");
        
        // Zeichen-Validierung
        boolean hasAlphabetic = pdfText.matches(".*[a-zA-ZäöüÄÖÜß].*");
        assertTrue(hasAlphabetic, "PDF sollte alphabetische Zeichen enthalten");
        
        System.out.println("✓ PDF-Struktur validiert:");
        System.out.println("  Seiten: " + pageCount);
        System.out.println("  Zeichen: " + pdfText.length());
        System.out.println("  Deutsche Inhalte erkannt: " + (hasGermanText ? "Ja" : "Nein"));
        
        // Textzeilen extrahieren für Strukturanalyse
        var lines = PdfContentValidator.extractLines(outputFile);
        System.out.println("  Textzeilen: " + lines.size());
        
        if (!lines.isEmpty()) {
            System.out.println("  Erste Zeile: \"" + lines.get(0) + "\"");
            if (lines.size() > 1) {
                System.out.println("  Zweite Zeile: \"" + lines.get(1) + "\"");
            }
        }
    }
    
    @Test
    @DisplayName("Detaillierte Inhaltsanalyse")
    public void testDetailedContentAnalysis() throws Exception {
        if (!inputFile.exists()) {
            System.out.println("Warnung: Generalversammlung.doc nicht gefunden");
            return;
        }
        
        converter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
        
        String pdfText = PdfContentValidator.extractTextFromPdf(outputFile);
        var lines = PdfContentValidator.extractLines(outputFile);
        
        System.out.println("📋 DETAILLIERTE INHALTSANALYSE");
        System.out.println("═".repeat(60));
        
        // Allgemeine Statistiken
        System.out.println("📊 Statistiken:");
        System.out.println("   Gesamtzeichen: " + pdfText.length());
        System.out.println("   Textzeilen: " + lines.size());
        System.out.println("   Wörter (geschätzt): " + pdfText.split("\\s+").length);
        
        // Häufige deutsche Wörter in Dokumenten suchen
        String[] commonWords = {"und", "der", "die", "das", "von", "für", "mit", "auf", "zu", "in"};
        int germanWordCount = 0;
        for (String word : commonWords) {
            if (pdfText.toLowerCase().contains(word)) {
                germanWordCount++;
            }
        }
        System.out.println("   Deutsche Grundwörter gefunden: " + germanWordCount + "/" + commonWords.length);
        
        // Zeige alle Textzeilen (begrenzt)
        System.out.println("\n📄 Textinhalt (erste 15 Zeilen):");
        for (int i = 0; i < Math.min(15, lines.size()); i++) {
            String line = lines.get(i);
            if (line.length() > 80) {
                line = line.substring(0, 77) + "...";
            }
            System.out.println(String.format("   %2d: %s", i + 1, line));
        }
        
        if (lines.size() > 15) {
            System.out.println("   ... (" + (lines.size() - 15) + " weitere Zeilen)");
        }
        
        System.out.println("═".repeat(60));
    }
}
