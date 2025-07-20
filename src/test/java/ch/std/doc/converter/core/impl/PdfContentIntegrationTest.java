package ch.std.doc.converter.core.impl;

import ch.std.doc.converter.core.DocumentConverter;
import ch.std.doc.converter.core.DocumentConverterFactory;
import ch.std.doc.converter.utils.PdfContentValidator;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.io.TempDir;
import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Path;

/**
 * Integrierte Tests für PDF-Inhaltsvalidierung mit kontrollierten Testdaten
 */
@DisplayName("PDF Content Integration Tests")
public class PdfContentIntegrationTest {
    
    @TempDir
    Path tempDir;
    
    private File testDocxFile;
    private File outputPdfFile;
    
    @BeforeEach
    void setUp() throws Exception {
        // Erstelle eine kontrollierte Test-DOCX-Datei
        testDocxFile = tempDir.resolve("test-content.docx").toFile();
        outputPdfFile = tempDir.resolve("test-output.pdf").toFile();
        
        // Erstelle Test-DOCX direkt hier
        createTestDocx(testDocxFile);
    }
    
    /**
     * Erstellt eine Test-DOCX-Datei mit bekanntem Inhalt
     */
    private void createTestDocx(File file) throws Exception {
        try (XWPFDocument document = new XWPFDocument();
             FileOutputStream out = new FileOutputStream(file)) {
            
            // Überschrift
            XWPFParagraph title = document.createParagraph();
            title.createRun().setText("Beispiel Dokument - DOCX zu PDF Konvertierung");
            title.getRuns().get(0).setBold(true);
            title.getRuns().get(0).setFontSize(18);
            
            // Absatz 1
            XWPFParagraph para1 = document.createParagraph();
            para1.createRun().setText("Dies ist ein Testdokument für die Inhaltsvalidierung.");
            
            // Zwischenüberschrift
            XWPFParagraph subtitle = document.createParagraph();
            subtitle.createRun().setText("Features des DocConverters");
            subtitle.getRuns().get(0).setBold(true);
            subtitle.getRuns().get(0).setFontSize(14);
            
            // Feature-Liste
            XWPFParagraph feature1 = document.createParagraph();
            feature1.createRun().setText("• Konvertierung von DOCX zu PDF");
            
            XWPFParagraph feature2 = document.createParagraph();
            feature2.createRun().setText("• Erhaltung der Grundformatierung");
            
            document.write(out);
        }
    }
    
    @Test
    @DisplayName("DOCX-zu-PDF Konvertierung behält Textinhalt bei")
    public void testDocxContentPreservation() throws Exception {
        DocumentConverter converter = DocumentConverterFactory.createConverter("test.docx");
        
        // Konvertierung durchführen
        converter.convertToPdf(testDocxFile.getAbsolutePath(), outputPdfFile.getAbsolutePath());
        
        // Bekannte Inhalte aus der Test-DOCX prüfen
        assertTrue(PdfContentValidator.containsText(outputPdfFile, "Beispiel Dokument"), 
                  "PDF sollte Titel enthalten");
        
        assertTrue(PdfContentValidator.containsText(outputPdfFile, "DOCX zu PDF Konvertierung"), 
                  "PDF sollte Haupttitel enthalten");
        
        assertTrue(PdfContentValidator.containsText(outputPdfFile, "Features des DocConverters"), 
                  "PDF sollte Zwischenüberschrift enthalten");
        
        assertTrue(PdfContentValidator.containsText(outputPdfFile, "Konvertierung von DOCX zu PDF"), 
                  "PDF sollte Feature-Liste enthalten");
        
        // Alle erwarteten Texte in einem Zug prüfen
        boolean hasAllContent = PdfContentValidator.containsAllTexts(outputPdfFile,
            "Beispiel Dokument",
            "Features des DocConverters", 
            "Konvertierung von DOCX zu PDF"
        );
        assertTrue(hasAllContent, "PDF sollte alle erwarteten Textinhalte enthalten");
        
        System.out.println("✓ Alle erwarteten Textinhalte in PDF gefunden");
    }
    
    @Test
    @DisplayName("PDF-Textstruktur entspricht DOCX-Struktur")
    public void testStructuralIntegrity() throws Exception {
        DocumentConverter converter = DocumentConverterFactory.createConverter("test.docx");
        converter.convertToPdf(testDocxFile.getAbsolutePath(), outputPdfFile.getAbsolutePath());
        
        // Extrahiere Textzeilen
        var pdfLines = PdfContentValidator.extractLines(outputPdfFile);
        
        // Strukturvalidierung
        assertFalse(pdfLines.isEmpty(), "PDF sollte Textzeilen enthalten");
        
        // Prüfe ob wichtige Strukturelemente vorhanden sind
        boolean hasTitleLine = pdfLines.stream()
            .anyMatch(line -> line.contains("Beispiel Dokument"));
        assertTrue(hasTitleLine, "PDF sollte Titel-Zeile enthalten");
        
        boolean hasFeatureSection = pdfLines.stream()
            .anyMatch(line -> line.contains("Features"));
        assertTrue(hasFeatureSection, "PDF sollte Features-Sektion enthalten");
        
        System.out.println("✓ PDF-Struktur validiert: " + pdfLines.size() + " Textzeilen");
    }
    
    @Test
    @DisplayName("PDF-Qualitätsmetriken")
    public void testPdfQualityMetrics() throws Exception {
        DocumentConverter converter = DocumentConverterFactory.createConverter("test.docx");
        converter.convertToPdf(testDocxFile.getAbsolutePath(), outputPdfFile.getAbsolutePath());
        
        // Seitenzahl prüfen
        int pageCount = PdfContentValidator.getPageCount(outputPdfFile);
        assertEquals(1, pageCount, "Test-Dokument sollte genau eine Seite haben");
        
        // Textlänge prüfen
        String pdfText = PdfContentValidator.extractTextFromPdf(outputPdfFile);
        assertTrue(pdfText.length() > 100, "PDF sollte mindestens 100 Zeichen Text enthalten");
        assertTrue(pdfText.length() < 2000, "PDF sollte nicht übermäßig lang sein");
        
        // Zeichen-Validierung
        boolean hasAlphabetic = pdfText.matches(".*[a-zA-Z].*");
        assertTrue(hasAlphabetic, "PDF sollte alphabetische Zeichen enthalten");
        
        boolean hasSpaces = pdfText.contains(" ");
        assertTrue(hasSpaces, "PDF sollte Leerzeichen enthalten (Wort-Trennung)");
        
        System.out.println("✓ PDF-Qualität validiert: " + pageCount + " Seite(n), " + 
                          pdfText.length() + " Zeichen");
    }
    
    @Test
    @DisplayName("Fehlerresistenz bei beschädigten Eingaben")
    public void testCorruptedInputHandling() throws Exception {
        DocumentConverter converter = DocumentConverterFactory.createConverter("test.docx");
        
        // Erstelle eine leere "beschädigte" Datei
        File corruptedFile = tempDir.resolve("corrupted.docx").toFile();
        corruptedFile.createNewFile(); // Leere Datei
        
        // Sollte Exception werfen
        assertThrows(Exception.class, () -> {
            converter.convertToPdf(corruptedFile.getAbsolutePath(), outputPdfFile.getAbsolutePath());
        }, "Beschädigte DOCX-Datei sollte Exception werfen");
        
        // PDF sollte nicht erstellt werden
        assertFalse(outputPdfFile.exists(), "Keine PDF sollte für beschädigte Eingabe erstellt werden");
        
        System.out.println("✓ Fehlerbehandlung für beschädigte Dateien funktioniert");
    }
}
