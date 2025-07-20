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

@DisplayName("ExcelToPdfConverter Tests")
public class ExcelToPdfConverterTest {
    
    @TempDir
    Path tempDir;
    
    private DocumentConverter converter;
    private File inputFile;
    private File outputFile;
    
    @BeforeEach
    void setUp() {
        converter = DocumentConverterFactory.createConverter("test.xlsx");
        assertNotNull(converter);
        assertInstanceOf(ExcelToPdfConverter.class, converter);
        
        inputFile = new File("/workspaces/docconverter/test-verkaufsdaten.xlsx");
        outputFile = tempDir.resolve("output.pdf").toFile();
    }
    
    @Test
    @DisplayName("Konverter-Eigenschaften sind korrekt")
    public void testConverterProperties() {
        assertEquals("Excel-Konverter", converter.getConverterName());
        String[] supportedExtensions = converter.getSupportedExtensions();
        assertEquals(2, supportedExtensions.length);
    }
    
    @Test
    @DisplayName("Unterstützte Dateiformate werden korrekt erkannt")
    public void testSupportedFiles() {
        assertTrue(converter.supportsFile("spreadsheet.xlsx"));
        assertTrue(converter.supportsFile("SPREADSHEET.XLSX"));
        assertTrue(converter.supportsFile("spreadsheet.xls"));
        
        assertFalse(converter.supportsFile("document.docx"));
        assertFalse(converter.supportsFile("document.doc"));
        assertFalse(converter.supportsFile("data.csv"));
        assertFalse(converter.supportsFile("file.txt"));
    }
    
    @Test
    @DisplayName("Excel-zu-PDF Konvertierung funktioniert")
    public void testBasicConversion() throws Exception {
        if (!inputFile.exists()) {
            System.out.println("Warnung: test-verkaufsdaten.xlsx nicht gefunden, überspringe Konvertierungstest");
            return;
        }
        
        assertDoesNotThrow(() -> {
            converter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
        }, "Konvertierung sollte ohne Exception erfolgreich sein");
        
        assertTrue(outputFile.exists(), "PDF-Datei sollte erstellt werden");
        assertTrue(outputFile.length() > 0, "PDF-Datei sollte nicht leer sein");
    }
    
    @Test
    @DisplayName("Ungültige Excel-Eingaben werden behandelt")
    public void testInvalidExcelInputHandling() {
        // Nicht existierende Datei
        assertThrows(Exception.class, () -> {
            converter.convertToPdf("nonexistent.xlsx", outputFile.getAbsolutePath());
        }, "Nicht existierende Datei sollte Exception werfen");
        
        // Null-Eingaben
        assertThrows(Exception.class, () -> {
            converter.convertToPdf(null, outputFile.getAbsolutePath());
        }, "Null-Input sollte Exception werfen");
        
        assertThrows(Exception.class, () -> {
            converter.convertToPdf(inputFile.getAbsolutePath(), null);
        }, "Null-Output sollte Exception werfen");
        
        // Leere Strings
        assertThrows(Exception.class, () -> {
            converter.convertToPdf("", outputFile.getAbsolutePath());
        }, "Leerer Input sollte Exception werfen");
        
        assertThrows(Exception.class, () -> {
            converter.convertToPdf(inputFile.getAbsolutePath(), "");
        }, "Leerer Output sollte Exception werfen");
    }
    
    @Test
    @DisplayName("Excel-PDF enthält Tabelleninhalt")
    public void testExcelPdfContentValidation() throws Exception {
        if (!inputFile.exists()) {
            System.out.println("Warnung: test-verkaufsdaten.xlsx nicht gefunden, überspringe Excel-Inhaltstest");
            return;
        }
        
        converter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
        
        // PDF sollte existieren und nicht leer sein
        assertTrue(outputFile.exists(), "PDF-Datei sollte erstellt werden");
        assertTrue(outputFile.length() > 0, "PDF-Datei sollte nicht leer sein");
        
        // Text aus PDF extrahieren
        String pdfText = PdfContentValidator.extractTextFromPdf(outputFile);
        assertNotNull(pdfText, "PDF-Text sollte extrahiert werden können");
        assertFalse(pdfText.trim().isEmpty(), "PDF sollte Text enthalten");
        
        // Excel-spezifische Validierung
        // Suche nach typischen Excel-Inhalten (Zahlen, Spaltenstruktur)
        boolean hasNumbers = pdfText.matches(".*\\d+.*");
        assertTrue(hasNumbers, "Excel-PDF sollte Zahlen enthalten");
        
        System.out.println("Excel-PDF-Text (erste 200 Zeichen): " + 
                          pdfText.substring(0, Math.min(200, pdfText.length())));
    }
    
    @Test
    @DisplayName("Excel-PDF Seitenstruktur validieren")
    public void testExcelPdfStructure() throws Exception {
        if (!inputFile.exists()) {
            System.out.println("Warnung: test-verkaufsdaten.xlsx nicht gefunden, überspringe Strukturtest");
            return;
        }
        
        converter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
        
        // Seitenanzahl prüfen
        int pageCount = PdfContentValidator.getPageCount(outputFile);
        assertTrue(pageCount > 0, "Excel-PDF sollte mindestens eine Seite haben");
        
        // Text von erster Seite extrahieren
        String firstPageText = PdfContentValidator.extractTextFromPage(outputFile, 1);
        assertNotNull(firstPageText, "Erste Seite sollte Text enthalten");
        
        System.out.println("Excel-PDF hat " + pageCount + " Seite(n)");
        System.out.println("Erste Seite Inhalt: " + firstPageText.substring(0, Math.min(100, firstPageText.length())));
    }
    
    @Test
    @DisplayName("Excel-PDF enthält erwartete Verkaufsdaten")
    public void testSalesDataContent() throws Exception {
        if (!inputFile.exists()) {
            System.out.println("Warnung: test-verkaufsdaten.xlsx nicht gefunden, überspringe Verkaufsdaten-Test");
            return;
        }
        
        converter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
        
        String pdfText = PdfContentValidator.extractTextFromPdf(outputFile);
        
        // Basis-Validierung für Verkaufsdaten
        assertTrue(pdfText.length() > 20, "PDF sollte substantiellen Inhalt haben");
        
        // Test auf typische Excel-Strukturen
        boolean hasTabularData = pdfText.contains(" ") && pdfText.length() > 50;
        assertTrue(hasTabularData, "PDF sollte tabellarische Daten enthalten");
        
        System.out.println("Verkaufsdaten-PDF validiert - " + pdfText.length() + " Zeichen");
    }
}
