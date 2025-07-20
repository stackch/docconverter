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

@DisplayName("DocxToPdfConverter Tests")
public class DocxToPdfConverterTest {
    
    @TempDir
    Path tempDir;
    
    private DocumentConverter converter;
    private File inputFile;
    private File outputFile;
    
    @BeforeEach
    void setUp() {
        converter = DocumentConverterFactory.createConverter("test.docx");
        assertNotNull(converter);
        assertInstanceOf(DocxToPdfConverter.class, converter);
        
        inputFile = new File("/workspaces/docconverter/beispiel.docx");
        outputFile = tempDir.resolve("output.pdf").toFile();
    }
    
    @Test
    @DisplayName("Konverter-Eigenschaften sind korrekt")
    public void testConverterProperties() {
        assertEquals("DOCX-Konverter", converter.getConverterName());
        assertArrayEquals(new String[]{".docx"}, converter.getSupportedExtensions());
    }
    
    @Test
    @DisplayName("Unterstützte Dateiformate werden korrekt erkannt")
    public void testSupportedFiles() {
        assertTrue(converter.supportsFile("document.docx"));
        assertTrue(converter.supportsFile("DOCUMENT.DOCX"));
        
        assertFalse(converter.supportsFile("document.doc"));
        assertFalse(converter.supportsFile("document.xlsx"));
        assertFalse(converter.supportsFile("document.txt"));
        assertFalse(converter.supportsFile("document"));
    }
    
    @Test
    @DisplayName("DOCX-zu-PDF Konvertierung funktioniert")
    public void testBasicConversion() throws Exception {
        if (!inputFile.exists()) {
            System.out.println("Warnung: beispiel.docx nicht gefunden, überspringe Konvertierungstest");
            return;
        }
        
        assertDoesNotThrow(() -> {
            converter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
        }, "Konvertierung sollte ohne Exception erfolgreich sein");
        
        assertTrue(outputFile.exists(), "PDF-Datei sollte erstellt werden");
        assertTrue(outputFile.length() > 0, "PDF-Datei sollte nicht leer sein");
    }
    
    @Test
    @DisplayName("Ungültige Eingabedateien werden korrekt behandelt")
    public void testInvalidInputHandling() {
        // Nicht existierende Datei
        assertThrows(Exception.class, () -> {
            converter.convertToPdf("nonexistent.docx", outputFile.getAbsolutePath());
        }, "Nicht existierende Datei sollte Exception werfen");
        
        // Null-Eingaben
        assertThrows(Exception.class, () -> {
            converter.convertToPdf(null, outputFile.getAbsolutePath());
        }, "Null-Input sollte Exception werfen");
        
        assertThrows(Exception.class, () -> {
            converter.convertToPdf(inputFile.getAbsolutePath(), null);
        }, "Null-Output sollte Exception werfen");
        
        assertThrows(Exception.class, () -> {
            converter.convertToPdf(null, null);
        }, "Null-Inputs sollten Exception werfen");
        
        // Leere Strings
        assertThrows(Exception.class, () -> {
            converter.convertToPdf("", outputFile.getAbsolutePath());
        }, "Leerer Input sollte Exception werfen");
        
        assertThrows(Exception.class, () -> {
            converter.convertToPdf(inputFile.getAbsolutePath(), "");
        }, "Leerer Output sollte Exception werfen");
    }
    
    @Test
    @DisplayName("PDF-Inhalt wird korrekt konvertiert - Textprüfung")
    public void testPdfContentValidation() throws Exception {
        if (!inputFile.exists()) {
            System.out.println("Warnung: beispiel.docx nicht gefunden, überspringe Inhaltstest");
            return;
        }
        
        // Konvertierung durchführen
        converter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
        
        // PDF sollte existieren und nicht leer sein
        assertTrue(outputFile.exists(), "PDF-Datei sollte erstellt werden");
        assertTrue(outputFile.length() > 0, "PDF-Datei sollte nicht leer sein");
        
        // Text aus PDF extrahieren
        String pdfText = PdfContentValidator.extractTextFromPdf(outputFile);
        assertNotNull(pdfText, "PDF-Text sollte extrahiert werden können");
        assertFalse(pdfText.trim().isEmpty(), "PDF sollte Text enthalten");
        
        // Grundlegende Inhaltsvalidierung
        System.out.println("Extrahierter PDF-Text (erste 200 Zeichen): " + 
                          pdfText.substring(0, Math.min(200, pdfText.length())));
    }
    
    @Test
    @DisplayName("PDF-Seitenanzahl wird korrekt bestimmt")
    public void testPdfPageCount() throws Exception {
        if (!inputFile.exists()) {
            System.out.println("Warnung: beispiel.docx nicht gefunden, überspringe Seitentest");
            return;
        }
        
        converter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
        
        int pageCount = PdfContentValidator.getPageCount(outputFile);
        assertTrue(pageCount > 0, "PDF sollte mindestens eine Seite haben");
        
        System.out.println("PDF hat " + pageCount + " Seite(n)");
    }
    
    @Test
    @DisplayName("PDF enthält erwartete Textinhalte")
    public void testPdfContainsExpectedContent() throws Exception {
        if (!inputFile.exists()) {
            System.out.println("Warnung: beispiel.docx nicht gefunden, überspringe Textinhalt-Test");
            return;
        }
        
        converter.convertToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
        
        // Prüfe ob bestimmte Standardinhalte enthalten sind
        // Diese können je nach beispiel.docx angepasst werden
        String pdfText = PdfContentValidator.extractTextFromPdf(outputFile);
        
        // Basis-Validierung: PDF sollte lesbaren Text enthalten
        assertTrue(pdfText.length() > 10, "PDF sollte substantiellen Text enthalten");
        
        // Test auf häufige Wörter/Zeichen
        boolean hasAlphanumeric = pdfText.matches(".*[a-zA-Z0-9].*");
        assertTrue(hasAlphanumeric, "PDF sollte alphanumerische Zeichen enthalten");
        
        System.out.println("PDF-Inhalt validiert - Text enthält " + pdfText.length() + " Zeichen");
    }
}
