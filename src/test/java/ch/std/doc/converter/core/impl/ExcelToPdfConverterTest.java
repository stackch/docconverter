package ch.std.doc.converter.core.impl;

import ch.std.doc.converter.core.DocumentConverter;
import ch.std.doc.converter.core.DocumentConverterFactory;

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
}
