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
}
