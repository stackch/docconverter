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

@DisplayName("DocToPdfConverter Tests")
public class DocToPdfConverterTest {
    
    @TempDir
    Path tempDir;
    
    private DocumentConverter converter;
    private File outputFile;
    
    @BeforeEach
    void setUp() {
        converter = DocumentConverterFactory.createConverter("test.doc");
        assertNotNull(converter);
        assertInstanceOf(DocToPdfConverter.class, converter);
        
        outputFile = tempDir.resolve("output.pdf").toFile();
    }
    
    @Test
    @DisplayName("Konverter-Eigenschaften sind korrekt")
    public void testConverterProperties() {
        assertEquals("Word-Doc-Konverter", converter.getConverterName());
        assertArrayEquals(new String[]{".doc"}, converter.getSupportedExtensions());
    }
    
    @Test
    @DisplayName("Unterstützte Dateiformate werden korrekt erkannt")
    public void testSupportedFiles() {
        assertTrue(converter.supportsFile("document.doc"));
        assertTrue(converter.supportsFile("DOCUMENT.DOC"));
        
        assertFalse(converter.supportsFile("document.docx"));
        assertFalse(converter.supportsFile("document.xlsx"));
        assertFalse(converter.supportsFile("document.txt"));
        assertFalse(converter.supportsFile("document"));
    }
    
    @Test
    @DisplayName("Nicht existierende DOC-Datei wird korrekt behandelt")
    public void testNonExistentDocFile() {
        File nonExistentDoc = new File("/workspaces/docconverter/nonexistent.doc");
        assertFalse(nonExistentDoc.exists(), "Test-Setup: DOC-Datei sollte nicht existieren");
        
        assertThrows(Exception.class, () -> {
            converter.convertToPdf(nonExistentDoc.getAbsolutePath(), outputFile.getAbsolutePath());
        }, "Konvertierung nicht existierender Datei sollte Exception werfen");
        
        assertFalse(outputFile.exists(), "Keine PDF sollte für nicht existierende Eingabe erstellt werden");
    }
    
    @Test
    @DisplayName("DOC-Konverter behandelt Legacy-Format-Probleme")
    public void testLegacyFormatHandling() {
        // Test mit verschiedenen ungültigen Eingaben
        assertThrows(Exception.class, () -> {
            converter.convertToPdf("invalid.doc", outputFile.getAbsolutePath());
        }, "Ungültige DOC-Datei sollte Exception werfen");
        
        assertThrows(Exception.class, () -> {
            converter.convertToPdf("", outputFile.getAbsolutePath());
        }, "Leerer Input sollte Exception werfen");
        
        assertThrows(Exception.class, () -> {
            converter.convertToPdf(null, outputFile.getAbsolutePath());
        }, "Null-Input sollte Exception werfen");
        
        // Test mit ungültigen Ausgabepfaden
        assertThrows(Exception.class, () -> {
            converter.convertToPdf("test.doc", "");
        }, "Leerer Output sollte Exception werfen");
        
        assertThrows(Exception.class, () -> {
            converter.convertToPdf("test.doc", null);
        }, "Null-Output sollte Exception werfen");
        
        assertThrows(Exception.class, () -> {
            converter.convertToPdf("test.doc", "/invalid/path/output.pdf");
        }, "Ungültiger Ausgabepfad sollte Exception werfen");
        
        System.out.println("DOC-Konverter behandelt ungültige Eingaben korrekt");
    }
}
