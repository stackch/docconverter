package ch.std.doc.converter;

import ch.std.doc.converter.core.DocumentConverter;
import ch.std.doc.converter.core.DocumentConverterFactory;
import ch.std.doc.converter.core.impl.DocxToPdfConverter;
import ch.std.doc.converter.core.impl.ExcelToPdfConverter;
import ch.std.doc.converter.core.impl.DocToPdfConverter;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Set;

/**
 * Tests für das neue Factory-basierte Dokumentkonverter-System
 */
public class DocumentConverterFactoryTest {
    
    @Test
    public void testFactoryExists() {
        // Teste dass die Factory-Klasse existiert
        assertNotNull(DocumentConverterFactory.class);
    }
    
    @Test
    public void testSupportedExtensions() {
        // Teste dass alle erwarteten Formate unterstützt werden
        Set<DocumentConverter> converters = DocumentConverterFactory.getRegisteredConverters();
        assertTrue(converters.size() >= 3, "Mindestens 3 Konverter sollten registriert sein");
        
        // Teste spezifische Formate
        assertTrue(DocumentConverterFactory.isSupported("test.docx"));
        assertTrue(DocumentConverterFactory.isSupported("test.xlsx"));
        assertTrue(DocumentConverterFactory.isSupported("test.doc"));
        assertFalse(DocumentConverterFactory.isSupported("test.txt"));
    }
    
    @Test
    public void testConverterCreation() {
        // Teste Konverter-Erstellung für verschiedene Formate
        DocumentConverter docxConverter = DocumentConverterFactory.createConverter("test.docx");
        assertNotNull(docxConverter, "DOCX-Konverter sollte erstellt werden");
        assertTrue(docxConverter instanceof DocxToPdfConverter);
        
        DocumentConverter excelConverter = DocumentConverterFactory.createConverter("test.xlsx");
        assertNotNull(excelConverter, "Excel-Konverter sollte erstellt werden");
        assertTrue(excelConverter instanceof ExcelToPdfConverter);
        
        DocumentConverter docConverter = DocumentConverterFactory.createConverter("test.doc");
        assertNotNull(docConverter, "DOC-Konverter sollte erstellt werden");
        assertTrue(docConverter instanceof DocToPdfConverter);
        
        DocumentConverter nullConverter = DocumentConverterFactory.createConverter("test.txt");
        assertNull(nullConverter, "Unsupported Format sollte null zurückgeben");
    }
    
    @Test
    public void testCaseInsensitiveExtensions() {
        // Teste dass Groß-/Kleinschreibung bei Erweiterungen ignoriert wird
        assertNotNull(DocumentConverterFactory.createConverter("TEST.DOCX"));
        assertNotNull(DocumentConverterFactory.createConverter("test.XLSX"));
        assertNotNull(DocumentConverterFactory.createConverter("Test.Doc"));
    }
    
    @Test
    public void testFormatOverview() {
        // Teste die Formatübersicht
        String overview = DocumentConverterFactory.getSupportedFormatsOverview();
        assertNotNull(overview);
        assertFalse(overview.isEmpty());
        assertTrue(overview.contains("DOCX-Konverter"));
        assertTrue(overview.contains("Excel-Konverter"));
        assertTrue(overview.contains("Word-Doc-Konverter"));
    }
    
    @Test
    public void testConverterNames() {
        // Teste dass alle Konverter sinnvolle Namen haben
        DocumentConverter docxConverter = DocumentConverterFactory.createConverter("test.docx");
        assertEquals("DOCX-Konverter", docxConverter.getConverterName());
        
        DocumentConverter excelConverter = DocumentConverterFactory.createConverter("test.xlsx");
        assertEquals("Excel-Konverter", excelConverter.getConverterName());
        
        DocumentConverter docConverter = DocumentConverterFactory.createConverter("test.doc");
        assertEquals("Word-Doc-Konverter", docConverter.getConverterName());
    }
    
    @Test
    public void testSupportedExtensionsPerConverter() {
        // Teste dass jeder Konverter seine Erweiterungen korrekt definiert
        DocumentConverter docxConverter = DocumentConverterFactory.createConverter("test.docx");
        String[] docxExtensions = docxConverter.getSupportedExtensions();
        assertEquals(1, docxExtensions.length);
        assertEquals(".docx", docxExtensions[0]);
        
        DocumentConverter excelConverter = DocumentConverterFactory.createConverter("test.xlsx");
        String[] excelExtensions = excelConverter.getSupportedExtensions();
        assertEquals(2, excelExtensions.length);
        assertTrue(java.util.Arrays.asList(excelExtensions).contains(".xlsx"));
        assertTrue(java.util.Arrays.asList(excelExtensions).contains(".xls"));
        
        DocumentConverter docConverter = DocumentConverterFactory.createConverter("test.doc");
        String[] docExtensions = docConverter.getSupportedExtensions();
        assertEquals(1, docExtensions.length);
        assertEquals(".doc", docExtensions[0]);
    }
    
    @Test
    public void testDocxConverterDescription() {
        // Teste den beschreibenden Text für DOCX-Format
        DocumentConverter docxConverter = DocumentConverterFactory.createConverter("test.docx");
        assertNotNull(docxConverter);
        assertEquals("DOCX-Konverter", docxConverter.getConverterName());
        assertTrue(docxConverter.getSupportedExtensions().length > 0);
        assertEquals(".docx", docxConverter.getSupportedExtensions()[0]);
    }
    
    @Test 
    public void testExcelConverterDescription() {
        // Teste den beschreibenden Text für Excel-Format
        DocumentConverter excelConverter = DocumentConverterFactory.createConverter("test.xlsx");
        assertNotNull(excelConverter);
        assertEquals("Excel-Konverter", excelConverter.getConverterName());
        String[] extensions = excelConverter.getSupportedExtensions();
        assertEquals(2, extensions.length);
        assertTrue(java.util.Arrays.asList(extensions).contains(".xlsx"));
        assertTrue(java.util.Arrays.asList(extensions).contains(".xls"));
    }
    
    @Test
    public void testDocConverterDescription() {
        // Teste den beschreibenden Text für DOC-Format  
        DocumentConverter docConverter = DocumentConverterFactory.createConverter("test.doc");
        assertNotNull(docConverter);
        assertEquals("Word-Doc-Konverter", docConverter.getConverterName());
        String[] extensions = docConverter.getSupportedExtensions();
        assertEquals(1, extensions.length);
        assertEquals(".doc", extensions[0]);
    }
    
    @Test
    public void testIntegrationWithRealFiles(@TempDir Path tempDir) throws IOException {
        // Teste mit echten Eingabedateien falls vorhanden
        java.io.File workspaceRoot = new java.io.File("/workspaces/docconverter");
        java.io.File beispielDocx = new java.io.File(workspaceRoot, "beispiel.docx");
        java.io.File testXlsx = new java.io.File(workspaceRoot, "test-verkaufsdaten.xlsx");
        
        if (beispielDocx.exists()) {
            // Test DOCX-Konvertierung
            DocumentConverter docxConverter = DocumentConverterFactory.createConverter(beispielDocx.getAbsolutePath());
            assertNotNull(docxConverter);
            assertTrue(docxConverter instanceof DocxToPdfConverter);
            
            java.io.File outputPdf = tempDir.resolve("integration_test.pdf").toFile();
            assertDoesNotThrow(() -> {
                docxConverter.convertToPdf(beispielDocx.getAbsolutePath(), outputPdf.getAbsolutePath());
            });
            assertTrue(outputPdf.exists());
            assertTrue(outputPdf.length() > 0);
        }
        
        if (testXlsx.exists()) {
            // Test Excel-Konvertierung  
            DocumentConverter excelConverter = DocumentConverterFactory.createConverter(testXlsx.getAbsolutePath());
            assertNotNull(excelConverter);
            assertTrue(excelConverter instanceof ExcelToPdfConverter);
            
            java.io.File outputPdf = tempDir.resolve("integration_excel_test.pdf").toFile();
            assertDoesNotThrow(() -> {
                excelConverter.convertToPdf(testXlsx.getAbsolutePath(), outputPdf.getAbsolutePath());
            });
            assertTrue(outputPdf.exists());
            assertTrue(outputPdf.length() > 0);
        }
    }
}
