package ch.std.doc.converter.core;

import ch.std.doc.converter.core.impl.DocxToPdfConverter;
import ch.std.doc.converter.core.impl.ExcelToPdfConverter;
import ch.std.doc.converter.core.impl.DocToPdfConverter;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.DisplayName;
import static org.junit.jupiter.api.Assertions.*;

import java.util.Set;

/**
 * Tests für die DocumentConverterFactory
 */
@DisplayName("DocumentConverterFactory Tests")
public class DocumentConverterFactoryTest {
    
    @Test
    @DisplayName("Factory-Klasse existiert und ist funktionsfähig")
    public void testFactoryExists() {
        assertNotNull(DocumentConverterFactory.class);
        assertDoesNotThrow(() -> {
            DocumentConverterFactory.getRegisteredConverters();
        });
    }
    
    @Test
    @DisplayName("Alle erwarteten Konverter sind registriert")
    public void testConverterRegistration() {
        Set<DocumentConverter> converters = DocumentConverterFactory.getRegisteredConverters();
        
        assertEquals(3, converters.size(), "Genau 3 Konverter sollten registriert sein");
        
        // Prüfe dass alle Converter-Typen vorhanden sind
        boolean hasDocxConverter = converters.stream()
                .anyMatch(c -> c instanceof DocxToPdfConverter);
        boolean hasExcelConverter = converters.stream()
                .anyMatch(c -> c instanceof ExcelToPdfConverter);
        boolean hasDocConverter = converters.stream()
                .anyMatch(c -> c instanceof DocToPdfConverter);
        
        assertTrue(hasDocxConverter, "DOCX-Konverter sollte registriert sein");
        assertTrue(hasExcelConverter, "Excel-Konverter sollte registriert sein");
        assertTrue(hasDocConverter, "DOC-Konverter sollte registriert sein");
    }
    
    @Test
    @DisplayName("Unterstützte Dateiformate werden korrekt erkannt")
    public void testSupportedExtensions() {
        // Unterstützte Formate
        assertTrue(DocumentConverterFactory.isSupported("document.docx"));
        assertTrue(DocumentConverterFactory.isSupported("data.xlsx"));
        assertTrue(DocumentConverterFactory.isSupported("legacy.doc"));
        assertTrue(DocumentConverterFactory.isSupported("spreadsheet.xls"));
        
        // Groß-/Kleinschreibung
        assertTrue(DocumentConverterFactory.isSupported("DOCUMENT.DOCX"));
        assertTrue(DocumentConverterFactory.isSupported("Data.XLSX"));
        assertTrue(DocumentConverterFactory.isSupported("Legacy.DOC"));
        
        // Nicht unterstützte Formate
        assertFalse(DocumentConverterFactory.isSupported("text.txt"));
        assertFalse(DocumentConverterFactory.isSupported("image.jpg"));
        assertFalse(DocumentConverterFactory.isSupported("archive.zip"));
        assertFalse(DocumentConverterFactory.isSupported("unknown.xyz"));
    }
    
    @Test
    @DisplayName("Konverter-Erstellung funktioniert korrekt")
    public void testConverterCreation() {
        // DOCX-Konverter
        DocumentConverter docxConverter = DocumentConverterFactory.createConverter("test.docx");
        assertNotNull(docxConverter);
        assertInstanceOf(DocxToPdfConverter.class, docxConverter);
        assertEquals("DOCX-Konverter", docxConverter.getConverterName());
        
        // Excel-Konverter (XLSX)
        DocumentConverter xlsxConverter = DocumentConverterFactory.createConverter("data.xlsx");
        assertNotNull(xlsxConverter);
        assertInstanceOf(ExcelToPdfConverter.class, xlsxConverter);
        assertEquals("Excel-Konverter", xlsxConverter.getConverterName());
        
        // Excel-Konverter (XLS)
        DocumentConverter xlsConverter = DocumentConverterFactory.createConverter("data.xls");
        assertNotNull(xlsConverter);
        assertInstanceOf(ExcelToPdfConverter.class, xlsConverter);
        
        // DOC-Konverter
        DocumentConverter docConverter = DocumentConverterFactory.createConverter("legacy.doc");
        assertNotNull(docConverter);
        assertInstanceOf(DocToPdfConverter.class, docConverter);
        assertEquals("Word-Doc-Konverter", docConverter.getConverterName());
        
        // Nicht unterstützte Formate
        assertNull(DocumentConverterFactory.createConverter("text.txt"));
        assertNull(DocumentConverterFactory.createConverter("image.png"));
    }
    
    @Test
    @DisplayName("Konverter-Namen sind eindeutig")
    public void testConverterNamesAreUnique() {
        Set<DocumentConverter> converters = DocumentConverterFactory.getRegisteredConverters();
        
        Set<String> names = converters.stream()
                .map(DocumentConverter::getConverterName)
                .collect(java.util.stream.Collectors.toSet());
        
        assertEquals(converters.size(), names.size(), 
                "Alle Konverter-Namen sollten eindeutig sein");
    }
    
    @Test
    @DisplayName("Unterstützte Erweiterungen sind korrekt definiert")
    public void testSupportedExtensionsPerConverter() {
        // DOCX-Konverter
        DocumentConverter docxConverter = DocumentConverterFactory.createConverter("test.docx");
        String[] docxExtensions = docxConverter.getSupportedExtensions();
        assertEquals(1, docxExtensions.length);
        assertEquals(".docx", docxExtensions[0]);
        
        // Excel-Konverter
        DocumentConverter excelConverter = DocumentConverterFactory.createConverter("test.xlsx");
        String[] excelExtensions = excelConverter.getSupportedExtensions();
        assertEquals(2, excelExtensions.length);
        assertTrue(java.util.Arrays.asList(excelExtensions).contains(".xlsx"));
        assertTrue(java.util.Arrays.asList(excelExtensions).contains(".xls"));
        
        // DOC-Konverter
        DocumentConverter docConverter = DocumentConverterFactory.createConverter("test.doc");
        String[] docExtensions = docConverter.getSupportedExtensions();
        assertEquals(1, docExtensions.length);
        assertEquals(".doc", docExtensions[0]);
    }
    
    @Test
    @DisplayName("Formatübersicht ist vollständig und informativ")
    public void testFormatOverview() {
        String overview = DocumentConverterFactory.getSupportedFormatsOverview();
        
        assertNotNull(overview);
        assertFalse(overview.isEmpty());
        
        // Prüfe dass alle Konverter in der Übersicht erwähnt werden
        assertTrue(overview.contains("DOCX-Konverter"));
        assertTrue(overview.contains("Excel-Konverter"));
        assertTrue(overview.contains("Word-Doc-Konverter"));
        
        // Prüfe dass alle Dateierweiterungen erwähnt werden
        assertTrue(overview.contains(".docx"));
        assertTrue(overview.contains(".xlsx"));
        assertTrue(overview.contains(".xls"));
        assertTrue(overview.contains(".doc"));
        
        // Prüfe dass Beschreibungen vorhanden sind
        assertTrue(overview.contains("Beschreibung:"));
        assertTrue(overview.contains("Microsoft Word"));
        assertTrue(overview.contains("Microsoft Excel"));
    }
    
    @Test
    @DisplayName("Beschreibungen sind aussagekräftig")
    public void testConverterDescriptions() {
        // DOCX-Konverter
        DocumentConverter docxConverter = DocumentConverterFactory.createConverter("test.docx");
        String docxDescription = docxConverter.getDescription();
        assertNotNull(docxDescription);
        assertFalse(docxDescription.isEmpty());
        assertTrue(docxDescription.contains("DOCX"));
        assertTrue(docxDescription.contains("PDF"));
        
        // Excel-Konverter
        DocumentConverter excelConverter = DocumentConverterFactory.createConverter("test.xlsx");
        String excelDescription = excelConverter.getDescription();
        assertNotNull(excelDescription);
        assertFalse(excelDescription.isEmpty());
        assertTrue(excelDescription.contains("Excel"));
        assertTrue(excelDescription.contains("PDF"));
        
        // DOC-Konverter
        DocumentConverter docConverter = DocumentConverterFactory.createConverter("test.doc");
        String docDescription = docConverter.getDescription();
        assertNotNull(docDescription);
        assertFalse(docDescription.isEmpty());
        assertTrue(docDescription.contains("DOC"));
        assertTrue(docDescription.contains("PDF"));
    }
    
    @Test
    @DisplayName("Null- und Edge-Cases werden korrekt behandelt")
    public void testEdgeCases() {
        // Null-Eingaben
        assertNull(DocumentConverterFactory.createConverter(null));
        assertFalse(DocumentConverterFactory.isSupported(null));
        
        // Leere Strings
        assertNull(DocumentConverterFactory.createConverter(""));
        assertFalse(DocumentConverterFactory.isSupported(""));
        
        // Dateien ohne Erweiterung
        assertNull(DocumentConverterFactory.createConverter("filename"));
        assertFalse(DocumentConverterFactory.isSupported("filename"));
        
        // Mehrere Punkte
        DocumentConverter converter = DocumentConverterFactory.createConverter("file.backup.docx");
        assertNotNull(converter);
        assertInstanceOf(DocxToPdfConverter.class, converter);
    }
}
