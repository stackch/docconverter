package ch.std.doc.converter;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.io.TempDir;
import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import javax.imageio.ImageIO;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.util.Units;

/**
 * Tests für die DocConverter Klasse mit erweiterten Features
 */
public class DocConverterTest {
    
    private DocConverter converter;
    
    @BeforeEach
    public void setUp() {
        converter = new DocConverter();
    }
    
    @Test
    public void testDocConverterExists() {
        assertNotNull(converter);
    }
    
    @Test
    public void testConvertDocxToPdf(@TempDir Path tempDir) throws IOException {
        // Erstelle eine Test-DOCX-Datei
        File inputFile = tempDir.resolve("test.docx").toFile();
        File outputFile = tempDir.resolve("output.pdf").toFile();
        
        createTestDocxFile(inputFile);
        
        // Führe die Konvertierung durch
        assertDoesNotThrow(() -> {
            converter.convertDocxToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
        });
        
        // Prüfe, ob die PDF-Datei erstellt wurde
        assertTrue(outputFile.exists());
        assertTrue(outputFile.length() > 0);
    }
    
    @Test
    public void testConvertComplexDocxToPdf(@TempDir Path tempDir) throws IOException {
        // Erstelle eine komplexe Test-DOCX-Datei mit Tabellen und Bildern
        File inputFile = tempDir.resolve("complex_test.docx").toFile();
        File outputFile = tempDir.resolve("complex_output.pdf").toFile();
        
        createComplexTestDocxFile(inputFile);
        
        // Führe die Konvertierung durch
        assertDoesNotThrow(() -> {
            converter.convertDocxToPdf(inputFile.getAbsolutePath(), outputFile.getAbsolutePath());
        });
        
        // Prüfe, ob die PDF-Datei erstellt wurde und größer ist (wegen Tabellen/Bildern)
        assertTrue(outputFile.exists());
        assertTrue(outputFile.length() > 1000); // Komplexere Datei sollte größer sein
    }
    
    @Test
    public void testConvertDocumentLegacy() {
        // Test für die Legacy-Methode
        assertDoesNotThrow(() -> {
            converter.convertDocument("test.docx", "output.pdf", "PDF");
        });
    }
    
    /**
     * Erstellt eine einfache Test-DOCX-Datei
     */
    private void createTestDocxFile(File file) throws IOException {
        try (XWPFDocument document = new XWPFDocument();
             FileOutputStream out = new FileOutputStream(file)) {
            
            // Überschrift hinzufügen
            XWPFParagraph title = document.createParagraph();
            title.createRun().setText("Test Dokument");
            title.getRuns().get(0).setBold(true);
            title.getRuns().get(0).setFontSize(16);
            
            // Normaler Text hinzufügen
            XWPFParagraph content = document.createParagraph();
            content.createRun().setText("Dies ist ein Test-Dokument für die DOCX zu PDF Konvertierung.");
            
            // Weiterer Absatz
            XWPFParagraph content2 = document.createParagraph();
            content2.createRun().setText("Dieser zweite Absatz testet die Formatierung und den Umgang mit mehreren Paragraphen.");
            
            document.write(out);
        }
    }
    
    /**
     * Erstellt eine komplexe Test-DOCX-Datei mit Tabellen und Bildern
     */
    private void createComplexTestDocxFile(File file) throws IOException {
        try (XWPFDocument document = new XWPFDocument();
             FileOutputStream out = new FileOutputStream(file)) {
            
            // Titel
            XWPFParagraph title = document.createParagraph();
            title.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun titleRun = title.createRun();
            titleRun.setText("Komplexes Test-Dokument");
            titleRun.setBold(true);
            titleRun.setFontSize(18);
            
            // Beschreibung
            XWPFParagraph desc = document.createParagraph();
            XWPFRun descRun = desc.createRun();
            descRun.setText("Dieses Dokument testet die Konvertierung von Tabellen und Bildern.");
            
            // Tabelle erstellen
            createTestTable(document);
            
            // Bild hinzufügen
            addTestImage(document);
            
            // Abschlusstext
            XWPFParagraph conclusion = document.createParagraph();
            XWPFRun conclusionRun = conclusion.createRun();
            conclusionRun.setText("Ende des Test-Dokuments mit Tabellen und Bildern.");
            
            document.write(out);
        }
    }
    
    /**
     * Erstellt eine Test-Tabelle mit Farben
     */
    private void createTestTable(XWPFDocument document) {
        XWPFTable table = document.createTable(3, 3);
        
        // Header-Zeile mit blauem Hintergrund
        XWPFTableRow headerRow = table.getRow(0);
        
        // Header-Zellen mit Farbe formatieren
        formatHeaderCell(headerRow.getCell(0), "Name");
        formatHeaderCell(headerRow.getCell(1), "Alter");
        formatHeaderCell(headerRow.getCell(2), "Stadt");
        
        // Datenzeilen mit abwechselnden Farben
        XWPFTableRow row1 = table.getRow(1);
        row1.getCell(0).setText("Max Mustermann");
        row1.getCell(1).setText("25");
        row1.getCell(2).setText("Berlin");
        // Hellgrauer Hintergrund für ungerade Zeilen
        setRowBackgroundColor(row1, "F2F2F2");
        
        XWPFTableRow row2 = table.getRow(2);
        row2.getCell(0).setText("Anna Schmidt");
        row2.getCell(1).setText("30");
        row2.getCell(2).setText("München");
        // Weiß bleibt Standard
    }
    
    /**
     * Formatiert eine Header-Zelle mit blauem Hintergrund und weißer Schrift
     */
    private void formatHeaderCell(XWPFTableCell cell, String text) {
        cell.setText(text);
        
        // Blauen Hintergrund setzen
        cell.setColor("4472C4");
        
        // Text formatieren
        java.util.List<XWPFParagraph> paragraphs = cell.getParagraphs();
        if (!paragraphs.isEmpty()) {
            XWPFParagraph para = paragraphs.get(0);
            para.setAlignment(ParagraphAlignment.CENTER);
            
            java.util.List<XWPFRun> runs = para.getRuns();
            if (!runs.isEmpty()) {
                XWPFRun run = runs.get(0);
                run.setBold(true);
                run.setColor("FFFFFF"); // Weiße Schrift
            }
        }
    }
    
    /**
     * Setzt Hintergrundfarbe für eine ganze Zeile
     */
    private void setRowBackgroundColor(XWPFTableRow row, String colorHex) {
        java.util.List<XWPFTableCell> cells = row.getTableCells();
        for (XWPFTableCell cell : cells) {
            cell.setColor(colorHex);
        }
    }
    
    /**
     * Fügt ein Test-Bild hinzu
     */
    private void addTestImage(XWPFDocument document) throws IOException {
        // Einfaches Test-Bild erstellen
        BufferedImage image = new BufferedImage(200, 100, BufferedImage.TYPE_INT_RGB);
        Graphics2D g2d = image.createGraphics();
        g2d.setColor(Color.BLUE);
        g2d.fillRect(0, 0, 200, 100);
        g2d.setColor(Color.WHITE);
        g2d.drawString("Test-Bild", 70, 50);
        g2d.dispose();
        
        // Bild als Byte-Array
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(image, "PNG", baos);
        byte[] imageData = baos.toByteArray();
        
        // Bild zum Dokument hinzufügen
        XWPFParagraph imagePara = document.createParagraph();
        imagePara.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun imageRun = imagePara.createRun();
        
        try {
            imageRun.addPicture(new java.io.ByteArrayInputStream(imageData),
                               XWPFDocument.PICTURE_TYPE_PNG, "test-image.png",
                               Units.toEMU(200), Units.toEMU(100));
        } catch (Exception e) {
            // Falls Bild nicht hinzugefügt werden kann, Text als Alternative
            imageRun.setText("[Test-Bild konnte nicht hinzugefügt werden]");
        }
    }
}
