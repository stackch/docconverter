package ch.std.doc.converter.core.impl;

import ch.std.doc.converter.core.DocumentConverter;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Text;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.events.PdfDocumentEvent;
import com.itextpdf.kernel.events.IEventHandler;
import com.itextpdf.kernel.geom.Rectangle;
import com.itextpdf.kernel.pdf.canvas.PdfCanvas;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.io.font.constants.StandardFonts;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Spezialisierter Konverter für legacy Word-Dateien (.doc) zu PDF
 */
public class DocToPdfConverter extends DocumentConverter {
    
    private static final String[] SUPPORTED_EXTENSIONS = {".doc"};
    private static final String CONVERTER_NAME = "Word-Doc-Konverter";
    
    @Override
    public void convertToPdf(String inputFile, String outputFile) throws IOException {
        validateFiles(inputFile, outputFile);
        logConversion(inputFile, outputFile);
        
        convertDocToPdf(inputFile, outputFile);
    }
    
    @Override
    public String[] getSupportedExtensions() {
        return SUPPORTED_EXTENSIONS.clone();
    }
    
    @Override
    public String getConverterName() {
        return CONVERTER_NAME;
    }
    
    @Override
    public String getDescription() {
        return "Konvertiert legacy Microsoft Word DOC-Dateien zu PDF. " +
               "Extrahiert Text und grundlegende Formatierung aus älteren Word-Formaten. " +
               "Fallback-Mechanismus für komplexe Dokumente mit reiner Textextraktion.";
    }
    
    /**
     * Event Handler für Seitenzahlen
     */
    private static class DocPageNumberEventHandler implements IEventHandler {
        private PdfFont font;
        private int totalPages;
        
        public DocPageNumberEventHandler(int totalPages) {
            this.totalPages = totalPages;
            try {
                this.font = PdfFontFactory.createFont(StandardFonts.HELVETICA);
            } catch (Exception e) {
                System.err.println("Fehler beim Laden der Schriftart: " + e.getMessage());
            }
        }
        
        @Override
        public void handleEvent(com.itextpdf.kernel.events.Event event) {
            PdfDocumentEvent docEvent = (PdfDocumentEvent) event;
            PdfCanvas canvas = new PdfCanvas(docEvent.getPage());
            Rectangle pageSize = docEvent.getPage().getPageSize();
            
            if (font != null) {
                int pageNumber = docEvent.getDocument().getPageNumber(docEvent.getPage());
                String pageText = "Seite " + pageNumber + " von " + totalPages;
                
                float x = pageSize.getWidth() / 2;
                float y = 20;
                float textWidth = pageText.length() * 3.0f;
                
                canvas.beginText()
                      .setFontAndSize(font, 9)
                      .moveText(x - (textWidth / 2), y)
                      .showText(pageText)
                      .endText();
            }
            
            canvas.release();
        }
    }
    
    private void convertDocToPdf(String inputFile, String outputFile) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputFile);
             HWPFDocument docFile = new HWPFDocument(fis);
             FileOutputStream fos = new FileOutputStream(outputFile)) {
            
            PdfWriter writer = new PdfWriter(fos);
            PdfDocument pdfDoc = new PdfDocument(writer);
            Document document = new Document(pdfDoc, PageSize.A4);
            
            document.setMargins(72, 72, 72, 72);
            
            // Einfache Textextraktion für legacy .doc Dateien
            try {
                Range range = docFile.getRange();
                processDocRange(range, document);
            } catch (Exception e) {
                // Fallback: Einfache Textextraktion
                System.err.println("Warnung: Formatierte Extraktion fehlgeschlagen, verwende einfache Textextraktion: " + e.getMessage());
                extractPlainText(docFile, document);
            }
            
            // Seitenzahlen hinzufügen
            int totalPages = pdfDoc.getNumberOfPages();
            if (totalPages > 0) {
                pdfDoc.addEventHandler(PdfDocumentEvent.END_PAGE, new DocPageNumberEventHandler(totalPages));
            }
            
            document.close();
        }
    }
    
    private void processDocRange(Range range, Document document) {
        int numParagraphs = range.numParagraphs();
        
        for (int i = 0; i < numParagraphs; i++) {
            try {
                org.apache.poi.hwpf.usermodel.Paragraph para = range.getParagraph(i);
                com.itextpdf.layout.element.Paragraph pdfParagraph = new com.itextpdf.layout.element.Paragraph();
                
                // Charaktere in diesem Absatz verarbeiten
                int numCharacterRuns = para.numCharacterRuns();
                for (int j = 0; j < numCharacterRuns; j++) {
                    try {
                        CharacterRun run = para.getCharacterRun(j);
                        String text = run.text();
                        
                        if (text != null && !text.trim().isEmpty()) {
                            Text textElement = new Text(text);
                            
                            // Formatierung anwenden
                            if (run.isBold()) {
                                textElement.setBold();
                            }
                            if (run.isItalic()) {
                                textElement.setItalic();
                            }
                            
                            // Schriftgröße
                            int fontSize = run.getFontSize() / 2; // POI gibt Schriftgröße in halben Punkten zurück
                            if (fontSize > 6 && fontSize < 72) {
                                textElement.setFontSize(fontSize);
                            }
                            
                            pdfParagraph.add(textElement);
                        }
                    } catch (Exception e) {
                        // Ignoriere problematische Character Runs
                        System.err.println("Warnung: Fehler beim Verarbeiten eines Character Runs: " + e.getMessage());
                    }
                }
                
                // Nur nicht-leere Absätze hinzufügen
                if (!pdfParagraph.isEmpty()) {
                    document.add(pdfParagraph);
                } else if (para.text() != null && !para.text().trim().isEmpty()) {
                    // Fallback für Absätze ohne Character Runs
                    document.add(new com.itextpdf.layout.element.Paragraph(para.text()));
                }
                
            } catch (Exception e) {
                // Ignoriere problematische Absätze
                System.err.println("Warnung: Fehler beim Verarbeiten eines Absatzes: " + e.getMessage());
            }
        }
    }
    
    private void extractPlainText(HWPFDocument docFile, Document document) throws IOException {
        try (WordExtractor extractor = new WordExtractor(docFile)) {
            String text = extractor.getText();
            
            if (text != null && !text.trim().isEmpty()) {
                String[] lines = text.split("\n");
                for (String line : lines) {
                    if (!line.trim().isEmpty()) {
                        document.add(new com.itextpdf.layout.element.Paragraph(line));
                    }
                }
            } else {
                document.add(new com.itextpdf.layout.element.Paragraph("(Kein Text im Dokument gefunden)"));
            }
        }
    }
}
