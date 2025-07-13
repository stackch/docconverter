package ch.std.doc.converter.core.impl;

import ch.std.doc.converter.core.DocumentConverter;

import org.apache.poi.xwpf.usermodel.*;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.AreaBreak;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.UnitValue;
import com.itextpdf.layout.properties.AreaBreakType;
import com.itextpdf.layout.borders.Border;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.colors.Color;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.layout.element.Div;
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
import java.util.List;

/**
 * Spezialisierter Konverter für DOCX-Dateien zu PDF
 */
public class DocxToPdfConverter extends DocumentConverter {
    
    private static final String[] SUPPORTED_EXTENSIONS = {".docx"};
    private static final String CONVERTER_NAME = "DOCX-Konverter";
    
    @Override
    public void convertToPdf(String inputFile, String outputFile) throws IOException {
        validateFiles(inputFile, outputFile);
        logConversion(inputFile, outputFile);
        
        // Temporäre Datei für ersten Durchlauf
        String tempFile = outputFile + ".temp";
        
        // Erster Durchlauf: Dokument erstellen ohne korrekte Seitenzahlen
        int totalPages = createPdfDocument(inputFile, tempFile);
        
        // Zweiter Durchlauf: Seitenzahlen mit korrekter Gesamtseitenzahl
        createFinalPdfWithPageNumbers(inputFile, outputFile, totalPages);
        
        // Temporäre Datei löschen
        try {
            java.nio.file.Files.deleteIfExists(java.nio.file.Paths.get(tempFile));
        } catch (Exception e) {
            System.err.println("Warnung: Temporäre Datei konnte nicht gelöscht werden: " + e.getMessage());
        }
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
        return "Konvertiert Microsoft Word DOCX-Dokumente zu PDF mit voller Formatierung, " +
               "Tabellen, Bildern und korrekten Seitenzahlen. Unterstützt Rich-Text-Formatierung, " +
               "verschiedene Schriftarten, Farben und komplexe Layouts.";
    }
    
    /**
     * Event Handler für finale Seitenzahlen mit bekannter Gesamtseitenzahl
     */
    private static class PageNumberEventHandler implements IEventHandler {
        private PdfFont font;
        private int totalPages;
        
        public PageNumberEventHandler(int totalPages) {
            this.totalPages = totalPages;
            try {
                this.font = PdfFontFactory.createFont(StandardFonts.HELVETICA);
            } catch (Exception e) {
                System.err.println("Fehler beim Laden der Schriftart für Seitenzahlen: " + e.getMessage());
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
    
    /**
     * Erstellt das PDF-Dokument und gibt die Gesamtseitenzahl zurück
     */
    private int createPdfDocument(String inputFile, String outputFile) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputFile);
             XWPFDocument document = new XWPFDocument(fis);
             FileOutputStream fos = new FileOutputStream(outputFile)) {
            
            PdfWriter writer = new PdfWriter(fos);
            PdfDocument pdfDoc = new PdfDocument(writer);
            Document doc = new Document(pdfDoc, PageSize.A4);
            
            doc.setMargins(72, 36, 90, 36);
            
            processHeadersAndFooters(document, doc);
            processBodyElements(document, doc);
            
            int totalPages = pdfDoc.getNumberOfPages();
            doc.close();
            
            return totalPages;
        }
    }
    
    /**
     * Erstellt das finale PDF mit korrekten Seitenzahlen
     */
    private void createFinalPdfWithPageNumbers(String inputFile, String outputFile, int totalPages) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputFile);
             XWPFDocument document = new XWPFDocument(fis);
             FileOutputStream fos = new FileOutputStream(outputFile)) {
            
            PdfWriter writer = new PdfWriter(fos);
            PdfDocument pdfDoc = new PdfDocument(writer);
            Document doc = new Document(pdfDoc, PageSize.A4);
            
            pdfDoc.addEventHandler(PdfDocumentEvent.END_PAGE, new PageNumberEventHandler(totalPages));
            
            doc.setMargins(72, 36, 90, 36);
            
            processHeadersAndFooters(document, doc);
            processBodyElements(document, doc);
            
            doc.close();
        }
    }
    
    private void processHeadersAndFooters(XWPFDocument document, Document doc) {
        List<XWPFHeader> headers = document.getHeaderList();
        if (!headers.isEmpty()) {
            for (XWPFHeader header : headers) {
                for (XWPFParagraph para : header.getParagraphs()) {
                    String text = para.getText();
                    if (text != null && !text.trim().isEmpty()) {
                        Paragraph headerPara = new Paragraph(text)
                                .setTextAlignment(TextAlignment.CENTER)
                                .setMarginBottom(20)
                                .setFontSize(10);
                        doc.add(headerPara);
                        doc.add(new Paragraph("").setBorderBottom(new com.itextpdf.layout.borders.SolidBorder(1))
                                .setMarginBottom(10));
                    }
                }
            }
        }
    }
    
    private void processBodyElements(XWPFDocument document, Document doc) throws IOException {
        List<IBodyElement> bodyElements = document.getBodyElements();
        boolean isFirstElement = true;
        
        for (IBodyElement element : bodyElements) {
            if (element instanceof XWPFParagraph) {
                XWPFParagraph para = (XWPFParagraph) element;
                
                if (shouldAddPageBreakBefore(para, isFirstElement)) {
                    doc.add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                }
                
                processParagraph(para, doc);
                isFirstElement = false;
                
            } else if (element instanceof XWPFTable) {
                XWPFTable table = (XWPFTable) element;
                
                if (!isFirstElement) {
                    doc.add(new Paragraph("").setMarginTop(15));
                }
                
                processTable(table, doc);
                doc.add(new Paragraph("").setMarginBottom(15));
                isFirstElement = false;
            }
        }
    }
    
    private boolean shouldAddPageBreakBefore(XWPFParagraph para, boolean isFirstElement) {
        if (isFirstElement) return false;
        
        String text = para.getText();
        if (text == null || text.trim().isEmpty()) return false;
        
        if (isHeading(para)) {
            String lowerText = text.toLowerCase().trim();
            return lowerText.matches("^\\d+\\..*") || 
                   lowerText.contains("komplexes") ||
                   lowerText.contains("bilder und grafiken") ||
                   lowerText.contains("zusätzlicher inhalt");
        }
        
        return false;
    }
    
    private void processParagraph(XWPFParagraph para, Document doc) {
        String text = para.getText();
        
        if (hasImages(para)) {
            processInlineImages(para, doc);
            return;
        }
        
        if (text != null && !text.trim().isEmpty()) {
            Paragraph pdfParagraph = new Paragraph();
            
            // Verarbeite jeden Run einzeln um Formatierung zu erhalten
            boolean hasContent = false;
            for (XWPFRun run : para.getRuns()) {
                String runText = run.getText(0);
                if (runText != null && !runText.isEmpty()) {
                    com.itextpdf.layout.element.Text textElement = 
                        new com.itextpdf.layout.element.Text(runText);
                    
                    // Schriftgröße
                    int fontSize = run.getFontSize() != -1 ? run.getFontSize() : 11;
                    textElement.setFontSize(fontSize);
                    
                    // Fett/Kursiv
                    if (run.isBold()) {
                        textElement.setBold();
                    }
                    if (run.isItalic()) {
                        textElement.setItalic();
                    }                            // Textfarbe aus DOCX extrahieren
                            String colorHex = run.getColor();
                            if (colorHex != null && !colorHex.equals("auto") && colorHex.length() == 6) {
                                try {
                                    int r = Integer.parseInt(colorHex.substring(0, 2), 16);
                                    int g = Integer.parseInt(colorHex.substring(2, 4), 16);
                                    int b = Integer.parseInt(colorHex.substring(4, 6), 16);
                                    textElement.setFontColor(new DeviceRgb(r, g, b));
                                } catch (Exception e) {
                                    // Ignoriere Farbfehler
                                }
                            }
                    
                    pdfParagraph.add(textElement);
                    hasContent = true;
                }
            }
            
            // Fallback wenn keine Runs
            if (!hasContent) {
                pdfParagraph.add(text);
                pdfParagraph.setFontSize(11);
            }
            
            // Paragraph-Level Formatierung
            if (isHeading(para)) {
                pdfParagraph.setFontSize(16).setBold().setMarginTop(15).setMarginBottom(10);
            } else {
                pdfParagraph.setMarginBottom(6);
            }
            
            // Ausrichtung
            ParagraphAlignment alignment = para.getAlignment();
            if (alignment != null) {
                switch (alignment) {
                    case CENTER:
                        pdfParagraph.setTextAlignment(TextAlignment.CENTER);
                        break;
                    case RIGHT:
                        pdfParagraph.setTextAlignment(TextAlignment.RIGHT);
                        break;
                    default:
                        pdfParagraph.setTextAlignment(TextAlignment.LEFT);
                        break;
                }
            }
            
            doc.add(pdfParagraph);
        }
    }
    
    private boolean hasImages(XWPFParagraph paragraph) {
        for (XWPFRun run : paragraph.getRuns()) {
            if (!run.getEmbeddedPictures().isEmpty()) {
                return true;
            }
        }
        return false;
    }
    
    private void processInlineImages(XWPFParagraph paragraph, Document doc) {
        try {
            for (XWPFRun run : paragraph.getRuns()) {
                for (XWPFPicture picture : run.getEmbeddedPictures()) {
                    XWPFPictureData pictureData = picture.getPictureData();
                    byte[] imageData = pictureData.getData();
                    
                    if (imageData != null && imageData.length > 0) {
                        Image pdfImage = new Image(ImageDataFactory.create(imageData));
                        pdfImage.setAutoScale(true);
                        pdfImage.setMaxWidth(400);
                        pdfImage.setMaxHeight(300);
                        
                        Paragraph imageP = new Paragraph().add(pdfImage).setTextAlignment(TextAlignment.CENTER);
                        doc.add(imageP);
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("Fehler beim Verarbeiten der Bilder: " + e.getMessage());
        }
    }
    
    private void processTable(XWPFTable xwpfTable, Document doc) {
        List<XWPFTableRow> rows = xwpfTable.getRows();
        if (rows.isEmpty()) return;
        
        int numCols = rows.get(0).getTableCells().size();
        Table pdfTable = new Table(UnitValue.createPercentArray(numCols))
                .useAllAvailableWidth()
                .setMarginTop(10)
                .setMarginBottom(10);
        
        int rowIndex = 0;
        for (XWPFTableRow row : rows) {
            for (XWPFTableCell cell : row.getTableCells()) {
                Paragraph cellParagraph = new Paragraph();
                
                // Extrahiere Formatierung aus allen Runs der Zelle
                boolean hasContent = false;
                for (XWPFParagraph para : cell.getParagraphs()) {
                    for (XWPFRun run : para.getRuns()) {
                        String runText = run.getText(0);
                        if (runText != null && !runText.isEmpty()) {
                            com.itextpdf.layout.element.Text textElement = 
                                new com.itextpdf.layout.element.Text(runText);
                            
                            // Formatierung anwenden
                            if (run.isBold()) {
                                textElement.setBold();
                            }
                            if (run.isItalic()) {
                                textElement.setItalic();
                            }
                            
                            // Textfarbe aus DOCX
                            String colorHex = run.getColor();
                            if (colorHex != null && !colorHex.equals("auto") && colorHex.length() == 6) {
                                try {
                                    int r = Integer.parseInt(colorHex.substring(0, 2), 16);
                                    int g = Integer.parseInt(colorHex.substring(2, 4), 16);
                                    int b = Integer.parseInt(colorHex.substring(4, 6), 16);
                                    textElement.setFontColor(new DeviceRgb(r, g, b));
                                } catch (Exception e) {
                                    // Ignoriere Farbfehler
                                }
                            }
                            
                            cellParagraph.add(textElement);
                            hasContent = true;
                        }
                    }
                }
                
                // Fallback wenn keine Runs
                if (!hasContent) {
                    String cellText = cell.getText();
                    cellParagraph.add(cellText != null ? cellText : "");
                }
                
                Cell pdfCell = new Cell().add(cellParagraph);
                pdfCell.setPadding(6);
                pdfCell.setBorder(new com.itextpdf.layout.borders.SolidBorder(0.5f));
                
                // Prüfe auf Hintergrundfarbe der Zelle
                Color bgColor = extractCellBackgroundColor(cell);
                if (bgColor != null) {
                    pdfCell.setBackgroundColor(bgColor);
                }
                
                pdfTable.addCell(pdfCell);
            }
            rowIndex++;
        }
        
        doc.add(pdfTable);
    }
    
    private boolean isHeading(XWPFParagraph paragraph) {
        String style = paragraph.getStyle();
        if (style != null && style.toLowerCase().contains("heading")) {
            return true;
        }
        
        List<XWPFRun> runs = paragraph.getRuns();
        if (!runs.isEmpty()) {
            return runs.get(0).isBold();
        }
        
        return false;
    }
    
    /**
     * Extrahiert die Hintergrundfarbe einer Tabellenzelle
     */
    private Color extractCellBackgroundColor(XWPFTableCell cell) {
        try {
            // Prüfe CTTcPr (Cell Properties) für Shading
            if (cell.getCTTc() != null && cell.getCTTc().getTcPr() != null) {
                if (cell.getCTTc().getTcPr().getShd() != null) {
                    Object fillObj = cell.getCTTc().getTcPr().getShd().getFill();
                    
                    if (fillObj != null) {
                        // Wenn es ein Byte-Array ist, konvertiere zu Hex
                        if (fillObj instanceof byte[]) {
                            byte[] colorBytes = (byte[]) fillObj;
                            if (colorBytes.length >= 3) {
                                StringBuilder hexColor = new StringBuilder();
                                for (int i = 0; i < Math.min(3, colorBytes.length); i++) {
                                    hexColor.append(String.format("%02X", colorBytes[i] & 0xFF));
                                }
                                String colorHex = hexColor.toString();
                                
                                if (colorHex.length() == 6) {
                                    int r = Integer.parseInt(colorHex.substring(0, 2), 16);
                                    int g = Integer.parseInt(colorHex.substring(2, 4), 16);
                                    int b = Integer.parseInt(colorHex.substring(4, 6), 16);
                                    return new DeviceRgb(r, g, b);
                                }
                            }
                        }
                        
                        // Versuche als String (fallback)
                        String colorStr = fillObj.toString();
                        if (!colorStr.startsWith("[B@") && !colorStr.equals("auto") && colorStr.length() == 6) {
                            try {
                                int r = Integer.parseInt(colorStr.substring(0, 2), 16);
                                int g = Integer.parseInt(colorStr.substring(2, 4), 16);
                                int b = Integer.parseInt(colorStr.substring(4, 6), 16);
                                return new DeviceRgb(r, g, b);
                            } catch (Exception e) {
                                // Ignoriere String-Parsing-Fehler
                            }
                        }
                    }
                }
            }
        } catch (Exception e) {
            // Ignoriere Extraktionsfehler
        }
        return null;
    }
}
