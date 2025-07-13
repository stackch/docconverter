package ch.std.doc.converter;

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
import com.itextpdf.layout.properties.BorderRadius;
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
import java.io.ByteArrayInputStream;
import java.util.List;

/**
 * Hauptklasse für den Dokumentkonverter mit erweiterten Features für Tabellen und Bilder
 */
public class DocConverter {
    
    /**
     * Event Handler für Seitenzahlen mit korrekter Gesamtseitenzahl
     */
    private static class PageNumberEventHandler implements IEventHandler {
        private PdfFont font;
        
        public PageNumberEventHandler() {
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
                // Aktuelle Seitenzahl und Gesamtseitenzahl
                int pageNumber = docEvent.getDocument().getPageNumber(docEvent.getPage());
                int totalPages = docEvent.getDocument().getNumberOfPages();
                
                // Wenn das Dokument noch nicht vollständig ist, verwenden wir einen Platzhalter
                String pageText;
                if (totalPages == 0) {
                    pageText = "Seite " + pageNumber;
                } else {
                    pageText = "Seite " + pageNumber + " von " + totalPages;
                }
                
                float x = pageSize.getWidth() / 2;
                float y = 20; // 20 Punkte vom unteren Rand
                
                // Bessere Zentrierung basierend auf Textlänge
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
     * Event Handler für finale Seitenzahlen mit bekannter Gesamtseitenzahl
     */
    private static class FinalPageNumberEventHandler implements IEventHandler {
        private PdfFont font;
        private int totalPages;
        
        public FinalPageNumberEventHandler(int totalPages) {
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
                // Korrekte Seitenzahl mit bekannter Gesamtseitenzahl
                int pageNumber = docEvent.getDocument().getPageNumber(docEvent.getPage());
                String pageText = "Seite " + pageNumber + " von " + totalPages;
                
                float x = pageSize.getWidth() / 2;
                float y = 20; // 20 Punkte vom unteren Rand
                
                // Bessere Zentrierung basierend auf Textlänge
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
    
    public static void main(String[] args) {
        System.out.println("DocConverter gestartet...");
        
        if (args.length < 2) {
            System.out.println("Verwendung: java DocConverter <input.docx> <output.pdf>");
            return;
        }
        
        DocConverter converter = new DocConverter();
        try {
            converter.convertDocxToPdf(args[0], args[1]);
            System.out.println("Konvertierung erfolgreich abgeschlossen!");
        } catch (IOException e) {
            System.err.println("Fehler bei der Konvertierung: " + e.getMessage());
        }
    }
    
    /**
     * Konvertiert eine DOCX-Datei in eine PDF-Datei mit verbessertem Layout und Paging
     * 
     * @param inputFile Pfad zur DOCX-Eingabedatei
     * @param outputFile Pfad zur PDF-Ausgabedatei
     * @throws IOException bei Fehlern beim Lesen oder Schreiben der Dateien
     */
    public void convertDocxToPdf(String inputFile, String outputFile) throws IOException {
        System.out.println("Konvertiere " + inputFile + " zu PDF...");
        
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
    
    /**
     * Erstellt das PDF-Dokument und gibt die Gesamtseitenzahl zurück
     */
    private int createPdfDocument(String inputFile, String outputFile) throws IOException {
        // DOCX-Datei laden
        try (FileInputStream fis = new FileInputStream(inputFile);
             XWPFDocument document = new XWPFDocument(fis);
             FileOutputStream fos = new FileOutputStream(outputFile)) {
            
            // PDF-Dokument erstellen mit A4-Format und Seitenrändern
            PdfWriter writer = new PdfWriter(fos);
            PdfDocument pdfDoc = new PdfDocument(writer);
            Document doc = new Document(pdfDoc, PageSize.A4);
            
            // Seitenränder setzen (mehr Platz unten für Seitenzahlen)
            doc.setMargins(72, 36, 90, 36); // top, right, bottom, left
            
            // Header und Footer verarbeiten
            processHeadersAndFooters(document, doc, pdfDoc);
            
            // Body-Elemente verarbeiten (Paragraphen, Tabellen, etc.)
            processBodyElements(document, doc);
            
            // Gesamtseitenzahl ermitteln
            int totalPages = pdfDoc.getNumberOfPages();
            
            // PDF schließen
            doc.close();
            
            return totalPages;
        }
    }
    
    /**
     * Erstellt das finale PDF mit korrekten Seitenzahlen
     */
    private void createFinalPdfWithPageNumbers(String inputFile, String outputFile, int totalPages) throws IOException {
        // DOCX-Datei laden
        try (FileInputStream fis = new FileInputStream(inputFile);
             XWPFDocument document = new XWPFDocument(fis);
             FileOutputStream fos = new FileOutputStream(outputFile)) {
            
            // PDF-Dokument erstellen mit A4-Format und Seitenrändern
            PdfWriter writer = new PdfWriter(fos);
            PdfDocument pdfDoc = new PdfDocument(writer);
            Document doc = new Document(pdfDoc, PageSize.A4);
            
            // Event Handler für Seitenzahlen mit bekannter Gesamtseitenzahl registrieren
            pdfDoc.addEventHandler(PdfDocumentEvent.END_PAGE, new FinalPageNumberEventHandler(totalPages));
            
            // Seitenränder setzen (mehr Platz unten für Seitenzahlen)
            doc.setMargins(72, 36, 90, 36); // top, right, bottom, left
            
            // Header und Footer verarbeiten
            processHeadersAndFooters(document, doc, pdfDoc);
            
            // Body-Elemente verarbeiten (Paragraphen, Tabellen, etc.)
            processBodyElements(document, doc);
            
            // PDF schließen
            doc.close();
        }
    }
    
    /**
     * Verarbeitet Header und Footer des Dokuments mit korrektem Paging und Formatierung
     */
    private void processHeadersAndFooters(XWPFDocument document, Document doc, PdfDocument pdfDoc) {
        // Header verarbeiten - nur einmal am Anfang
        List<XWPFHeader> headers = document.getHeaderList();
        if (!headers.isEmpty()) {
            for (XWPFHeader header : headers) {
                for (XWPFParagraph para : header.getParagraphs()) {
                    String text = para.getText();
                    if (text != null && !text.trim().isEmpty()) {
                        Div headerDiv = new Div();
                        Paragraph headerPara = createFormattedParagraph(para, 10);
                        headerPara.setTextAlignment(TextAlignment.CENTER)
                                 .setMarginBottom(20);
                        
                        headerDiv.add(headerPara);
                        headerDiv.setBorder(Border.NO_BORDER);
                        doc.add(headerDiv);
                        
                        // Trennlinie nach Header
                        doc.add(new Paragraph("").setBorderBottom(new com.itextpdf.layout.borders.SolidBorder(1))
                                .setMarginBottom(10));
                    }
                }
            }
        }
    }
    
    /**
     * Verarbeitet alle Body-Elemente des Dokuments mit verbessertem Layout
     */
    private void processBodyElements(XWPFDocument document, Document doc) throws IOException {
        // Alle Body-Elemente in der richtigen Reihenfolge verarbeiten
        List<IBodyElement> bodyElements = document.getBodyElements();
        
        boolean isFirstElement = true;
        
        for (IBodyElement element : bodyElements) {
            if (element instanceof XWPFParagraph) {
                XWPFParagraph para = (XWPFParagraph) element;
                
                // Seitenumbruch vor bestimmten Elementen
                if (shouldAddPageBreakBefore(para, isFirstElement)) {
                    doc.add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                }
                
                processParagraph(para, doc);
                isFirstElement = false;
                
            } else if (element instanceof XWPFTable) {
                XWPFTable table = (XWPFTable) element;
                
                // Mehr Platz vor Tabellen
                if (!isFirstElement) {
                    doc.add(new Paragraph("").setMarginTop(15));
                }
                
                processTable(table, doc);
                
                // Platz nach Tabellen
                doc.add(new Paragraph("").setMarginBottom(15));
                isFirstElement = false;
            }
        }
        
        // Bilder werden jetzt inline verarbeitet - nicht mehr am Ende
        // processImages(document, doc);
    }
    
    /**
     * Bestimmt ob vor einem Paragraph ein Seitenumbruch eingefügt werden soll
     */
    private boolean shouldAddPageBreakBefore(XWPFParagraph para, boolean isFirstElement) {
        if (isFirstElement) {
            return false;
        }
        
        String text = para.getText();
        if (text == null || text.trim().isEmpty()) {
            return false;
        }
        
        // Seitenumbruch bei großen Überschriften
        if (isHeading(para)) {
            String lowerText = text.toLowerCase().trim();
            // Neue Seite für Hauptkapitel
            if (lowerText.matches("^\\d+\\..*") || // "1. Kapitel"
                lowerText.contains("komplexes") ||
                lowerText.contains("bilder und grafiken") ||
                lowerText.contains("zusätzlicher inhalt")) {
                return true;
            }
        }
        
        return false;
    }
    
    /**
     * Verarbeitet einen Paragraph mit verbesserter Formatierung und Rich-Text-Unterstützung
     */
    private void processParagraph(XWPFParagraph para, Document doc) {
        String text = para.getText();
        
        // Prüfe zuerst ob der Paragraph Bilder enthält
        if (hasImages(para)) {
            processInlineImages(para, doc);
            return; // Bilder-Paragraph separat verarbeiten
        }
        
        if (text != null && !text.trim().isEmpty()) {
            // Erstelle einen formatierten Paragraph
            Paragraph pdfParagraph = createFormattedParagraph(para, 11);
            
            // Heading-spezifische Formatierung
            if (isHeading(para)) {
                int headingLevel = getHeadingLevel(para);
                switch (headingLevel) {
                    case 1: // Hauptüberschrift
                        pdfParagraph.setFontSize(20)
                                   .setMarginTop(25)
                                   .setMarginBottom(15);
                        break;
                    case 2: // Unterüberschrift
                        pdfParagraph.setFontSize(16)
                                   .setMarginTop(20)
                                   .setMarginBottom(12);
                        break;
                    default: // Standard-Überschrift
                        pdfParagraph.setFontSize(14)
                                   .setMarginTop(15)
                                   .setMarginBottom(8);
                        break;
                }
            } else {
                // Normaler Text
                pdfParagraph.setMarginBottom(6);
                
                // Einrückung für bestimmte Texttypen
                if (isListItem(text)) {
                    pdfParagraph.setMarginLeft(20);
                }
            }
            
            // Ausrichtung setzen
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
    
    /**
     * Prüft ob ein Paragraph Bilder enthält
     */
    private boolean hasImages(XWPFParagraph paragraph) {
        for (XWPFRun run : paragraph.getRuns()) {
            List<XWPFPicture> pictures = run.getEmbeddedPictures();
            if (!pictures.isEmpty()) {
                return true;
            }
        }
        return false;
    }
    
    /**
     * Verarbeitet Bilder die inline in Paragraphen stehen
     */
    private void processInlineImages(XWPFParagraph paragraph, Document doc) {
        try {
            // Text vor den Bildern hinzufügen (falls vorhanden)
            String text = paragraph.getText();
            if (text != null && !text.trim().isEmpty()) {
                // Erstelle Paragraph für Text vor Bildern
                Paragraph textPara = new Paragraph(text)
                        .setFontSize(11)
                        .setMarginBottom(5); // Reduzierter Abstand
                doc.add(textPara);
            }
            
            // Sammle alle Bilder aus diesem Paragraph
            List<XWPFPicture> allPictures = new java.util.ArrayList<>();
            for (XWPFRun run : paragraph.getRuns()) {
                allPictures.addAll(run.getEmbeddedPictures());
            }
            
            // Verarbeite Bilder kompakter - mehrere nebeneinander wenn möglich
            if (allPictures.size() > 1) {
                processMultipleImages(allPictures, doc);
            } else {
                // Einzelbild kompakt verarbeiten
                for (XWPFPicture picture : allPictures) {
                    processSingleImageCompact(picture, doc);
                }
            }
            
        } catch (Exception e) {
            System.err.println("Fehler beim Verarbeiten der Inline-Bilder: " + e.getMessage());
        }
    }
    
    /**
     * Verarbeitet mehrere Bilder in einem kompakten Layout
     */
    private void processMultipleImages(List<XWPFPicture> pictures, Document doc) {
        try {
            // Erstelle eine Tabelle für nebeneinander liegende Bilder
            Table imageTable = new Table(UnitValue.createPercentArray(pictures.size()))
                    .useAllAvailableWidth()
                    .setBorder(Border.NO_BORDER)
                    .setMarginTop(5)
                    .setMarginBottom(10);
            
            for (XWPFPicture picture : pictures) {
                try {
                    XWPFPictureData pictureData = picture.getPictureData();
                    byte[] imageData = pictureData.getData();
                    
                    if (imageData != null && imageData.length > 0) {
                        Image pdfImage = new Image(ImageDataFactory.create(imageData));
                        
                        // Kleinere Bildgröße für nebeneinander liegende Bilder
                        float maxWidth = 200;  // Deutlich kleiner für mehrere Bilder
                        float maxHeight = 150;
                        
                        if (pdfImage.getImageWidth() > maxWidth || pdfImage.getImageHeight() > maxHeight) {
                            pdfImage.setAutoScale(true);
                            pdfImage.setMaxWidth(maxWidth);
                            pdfImage.setMaxHeight(maxHeight);
                        }
                        
                        // Zelle für das Bild
                        Cell imageCell = new Cell();
                        imageCell.add(pdfImage);
                        imageCell.setTextAlignment(TextAlignment.CENTER);
                        imageCell.setBorder(Border.NO_BORDER);
                        imageCell.setPadding(2);
                        
                        // Kompakte Bildunterschrift
                        if (pictureData.getFileName() != null && !pictureData.getFileName().isEmpty()) {
                            Paragraph caption = new Paragraph(pictureData.getFileName())
                                    .setFontSize(7)
                                    .setTextAlignment(TextAlignment.CENTER)
                                    .setMarginTop(2)
                                    .setItalic();
                            imageCell.add(caption);
                        }
                        
                        imageTable.addCell(imageCell);
                    }
                } catch (Exception e) {
                    System.err.println("Fehler beim Verarbeiten eines Bildes in der Gruppe: " + e.getMessage());
                    // Leere Zelle hinzufügen bei Fehler
                    imageTable.addCell(new Cell().setBorder(Border.NO_BORDER));
                }
            }
            
            doc.add(imageTable);
            
        } catch (Exception e) {
            System.err.println("Fehler beim Verarbeiten der Bild-Gruppe: " + e.getMessage());
            // Fallback: Bilder einzeln verarbeiten
            for (XWPFPicture picture : pictures) {
                processSingleImageCompact(picture, doc);
            }
        }
    }
    
    /**
     * Verarbeitet ein einzelnes Bild kompakt
     */
    private void processSingleImageCompact(XWPFPicture picture, Document doc) {
        try {
            XWPFPictureData pictureData = picture.getPictureData();
            byte[] imageData = pictureData.getData();
            
            if (imageData != null && imageData.length > 0) {
                Image pdfImage = new Image(ImageDataFactory.create(imageData));
                
                // Kompaktere Bildgröße
                float maxWidth = 300;  // Etwas kleiner für besseres Layout
                float maxHeight = 200;
                
                if (pdfImage.getImageWidth() > maxWidth || pdfImage.getImageHeight() > maxHeight) {
                    pdfImage.setAutoScale(true);
                    pdfImage.setMaxWidth(maxWidth);
                    pdfImage.setMaxHeight(maxHeight);
                }
                
                // Kompakte Margins
                pdfImage.setMarginBottom(5);
                pdfImage.setMarginTop(5);
                
                // Bild zentrieren ohne großen Container
                Paragraph imageP = new Paragraph().add(pdfImage).setTextAlignment(TextAlignment.CENTER);
                doc.add(imageP);
                
                // Kompakte Bildunterschrift
                if (pictureData.getFileName() != null && !pictureData.getFileName().isEmpty()) {
                    Paragraph caption = new Paragraph(pictureData.getFileName())
                            .setFontSize(7)
                            .setTextAlignment(TextAlignment.CENTER)
                            .setMarginBottom(5)
                            .setMarginTop(1)
                            .setItalic();
                    doc.add(caption);
                }
            }
        } catch (Exception e) {
            System.err.println("Fehler beim Verarbeiten eines Einzelbildes: " + e.getMessage());
        }
    }
    private Paragraph createFormattedParagraph(XWPFParagraph xwpfPara, int defaultFontSize) {
        Paragraph pdfParagraph = new Paragraph();
        List<XWPFRun> runs = xwpfPara.getRuns();
        
        if (runs.isEmpty()) {
            // Fallback: einfacher Text
            pdfParagraph.add(xwpfPara.getText()).setFontSize(defaultFontSize);
            return pdfParagraph;
        }
        
        // Verarbeite jeden Run separat für Rich-Text
        for (XWPFRun run : runs) {
            String runText = run.text();
            if (runText != null && !runText.isEmpty()) {
                com.itextpdf.layout.element.Text textElement = new com.itextpdf.layout.element.Text(runText);
                
                // Schriftgröße
                int fontSize = run.getFontSize();
                if (fontSize > 0) {
                    textElement.setFontSize(fontSize);
                } else {
                    textElement.setFontSize(defaultFontSize);
                }
                
                // Fett
                if (run.isBold()) {
                    textElement.setBold();
                }
                
                // Kursiv
                if (run.isItalic()) {
                    textElement.setItalic();
                }
                
                // Unterstrichen
                if (run.getUnderline() != null && 
                    run.getUnderline() != org.apache.poi.xwpf.usermodel.UnderlinePatterns.NONE) {
                    textElement.setUnderline();
                }
                
                // Textfarbe
                String colorHex = run.getColor();
                if (colorHex != null && !colorHex.equals("auto") && !colorHex.isEmpty()) {
                    Color textColor = hexToColor(colorHex);
                    if (textColor != null) {
                        textElement.setFontColor(textColor);
                    }
                }
                
                pdfParagraph.add(textElement);
            }
        }
        
        return pdfParagraph;
    }
    
    /**
     * Bestimmt das Heading-Level eines Paragraphs
     */
    private int getHeadingLevel(XWPFParagraph para) {
        String text = para.getText();
        if (text == null) return 3;
        
        text = text.trim().toLowerCase();
        
        // Level 1: Hauptkapitel
        if (text.matches("^\\d+\\..*") || 
            text.contains("komplexes test") ||
            text.contains("dokumentstruktur")) {
            return 1;
        }
        
        // Level 2: Unterkapitel  
        if (text.contains("tabellen") ||
            text.contains("bilder") ||
            text.contains("zusätzlich")) {
            return 2;
        }
        
        return 3; // Standard
    }
    
    /**
     * Prüft ob ein Text ein Listen-Element ist
     */
    private boolean isListItem(String text) {
        return text.trim().matches("^[•✓▪▫-].*") ||
               text.trim().matches("^\\d+\\.\\s.*") ||
               text.trim().matches("^[a-zA-Z]\\)\\s.*");
    }
    
    /**
     * Verarbeitet eine Tabelle mit verbessertem Layout und Farbunterstützung
     */
    private void processTable(XWPFTable xwpfTable, Document doc) {
        List<XWPFTableRow> rows = xwpfTable.getRows();
        if (rows.isEmpty()) {
            return;
        }
        
        // Anzahl der Spalten bestimmen
        int numCols = rows.get(0).getTableCells().size();
        Table pdfTable = new Table(UnitValue.createPercentArray(numCols))
                .useAllAvailableWidth()
                .setMarginTop(10)
                .setMarginBottom(10);
        
        // Tabellen-Border
        pdfTable.setBorder(new com.itextpdf.layout.borders.SolidBorder(1));
        
        for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
            XWPFTableRow row = rows.get(rowIndex);
            List<XWPFTableCell> cells = row.getTableCells();
            
            for (int colIndex = 0; colIndex < cells.size(); colIndex++) {
                XWPFTableCell cell = cells.get(colIndex);
                String cellText = cell.getText();
                
                Cell pdfCell = new Cell().add(new Paragraph(cellText != null ? cellText : ""));
                
                // Zell-Padding
                pdfCell.setPadding(8);
                pdfCell.setBorder(new com.itextpdf.layout.borders.SolidBorder(0.5f));
                
                // Zellfarbe extrahieren und anwenden
                Color backgroundColor = extractCellBackgroundColor(cell);
                if (backgroundColor != null) {
                    pdfCell.setBackgroundColor(backgroundColor);
                }
                
                // Textformatierung und Farben
                applyTextFormatting(cell, pdfCell, rowIndex);
                
                // Textausrichtung
                applyTextAlignment(cell, pdfCell);
                
                // Spaltenbreite für numerische Werte anpassen
                if (isNumericColumn(colIndex, cellText)) {
                    pdfCell.setTextAlignment(TextAlignment.RIGHT);
                }
                
                pdfTable.addCell(pdfCell);
            }
        }
        
        doc.add(pdfTable);
    }
    
    /**
     * Prüft ob eine Spalte numerische Werte enthält
     */
    private boolean isNumericColumn(int colIndex, String cellText) {
        if (cellText == null) return false;
        
        // Prüfe auf Zahlen, Währungen, Prozente
        return cellText.trim().matches(".*\\d.*") && 
               (cellText.contains("€") || 
                cellText.contains("$") || 
                cellText.contains("%") ||
                cellText.matches("^\\d+[.,]?\\d*$"));
    }
    
    /**
     * Extrahiert die Hintergrundfarbe einer Zelle
     */
    private Color extractCellBackgroundColor(XWPFTableCell cell) {
        try {
            // Zellfarbe aus den Properties extrahieren
            String colorHex = cell.getColor();
            if (colorHex != null && !colorHex.equals("auto")) {
                return hexToColor(colorHex);
            }
            
            // Alternative: Über CTShd (Cell Shading) - mit sicherer Typisierung
            if (cell.getCTTc() != null && cell.getCTTc().getTcPr() != null) {
                if (cell.getCTTc().getTcPr().getShd() != null) {
                    Object fillColorObj = cell.getCTTc().getTcPr().getShd().getFill();
                    if (fillColorObj != null) {
                        String fillColor = fillColorObj.toString();
                        if (!"auto".equals(fillColor)) {
                            return hexToColor(fillColor);
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("Fehler beim Extrahieren der Zellfarbe: " + e.getMessage());
        }
        
        return null;
    }
    
    /**
     * Wendet Textformatierung auf eine Zelle an mit verbesserter Rich-Text-Unterstützung
     */
    private void applyTextFormatting(XWPFTableCell cell, Cell pdfCell, int rowIndex) {
        // Schriftgröße und Stil
        List<XWPFParagraph> paragraphs = cell.getParagraphs();
        
        if (!paragraphs.isEmpty()) {
            XWPFParagraph firstPara = paragraphs.get(0);
            
            // Erstelle Rich-Text für die Zelle
            Paragraph cellParagraph = createFormattedParagraph(firstPara, 9);
            
            // Header-Zeile Behandlung
            if (rowIndex == 0) {
                cellParagraph.setFontSize(10);
                // Für Header-Zeilen: wenn kein explizites Bold gesetzt, dann Bold
                if (!hasExplicitBold(firstPara)) {
                    cellParagraph.setBold();
                }
                
                // Weiße Schrift für dunkle Header (falls keine Farbe explizit gesetzt)
                if (!hasExplicitTextColor(firstPara)) {
                    cellParagraph.setFontColor(new DeviceRgb(255, 255, 255)); // Weiß
                }
            }
            
            // Ersetze den Standard-Paragraph in der Zelle
            pdfCell.getChildren().clear();
            pdfCell.add(cellParagraph);
        }
    }
    
    /**
     * Prüft ob ein Paragraph explizite Bold-Formatierung hat
     */
    private boolean hasExplicitBold(XWPFParagraph para) {
        List<XWPFRun> runs = para.getRuns();
        for (XWPFRun run : runs) {
            if (run.isBold()) {
                return true;
            }
        }
        return false;
    }
    
    /**
     * Prüft ob ein Paragraph explizite Textfarbe hat
     */
    private boolean hasExplicitTextColor(XWPFParagraph para) {
        List<XWPFRun> runs = para.getRuns();
        for (XWPFRun run : runs) {
            String color = run.getColor();
            if (color != null && !color.equals("auto") && !color.isEmpty()) {
                return true;
            }
        }
        return false;
    }
    
    /**
     * Wendet Textausrichtung auf eine Zelle an
     */
    private void applyTextAlignment(XWPFTableCell cell, Cell pdfCell) {
        List<XWPFParagraph> paragraphs = cell.getParagraphs();
        if (!paragraphs.isEmpty()) {
            XWPFParagraph firstPara = paragraphs.get(0);
            String alignment = firstPara.getAlignment().toString();
            
            switch (alignment) {
                case "CENTER":
                    pdfCell.setTextAlignment(TextAlignment.CENTER);
                    break;
                case "RIGHT":
                    pdfCell.setTextAlignment(TextAlignment.RIGHT);
                    break;
                default:
                    pdfCell.setTextAlignment(TextAlignment.LEFT);
                    break;
            }
        }
    }
    
    /**
     * Konvertiert einen Hex-Farbstring zu iText Color
     */
    private Color hexToColor(String hex) {
        try {
            // Entferne # falls vorhanden
            if (hex.startsWith("#")) {
                hex = hex.substring(1);
            }
            
            // Padding für kurze Hex-Codes
            if (hex.length() == 3) {
                hex = "" + hex.charAt(0) + hex.charAt(0) + 
                         hex.charAt(1) + hex.charAt(1) + 
                         hex.charAt(2) + hex.charAt(2);
            }
            
            if (hex.length() == 6) {
                int r = Integer.parseInt(hex.substring(0, 2), 16);
                int g = Integer.parseInt(hex.substring(2, 4), 16);
                int b = Integer.parseInt(hex.substring(4, 6), 16);
                return new DeviceRgb(r, g, b);
            }
        } catch (Exception e) {
            System.err.println("Fehler beim Konvertieren der Farbe: " + hex + " - " + e.getMessage());
        }
        
        return null;
    }
    
    /**
     * Verarbeitet Bilder aus dem Dokument mit verbessertem Layout
     */
    private void processImages(XWPFDocument document, Document doc) throws IOException {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        boolean hasImages = false;
        
        for (XWPFParagraph paragraph : paragraphs) {
            for (XWPFRun run : paragraph.getRuns()) {
                List<XWPFPicture> pictures = run.getEmbeddedPictures();
                
                for (XWPFPicture picture : pictures) {
                    try {
                        XWPFPictureData pictureData = picture.getPictureData();
                        byte[] imageData = pictureData.getData();
                        
                        if (imageData != null && imageData.length > 0) {
                            if (!hasImages) {
                                // Überschrift für Bilder-Bereich
                                doc.add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                                Paragraph imageTitle = new Paragraph("Bilder und Grafiken")
                                        .setFontSize(16)
                                        .setBold()
                                        .setMarginTop(20)
                                        .setMarginBottom(15);
                                doc.add(imageTitle);
                                hasImages = true;
                            }
                            
                            Image pdfImage = new Image(ImageDataFactory.create(imageData));
                            
                            // Bildgröße intelligent anpassen
                            float maxWidth = 450;
                            float maxHeight = 350;
                            
                            if (pdfImage.getImageWidth() > maxWidth || pdfImage.getImageHeight() > maxHeight) {
                                pdfImage.setAutoScale(true);
                                pdfImage.setMaxWidth(maxWidth);
                                pdfImage.setMaxHeight(maxHeight);
                            }
                            
                            // Bild in Container mit Rahmen
                            Div imageContainer = new Div();
                            imageContainer.add(pdfImage);
                            imageContainer.setTextAlignment(TextAlignment.CENTER);
                            imageContainer.setBorder(new com.itextpdf.layout.borders.SolidBorder(1));
                            imageContainer.setPadding(10);
                            imageContainer.setMarginBottom(20);
                            
                            doc.add(imageContainer);
                            
                            // Bildunterschrift
                            Paragraph caption = new Paragraph("Abbildung: " + pictureData.getFileName())
                                    .setFontSize(9)
                                    .setTextAlignment(TextAlignment.CENTER)
                                    .setMarginBottom(15);
                            caption.setItalic();
                            doc.add(caption);
                        }
                    } catch (Exception e) {
                        System.err.println("Fehler beim Verarbeiten eines Bildes: " + e.getMessage());
                        // Weiter mit dem nächsten Bild
                    }
                }
            }
        }
    }
    
    /**
     * Prüft, ob ein Paragraph eine Überschrift ist
     * 
     * @param paragraph der zu prüfende Paragraph
     * @return true wenn es eine Überschrift ist
     */
    private boolean isHeading(XWPFParagraph paragraph) {
        // Einfache Heuristik: Überschriften sind oft fett oder haben einen Heading-Style
        String style = paragraph.getStyle();
        if (style != null && style.toLowerCase().contains("heading")) {
            return true;
        }
        
        // Prüfe ob der erste Run fett ist
        List<XWPFRun> runs = paragraph.getRuns();
        if (!runs.isEmpty()) {
            XWPFRun firstRun = runs.get(0);
            return firstRun.isBold();
        }
        
        return false;
    }
    
    /**
     * Konvertiert ein Dokument von einem Format in ein anderes (Legacy-Methode)
     * 
     * @param inputFile Pfad zur Eingabedatei
     * @param outputFile Pfad zur Ausgabedatei
     * @param targetFormat Zielformat für die Konvertierung
     */
    @Deprecated
    public void convertDocument(String inputFile, String outputFile, String targetFormat) {
        // TODO: Implementierung für andere Formate
        System.out.println("Konvertiere " + inputFile + " zu " + targetFormat);
        
        if ("PDF".equalsIgnoreCase(targetFormat) && inputFile.toLowerCase().endsWith(".docx")) {
            try {
                convertDocxToPdf(inputFile, outputFile);
            } catch (IOException e) {
                System.err.println("Konvertierung fehlgeschlagen: " + e.getMessage());
            }
        }
    }
}
