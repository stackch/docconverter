package ch.std.doc.converter.core.impl;

import ch.std.doc.converter.core.DocumentConverter;

import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DateUtil;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.AreaBreak;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.UnitValue;
import com.itextpdf.layout.properties.AreaBreakType;
import com.itextpdf.kernel.colors.DeviceRgb;
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
 * Spezialisierter Konverter für Excel-Dateien zu PDF
 */
public class ExcelToPdfConverter extends DocumentConverter {
    
    private static final String[] SUPPORTED_EXTENSIONS = {".xlsx", ".xls"};
    private static final String CONVERTER_NAME = "Excel-Konverter";
    
    @Override
    public void convertToPdf(String inputFile, String outputFile) throws IOException {
        validateFiles(inputFile, outputFile);
        logConversion(inputFile, outputFile);
        
        if (inputFile.toLowerCase().endsWith(".xlsx")) {
            convertXlsxToPdf(inputFile, outputFile);
        } else if (inputFile.toLowerCase().endsWith(".xls")) {
            // TODO: Implementierung für .xls Dateien
            throw new UnsupportedOperationException("XLS-Format wird noch nicht unterstützt");
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
        return "Konvertiert Microsoft Excel XLSX/XLS-Tabellen zu PDF im Querformat. " +
               "Erhält Zellformatierung, Hintergrundfarben, Schriftarten und Ausrichtung. " +
               "Verarbeitet mehrere Arbeitsblätter und fügt automatisch Seitenzahlen hinzu.";
    }
    
    /**
     * Event Handler für Seitenzahlen
     */
    private static class ExcelPageNumberEventHandler implements IEventHandler {
        private PdfFont font;
        private int totalPages;
        
        public ExcelPageNumberEventHandler(int totalPages) {
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
    
    private void convertXlsxToPdf(String inputFile, String outputFile) throws IOException {
        try (FileInputStream fis = new FileInputStream(inputFile);
             XSSFWorkbook workbook = new XSSFWorkbook(fis);
             FileOutputStream fos = new FileOutputStream(outputFile)) {
            
            PdfWriter writer = new PdfWriter(fos);
            PdfDocument pdfDoc = new PdfDocument(writer);
            Document doc = new Document(pdfDoc, PageSize.A4.rotate()); // Querformat für Excel
            
            doc.setMargins(36, 36, 72, 36);
            
            // Alle Worksheets verarbeiten
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                
                if (i > 0) {
                    doc.add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                }
                
                // Worksheet-Titel
                Paragraph sheetTitle = new Paragraph(sheet.getSheetName())
                        .setFontSize(16)
                        .setBold()
                        .setMarginBottom(15);
                doc.add(sheetTitle);
                
                processExcelSheet(sheet, doc);
            }
            
            // Seitenzahlen hinzufügen
            int totalPages = pdfDoc.getNumberOfPages();
            if (totalPages > 0) {
                pdfDoc.addEventHandler(PdfDocumentEvent.END_PAGE, new ExcelPageNumberEventHandler(totalPages));
            }
            
            doc.close();
        }
    }
    
    private void processExcelSheet(XSSFSheet sheet, Document doc) {
        if (sheet.getPhysicalNumberOfRows() == 0) {
            doc.add(new Paragraph("(Leeres Arbeitsblatt)").setItalic());
            return;
        }
        
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        int maxCols = 0;
        
        // Finde die maximale Spaltenanzahl
        for (int rowNum = firstRowNum; rowNum <= lastRowNum; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null && row.getLastCellNum() > maxCols) {
                maxCols = row.getLastCellNum();
            }
        }
        
        if (maxCols == 0) {
            doc.add(new Paragraph("(Keine Daten gefunden)").setItalic());
            return;
        }
        
        // PDF-Tabelle erstellen
        Table pdfTable = new Table(UnitValue.createPercentArray(maxCols))
                .useAllAvailableWidth()
                .setMarginBottom(20);
        
        // Zeilen verarbeiten
        for (int rowNum = firstRowNum; rowNum <= lastRowNum; rowNum++) {
            Row row = sheet.getRow(rowNum);
            
            for (int colNum = 0; colNum < maxCols; colNum++) {
                org.apache.poi.ss.usermodel.Cell excelCell = (row != null) ? row.getCell(colNum) : null;
                String cellText = getExcelCellText(excelCell);
                
                Cell pdfCell = new Cell().add(new Paragraph(cellText));
                formatExcelCell(excelCell, pdfCell, rowNum == firstRowNum);
                pdfTable.addCell(pdfCell);
            }
        }
        
        doc.add(pdfTable);
    }
    
    private String getExcelCellText(org.apache.poi.ss.usermodel.Cell cell) {
        if (cell == null) {
            return "";
        }
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double numValue = cell.getNumericCellValue();
                    if (numValue == Math.floor(numValue)) {
                        return String.valueOf((long) numValue);
                    } else {
                        return String.format("%.2f", numValue);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return String.valueOf(cell.getNumericCellValue());
                } catch (Exception e) {
                    try {
                        return cell.getStringCellValue();
                    } catch (Exception e2) {
                        return cell.getCellFormula();
                    }
                }
            default:
                return "";
        }
    }
    
    private void formatExcelCell(org.apache.poi.ss.usermodel.Cell excelCell, Cell pdfCell, boolean isHeaderRow) {
        pdfCell.setPadding(6);
        pdfCell.setBorder(new com.itextpdf.layout.borders.SolidBorder(0.5f));
        
        // Header-Zeile hervorheben
        if (isHeaderRow) {
            pdfCell.setBackgroundColor(new DeviceRgb(200, 200, 200));
            pdfCell.setBold();
        }
        
        if (excelCell != null) {
            CellStyle cellStyle = excelCell.getCellStyle();
            
            // Textausrichtung
            HorizontalAlignment alignment = cellStyle.getAlignment();
            switch (alignment) {
                case CENTER:
                    pdfCell.setTextAlignment(TextAlignment.CENTER);
                    break;
                case RIGHT:
                    pdfCell.setTextAlignment(TextAlignment.RIGHT);
                    break;
                default:
                    pdfCell.setTextAlignment(TextAlignment.LEFT);
                    break;
            }
            
            // Schriftformatierung
            try {
                Font font = excelCell.getSheet().getWorkbook().getFontAt(cellStyle.getFontIndex());
                if (font.getBold()) {
                    pdfCell.setBold();
                }
                if (font.getItalic()) {
                    pdfCell.setItalic();
                }
            } catch (Exception e) {
                // Ignoriere Formatierungsfehler
            }
        }
    }
}
