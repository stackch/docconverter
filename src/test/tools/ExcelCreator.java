package ch.std.doc.converter.tools;

import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Klasse zum Erstellen von Test-Excel-Dateien
 */
public class ExcelCreator {
    
    /**
     * Erstellt eine einfache Excel-Datei mit Testdaten
     * 
     * @param filename Der Dateiname für die Excel-Datei
     * @throws IOException bei Fehlern beim Schreiben
     */
    public static void createSimpleExcel(String filename) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Verkaufsdaten");
        
        // Header-Zeile
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Produkt");
        headerRow.createCell(1).setCellValue("Anzahl");
        headerRow.createCell(2).setCellValue("Preis pro Stück");
        headerRow.createCell(3).setCellValue("Gesamt");
        headerRow.createCell(4).setCellValue("Kategorie");
        
        // Header-Style
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        // Header-Style anwenden
        for (int i = 0; i < 5; i++) {
            headerRow.getCell(i).setCellStyle(headerStyle);
        }
        
        // Testdaten
        Object[][] data = {
            {"Laptop", 5, 899.99, 4499.95, "Elektronik"},
            {"Maus", 25, 15.50, 387.50, "Zubehör"},
            {"Tastatur", 12, 45.00, 540.00, "Zubehör"},
            {"Monitor", 8, 299.99, 2399.92, "Elektronik"},
            {"USB-Stick", 50, 12.99, 649.50, "Speicher"},
            {"Drucker", 3, 199.99, 599.97, "Büro"},
            {"Papier A4", 100, 4.50, 450.00, "Bürobedarf"},
            {"Tintenpatrone", 20, 25.99, 519.80, "Bürobedarf"}
        };
        
        // Datenzeilen erstellen
        for (int i = 0; i < data.length; i++) {
            Row row = sheet.createRow(i + 1);
            Object[] rowData = data[i];
            
            for (int j = 0; j < rowData.length; j++) {
                Cell cell = row.createCell(j);
                if (rowData[j] instanceof String) {
                    cell.setCellValue((String) rowData[j]);
                } else if (rowData[j] instanceof Integer) {
                    cell.setCellValue((Integer) rowData[j]);
                } else if (rowData[j] instanceof Double) {
                    cell.setCellValue((Double) rowData[j]);
                    
                    // Währungsformatierung für Preise
                    if (j == 2 || j == 3) {
                        CellStyle currencyStyle = workbook.createCellStyle();
                        currencyStyle.setDataFormat(workbook.createDataFormat().getFormat("€#,##0.00"));
                        cell.setCellStyle(currencyStyle);
                    }
                }
            }
        }
        
        // Summenzeile
        Row sumRow = sheet.createRow(data.length + 2);
        sumRow.createCell(0).setCellValue("GESAMT:");
        Cell totalCell = sumRow.createCell(3);
        totalCell.setCellFormula("SUM(D2:D" + (data.length + 1) + ")");
        
        // Summenzeile formatieren
        CellStyle sumStyle = workbook.createCellStyle();
        Font sumFont = workbook.createFont();
        sumFont.setBold(true);
        sumStyle.setFont(sumFont);
        sumStyle.setDataFormat(workbook.createDataFormat().getFormat("€#,##0.00"));
        
        sumRow.getCell(0).setCellStyle(sumStyle);
        totalCell.setCellStyle(sumStyle);
        
        // Spaltenbreiten anpassen
        for (int i = 0; i < 5; i++) {
            sheet.autoSizeColumn(i);
        }
        
        // Zusätzliches Arbeitsblatt
        XSSFSheet sheet2 = workbook.createSheet("Monatsübersicht");
        Row monthHeader = sheet2.createRow(0);
        monthHeader.createCell(0).setCellValue("Monat");
        monthHeader.createCell(1).setCellValue("Umsatz");
        monthHeader.createCell(2).setCellValue("Gewinn");
        
        // Monatsdaten
        String[] months = {"Januar", "Februar", "März", "April", "Mai", "Juni"};
        double[] revenue = {12500.50, 15200.75, 18900.25, 16750.00, 21300.80, 19850.40};
        double[] profit = {3125.13, 3800.19, 4725.06, 4187.50, 5325.20, 4962.60};
        
        for (int i = 0; i < months.length; i++) {
            Row row = sheet2.createRow(i + 1);
            row.createCell(0).setCellValue(months[i]);
            row.createCell(1).setCellValue(revenue[i]);
            row.createCell(2).setCellValue(profit[i]);
        }
        
        // Spaltenbreiten anpassen
        for (int i = 0; i < 3; i++) {
            sheet2.autoSizeColumn(i);
        }
        
        // Datei speichern
        try (FileOutputStream fileOut = new FileOutputStream(filename)) {
            workbook.write(fileOut);
        }
        
        workbook.close();
        System.out.println("Excel-Datei '" + filename + "' wurde erstellt.");
    }
    
    public static void main(String[] args) {
        try {
            createSimpleExcel("test-verkaufsdaten.xlsx");
        } catch (IOException e) {
            System.err.println("Fehler beim Erstellen der Excel-Datei: " + e.getMessage());
        }
    }
}
