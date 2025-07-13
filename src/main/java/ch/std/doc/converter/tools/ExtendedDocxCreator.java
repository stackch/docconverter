package ch.std.doc.converter.tools;

import org.apache.poi.util.Units;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;

import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.imageio.ImageIO;

/**
 * Erweiterte Utility-Klasse zum Erstellen komplexer DOCX-Dateien
 * mit Tabellen, Header/Footer, mehreren Seiten und Bildern.
 */
public class ExtendedDocxCreator {

    /**
     * Erstellt ein umfangreiches DOCX-Dokument mit verschiedenen Elementen
     * 
     * @param filename Der Dateiname für das zu erstellende DOCX-Dokument
     * @throws IOException Bei Problemen mit der Dateierstellung
     */
    public static void createComplexDocx(String filename) throws IOException {
        XWPFDocument document = new XWPFDocument();

        try {
            // Header und Footer erstellen
            createHeaderAndFooter(document);

            // Seite 1: Titelseite
            createTitlePage(document);
            addPageBreak(document);

            // Seite 2: Tabellen-Demo
            createTablePage(document);
            addPageBreak(document);

            // Seite 3: Bilder und Text
            createImagePage(document);
            addPageBreak(document);

            // Seite 4: Zusätzlicher Inhalt
            createAdditionalContentPage(document);

            // Dokument speichern
            try (FileOutputStream out = new FileOutputStream(filename)) {
                document.write(out);
            }

            System.out.println("Komplexes DOCX-Dokument erfolgreich erstellt: " + filename);

        } finally {
            document.close();
        }
    }

    /**
     * Erstellt Header und Footer für das Dokument
     */
    private static void createHeaderAndFooter(XWPFDocument document) {
        // Header erstellen
        XWPFHeader header = document.createHeader(HeaderFooterType.DEFAULT);
        XWPFParagraph headerPara = header.createParagraph();
        headerPara.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun headerRun = headerPara.createRun();
        headerRun.setText("DocConverter - Beispieldokument");
        headerRun.setBold(true);
        headerRun.setFontSize(12);

        // Footer erstellen
        XWPFFooter footer = document.createFooter(HeaderFooterType.DEFAULT);
        XWPFParagraph footerPara = footer.createParagraph();
        footerPara.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun footerRun = footerPara.createRun();
        footerRun.setText("Seite ");
        
        // Seitenzahl hinzufügen
        footerPara.getCTP().addNewFldSimple().setInstr("PAGE");
        
        XWPFRun footerRun2 = footerPara.createRun();
        footerRun2.setText(" von ");
        
        footerPara.getCTP().addNewFldSimple().setInstr("NUMPAGES");
    }

    /**
     * Erstellt die Titelseite
     */
    private static void createTitlePage(XWPFDocument document) {
        // Titel
        XWPFParagraph titlePara = document.createParagraph();
        titlePara.setAlignment(ParagraphAlignment.CENTER);
        titlePara.setSpacingAfter(400);
        XWPFRun titleRun = titlePara.createRun();
        titleRun.setText("Komplexes Testdokument");
        titleRun.setBold(true);
        titleRun.setFontSize(24);
        titleRun.setColor("2E74B5");

        // Untertitel
        XWPFParagraph subtitlePara = document.createParagraph();
        subtitlePara.setAlignment(ParagraphAlignment.CENTER);
        subtitlePara.setSpacingAfter(600);
        XWPFRun subtitleRun = subtitlePara.createRun();
        subtitleRun.setText("Demonstration verschiedener DOCX-Features");
        subtitleRun.setFontSize(16);
        subtitleRun.setItalic(true);

        // Inhaltsverzeichnis
        XWPFParagraph tocTitle = document.createParagraph();
        tocTitle.setAlignment(ParagraphAlignment.LEFT);
        tocTitle.setSpacingAfter(200);
        XWPFRun tocTitleRun = tocTitle.createRun();
        tocTitleRun.setText("Inhaltsverzeichnis");
        tocTitleRun.setBold(true);
        tocTitleRun.setFontSize(16);

        String[] tocItems = {
            "1. Tabellenbeispiele ........................... Seite 2",
            "2. Bilder und Grafiken ........................ Seite 3", 
            "3. Zusätzlicher Inhalt ........................ Seite 4"
        };

        for (String item : tocItems) {
            XWPFParagraph tocItem = document.createParagraph();
            tocItem.setSpacingAfter(100);
            XWPFRun tocRun = tocItem.createRun();
            tocRun.setText(item);
            tocRun.setFontSize(12);
        }
    }

    /**
     * Erstellt eine Seite mit verschiedenen Tabellen
     */
    private static void createTablePage(XWPFDocument document) {
        // Seitentitel
        XWPFParagraph title = document.createParagraph();
        title.setSpacingAfter(300);
        XWPFRun titleRun = title.createRun();
        titleRun.setText("1. Tabellenbeispiele");
        titleRun.setBold(true);
        titleRun.setFontSize(18);

        // Erste Tabelle: Mitarbeiterdaten
        createEmployeeTable(document);

        // Zweite Tabelle: Verkaufsdaten
        createSalesTable(document);

        // Beschreibungstext
        XWPFParagraph description = document.createParagraph();
        description.setSpacingBefore(300);
        XWPFRun descRun = description.createRun();
        descRun.setText("Die obigen Tabellen demonstrieren verschiedene Formatierungsoptionen " +
                       "wie farbige Zellen, verschiedene Schriftgrößen und Zellenausrichtung. " +
                       "Diese Elemente werden beim PDF-Export entsprechend konvertiert.");
        descRun.setFontSize(11);
    }

    /**
     * Erstellt eine Mitarbeitertabelle
     */
    private static void createEmployeeTable(XWPFDocument document) {
        XWPFParagraph tableTitle = document.createParagraph();
        tableTitle.setSpacingAfter(100);
        XWPFRun tableTitleRun = tableTitle.createRun();
        tableTitleRun.setText("Mitarbeiterdaten:");
        tableTitleRun.setBold(true);
        tableTitleRun.setFontSize(14);

        XWPFTable table = document.createTable(4, 4);
        table.setWidth("100%");

        // Header-Zeile
        XWPFTableRow headerRow = table.getRow(0);
        String[] headers = {"Name", "Position", "Abteilung", "Gehalt"};
        for (int i = 0; i < headers.length; i++) {
            XWPFTableCell cell = headerRow.getCell(i);
            cell.setColor("4472C4");
            XWPFParagraph para = cell.getParagraphs().get(0);
            XWPFRun run = para.createRun();
            run.setText(headers[i]);
            run.setBold(true);
            run.setColor("FFFFFF");
            para.setAlignment(ParagraphAlignment.CENTER);
        }

        // Datenzeilen
        String[][] data = {
            {"Anna Müller", "Entwicklerin", "IT", "65.000 €"},
            {"Thomas Schmidt", "Manager", "Vertrieb", "75.000 €"},
            {"Lisa Weber", "Designerin", "Marketing", "55.000 €"}
        };

        for (int i = 0; i < data.length; i++) {
            XWPFTableRow row = table.getRow(i + 1);
            for (int j = 0; j < data[i].length; j++) {
                XWPFTableCell cell = row.getCell(j);
                if (i % 2 == 0) {
                    cell.setColor("F2F2F2");
                }
                XWPFParagraph para = cell.getParagraphs().get(0);
                XWPFRun run = para.createRun();
                run.setText(data[i][j]);
                if (j == 3) { // Gehalt rechtsbündig
                    para.setAlignment(ParagraphAlignment.RIGHT);
                }
            }
        }
    }

    /**
     * Erstellt eine Verkaufstabelle
     */
    private static void createSalesTable(XWPFDocument document) {
        XWPFParagraph tableTitle = document.createParagraph();
        tableTitle.setSpacingBefore(400);
        tableTitle.setSpacingAfter(100);
        XWPFRun tableTitleRun = tableTitle.createRun();
        tableTitleRun.setText("Quartalsumsätze 2024:");
        tableTitleRun.setBold(true);
        tableTitleRun.setFontSize(14);

        XWPFTable table = document.createTable(5, 5);
        table.setWidth("100%");

        // Header
        XWPFTableRow headerRow = table.getRow(0);
        String[] headers = {"Produkt", "Q1", "Q2", "Q3", "Q4"};
        for (int i = 0; i < headers.length; i++) {
            XWPFTableCell cell = headerRow.getCell(i);
            cell.setColor("70AD47");
            XWPFParagraph para = cell.getParagraphs().get(0);
            XWPFRun run = para.createRun();
            run.setText(headers[i]);
            run.setBold(true);
            run.setColor("FFFFFF");
            para.setAlignment(ParagraphAlignment.CENTER);
        }

        // Daten
        String[][] salesData = {
            {"Produkt A", "125.000", "138.000", "142.000", "155.000"},
            {"Produkt B", "89.000", "95.000", "101.000", "98.000"},
            {"Produkt C", "67.000", "72.000", "78.000", "85.000"},
            {"Gesamt", "281.000", "305.000", "321.000", "338.000"}
        };

        for (int i = 0; i < salesData.length; i++) {
            XWPFTableRow row = table.getRow(i + 1);
            for (int j = 0; j < salesData[i].length; j++) {
                XWPFTableCell cell = row.getCell(j);
                if (i == salesData.length - 1) { // Gesamt-Zeile
                    cell.setColor("E2EFDA");
                }
                XWPFParagraph para = cell.getParagraphs().get(0);
                XWPFRun run = para.createRun();
                run.setText(salesData[i][j]);
                if (i == salesData.length - 1) {
                    run.setBold(true);
                }
                if (j > 0) { // Zahlen rechtsbündig
                    para.setAlignment(ParagraphAlignment.RIGHT);
                }
            }
        }
    }

    /**
     * Erstellt eine Seite mit Bildern
     */
    private static void createImagePage(XWPFDocument document) throws IOException {
        // Seitentitel
        XWPFParagraph title = document.createParagraph();
        title.setSpacingAfter(300);
        XWPFRun titleRun = title.createRun();
        titleRun.setText("2. Bilder und Grafiken");
        titleRun.setBold(true);
        titleRun.setFontSize(18);

        // Erstes Bild: Einfaches Diagramm
        XWPFParagraph imgPara1 = document.createParagraph();
        imgPara1.setAlignment(ParagraphAlignment.CENTER);
        imgPara1.setSpacingAfter(200);
        XWPFRun imgRun1 = imgPara1.createRun();
        
        try {
            BufferedImage chart = createSimpleChart();
            ByteArrayOutputStream chartBaos = new ByteArrayOutputStream();
            ImageIO.write(chart, "PNG", chartBaos);
            imgRun1.addPicture(new java.io.ByteArrayInputStream(chartBaos.toByteArray()),
                              XWPFDocument.PICTURE_TYPE_PNG, "chart.png",
                              Units.toEMU(400), Units.toEMU(250));
        } catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException e) {
            System.err.println("Fehler beim Hinzufügen des Diagramms: " + e.getMessage());
        }

        // Bildunterschrift
        XWPFParagraph caption1 = document.createParagraph();
        caption1.setAlignment(ParagraphAlignment.CENTER);
        caption1.setSpacingAfter(300);
        XWPFRun captionRun1 = caption1.createRun();
        captionRun1.setText("Abbildung 1: Beispiel-Balkendiagramm");
        captionRun1.setItalic(true);
        captionRun1.setFontSize(10);

        // Zweites Bild: Geometrische Form
        XWPFParagraph imgPara2 = document.createParagraph();
        imgPara2.setAlignment(ParagraphAlignment.CENTER);
        imgPara2.setSpacingAfter(200);
        XWPFRun imgRun2 = imgPara2.createRun();
        
        try {
            BufferedImage geometry = createGeometricShape();
            ByteArrayOutputStream geomBaos = new ByteArrayOutputStream();
            ImageIO.write(geometry, "PNG", geomBaos);
            imgRun2.addPicture(new java.io.ByteArrayInputStream(geomBaos.toByteArray()),
                              XWPFDocument.PICTURE_TYPE_PNG, "geometry.png",
                              Units.toEMU(300), Units.toEMU(300));
        } catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException e) {
            System.err.println("Fehler beim Hinzufügen der geometrischen Form: " + e.getMessage());
        }

        // Bildunterschrift
        XWPFParagraph caption2 = document.createParagraph();
        caption2.setAlignment(ParagraphAlignment.CENTER);
        caption2.setSpacingAfter(300);
        XWPFRun captionRun2 = caption2.createRun();
        captionRun2.setText("Abbildung 2: Geometrische Formen");
        captionRun2.setItalic(true);
        captionRun2.setFontSize(10);

        // Beschreibungstext
        XWPFParagraph description = document.createParagraph();
        XWPFRun descRun = description.createRun();
        descRun.setText("Die obigen Grafiken wurden programmatisch erstellt und zeigen die " +
                       "Möglichkeiten der Bildintegration in DOCX-Dokumenten. Beim Export zu PDF " +
                       "werden diese Bilder entsprechend konvertiert und beibehalten.");
        descRun.setFontSize(11);
    }

    /**
     * Erstellt ein einfaches Balkendiagramm
     */
    private static BufferedImage createSimpleChart() {
        int width = 400;
        int height = 250;
        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g2d = image.createGraphics();
        
        // Antialiasing für bessere Qualität
        g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        
        // Hintergrund
        g2d.setColor(Color.WHITE);
        g2d.fillRect(0, 0, width, height);
        
        // Achsen
        g2d.setColor(Color.BLACK);
        g2d.drawLine(50, height - 50, width - 20, height - 50); // X-Achse
        g2d.drawLine(50, height - 50, 50, 20); // Y-Achse
        
        // Balken
        int[] values = {120, 80, 150, 90};
        String[] labels = {"Q1", "Q2", "Q3", "Q4"};
        Color[] colors = {Color.BLUE, Color.RED, Color.GREEN, Color.ORANGE};
        
        int barWidth = 60;
        int startX = 80;
        
        for (int i = 0; i < values.length; i++) {
            int barHeight = values[i];
            int x = startX + i * (barWidth + 20);
            int y = height - 50 - barHeight;
            
            g2d.setColor(colors[i]);
            g2d.fillRect(x, y, barWidth, barHeight);
            
            // Label
            g2d.setColor(Color.BLACK);
            g2d.drawString(labels[i], x + 20, height - 30);
            g2d.drawString(String.valueOf(values[i]), x + 15, y - 5);
        }
        
        // Titel
        g2d.setFont(new Font("Arial", Font.BOLD, 14));
        g2d.drawString("Quartalsumsätze (in Tausend €)", 120, 15);
        
        g2d.dispose();
        return image;
    }

    /**
     * Erstellt geometrische Formen
     */
    private static BufferedImage createGeometricShape() {
        int size = 300;
        BufferedImage image = new BufferedImage(size, size, BufferedImage.TYPE_INT_RGB);
        Graphics2D g2d = image.createGraphics();
        
        g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        
        // Hintergrund
        g2d.setColor(Color.WHITE);
        g2d.fillRect(0, 0, size, size);
        
        // Kreis
        g2d.setColor(new Color(70, 130, 180, 150));
        g2d.fillOval(50, 50, 100, 100);
        g2d.setColor(Color.BLUE);
        g2d.drawOval(50, 50, 100, 100);
        
        // Rechteck
        g2d.setColor(new Color(255, 69, 0, 150));
        g2d.fillRect(150, 50, 100, 100);
        g2d.setColor(Color.RED);
        g2d.drawRect(150, 50, 100, 100);
        
        // Dreieck
        int[] xPoints = {100, 150, 200};
        int[] yPoints = {200, 170, 200};
        g2d.setColor(new Color(50, 205, 50, 150));
        g2d.fillPolygon(xPoints, yPoints, 3);
        g2d.setColor(Color.GREEN);
        g2d.drawPolygon(xPoints, yPoints, 3);
        
        g2d.dispose();
        return image;
    }

    /**
     * Erstellt zusätzlichen Inhalt für die vierte Seite
     */
    private static void createAdditionalContentPage(XWPFDocument document) {
        // Seitentitel
        XWPFParagraph title = document.createParagraph();
        title.setSpacingAfter(300);
        XWPFRun titleRun = title.createRun();
        titleRun.setText("3. Zusätzlicher Inhalt");
        titleRun.setBold(true);
        titleRun.setFontSize(18);

        // Aufzählung
        XWPFParagraph listTitle = document.createParagraph();
        listTitle.setSpacingAfter(100);
        XWPFRun listTitleRun = listTitle.createRun();
        listTitleRun.setText("Features dieses Dokuments:");
        listTitleRun.setBold(true);
        listTitleRun.setFontSize(14);

        String[] features = {
            "✓ Header und Footer auf allen Seiten",
            "✓ Automatische Seitennummerierung", 
            "✓ Verschiedene Tabellenformate",
            "✓ Eingebettete Bilder und Grafiken",
            "✓ Unterschiedliche Textformatierungen",
            "✓ Farbige Elemente und Hervorhebungen"
        };

        for (String feature : features) {
            XWPFParagraph listItem = document.createParagraph();
            listItem.setSpacingAfter(100);
            XWPFRun listRun = listItem.createRun();
            listRun.setText(feature);
            listRun.setFontSize(12);
        }

        // Abschlusstext
        XWPFParagraph conclusion = document.createParagraph();
        conclusion.setSpacingBefore(400);
        XWPFRun conclusionRun = conclusion.createRun();
        conclusionRun.setText("Dieses Dokument demonstriert die umfangreichen Möglichkeiten " +
                             "der DOCX-zu-PDF-Konvertierung. Alle hier gezeigten Elemente werden " +
                             "beim Export entsprechend übertragen und formatiert beibehalten. " +
                             "Der DocConverter kann somit auch komplexe Dokumentstrukturen " +
                             "erfolgreich verarbeiten.");
        conclusionRun.setFontSize(11);

        // Datum und Unterschrift
        XWPFParagraph signature = document.createParagraph();
        signature.setSpacingBefore(600);
        signature.setAlignment(ParagraphAlignment.RIGHT);
        XWPFRun sigRun = signature.createRun();
        sigRun.setText("Erstellt am: " + java.time.LocalDate.now().toString());
        sigRun.setFontSize(10);
        sigRun.setItalic(true);
    }

    /**
     * Fügt einen Seitenumbruch hinzu
     */
    private static void addPageBreak(XWPFDocument document) {
        XWPFParagraph pageBreak = document.createParagraph();
        XWPFRun pageBreakRun = pageBreak.createRun();
        pageBreakRun.addBreak(BreakType.PAGE);
    }

    /**
     * Main-Methode zum direkten Ausführen
     */
    public static void main(String[] args) {
        String filename = args.length > 0 ? args[0] : "komplexes-beispiel.docx";
        
        try {
            createComplexDocx(filename);
        } catch (IOException e) {
            System.err.println("Fehler beim Erstellen des DOCX-Dokuments: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
