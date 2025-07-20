package ch.std.doc.converter.utils;

import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.canvas.parser.PdfTextExtractor;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Hilfsklasse zur Validierung von PDF-Inhalten in Tests
 */
public class PdfContentValidator {
    
    /**
     * Extrahiert den gesamten Text aus einer PDF-Datei
     */
    public static String extractTextFromPdf(File pdfFile) throws IOException {
        if (!pdfFile.exists()) {
            throw new IllegalArgumentException("PDF-Datei existiert nicht: " + pdfFile.getAbsolutePath());
        }
        
        StringBuilder extractedText = new StringBuilder();
        
        try (PdfDocument pdfDoc = new PdfDocument(new PdfReader(pdfFile.getAbsolutePath()))) {
            int numberOfPages = pdfDoc.getNumberOfPages();
            
            for (int i = 1; i <= numberOfPages; i++) {
                String pageText = PdfTextExtractor.getTextFromPage(pdfDoc.getPage(i));
                extractedText.append(pageText);
                if (i < numberOfPages) {
                    extractedText.append("\n");
                }
            }
        }
        
        return extractedText.toString();
    }
    
    /**
     * Extrahiert Text von einer bestimmten Seite
     */
    public static String extractTextFromPage(File pdfFile, int pageNumber) throws IOException {
        try (PdfDocument pdfDoc = new PdfDocument(new PdfReader(pdfFile.getAbsolutePath()))) {
            if (pageNumber < 1 || pageNumber > pdfDoc.getNumberOfPages()) {
                throw new IllegalArgumentException("Ungültige Seitenzahl: " + pageNumber);
            }
            return PdfTextExtractor.getTextFromPage(pdfDoc.getPage(pageNumber));
        }
    }
    
    /**
     * Zählt die Anzahl der Seiten in der PDF
     */
    public static int getPageCount(File pdfFile) throws IOException {
        try (PdfDocument pdfDoc = new PdfDocument(new PdfReader(pdfFile.getAbsolutePath()))) {
            return pdfDoc.getNumberOfPages();
        }
    }
    
    /**
     * Prüft ob ein bestimmter Text in der PDF enthalten ist
     */
    public static boolean containsText(File pdfFile, String searchText) throws IOException {
        String pdfText = extractTextFromPdf(pdfFile);
        return pdfText.contains(searchText);
    }
    
    /**
     * Prüft ob alle angegebenen Texte in der PDF enthalten sind
     */
    public static boolean containsAllTexts(File pdfFile, String... searchTexts) throws IOException {
        String pdfText = extractTextFromPdf(pdfFile);
        
        for (String searchText : searchTexts) {
            if (!pdfText.contains(searchText)) {
                return false;
            }
        }
        return true;
    }
    
    /**
     * Extrahiert alle Textzeilen aus der PDF
     */
    public static List<String> extractLines(File pdfFile) throws IOException {
        String pdfText = extractTextFromPdf(pdfFile);
        List<String> lines = new ArrayList<>();
        
        String[] textLines = pdfText.split("\n");
        for (String line : textLines) {
            String trimmedLine = line.trim();
            if (!trimmedLine.isEmpty()) {
                lines.add(trimmedLine);
            }
        }
        
        return lines;
    }
    
    /**
     * Berechnet eine grobe Ähnlichkeit zwischen zwei Texten (0.0 - 1.0)
     */
    public static double calculateTextSimilarity(String text1, String text2) {
        if (text1 == null || text2 == null) {
            return 0.0;
        }
        
        // Normalisiere Texte (Leerzeichen, Zeilenumbrüche)
        String normalized1 = text1.replaceAll("\\s+", " ").trim().toLowerCase();
        String normalized2 = text2.replaceAll("\\s+", " ").trim().toLowerCase();
        
        if (normalized1.equals(normalized2)) {
            return 1.0;
        }
        
        // Einfache Wort-basierte Ähnlichkeit
        String[] words1 = normalized1.split(" ");
        String[] words2 = normalized2.split(" ");
        
        int commonWords = 0;
        for (String word1 : words1) {
            for (String word2 : words2) {
                if (word1.equals(word2)) {
                    commonWords++;
                    break;
                }
            }
        }
        
        int totalWords = Math.max(words1.length, words2.length);
        return totalWords == 0 ? 0.0 : (double) commonWords / totalWords;
    }
}
