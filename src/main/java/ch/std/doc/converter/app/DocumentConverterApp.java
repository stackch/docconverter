package ch.std.doc.converter.app;

import ch.std.doc.converter.core.DocumentConverter;
import ch.std.doc.converter.core.DocumentConverterFactory;

import java.io.IOException;
import java.util.Arrays;

/**
 * Hauptklasse für die Dokumentkonvertierung mit Factory-Pattern
 */
public class DocumentConverterApp {
    
    public static void main(String[] args) {
        if (args.length != 2) {
            showUsage();
            System.exit(1);
        }
        
        String inputFile = args[0];
        String outputFile = args[1];
        
        try {
            // Factory-Pattern: Automatische Konverter-Auswahl
            DocumentConverter converter = DocumentConverterFactory.createConverter(inputFile);
            
            if (converter == null) {
                System.err.println("Fehler: Dateiformat wird nicht unterstützt: " + inputFile);
                System.err.println("\nUnterstützte Formate:");
                showSupportedFormats();
                System.exit(1);
            }
            
            System.out.println("Verwende " + converter.getConverterName() + " für: " + inputFile);
            
            converter.convertToPdf(inputFile, outputFile);
            
            System.out.println("Konvertierung erfolgreich abgeschlossen!");
            System.out.println("PDF erstellt: " + outputFile);
            
        } catch (IOException e) {
            System.err.println("Fehler bei der Konvertierung: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        } catch (Exception e) {
            System.err.println("Unerwarteter Fehler: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
    
    private static void showUsage() {
        System.out.println("Document Converter - Factory-basierte Architektur");
        System.out.println("================================================");
        System.out.println();
        System.out.println("Verwendung: java -jar docconverter.jar <eingabe-datei> <ausgabe-pdf>");
        System.out.println();
        System.out.println("Beispiele:");
        System.out.println("  java -jar docconverter.jar document.docx output.pdf");
        System.out.println("  java -jar docconverter.jar data.xlsx report.pdf");
        System.out.println("  java -jar docconverter.jar legacy.doc converted.pdf");
        System.out.println();
        showSupportedFormats();
    }
    
    private static void showSupportedFormats() {
        System.out.println("Unterstützte Formate:");
        System.out.println("====================");
        
        String formatOverview = DocumentConverterFactory.getSupportedFormatsOverview();
        System.out.println(formatOverview);
        
        System.out.println();
        System.out.println("Alle unterstützten Dateierweiterungen:");
        String[] extensions = DocumentConverterFactory.getSupportedExtensions();
        Arrays.sort(extensions);
        for (String ext : extensions) {
            System.out.println("  " + ext);
        }
    }
}
