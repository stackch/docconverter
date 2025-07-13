package ch.std.doc.converter.core;

import java.io.IOException;

/**
 * Abstrakte Basisklasse für alle Dokumentkonverter
 */
public abstract class DocumentConverter {
    
    /**
     * Konvertiert ein Dokument zu PDF
     * 
     * @param inputFile Pfad zur Eingabedatei
     * @param outputFile Pfad zur PDF-Ausgabedatei
     * @throws IOException bei Fehlern beim Lesen oder Schreiben der Dateien
     */
    public abstract void convertToPdf(String inputFile, String outputFile) throws IOException;
    
    /**
     * Gibt die unterstützten Dateierweiterungen zurück
     * 
     * @return Array von unterstützten Dateierweiterungen (z.B. {".docx", ".doc"})
     */
    public abstract String[] getSupportedExtensions();
    
    /**
     * Gibt den Namen des Konverters zurück
     * 
     * @return Name des Konverters
     */
    public abstract String getConverterName();
    
    /**
     * Gibt eine detaillierte Beschreibung des Konverters zurück
     * 
     * @return Beschreibungstext
     */
    public abstract String getDescription();
    
    /**
     * Prüft ob eine Datei von diesem Konverter unterstützt wird
     * 
     * @param filename Dateiname oder Pfad
     * @return true wenn unterstützt, false sonst
     */
    public boolean supportsFile(String filename) {
        if (filename == null) {
            return false;
        }
        
        String lowerFilename = filename.toLowerCase();
        for (String extension : getSupportedExtensions()) {
            if (lowerFilename.endsWith(extension.toLowerCase())) {
                return true;
            }
        }
        return false;
    }
    
    /**
     * Validiert die Eingabe- und Ausgabedateien
     * 
     * @param inputFile Eingabedatei
     * @param outputFile Ausgabedatei
     * @throws IllegalArgumentException bei ungültigen Parametern
     */
    protected void validateFiles(String inputFile, String outputFile) {
        if (inputFile == null || inputFile.trim().isEmpty()) {
            throw new IllegalArgumentException("Eingabedatei darf nicht null oder leer sein");
        }
        
        if (outputFile == null || outputFile.trim().isEmpty()) {
            throw new IllegalArgumentException("Ausgabedatei darf nicht null oder leer sein");
        }
        
        if (!supportsFile(inputFile)) {
            throw new IllegalArgumentException("Dateierweiterung wird nicht unterstützt: " + inputFile);
        }
    }
    
    /**
     * Loggt eine Konvertierungsmeldung
     * 
     * @param inputFile Eingabedatei
     * @param outputFile Ausgabedatei
     */
    protected void logConversion(String inputFile, String outputFile) {
        System.out.println(String.format("[%s] Konvertiere '%s' zu '%s'", 
                          getConverterName(), inputFile, outputFile));
    }
}
