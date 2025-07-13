package ch.std.doc.converter.core;

import ch.std.doc.converter.core.impl.DocxToPdfConverter;
import ch.std.doc.converter.core.impl.ExcelToPdfConverter;
import ch.std.doc.converter.core.impl.DocToPdfConverter;

import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.HashSet;

/**
 * Factory für die Erstellung von Dokumentkonvertern
 */
public class DocumentConverterFactory {
    
    private static final Map<String, DocumentConverter> converterRegistry = new HashMap<>();
    
    static {
        // Registriere alle verfügbaren Konverter
        registerConverter(new DocxToPdfConverter());
        registerConverter(new ExcelToPdfConverter());
        registerConverter(new DocToPdfConverter());
    }
    
    /**
     * Registriert einen neuen Konverter
     * 
     * @param converter Der zu registrierende Konverter
     */
    public static void registerConverter(DocumentConverter converter) {
        for (String extension : converter.getSupportedExtensions()) {
            converterRegistry.put(extension.toLowerCase(), converter);
        }
    }
    
    /**
     * Erstellt einen Konverter für die gegebene Datei
     * 
     * @param filename Dateiname oder Pfad
     * @return Passender DocumentConverter oder null wenn nicht unterstützt
     */
    public static DocumentConverter createConverter(String filename) {
        if (filename == null) {
            return null;
        }
        
        String lowerFilename = filename.toLowerCase();
        for (Map.Entry<String, DocumentConverter> entry : converterRegistry.entrySet()) {
            if (lowerFilename.endsWith(entry.getKey())) {
                return entry.getValue();
            }
        }
        
        return null;
    }
    
    /**
     * Gibt alle registrierten Konverter zurück
     * 
     * @return Set aller registrierten Konverter
     */
    public static Set<DocumentConverter> getRegisteredConverters() {
        return new HashSet<>(converterRegistry.values());
    }
    
    /**
     * Prüft ob eine Datei unterstützt wird
     * 
     * @param filename Dateiname oder Pfad
     * @return true wenn unterstützt, false sonst
     */
    public static boolean isSupported(String filename) {
        return createConverter(filename) != null;
    }
    
    /**
     * Gibt alle unterstützten Dateierweiterungen zurück
     * 
     * @return Array aller unterstützten Erweiterungen
     */
    public static String[] getSupportedExtensions() {
        return converterRegistry.keySet().toArray(new String[0]);
    }
    
    /**
     * Gibt eine detaillierte Übersicht aller unterstützten Formate zurück
     * 
     * @return Formatierte Übersicht mit Beschreibungen
     */
    public static String getSupportedFormatsOverview() {
        StringBuilder sb = new StringBuilder();
        sb.append("Unterstützte Dokumentformate:\n");
        sb.append("=============================\n\n");
        
        for (DocumentConverter converter : getRegisteredConverters()) {
            sb.append("🔹 ").append(converter.getConverterName()).append("\n");
            sb.append("   Formate: ");
            String[] extensions = converter.getSupportedExtensions();
            for (int i = 0; i < extensions.length; i++) {
                sb.append(extensions[i]);
                if (i < extensions.length - 1) {
                    sb.append(", ");
                }
            }
            sb.append("\n");
            sb.append("   Beschreibung: ").append(converter.getDescription()).append("\n\n");
        }
        
        return sb.toString();
    }
}
