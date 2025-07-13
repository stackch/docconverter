package ch.std.doc.converter.tools;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Beispielprogramm zum Erstellen einer Test-DOCX-Datei
 */
public class DocxCreator {
    
    public static void main(String[] args) throws IOException {
        if (args.length < 1) {
            System.out.println("Verwendung: java DocxCreator <dateiname.docx>");
            return;
        }
        
        String filename = args[0];
        createSampleDocx(filename);
        System.out.println("Test-DOCX-Datei erstellt: " + filename);
    }
    
    public static void createSampleDocx(String filename) throws IOException {
        try (XWPFDocument document = new XWPFDocument();
             FileOutputStream out = new FileOutputStream(filename)) {
            
            // Überschrift
            XWPFParagraph title = document.createParagraph();
            title.createRun().setText("Beispiel Dokument - DOCX zu PDF Konvertierung");
            title.getRuns().get(0).setBold(true);
            title.getRuns().get(0).setFontSize(18);
            
            // Absatz 1
            XWPFParagraph para1 = document.createParagraph();
            para1.createRun().setText("Dies ist ein Beispieldokument, das demonstriert, wie ein DOCX-Dokument in ein PDF konvertiert werden kann.");
            
            // Zwischenüberschrift
            XWPFParagraph subtitle = document.createParagraph();
            subtitle.createRun().setText("Features des DocConverters");
            subtitle.getRuns().get(0).setBold(true);
            subtitle.getRuns().get(0).setFontSize(14);
            
            // Liste von Features
            XWPFParagraph feature1 = document.createParagraph();
            feature1.createRun().setText("• Konvertierung von DOCX zu PDF");
            
            XWPFParagraph feature2 = document.createParagraph();
            feature2.createRun().setText("• Erhaltung der Grundformatierung");
            
            XWPFParagraph feature3 = document.createParagraph();
            feature3.createRun().setText("• Erkennung von Überschriften (fett gedruckte Texte)");
            
            XWPFParagraph feature4 = document.createParagraph();
            feature4.createRun().setText("• Einfache Kommandozeilen-Verwendung");
            
            // Schlussbemerkung
            XWPFParagraph conclusion = document.createParagraph();
            conclusion.createRun().setText("Der DocConverter ist ein einfaches, aber effektives Tool zur Dokumentkonvertierung, das auf bewährten Java-Bibliotheken basiert.");
            
            document.write(out);
        }
    }
}
