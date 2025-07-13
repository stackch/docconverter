package ch.std.doc.converter;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.BeforeEach;
import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests für die DocConverter Klasse
 */
public class DocConverterTest {
    
    private DocConverter converter;
    
    @BeforeEach
    public void setUp() {
        converter = new DocConverter();
    }
    
    @Test
    public void testDocConverterExists() {
        assertNotNull(converter);
    }
    
    @Test
    public void testConvertDocument() {
        // TODO: Implementiere Tests für die convertDocument Methode
        assertDoesNotThrow(() -> {
            converter.convertDocument("test.txt", "output.pdf", "PDF");
        });
    }
}
