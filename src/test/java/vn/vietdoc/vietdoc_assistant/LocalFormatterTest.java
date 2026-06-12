package vn.vietdoc.vietdoc_assistant;

import org.junit.jupiter.api.Test;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import static org.junit.jupiter.api.Assertions.assertTrue;

public class LocalFormatterTest {

    @Test
    public void testFormatLocalDocument() throws Exception {
        // Path to the target docx file in the parent workspace directory
        Path docxPath = Paths.get("../BTL_quanLyGiaDung.docx").toAbsolutePath().normalize();
        System.out.println("Processing target document at path: " + docxPath);

        assertTrue(Files.exists(docxPath), "Target docx file does not exist: " + docxPath);

        byte[] inputBytes = Files.readAllBytes(docxPath);

        // Create the formatting parameters matching HAUI requirements
        FormattingParameters params = VietDocStyleConfig.createDefaultParameters();
        
        System.out.println("Using formatting parameters: " + params.toString());

        // Process document
        byte[] outputBytes = VietDocEngine.processForWeb(inputBytes, params);

        // Backup original document just in case
        Path backupPath = Paths.get("../BTL_quanLyGiaDung_backup.docx").toAbsolutePath().normalize();
        if (!Files.exists(backupPath)) {
            Files.write(backupPath, inputBytes);
            System.out.println("Original document backed up to: " + backupPath);
        }

        // Save formatted document back to original location
        Files.write(docxPath, outputBytes);
        System.out.println("Formatted document successfully saved to: " + docxPath);
    }
}
