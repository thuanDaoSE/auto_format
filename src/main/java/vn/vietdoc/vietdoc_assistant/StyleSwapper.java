package vn.vietdoc.vietdoc_assistant;

import java.io.IOException;
import java.nio.file.*;
import java.net.URI;
import java.util.HashMap;
import java.util.Map;

public class StyleSwapper {

    public static void swapStyles(String docxPath, String newStyleXmlPath) {
        System.out.println(">>> [ZIP SWAP] Đang thay thế file styles.xml bên trong DOCX...");
        
        // Cấu hình để mở file Zip (docx thực chất là zip)
        Map<String, String> env = new HashMap<>();
        env.put("create", "false");
        
        // Chuyển đường dẫn file thành URI
        Path path = Paths.get(docxPath);
        URI uri = URI.create("jar:" + path.toUri());

        try (FileSystem fs = FileSystems.newFileSystem(uri, env)) {
            // 1. Xác định đường dẫn file styles.xml bên trong file Docx
            Path internalStylePath = fs.getPath("/word/styles.xml");

            // 2. Xác định đường dẫn file XML bên ngoài (của bạn)
            Path externalXmlPath = Paths.get(newStyleXmlPath);

            // 3. Thực hiện GHI ĐÈ (REPLACE)
            // Lệnh này copy file ngoài đè vào file trong, giữ nguyên từng byte
            Files.copy(externalXmlPath, internalStylePath, StandardCopyOption.REPLACE_EXISTING);
            
            System.out.println(">>> [ZIP SWAP] Thành công! Đã nạp Style XML chuẩn vào file.");
            
        } catch (IOException e) {
            System.err.println("!!! LỖI KHI TRÁO STYLE: " + e.getMessage());
            e.printStackTrace();
        }
    }
}