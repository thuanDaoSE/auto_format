package vn.vietdoc.vietdoc_assistant;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class VietDocEngine {

    public static void main(String[] args) {
        String inputFile = "D:\\\\Workspaces\\\\vietdoc-assistant\\\\nhoCaiDMM.docx";
        String outputFile = "D:\\\\Workspaces\\\\vietdoc-assistant\\\\nhoCaiDMM_res.docx";

        try (FileInputStream fis = new FileInputStream(inputFile);
             XWPFDocument document = new XWPFDocument(fis)) {

            // BƯỚC 1: Setup Margins (Toàn bộ file)
            GeneralFormatter.setupPageMargins(document);

            // BƯỚC 2: Tìm vị trí kết thúc trang bìa (Logic VietDocCoverSolution)
            // Ta cần hàm này trả về index (int) của đoạn văn bắt đầu nội dung chính
            int bodyStartIndex = CoverPageDetector.detect(document) + 1; 
            System.out.println("Nội dung chính bắt đầu từ paragraph index: " + bodyStartIndex);

            // BƯỚC 3: Xử lý Bảng & Ảnh + Caption (Logic ReFormat + SetImageName)
            // Truyền bodyStartIndex vào để không sửa bảng/ảnh ở bìa
            TableImageProcessor.process(document, bodyStartIndex);

            // BƯỚC 4: Xử lý Heading (Logic AutoHeading)
            // Chỉ detect và style heading từ bodyStartIndex trở đi
            HeadingProcessor.process(document, bodyStartIndex);

            // BƯỚC 5: Format văn bản thường (Normal Style - Hard Rules)
            // Áp dụng cho các đoạn KHÔNG PHẢI là Heading, Bảng, Caption
            GeneralFormatter.formatNormalText(document, bodyStartIndex);

            // BƯỚC 6: Tạo Mục lục (Cuối cùng)
            HeadingProcessor.createTOC(document, bodyStartIndex);

            // Lưu file
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                document.write(fos);
                System.out.println("Hoàn tất xử lý!");
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}