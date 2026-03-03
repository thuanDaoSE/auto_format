package vn.vietdoc.vietdoc_assistant.oldFiles;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.regex.Pattern;

public class AutoHeadingFinalV2 {

    // --- CẤU HÌNH LOGIC ---
    private static final String[] LVL1_KEYWORDS = {
            "CHƯƠNG ", "MỞ ĐẦU", "KẾT LUẬN", "LỜI CẢM ƠN", "LỜI MỞ ĐẦU",
            "DANH MỤC", "TÀI LIỆU THAM KHẢO", "PHỤ LỤC", "TÓM TẮT"
    };

    private static final Pattern LVL4_REGEX = Pattern.compile("^\\s*\\d+\\.\\d+\\.\\d+\\.\\d+.*");
    private static final Pattern LVL3_REGEX = Pattern.compile("^\\s*\\d+\\.\\d+\\.\\d+.*");
    private static final Pattern LVL2_REGEX = Pattern.compile("^\\s*\\d+\\.\\d+.*");
    private static final Pattern LVL1_REGEX_NUM = Pattern.compile("^\\s*\\d+\\.\\s+.*");

    public static void main(String[] args) {
        String inputFile = "D:\\Workspaces\\vietdoc-assistant\\nhoCaiDMM.docx";
        String outputFile = "D:\\Workspaces\\vietdoc-assistant\\nhoCaiDMM_Done.docx";

        processDocx(inputFile, outputFile);
    }

    public static void processDocx(String inputPath, String outputPath) {
        System.out.println("... Đang xử lý file: " + inputPath);

        try (FileInputStream fis = new FileInputStream(inputPath);
             XWPFDocument document = new XWPFDocument(fis)) {

            List<XWPFParagraph> paragraphs = document.getParagraphs();

            // ========================================================
            // BƯỚC 1: HARD RESET (Xóa Heading cũ + Xóa Outline Level)
            // ========================================================
            for (XWPFParagraph p : paragraphs) {
                // Xóa Outline Level (Fix lỗi rác TOC)
                CTP ctp = p.getCTP();
                if (ctp.getPPr() != null && ctp.getPPr().isSetOutlineLvl()) {
                    ctp.getPPr().unsetOutlineLvl();
                }
                // Reset Style Heading về Normal
                String styleID = p.getStyleID();
                if (styleID != null && styleID.startsWith("Heading")) {
                    p.setStyle("Normal");
                }
            }

            // ========================================================
            // BƯỚC 2: NHẬN DIỆN & GÁN HEADING MỚI
            // ========================================================
            int appliedCount = 0;
            for (XWPFParagraph p : paragraphs) {
                String text = p.getText();
                int level = detectLevel(text);

                if (level > 0) {
                    // Xóa định dạng thủ công (Bold/Size...) để ăn theo Style mới
                    // Dùng XmlCursor -> KHÔNG LO LỖI THƯ VIỆN LITE/FULL
                    resetRunFormatting(p);

                    // Gán Style
                    p.setStyle("Heading" + level);
                    appliedCount++;
                }
            }
            System.out.println("-> Đã tái cấu trúc " + appliedCount + " headings.");

            // ========================================================
            // BƯỚC 3: XỬ LÝ MỤC LỤC & CÀI ĐẶT UPDATE
            // ========================================================
            updateOrInsertTOC(document);

            // QUAN TRỌNG: Lệnh này ép Word cập nhật toàn bộ Field khi mở file
            // Nó sẽ hiện thông báo "Update Document?" hoặc tự động update tùy phiên bản Word
            document.enforceUpdateFields(); 

            // ========================================================
            // BƯỚC 4: LƯU FILE
            // ========================================================
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                document.write(fos);
                System.out.println("-> Hoàn tất! File tại: " + outputPath);
                System.out.println("-> KHI MỞ FILE: Nếu Word hỏi 'Do you want to update fields?', hãy chọn YES.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Hàm xóa định dạng thủ công bằng XmlCursor (Chạy được trên mọi bản POI)
     */
    private static void resetRunFormatting(XWPFParagraph p) {
        for (XWPFRun run : p.getRuns()) {
            CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : null;
            if (rpr != null) {
                // Xóa sạch thẻ rPr (Run Properties) trong XML
                XmlCursor cursor = rpr.newCursor();
                cursor.removeXml(); // Bùm! Mất hết định dạng rác
            }
        }
    }

    /**
     * Hàm xử lý Mục Lục
     */
    public static void updateOrInsertTOC(XWPFDocument doc) {
        // Kiểm tra xem đã có TOC chưa
        boolean tocExists = false;
        for (XWPFParagraph p : doc.getParagraphs()) {
            if (p.getCTP().toString().contains("TOC \\o")) {
                tocExists = true;
                break;
            }
        }

        if (!tocExists) {
            // Tìm vị trí chèn (Trước Heading đầu tiên)
            int targetPos = 0;
            for (int i = 0; i < doc.getParagraphs().size(); i++) {
                XWPFParagraph p = doc.getParagraphs().get(i);
                if (p.getStyleID() != null && p.getStyleID().startsWith("Heading")) {
                    targetPos = i;
                    break;
                }
            }
            // An toàn
            if (targetPos < 0) targetPos = 0;

            // Tạo tiêu đề
            XWPFParagraph titleToc = doc.insertNewParagraph(doc.getParagraphs().get(targetPos).getCTP().newCursor());
            titleToc.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun run = titleToc.createRun();
            run.setText("MỤC LỤC TỰ ĐỘNG");
            run.setBold(true);
            run.setFontSize(14);
            run.setColor("2E74B5"); // Màu xanh điểm nhấn

            // Tạo Field TOC
            XWPFParagraph tocPara = doc.insertNewParagraph(doc.getParagraphs().get(targetPos + 1).getCTP().newCursor());
            CTSimpleField toc = tocPara.getCTP().addNewFldSimple();
            toc.setInstr("TOC \\o \"1-3\" \\h \\z \\u");
            
            // Set dirty = true (Đánh dấu cần update)
            // Vì dùng thư viện POI thường, có thể cần set thông qua XML nếu hàm setDirty(boolean) lỗi
            // Nhưng với document.enforceUpdateFields() ở trên thì dòng này là phụ trợ thôi.
            try {
                toc.setDirty(true);
            } catch (Exception e) {
                 // Nếu thư viện cũ quá không có setDirty, bỏ qua vì đã có enforceUpdateFields
            }
        }
    }

    public static int detectLevel(String text) {
        if (text == null || text.trim().isEmpty()) return 0;
        String rawText = text.trim();
        String upperText = rawText.toUpperCase();

        for (String keyword : LVL1_KEYWORDS) {
            if (upperText.startsWith(keyword)) return 1;
        }
        if (LVL4_REGEX.matcher(rawText).matches()) return 4;
        if (LVL3_REGEX.matcher(rawText).matches()) return 3;
        if (LVL2_REGEX.matcher(rawText).matches()) return 2;
        if (LVL1_REGEX_NUM.matcher(rawText).matches()) return 1;

        return 0;
    }
}