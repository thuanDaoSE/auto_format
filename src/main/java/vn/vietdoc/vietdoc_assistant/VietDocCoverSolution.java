
package vn.vietdoc.vietdoc_assistant;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.regex.Pattern;

/**
 * VIETDOC STARTUP - CORE ENGINE
 * Module: Cover Page Detector (Run on Real File)
 */
public class VietDocCoverSolution {

    public static void main(String[] args) {
        // ĐƯỜNG DẪN ĐẾN FILE CỦA BẠN
        // Lưu ý: Trong Java string, dấu \ phải được escape thành \\
        String filePath = "D:\\Workspaces\\vietdoc-assistant\\coverPageInput.docx";

        System.out.println("==================================================");
        System.out.println("   VIETDOC AUTO-FORMAT: REAL FILE TEST");
        System.out.println("==================================================\n");

        File file = new File(filePath);
        if (!file.exists()) {
            System.err.println("LỖI: Không tìm thấy file tại đường dẫn: " + filePath);
            return;
        }

        try (FileInputStream fis = new FileInputStream(file);
             XWPFDocument doc = new XWPFDocument(fis)) {

            System.out.println(">>> Đang đọc file: " + file.getName());
            
            // CHẠY THUẬT TOÁN
            long startTime = System.currentTimeMillis();
            int cutIndex = CoverPageDetector.detect(doc);
            long endTime = System.currentTimeMillis();

            // IN KẾT QUẢ
            if (cutIndex != -1) {
                System.out.println("   [SUCCESS]: Tìm thấy điểm kết thúc trang bìa.");
                System.out.println("   [INDEX CẮT]: " + cutIndex);
                
                // In nội dung dòng cắt để kiểm tra
                String contentAtCut = doc.getParagraphs().get(cutIndex).getText();
                System.out.println("   [NỘI DUNG DÒNG CẮT]: \"" + contentAtCut + "\"");
                
                System.out.println("   -> Đề xuất: Chèn Section Break (Next Page) vào sau dòng này.");
            } else {
                System.out.println("   [FAILED]: Không nhận diện được trang bìa (hoặc file không có bìa).");
            }
            
            System.out.println("   [THỜI GIAN XỬ LÝ]: " + (endTime - startTime) + "ms");

        } catch (IOException e) {
            System.err.println("LỖI ĐỌC FILE: " + e.getMessage());
            e.printStackTrace();
        }
    }

    // =========================================================================
    // CORE LOGIC (GIỮ NGUYÊN NHƯ CŨ)
    // =========================================================================
    static class CoverPageDetector {

        private static final int SCAN_LIMIT = 80;
        private static final int MAX_CHAR_DENSITY = 250;

        // Regex tìm Năm: "Hà Nội... năm 20xx" hoặc "Tháng... năm..."
        private static final Pattern YEAR_PATTERN = Pattern.compile("(?i).*(năm|year|date).*20[2-9][0-9].*");
        
        // Regex tìm các từ khóa báo hiệu "Hết bìa"
        private static final Pattern STOP_KEYWORDS = Pattern.compile("(?i)^(lời cảm ơn|mục lục|chương|phần|mở đầu|lời nói đầu).*");

        public static int detect(XWPFDocument document) {
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            int limit = Math.min(paragraphs.size(), SCAN_LIMIT);

            boolean hasContentStarted = false;
            int lastDateIndex = -1;

            for (int i = 0; i < limit; i++) {
                XWPFParagraph p = paragraphs.get(i);
                String text = p.getText().trim();

                if (!hasContentStarted) {
                    if (text.isEmpty()) continue;
                    hasContentStarted = true;
                }

                if (text.length() > MAX_CHAR_DENSITY) {
                    return (lastDateIndex != -1) ? lastDateIndex : -1;
                }

                if (STOP_KEYWORDS.matcher(text).matches()) {
                     return (i > 0) ? i - 1 : -1;
                }

                if (hasSectionBreak(p) || p.isPageBreak()) {
                    return i; 
                }

                if (YEAR_PATTERN.matcher(text).matches()) {
                    lastDateIndex = i;
                }
            }

            if (lastDateIndex != -1) {
                return lastDateIndex;
            }

            return -1;
        }

        private static boolean hasSectionBreak(XWPFParagraph p) {
            if (p.getCTP() == null || p.getCTP().getPPr() == null) return false;
            CTPPr pPr = p.getCTP().getPPr();
            return pPr.isSetSectPr();
        }
    }
}