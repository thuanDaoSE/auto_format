package vn.vietdoc.vietdoc_assistant;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.util.List;
import java.util.regex.Pattern;

/**
 * MODULE: COVER PAGE DETECTOR (NÂNG CẤP)
 * Fix lỗi: Không nhận diện được dòng "Bắc Ninh - 2025" (thiếu chữ "Năm")
 */
public class CoverPageDetector {

    private static final int SCAN_LIMIT = 80;
    private static final int MAX_CHAR_DENSITY = 250;

    // Regex 1: Có chữ "Năm/Year" + Số năm (VD: "Hà Nội, năm 2025")
    private static final Pattern YEAR_KEYWORD_PATTERN = Pattern.compile("(?i).*(năm|year|date).*20[2-9][0-9].*");
    
    // Regex 2 (MỚI): Dạng địa danh + năm (VD: "Bắc Ninh - 2025", "TP.HCM 2026")
    // Tìm chuỗi kết thúc bằng 4 số năm 20xx
    private static final Pattern YEAR_ONLY_PATTERN = Pattern.compile(".*\\b20[2-9][0-9]\\b.*");

    // Từ khóa dừng (Bắt đầu nội dung)
    private static final Pattern STOP_KEYWORDS = Pattern.compile("(?i)^(lời cảm ơn|lời mở đầu|mở đầu|lời cam đoan|mục lục|chương|tóm tắt|nhận xét|phiếu chấm).*");

    public static int detect(XWPFDocument document) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        int limit = Math.min(paragraphs.size(), SCAN_LIMIT);
        int lastDateIndex = -1;

        for (int i = 0; i < limit; i++) {
            XWPFParagraph p = paragraphs.get(i);
            String text = p.getText().trim();

            if (text.isEmpty()) continue;

            // 1. Nếu gặp từ khóa BẮT ĐẦU NỘI DUNG (VD: "Lời mở đầu")
            // -> Thì dòng NGAY TRƯỚC ĐÓ chính là kết thúc bìa.
            if (STOP_KEYWORDS.matcher(text).matches()) {
                // Nếu tìm thấy "Lời mở đầu" ở dòng i, thì bìa kết thúc ở i-1
                return Math.max(0, i - 1);
            }
            
            // 2. Nếu dòng quá dài -> Chắc chắn là nội dung -> Dừng quét
            if (text.length() > MAX_CHAR_DENSITY) {
                break;
            }

            // 3. Detect dòng chứa NĂM (Dấu hiệu cuối bìa)
            // Ưu tiên dòng có chữ "Năm", nếu không thì tìm dòng chứa số 20xx
            if (YEAR_KEYWORD_PATTERN.matcher(text).matches() || YEAR_ONLY_PATTERN.matcher(text).matches()) {
                lastDateIndex = i;
            }
            
            // 4. Nếu có Section Break hoặc Page Break cứng
            if (p.isPageBreak()) {
                return i;
            }
        }

        // Nếu tìm thấy dòng chứa năm, trả về index đó
        if (lastDateIndex != -1) {
            return lastDateIndex;
        }

        // Fallback: Nếu không tìm thấy gì, trả về 0 (coi như không có bìa hoặc bìa chỉ 1 dòng)
        return 0;
    }
}