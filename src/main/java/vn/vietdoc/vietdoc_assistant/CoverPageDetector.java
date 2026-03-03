package vn.vietdoc.vietdoc_assistant;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

public class CoverPageDetector {

    private static final int SCAN_LIMIT = 24; 

    private static final Pattern YEAR_PATTERN = Pattern.compile(".*\\b20[0-3][0-9]\\b.*");
    private static final Pattern BLACKLIST_KEYWORDS = Pattern.compile("(?i).*(mã|sv|student|id|lớp|class|đề tài|topic|gvhd|giảng viên).*");
    private static final Pattern STOP_KEYWORDS = Pattern.compile("(?i)^(lời mở đầu|mở đầu|lời cảm ơn|tóm tắt|abstract|chương 1|phần 1).*");

    // Danh sách 63 tỉnh thành chuẩn
    private static final String PROVINCES_REGEX = "AN GIANG|BÀ RỊA - VŨNG TÀU|BÀ RỊA VŨNG TÀU|BẠC LIÊU|BẮC GIANG|BẮC KẠN|BẮC NINH|BẾN TRE|BÌNH DƯƠNG|BÌNH ĐỊNH|BÌNH PHƯỚC|BÌNH THUẬN|CÀ MAU|CAO BẰNG|CẦN THƠ|ĐÀ NẴNG|ĐẮK LẮK|ĐẮK NÔNG|ĐIỆN BIÊN|ĐỒNG NAI|ĐỒNG THÁP|GIA LAI|HÀ GIANG|HÀ NAM|HÀ NỘI|HÀ TĨNH|HẢI DƯƠNG|HẢI PHÒNG|HẬU GIANG|HÒA BÌNH|HƯNG YÊN|KHÁNH HÒA|KIÊN GIANG|KON TUM|LAI CHÂU|LẠNG SƠN|LÀO CAI|LÂM ĐỒNG|LONG AN|NAM ĐỊNH|NGHỆ AN|NINH BÌNH|NINH THUẬN|PHÚ THỌ|PHÚ YÊN|QUẢNG BÌNH|QUẢNG NAM|QUẢNG NGÃI|QUẢNG NINH|QUẢNG TRỊ|SÓC TRĂNG|SƠN LA|TÂY NINH|THÁI BÌNH|THÁI NGUYÊN|THANH HÓA|THỪA THIÊN HUẾ|HUẾ|TIỀN GIANG|TP HỒ CHÍ MINH|HỒ CHÍ MINH|TPHCM|TP\\.HCM|TP\\. HỒ CHÍ MINH|TRÀ VINH|TUYÊN QUANG|VĨNH LONG|VĨNH PHÚC|YÊN BÁI";
    private static final Pattern PROVINCE_PATTERN = Pattern.compile(".*\\b(" + PROVINCES_REGEX + ")\\b.*", Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE);


    // ==========================================
    // DỌN SẠCH TOC KHỎI TRANG BÌA
    // ==========================================
    public static int removeHeadingsFromCover(XWPFDocument doc) {
        int coverEndElementIndex = detect(doc); 
        System.out.println(">>> [CoverPageDetector] Trang bìa kết thúc tại Element index: " + coverEndElementIndex);
        
        int currentElementCount = 0; 
        for (int i = 0; i < doc.getBodyElements().size(); i++) {
            IBodyElement elem = doc.getBodyElements().get(i);
            
            if (elem instanceof XWPFParagraph) {
                cleanCoverParagraph((XWPFParagraph) elem);
            } else if (elem instanceof XWPFTable) {
                XWPFTable table = (XWPFTable) elem;
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph p : cell.getParagraphs()) {
                            cleanCoverParagraph(p); 
                        }
                    }
                }
            }
            if (currentElementCount == coverEndElementIndex) break; 
            currentElementCount++;
        }
        return coverEndElementIndex;
    }

    private static void cleanCoverParagraph(XWPFParagraph p) {
        if (p.getCTP().isSetPPr() && p.getCTP().getPPr().isSetOutlineLvl()) {
            p.getCTP().getPPr().unsetOutlineLvl(); 
        }
        String style = p.getStyleID();
        if (style != null && style.toLowerCase().contains("heading")) {
            p.setStyle("Normal"); 
        }
    }


    // ==========================================
    // THUẬT TOÁN NHẬN DIỆN BÌA (XUYÊN THẤU BẢNG)
    // ==========================================
    public static int detect(XWPFDocument doc) {
        List<IBodyElement> elements = doc.getBodyElements();
        int limit = Math.min(elements.size(), SCAN_LIMIT);

        int explicitBreakIndex = -1;
        int keywordIndex = -1; 

        for (int i = 0; i < limit; i++) {
            IBodyElement elem = elements.get(i);
            
            // [GIẢI PHÁP MẤU CHỐT]: Bóc tách element ra thành danh sách các Paragraph con
            List<XWPFParagraph> parasToCheck = new ArrayList<>();
            if (elem instanceof XWPFParagraph) {
                parasToCheck.add((XWPFParagraph) elem);
            } else if (elem instanceof XWPFTable) {
                XWPFTable table = (XWPFTable) elem;
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        parasToCheck.addAll(cell.getParagraphs());
                    }
                }
            }

            boolean stopFound = false;

            // Quét từng đoạn văn nhỏ (Cho dù nó nằm trong bảng hay bên ngoài)
            for (XWPFParagraph p : parasToCheck) {
                String text = p.getText().trim();
                String upperText = text.toUpperCase();

                if (text.isEmpty()) continue;

                // 1. Chặn lại ngay lập tức nếu gặp phần nội dung (Lời cảm ơn, Mục lục...)
                if (STOP_KEYWORDS.matcher(text).matches()) {
                    stopFound = true;
                    break;
                }

                // 2. Ghi nhận Break cứng
                if (containsPageBreak(p)) {
                    explicitBreakIndex = i; // Lưu lại vị trí Element (i) chứa lệnh ngắt trang
                }

                // 3. Ghi nhận Năm, SV
                if (YEAR_PATTERN.matcher(text).matches() || BLACKLIST_KEYWORDS.matcher(text).matches()) {
                    keywordIndex = i;
                }

                // 4. Nhận diện Tỉnh thành độc lập từng dòng
                if (PROVINCE_PATTERN.matcher(upperText).matches()) {
                    if (!upperText.startsWith("TRƯỜNG") && !upperText.startsWith("ĐẠI HỌC")) {
                        keywordIndex = i;
                    }
                }
            }

            if (stopFound) {
                limit = i; 
                break;
            }
        }

        // --- ĐÁNH GIÁ ƯU TIÊN CUỐI CÙNG ---
        if (explicitBreakIndex != -1 && explicitBreakIndex >= keywordIndex) {
            return explicitBreakIndex;
        }
        
        if (keywordIndex != -1) {
            return keywordIndex;
        }

        return limit > 0 ? limit - 1 : 0;
    }

    // --- HÀM BẮT BREAK CHUYÊN SÂU ---
    private static boolean containsPageBreak(XWPFParagraph p) {
        if (p.getCTP().getPPr() != null && p.getCTP().getPPr().isSetSectPr()) return true;
        if (p.isPageBreak()) return true;
        if (p.getCTP().getPPr() != null && p.getCTP().getPPr().isSetPageBreakBefore()) return true; 
        
        for (XWPFRun r : p.getRuns()) {
            if (r.getCTR() != null) {
                List<CTBr> brList = r.getCTR().getBrList();
                if (brList != null) {
                    for (CTBr br : brList) {
                        if (br.getType() == STBrType.PAGE) {
                            return true;
                        }
                    }
                }
            }
        }
        return false;
    }
}