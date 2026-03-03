package vn.vietdoc.vietdoc_assistant;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigInteger;
import java.util.List;
import java.util.regex.Pattern;

public class PageNumberFormatter {

private static final Pattern CHAPTER_1_REGEX = Pattern.compile("^[\\s\\u00A0]*(?:(?:CHƯƠNG|PHẦN)[\\s\\u00A0.:\\-]+)?([1I])(?:\\b|[^0-9IVX]).*", Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE);
    public static void formatPageNumbers(XWPFDocument doc) {
        System.out.println(">>> Bắt đầu chia Section và đánh số trang an toàn...");
        List<XWPFParagraph> paras = doc.getParagraphs();
        
        int splitIdx = -1;
        boolean hasChapter1 = false;

        // 1. TÌM ĐIỂM CẮT (CHƯƠNG 1 hoặc LỜI MỞ ĐẦU)
        for (int i = 0; i < paras.size(); i++) {
            XWPFParagraph p = paras.get(i);
            String text = p.getText().trim().toUpperCase();
            String styleId = p.getStyleID() != null ? p.getStyleID() : "";

            if (styleId.equalsIgnoreCase("Heading1") || styleId.equalsIgnoreCase("Heading 1")) {
                if (CHAPTER_1_REGEX.matcher(text).matches()) {
                    splitIdx = i;
                    hasChapter1 = true;
                    break; // Ưu tiên cao nhất, thấy Chương 1 là chốt luôn
                }
                // Nếu chưa có Chương 1, lấy Lời mở đầu làm điểm cắt dự phòng
                if (!hasChapter1 && (text.equals("LỜI MỞ ĐẦU") || text.equals("MỞ ĐẦU") || text.equals("ĐẶT VẤN ĐỀ"))) {
                    splitIdx = i; 
                }
            }
        }

        CTSectPr originalBodySect = doc.getDocument().getBody().getSectPr();

        // 2. XỬ LÝ CHIA SECTION
        if (splitIdx > 0) {
            System.out.println("   -> Tìm thấy điểm cắt Main Body tại index: " + splitIdx + ". Đang bẻ Section...");

            // [QUAN TRỌNG NHẤT] XÓA PAGE BREAK BEFORE ở Heading để tránh sinh 2 trang trắng 
            // và nhường quyền ngắt trang hoàn toàn cho Section Break
            XWPFParagraph splitPara = paras.get(splitIdx);
            if (splitPara.getCTP().isSetPPr() && splitPara.getCTP().getPPr().isSetPageBreakBefore()) {
                splitPara.getCTP().getPPr().unsetPageBreakBefore();
            }

            // Lấy đoạn văn ngay trước điểm cắt làm mốc để bẻ Section
            XWPFParagraph prevPara = paras.get(splitIdx - 1);
            CTPPr ppr = prevPara.getCTP().isSetPPr() ? prevPara.getCTP().getPPr() : prevPara.getCTP().addNewPPr();
            CTSectPr frontSectPr = ppr.isSetSectPr() ? ppr.getSectPr() : ppr.addNewSectPr();

            // Kế thừa khổ giấy và lề từ Body để chống vỡ form
            if (originalBodySect != null) {
                if (originalBodySect.isSetPgSz() && !frontSectPr.isSetPgSz()) {
                    frontSectPr.setPgSz((CTPageSz) originalBodySect.getPgSz().copy());
                }
                if (originalBodySect.isSetPgMar() && !frontSectPr.isSetPgMar()) {
                    frontSectPr.setPgMar((CTPageMar) originalBodySect.getPgMar().copy());
                }
            }

            // --- A. CẤU HÌNH PHẦN FRONT MATTER (LA MÃ HOA) ---
            CTSectType sectType = frontSectPr.isSetType() ? frontSectPr.getType() : frontSectPr.addNewType();
            sectType.setVal(STSectionMark.NEXT_PAGE);

            CTPageNumber pgNumFront = frontSectPr.isSetPgNumType() ? frontSectPr.getPgNumType() : frontSectPr.addNewPgNumType();
            pgNumFront.setFmt(STNumberFormat.UPPER_ROMAN); // Số La Mã I, II, III...
            pgNumFront.setStart(BigInteger.ONE); // Bắt đầu từ 1

            // Bỏ giấu trang đầu vì Trang Bìa đã được tách thành một Section riêng bởi CoverPageDetector rồi
            if (frontSectPr.isSetTitlePg()) frontSectPr.unsetTitlePg();
            
            createFooter(doc, frontSectPr);

            // --- B. CẤU HÌNH PHẦN MAIN BODY (SỐ THƯỜNG) ---
            CTSectPr bodySectPr = originalBodySect != null ? originalBodySect : doc.getDocument().getBody().addNewSectPr();
            
            CTPageNumber pgNumBody = bodySectPr.isSetPgNumType() ? bodySectPr.getPgNumType() : bodySectPr.addNewPgNumType();
            pgNumBody.setFmt(STNumberFormat.DECIMAL); // Số thường 1, 2, 3...
            pgNumBody.setStart(BigInteger.ONE); // Bắt đầu lại từ 1

            if (bodySectPr.isSetTitlePg()) bodySectPr.unsetTitlePg();
            
            createFooter(doc, bodySectPr);

        } else {
            // NẾU KHÔNG CÓ CẢ CHƯƠNG 1 LẪN LỜI MỞ ĐẦU (Văn bản liền mạch)
            System.out.println("   -> Không tìm thấy điểm cắt. Format số thường cho toàn bộ.");
            CTSectPr bodySectPr = originalBodySect != null ? originalBodySect : doc.getDocument().getBody().addNewSectPr();
            CTPageNumber pgNumBody = bodySectPr.isSetPgNumType() ? bodySectPr.getPgNumType() : bodySectPr.addNewPgNumType();
            pgNumBody.setFmt(STNumberFormat.DECIMAL);
            pgNumBody.setStart(BigInteger.ONE);
            
            if (bodySectPr.isSetTitlePg()) bodySectPr.unsetTitlePg();
            createFooter(doc, bodySectPr);
        }

        // ========================================================
        // --- C. FIX LỖI TRÔI VIỀN TRANG BÌA (BORDER LEAKING) ---
        // ========================================================
        CTSectPr finalBodySect = doc.getDocument().getBody().getSectPr();
        if (finalBodySect != null && finalBodySect.isSetPgBorders()) {
            // 1. Copy thuộc tính viền (Border) từ Section cuối cùng
            CTPageBorders originalBorders = finalBodySect.getPgBorders();

            // 2. Tìm Section Break đầu tiên (của trang bìa) để gán viền trả lại
            boolean isFirstSection = true;
            for (XWPFParagraph p : doc.getParagraphs()) {
                if (p.getCTP().isSetPPr() && p.getCTP().getPPr().isSetSectPr()) {
                    CTSectPr pSect = p.getCTP().getPPr().getSectPr();
                    if (isFirstSection) {
                        pSect.setPgBorders((CTPageBorders) originalBorders.copy());
                        isFirstSection = false;
                    } else {
                        // Quét sạch viền lặp lại ở các Section nằm giữa (Phần La Mã)
                        if (pSect.isSetPgBorders()) pSect.unsetPgBorders();
                    }
                }
            }
            // 3. Xóa viền ở Section cuối (Phần Số thường) để không bị hiện lên ở Chương 1
            finalBodySect.unsetPgBorders();
            System.out.println("   -> Đã luân chuyển Page Border về đúng Trang bìa.");
        }

        System.out.println(">>> [DONE] Định dạng số trang và A4 hoàn tất.");
    }

    // --- HÀM TẠO FOOTER SỐ TRANG CHUẨN ---
    private static void createFooter(XWPFDocument doc, CTSectPr sectPr) {
        try {
            XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(doc, sectPr);
            // 1. Xóa rác Footer cũ
            if (policy.getDefaultFooter() != null) policy.getDefaultFooter().clearHeaderFooter();
            if (policy.getFirstPageFooter() != null) policy.getFirstPageFooter().clearHeaderFooter(); 
            if (policy.getEvenPageFooter() != null) policy.getEvenPageFooter().clearHeaderFooter();

            // ==========================================================
            // 2. [THÊM MỚI] XÓA SẠCH HEADER CŨ ĐỂ KHÔNG BỊ LƯU SỐ TRANG TRÊN ĐẦU
            // ==========================================================
            if (policy.getDefaultHeader() != null) policy.getDefaultHeader().clearHeaderFooter();
            if (policy.getFirstPageHeader() != null) policy.getFirstPageHeader().clearHeaderFooter();
            if (policy.getEvenPageHeader() != null) policy.getEvenPageHeader().clearHeaderFooter();

            XWPFFooter footer = policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
            XWPFParagraph p = footer.createParagraph();
            p.setAlignment(ParagraphAlignment.CENTER);

            // Sinh Field PAGE
            XWPFRun r1 = p.createRun();
            r1.getCTR().addNewFldChar().setFldCharType(STFldCharType.BEGIN);
            p.createRun().getCTR().addNewInstrText().setStringValue(" PAGE ");
            p.createRun().getCTR().addNewFldChar().setFldCharType(STFldCharType.SEPARATE);
            p.createRun().getCTR().addNewFldChar().setFldCharType(STFldCharType.END);

            for (XWPFRun r : p.getRuns()) {
                r.setFontFamily("Times New Roman");
                r.setFontSize(13);
                r.setColor("000000");
                r.setBold(false);
            }
        } catch (Exception e) {
            System.err.println("Lỗi tạo Footer: " + e.getMessage());
        }
    }
}