package vn.vietdoc.vietdoc_assistant;

import org.apache.poi.xwpf.usermodel.*;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.SectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigInteger;
import java.util.List;

public class GeneralFormatter {

    // 1.27 cm = 720 twips
    private static final int INDENT_FIRST_LINE = 567; 
    private static final int SPACING_BEFORE = 120; // 6pt
    private static final int SPACING_AFTER = 0;    // 0pt
    
    // Ngưỡng ký tự để quyết định Justify (Khoảng 1 dòng A4 font 13 là 70-80 ký tự)
    private static final int JUSTIFY_THRESHOLD = 70; 

    public static void formatNormalText(XWPFDocument document, int startIdx, FormattingParameters params) {
        System.out.println(">>> Đang chuẩn hóa Normal (Hỗ trợ List Item + Fix Soft Break)...");
        
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (int i = startIdx; i < paragraphs.size(); i++) {
            XWPFParagraph p = paragraphs.get(i);
            String styleID = p.getStyleID();
            String text = p.getText();

            // --- 1. BỘ LỌC CƠ BẢN ---
            if (text == null || text.trim().isEmpty()) continue;
            
            // Bỏ qua Heading, Title, TOC
            if (styleID != null && (styleID.startsWith("Heading") || styleID.startsWith("Title") || styleID.startsWith("TOC") || styleID.startsWith("FakeHeading"))) {
                continue;
            }
            // Bỏ qua đoạn văn căn giữa (chú thích ảnh)
            if (p.getAlignment() == ParagraphAlignment.CENTER) continue;

            // Kiểm tra xem có phải là List Item (Bullet/Numbering) không
            boolean isList = (p.getNumID() != null);

            // --- 2. THIẾT LẬP THỤT ĐẦU DÒNG (INDENT) ---
            // [QUAN TRỌNG] Chỉ Reset Indent với đoạn văn thường. 
            // KHÔNG can thiệp vào List Item vì sẽ làm hỏng style của Bullet.
            CTPPr ppr = p.getCTP().isSetPPr() ? p.getCTP().getPPr() : p.getCTP().addNewPPr();
            
            if (!isList) {
                if (ppr.isSetInd()) ppr.unsetInd(); // Xóa sạch cái cũ
                CTInd ind = ppr.addNewInd();
                ind.setFirstLine(BigInteger.valueOf((long) (params.getIndentationFirstLine() * 567)));
            }

            // --- 3. XỬ LÝ ALIGNMENT & SOFT BREAK (ÁP DỤNG CHO CẢ LIST) ---
            // Logic này giờ sẽ chạy cho cả List Item để fix lỗi "Ảnh hưởng... Accuracy:"
            
            String cleanText = text.trim().replace('\u00A0', ' '); 
            boolean hasSoftBreak = hasManualLineBreak(p); // Phát hiện Shift+Enter (<w:br>)
            boolean endsWithColon = cleanText.endsWith(":") || cleanText.endsWith("：");
            boolean isShortLine = cleanText.length() < JUSTIFY_THRESHOLD;

            if (hasSoftBreak || endsWithColon || isShortLine) {
                // Nếu dính Soft Break hoặc dấu 2 chấm -> Bắt buộc Left để tránh giãn dòng
                p.setAlignment(ParagraphAlignment.LEFT);
            } else {
                p.setAlignment(ParagraphAlignment.BOTH); 
            }

            // --- 4. RESET SPACING (6pt - 0pt) ---
            // Áp dụng cho cả List luôn để đồng bộ
            if (ppr.isSetSpacing()) {
                CTSpacing spacing = ppr.getSpacing();
                spacing.setBefore(BigInteger.valueOf(SPACING_BEFORE));
                spacing.setAfter(BigInteger.valueOf(SPACING_AFTER));
            } else {
                CTSpacing spacing = ppr.addNewSpacing();
                spacing.setBefore(BigInteger.valueOf(SPACING_BEFORE));
                spacing.setAfter(BigInteger.valueOf(SPACING_AFTER));
            }



            // --- 5. CHUẨN HÓA FONT ---
            for (XWPFRun run : p.getRuns()) {
                run.setFontFamily("Times New Roman");
                run.setFontSize(params.getFontSizeBody());

                CTSpacing spacing = ppr.addNewSpacing();
                spacing.setLine(BigInteger.valueOf(360)); // 360 = 1.5 lines
                spacing.setLineRule(STLineSpacingRule.AUTO);
            }
        }
        System.out.println(">>> [DONE] Đã format xong Normal text.");
    }

    /**
     * Soi kính lúp vào XML để tìm thẻ <w:br> (Shift+Enter)
     */
    private static boolean hasManualLineBreak(XWPFParagraph p) {
        for (XWPFRun r : p.getRuns()) {
            if (r.getCTR().getBrList() != null && !r.getCTR().getBrList().isEmpty()) {
                return true;
            }
        }
        return false;
    }

    // // Giữ nguyên các hàm khác (enforceA4Paper)...
    // public static void enforceA4Paper(XWPFDocument doc) {
    //     CTSectPr sectPr = doc.getDocument().getBody().getSectPr();
    //     if (sectPr == null) sectPr = doc.getDocument().getBody().addNewSectPr();
        
    //     CTPageSz pageSize = sectPr.isSetPgSz() ? sectPr.getPgSz() : sectPr.addNewPgSz();
    //     pageSize.setW(BigInteger.valueOf(11906));
    //     pageSize.setH(BigInteger.valueOf(16838));
        
    //     CTPageMar pageMar = sectPr.isSetPgMar() ? sectPr.getPgMar() : sectPr.addNewPgMar();
    //     pageMar.setTop(BigInteger.valueOf(1134));    
    //     pageMar.setBottom(BigInteger.valueOf(1134)); 
    //     pageMar.setLeft(BigInteger.valueOf(1701));   
    //     pageMar.setRight(BigInteger.valueOf(1134));  
    // }

    public static void fixPageNumbering(WordprocessingMLPackage wordMLPackage) {
        try {
            List<Object> bodyContent = wordMLPackage.getMainDocumentPart().getContent();
            
            // 1. Duyệt qua Body để tìm SectPr cuối cùng (quan trọng nhất)
            SectPr bodySectPr = wordMLPackage.getMainDocumentPart().getJaxbElement().getBody().getSectPr();
            if (bodySectPr != null && bodySectPr.getPgNumType() != null) {
                // Xóa thiết lập bắt đầu từ 0 đi -> Để Word tự đếm từ 1
                bodySectPr.setPgNumType(null); 
                System.out.println(">>> [FIX] Đã xóa pgNumType=0 ở cuối tài liệu.");
            }
    
            // 2. Duyệt qua các Paragraph để tìm SectPr ẩn (nếu có chia section giữa bài)
            for (Object o : bodyContent) {
                if (o instanceof org.docx4j.wml.P) {
                    org.docx4j.wml.P p = (org.docx4j.wml.P) o;
                    if (p.getPPr() != null && p.getPPr().getSectPr() != null) {
                        SectPr pSectPr = p.getPPr().getSectPr();
                        if (pSectPr.getPgNumType() != null) {
                            pSectPr.setPgNumType(null);
                            System.out.println(">>> [FIX] Đã xóa pgNumType=0 ở giữa tài liệu.");
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.err.println(">>> Lỗi khi fix số trang: " + e.getMessage());
        }
    }
}