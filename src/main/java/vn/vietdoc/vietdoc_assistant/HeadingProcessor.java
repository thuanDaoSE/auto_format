package vn.vietdoc.vietdoc_assistant;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import java.math.BigInteger;
import java.util.List;
import java.util.regex.Pattern;

public class HeadingProcessor {

    private static final String[] LVL1_KEYWORDS = {
            "CHƯƠNG ", "MỞ ĐẦU", "KẾT LUẬN", "LỜI CẢM ƠN", "LỜI MỞ ĐẦU",
            "DANH MỤC", "TÀI LIỆU THAM KHẢO", "PHỤ LỤC", "TÓM TẮT", "LỜI CAM ĐOAN"
    };
    private static final Pattern LVL4_REGEX = Pattern.compile("^\\s*\\d+\\.\\d+\\.\\d+\\.\\d+.*");
    private static final Pattern LVL3_REGEX = Pattern.compile("^\\s*\\d+\\.\\d+\\.\\d+.*");
    private static final Pattern LVL2_REGEX = Pattern.compile("^\\s*\\d+\\.\\d+.*");

    public static void process(XWPFDocument doc, int startIndex) {
        List<XWPFParagraph> paragraphs = doc.getParagraphs();
        for (int i = startIndex; i < paragraphs.size(); i++) {
            XWPFParagraph p = paragraphs.get(i);
            int level = detectLevel(p.getText());

            if (level > 0) {
                p.setStyle("Heading" + level);
                
                // --- FIX: ÁP DỤNG SPACING CHO CẢ HEADING ---
                p.setSpacingBefore(VietDocStyleConfig.SPACING_AFTER); // Heading nên cách trên 1 chút
                p.setSpacingAfter(VietDocStyleConfig.SPACING_AFTER);
                p.setSpacingBetween(VietDocStyleConfig.LINE_SPACING, LineSpacingRule.AUTO);
                
                for (XWPFRun r : p.getRuns()) {
                    // Dùng hàm an toàn, font to hơn chút cho Heading 1
                    int size = (level == 1) ? VietDocStyleConfig.FONT_SIZE_HEADING1 : VietDocStyleConfig.FONT_SIZE_BODY;
                    safeSetFont(r, true, size);
                }
            }
        }
    }

    public static void createTOC(XWPFDocument doc, int startIndex) {
        cleanUpBreaks(doc, startIndex);
        if (startIndex >= doc.getParagraphs().size()) startIndex = Math.max(0, doc.getParagraphs().size() - 1);
        
        XWPFParagraph targetPara = doc.getParagraphs().isEmpty() ? doc.createParagraph() : doc.getParagraphs().get(startIndex);
        XmlCursor cursor = targetPara.getCTP().newCursor();

        // Tiêu đề
        XWPFParagraph titlePara = doc.insertNewParagraph(cursor);
        if (titlePara == null) titlePara = doc.createParagraph();
        titlePara.setAlignment(ParagraphAlignment.CENTER);
        titlePara.setStyle("Heading1");
        XWPFRun titleRun = titlePara.createRun();
        titleRun.setText("MỤC LỤC");
        safeSetFont(titleRun, true, VietDocStyleConfig.FONT_SIZE_HEADING1);
        titleRun.setColor("000000");
        titleRun.addBreak(); 

        // Field TOC
        cursor = targetPara.getCTP().newCursor();
        XWPFParagraph tocPara = doc.insertNewParagraph(cursor);
        if (tocPara == null) tocPara = doc.createParagraph();
        tocPara.setAlignment(ParagraphAlignment.BOTH);
        CTSimpleField toc = tocPara.getCTP().addNewFldSimple();
        toc.setInstr("TOC \\o \"1-3\" \\h \\z \\u");

        // Page Break
        cursor = targetPara.getCTP().newCursor();
        XWPFParagraph breakPara = doc.insertNewParagraph(cursor);
        if (breakPara == null) breakPara = doc.createParagraph();
        breakPara.setPageBreak(true);
    }

   /**
 * HÀM SET FONT AN TOÀN CHO BẢN POI LITE (Wipe & Recreate)
 * Tham khảo logic từ AutoHeadingFinalV2: Dùng XmlCursor xử lý để tránh lỗi thiếu hàm isSet...
 */
/**
 * HÀM SET FONT CHUẨN (Logic từ AutoHeadingFinalV2)
 * Khắc phục triệt để lỗi "The method isSet... undefined" trên bản POI Lite.
 */
private static void safeSetFont(XWPFRun run, boolean isBold, int fontSizePt) {
    if (run == null || run.getCTR() == null) return;

    // BƯỚC 1: Xóa sạch định dạng cũ (Dùng XmlCursor như AutoHeadingFinalV2)
    // CTR.isSetRPr() thì có sẵn, nhưng CTRPr.isSet... thì không -> Phải xóa từ gốc.
    if (run.getCTR().isSetRPr()) {
        CTRPr rPr = run.getCTR().getRPr();
        try (XmlCursor cursor = rPr.newCursor()) {
            cursor.removeXml(); // Xóa sạch thẻ rPr cũ
        }
    }

    // BƯỚC 2: Tạo mới rPr (Lúc này run sạch trơn, chỉ việc addNew)
    CTRPr rPr = run.getCTR().addNewRPr();

    // --- 1. Font Family ---
    // Không check isSet, cứ addNew vì rPr mới tinh
    CTFonts fonts = rPr.addNewRFonts();
    fonts.setAscii(VietDocStyleConfig.FONT_FAMILY);
    fonts.setHAnsi(VietDocStyleConfig.FONT_FAMILY);
    
    // LƯU Ý: Tuyệt đối KHÔNG gọi setCs() hoặc addNewSzCs() nếu dùng bản Lite
    // Vì nó sẽ gây lỗi NoSuchMethodError hoặc làm hỏng cấu trúc XML dẫn đến phình file.

    // --- 2. Font Size ---
    CTHpsMeasure sz = rPr.addNewSz();
    sz.setVal(BigInteger.valueOf(fontSizePt * 2L));

    // --- 3. Bold ---
    if (isBold) {
        CTOnOff bold = rPr.addNewB();
        // Với bản Lite, setVal(true) là an toàn nhất. 
        // Nếu vẫn lỗi, chỉ cần gọi addNewB() là đủ (mặc định là bật).
        bold.setVal(true); 
    }
}

    private static void cleanUpBreaks(XWPFDocument doc, int index) {
        int checkLimit = 3;
        while (checkLimit > 0 && index < doc.getParagraphs().size()) {
            XWPFParagraph p = doc.getParagraphs().get(index);
            String text = p.getText().trim();
            if (text.isEmpty() && p.getRuns().isEmpty()) {
                doc.removeBodyElement(doc.getPosOfParagraph(p));
            } else {
                break;
            }
            checkLimit--;
        }
    }

    private static int detectLevel(String text) {
        if (text == null || text.trim().isEmpty()) return 0;
        String rawText = text.trim();
        String upperText = rawText.toUpperCase();
        for (String keyword : LVL1_KEYWORDS) {
            if (upperText.startsWith(keyword)) return 1;
        }
        if (LVL4_REGEX.matcher(rawText).matches()) return 4;
        if (LVL3_REGEX.matcher(rawText).matches()) return 3;
        if (LVL2_REGEX.matcher(rawText).matches()) return 2;
        return 0;
    }
}