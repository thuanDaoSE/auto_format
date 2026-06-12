package vn.vietdoc.vietdoc_assistant;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.apache.xmlbeans.XmlCursor;
import java.math.BigInteger;
import java.util.List;
import java.util.regex.Pattern;

public class HeadingProcessor {

    // --- CONSTANTS ---
    private static final String[] LVL1_TEXT_KEYWORDS = {
            "CHƯƠNG ", "MỞ ĐẦU", "KẾT LUẬN", "LỜI CẢM ƠN", "LỜI MỞ ĐẦU",
            "DANH MỤC", "TÀI LIỆU THAM KHẢO", "PHỤ LỤC", "TÓM TẮT", "LỜI CAM ĐOAN"
    };
    private static final Pattern LVL1_NUMBER_REGEX = Pattern.compile("^\\s*\\d+\\.\\s+.*");
    private static final Pattern LVL4_REGEX = Pattern.compile("^\\s*\\d+\\.\\d+\\.\\d+\\.\\d+.*");
    private static final Pattern LVL3_REGEX = Pattern.compile("^\\s*\\d+\\.\\d+\\.\\d+.*");
    private static final Pattern LVL2_REGEX = Pattern.compile("^\\s*\\d+\\.\\d+.*");
    // --- KHAI BÁO REGEX NHẬN DIỆN CHƯƠNG (Để kích hoạt Main Body) ---
    private static final Pattern CHAPTER_START_REGEX = Pattern.compile("^[\\s\\u00A0]*(CHƯƠNG|PHẦN)[\\s\\u00A0.:\\-]*([0-9IVX]+).*", Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE);
    // --- DANH SÁCH CÁC TIÊU ĐỀ FRONT MATTER ĐƯỢC PHÉP LÀ HEADING 1 ---
    private static final String[] FRONT_MATTER_TITLES = {
            "LỜI CAM ĐOAN", "LỜI CẢM ƠN", "MỤC LỤC", 
            "DANH MỤC HÌNH ẢNH", "DANH MỤC CÁC HÌNH ẢNH", "DANH MỤC BẢNG BIỂU","DANH MỤC CÁC BẢNG BIỂU", "DANH MỤC TỪ VIẾT TẮT", 
            "LỜI MỞ ĐẦU", "MỞ ĐẦU", "TÓM TẮT", "NHẬN XÉT CỦA GIẢNG VIÊN"
    };

    // ==========================================
    // 1. MAIN PROCESS (XỬ LÝ HEADING)
    // ==========================================
    // ==========================================
    // 1. MAIN PROCESS (XỬ LÝ HEADING)
    // ==========================================
    public static void process(XWPFDocument doc, int startIndex) {
        List<XWPFParagraph> paragraphs = doc.getParagraphs();
        int total = paragraphs.size();
        if (startIndex >= total) return;

        // Biến cờ kiểm tra xem toàn bài có dùng Chương/Phần không
        boolean hasChapters = false;
        
        for (XWPFParagraph p : paragraphs) {
            demoteToBodyText(p);
            // Quét trước một lượt
            if (isChapterStart(p.getText().trim().toUpperCase())) {
                hasChapters = true;
            }
        }

        for (int i = startIndex; i < total; i++) {
            XWPFParagraph p = paragraphs.get(i);
            String text = p.getText().trim();
            
            // 1. ĐOÁN LEVEL HEADING
            int finalLevel = detectLevel(text, hasChapters);

            // 2. ÁP DỤNG LUẬT CỨNG ĐỂ CHẶN HEADING SAI
            if (finalLevel > 0) {
                if (isBulletOrList(p) || (text.contains(":") && finalLevel == 2)) {
                    finalLevel = 0;
                }
            }

            // 3. FORMAT
            if (finalLevel > 0) {
                if (p.getStyleID() != null) p.setStyle(""); 
                resetFormatting(p);
                
                // Cờ chặn đưa vào Mục Lục
                boolean skipTOC = false;
                if (hasChapters && LVL1_NUMBER_REGEX.matcher(text).matches()) {
                    skipTOC = true;
                }

                CTPPr ppr = p.getCTP().isSetPPr() ? p.getCTP().getPPr() : p.getCTP().addNewPPr();
                if (ppr.isSetOutlineLvl()) ppr.unsetOutlineLvl();
                
                // Gán Style và Outline Level
                if (!skipTOC) {
                    p.setStyle("Heading" + finalLevel); 
                    if (finalLevel == 1 && (isLevel1Text(text) || isChapterStart(text))) {
                        ppr.addNewOutlineLvl().setVal(BigInteger.ZERO); // Lọt vào TOC (Level 1)
                    } else {
                        ppr.addNewOutlineLvl().setVal(BigInteger.valueOf(finalLevel - 1));
                    }
                } else {
                    p.setStyle("Normal"); 
                }
                
                // [ĐÃ KHÔI PHỤC] CĂN LỀ VÀ NGẮT TRANG (PAGE BREAK)
                if (finalLevel == 1 && !skipTOC) {
                    if (isLevel1Text(text) || isChapterStart(text)) {
                        p.setAlignment(ParagraphAlignment.CENTER);
                    } else { 
                        p.setAlignment(ParagraphAlignment.LEFT); 
                        forceZeroIndent(p); 
                    }
                } else {
                    p.setAlignment(ParagraphAlignment.LEFT);
                    forceZeroIndent(p);
                }

                // Logic ngắt trang: Chỉ áp dụng cho Heading 1 KHÔNG bị skipTOC
                boolean shouldPageBreak = (finalLevel == 1 && !skipTOC);
                p.setPageBreak(shouldPageBreak); // Set lệnh ngắt trang cứng của POI
                fixHeadingLayout(p, shouldPageBreak);
                
                // Cấu hình Font chữ
                for (XWPFRun r : p.getRuns()) {
                    r.setFontFamily("Times New Roman"); 
                    r.setColor("000000");
                    if (finalLevel == 1 && !skipTOC) {
                        r.setBold(true); 
                        r.setFontSize(14);
                        String runText = r.getText(0);
                        if (runText != null && !runText.isEmpty()) {
                            r.setText(runText.toUpperCase(), 0);
                        }
                    } else {
                        r.setBold(true); // Đảm bảo Heading 2 và heading bị skip đều in đậm
                        r.setFontSize(14);
                    }
                }
            } 
        }
    }

    // ==========================================
    // 2. TOC HANDLING (ĐA MỤC LỤC THÔNG MINH)
    // ==========================================
    public static void addOrUpdateTOC(XWPFDocument doc, int startIndex) {
        // 1. DỌN DẸP SẠCH SẼ CÁC TOC CŨ (Bao gồm cả "Contents", "Mục lục" cũ)
        removeOldTOCs(doc);

        // 2. TÌM VỊ TRÍ CHÈN (Chèn TRƯỚC Lời mở đầu hoặc Chương 1)
        int insertPos = findInsertPosition(doc, startIndex);

        // Xác định đoạn văn mốc (Anchor) để chèn
        XWPFParagraph anchorPara;
        if (doc.getParagraphs().isEmpty() || insertPos >= doc.getParagraphs().size()) {
            anchorPara = doc.createParagraph(); 
        } else {
            anchorPara = doc.getParagraphs().get(insertPos);
        }

        // 3. CHÈN LẦN LƯỢT THEO THỨ TỰ TỪ TRÊN XUỐNG DƯỚI
        // [ĐÃ SỬA LỖI]: Phải chèn Mục lục trước, rồi mới đến Hình, rồi mới đến Bảng.
        
        // 3.1. Mục Lục Chính (Nằm trên cùng)
        createMainTOC(doc, anchorPara);

        // 3.2. Danh mục Hình ảnh (Nằm giữa)
        if (isStyleUsed(doc, "HinhStyle")) {
            createListByStyle(doc, anchorPara, "DANH MỤC CÁC HÌNH ẢNH", "HinhStyle");
        }

        // 3.3. Danh mục Bảng biểu (Nằm dưới cùng)
        if (isStyleUsed(doc, "BangStyle")) {
            createListByStyle(doc, anchorPara, "DANH MỤC BẢNG BIỂU", "BangStyle");
        }
    }

    // --- HÀM KIỂM TRA STYLE CÓ ĐƯỢC DÙNG KHÔNG ---
    private static boolean isStyleUsed(XWPFDocument doc, String styleId) {
        for (XWPFParagraph p : doc.getParagraphs()) {
            if (styleId.equals(p.getStyleID())) {
                return true;
            }
        }
        return false;
    }

    // --- TẠO MỤC LỤC CHÍNH ---
    private static void createMainTOC(XWPFDocument doc, XWPFParagraph anchor) {
        XmlCursor cursor = anchor.getCTP().newCursor();

        // 1. Chèn khoảng trắng mốc
        XWPFParagraph br = doc.insertNewParagraph(cursor);

        // 2. Chèn TOC
        XWPFParagraph tocPara = doc.insertNewParagraph(br.getCTP().newCursor());
        createTOCField(tocPara, "TOC \\o \"1-4\" \\h \\z \\u");

        // 3. Chèn Tiêu đề
        XWPFParagraph title = doc.insertNewParagraph(tocPara.getCTP().newCursor());
        setupTitle(title, "MỤC LỤC", true);
    }

    // --- TẠO DANH MỤC THEO STYLE (THỨ TỰ CHUẨN + CHECK PAGE BREAK) ---
    // --- TẠO DANH MỤC THEO STYLE ---
    private static XWPFParagraph createListByStyle(XWPFDocument doc, XWPFParagraph anchor, String titleText, String styleName) {
        XmlCursor cursor = anchor.getCTP().newCursor();

        // [SỬA] Bỏ lệnh setPageBreak(true) ở đây
        XWPFParagraph br = doc.insertNewParagraph(cursor);
        // br.setPageBreak(true); <--- XÓA DÒNG NÀY ĐI

        // 2. Chèn TOC
        XWPFParagraph tocPara = doc.insertNewParagraph(br.getCTP().newCursor());
        createTOCField(tocPara, "TOC \\h \\z \\t \"" + styleName + ",1\"");

        // 3. Chèn Tiêu đề
        XWPFParagraph title = doc.insertNewParagraph(tocPara.getCTP().newCursor());
        
        // Vẫn giữ true để tiêu đề này luôn bắt đầu ở trang mới
        setupTitle(title, titleText, true); 
        
        return br;
    }

    // --- HELPER: Setup Tiêu đề ---
    // --- HELPER: Setup Tiêu đề cho các Danh mục ---
    private static void setupTitle(XWPFParagraph p, String text, boolean pageBreakBefore) {
        // 1. Set Style là Heading 1
        p.setStyle("Heading1"); 
        p.setAlignment(ParagraphAlignment.CENTER);
        
        // 2. [QUAN TRỌNG] Gán Outline Level = 0 (Tương đương Level 1 để lọt vào Mục lục)
        org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr ppr = 
            p.getCTP().isSetPPr() ? p.getCTP().getPPr() : p.getCTP().addNewPPr();
        if (ppr.isSetOutlineLvl()) ppr.unsetOutlineLvl();
        ppr.addNewOutlineLvl().setVal(java.math.BigInteger.ZERO); 
        
        // 3. Layout và Format
        fixHeadingLayout(p, pageBreakBefore); 
        
        XWPFRun r = p.createRun();
        r.setText(text);
        r.setFontFamily("Times New Roman"); 
        r.setFontSize(14); 
        r.setBold(true); 
        r.setColor("000000");
    }

    // --- HELPER: Setup Field Code ---
    // Thay thế hàm createTOCField cũ
    private static void createTOCField(XWPFParagraph p, String instruction) {
        p.setAlignment(ParagraphAlignment.BOTH); 
        
        // Xóa sạch định dạng cũ (indent, tabs) để TOC thẳng hàng
        if (p.getCTP().getPPr() != null) {
            if (p.getCTP().getPPr().isSetInd()) p.getCTP().getPPr().unsetInd();
            if (p.getCTP().getPPr().isSetTabs()) p.getCTP().getPPr().unsetTabs();
        }
        
        // Run 1: BEGIN (Bắt đầu Field)
        XWPFRun r1 = p.createRun();
        CTFldChar fldChar = r1.getCTR().addNewFldChar();
        fldChar.setFldCharType(STFldCharType.BEGIN);
        fldChar.setDirty(true); // Cờ Dirty: Ép Word hiện popup Update khi mở file
        
        // Run 2: INSTR (Chứa câu lệnh TOC)
        XWPFRun r2 = p.createRun();
        CTText instrText = r2.getCTR().addNewInstrText();
        // Thêm khoảng trắng an toàn 2 đầu
        instrText.setStringValue(" " + instruction + " "); 
        
        // Run 3: SEPARATE (Ngăn cách) -> ĐÂY LÀ CÁI BẠN THIẾU
        // Nếu thiếu cái này, Word sẽ in instruction ra như văn bản thường
        XWPFRun r3 = p.createRun();
        r3.getCTR().addNewFldChar().setFldCharType(STFldCharType.SEPARATE);
        
        // Run 4: RESULT (Kết quả hiển thị tạm)
        XWPFRun rResult = p.createRun();
        rResult.setText(""); // Để trống cho đẹp
        
        // Run 5: END (Kết thúc Field)
        XWPFRun r4 = p.createRun();
        r4.getCTR().addNewFldChar().setFldCharType(STFldCharType.END);
    }

    // ==========================================
    // 3. HELPER METHODS (LAYOUT FIXES)
    // ==========================================
    private static void fixHeadingLayout(XWPFParagraph p, boolean isPageBreakBefore) {
        CTPPr ppr = p.getCTP().isSetPPr() ? p.getCTP().getPPr() : p.getCTP().addNewPPr();

        if (ppr.isSetKeepNext()) ppr.unsetKeepNext();

        if (isPageBreakBefore) {
            if (ppr.isSetPageBreakBefore()) ppr.unsetPageBreakBefore();
            ppr.addNewPageBreakBefore().setVal(true);
        } else {
            if (ppr.isSetPageBreakBefore()) ppr.unsetPageBreakBefore();
        }

        CTSpacing spacing = ppr.isSetSpacing() ? ppr.getSpacing() : ppr.addNewSpacing();
        spacing.setBeforeAutospacing(false);
        spacing.setAfterAutospacing(false);
        spacing.setBefore(BigInteger.valueOf(120)); 
        spacing.setAfter(BigInteger.ZERO);          
        
        spacing.setLine(BigInteger.valueOf(360)); 
        spacing.setLineRule(STLineSpacingRule.AUTO);
        
        if (p.getAlignment() != ParagraphAlignment.CENTER) {
            forceZeroIndent(p);
        }
    }

    private static void forceZeroIndent(XWPFParagraph p) {
        CTPPr ppr = p.getCTP().isSetPPr() ? p.getCTP().getPPr() : p.getCTP().addNewPPr();
        CTInd ind = ppr.isSetInd() ? ppr.getInd() : ppr.addNewInd();
        ind.setLeft(BigInteger.ZERO); 
        ind.setFirstLine(BigInteger.ZERO); 
        ind.setRight(BigInteger.ZERO);
        if(ind.isSetHanging()) ind.unsetHanging();
    }
    
    private static void demoteToBodyText(XWPFParagraph p) {
        if (p.getCTP().isSetPPr()) {
             CTPPr ppr = p.getCTP().getPPr();
             if (ppr.isSetOutlineLvl()) {
                ppr.unsetOutlineLvl();
             }
        }
    }
    
    private static void resetFormatting(XWPFParagraph p) {
        if (p.getCTP().isSetPPr()) {
            CTPPr ppr = p.getCTP().getPPr();
            if(ppr.isSetInd()) ppr.unsetInd();
            if(ppr.isSetJc()) ppr.unsetJc();
        }
    }
    
    // --- DETECTION ---
    private static boolean isLevel1Text(String text) {
        String up = text.toUpperCase();
        for (String k : LVL1_TEXT_KEYWORDS) if (up.startsWith(k)) return true;
        return false;
    }

    // [SỬA LẠI]: Thêm tham số boolean hasChapters
    private static int detectLevel(String text, boolean hasChapters) {
        if (text == null || text.trim().isEmpty()) return 0;
        String raw = text.trim();
        if (raw.contains("....")) return 0;

        if (LVL4_REGEX.matcher(raw).matches()) return 4;
        if (LVL3_REGEX.matcher(raw).matches()) return 3;
        if (LVL2_REGEX.matcher(raw).matches()) return 2;
        
        // [SỬA LẠI]
        if (LVL1_NUMBER_REGEX.matcher(raw).matches()) {
            // Chặn các dòng quá dài hoặc có dấu hai chấm (Liệt kê)
            if (raw.length() > 80 || raw.contains(":")) return 0; 
            
            // Hạ cấp xuống Level 2 nếu bài viết dùng hệ thống Chương/Phần
            if (hasChapters) return 2; 
            
            return 1;
        }

        if (isLevel1Text(raw) || isChapterStart(raw)) {
            // [CÚ CHỐT] BỘ LỌC ĐỘ DÀI (LENGTH FILTER)
            // Một Tiêu đề (Heading 1) chuẩn hiếm khi vượt quá 120 ký tự. 
            // Nếu dòng này bắt đầu bằng "PHỤ LỤC" nhưng lại dài ngoằng -> Đó là đoạn văn nội dung! 
            // Ta trả về 0 để ép nó thành Normal Text.
            if (raw.length() > 120) {
                return 0;
            }
            return 1;
        }
        
        return 0;
    }

    // ==========================================
    // TÌM TỌA ĐỘ CHÈN MỤC LỤC
    // ==========================================
    private static int findInsertPosition(XWPFDocument doc, int defaultStart) {
        // List<XWPFParagraph> paras = doc.getParagraphs();
        
        // // 1. Ưu tiên 1: Tìm "LỜI MỞ ĐẦU" hoặc "MỞ ĐẦU"
        // for (int i = defaultStart; i < paras.size(); i++) {
        //     String text = paras.get(i).getText().trim().toUpperCase();
        //     if (text.equals("LỜI MỞ ĐẦU") || text.equals("MỞ ĐẦU") || text.equals("ĐẶT VẤN ĐỀ")) {
        //         System.out.println("   -> Đã tìm thấy '" + text + "' tại index: " + i + ". Sẽ chèn MỤC LỤC lên TRÊN nó.");
        //         return i; // [SỬA LỖI] Trả về chính xác vị trí i để chèn TRƯỚC Lời mở đầu (Bỏ i + 1)
        //     }
        // }

        // // 2. Ưu tiên 2: Nếu không có Lời mở đầu, tìm Chương 1 / Phần 1
        // for (int i = defaultStart; i < paras.size(); i++) {
        //     String text = paras.get(i).getText().trim().toUpperCase();
        //     if (isChapterStart(text) || text.startsWith("CHƯƠNG 1") || text.startsWith("PHẦN 1")) {
        //         System.out.println("   -> Không có Lời mở đầu. Tìm thấy Chương 1 tại index: " + i + ". Sẽ chèn MỤC LỤC lên TRÊN nó.");
        //         return i; // [SỬA LỖI] Trả về chính xác vị trí i để chèn TRƯỚC Chương 1
        //     }
        // }

        return defaultStart;
    }

    // --- HÀM XÓA SẠCH CÁC TOC CŨ (QUÉT SÂU VÀO RUỘT CONTENT CONTROL) ---
    private static void removeOldTOCs(XWPFDocument doc) {
        System.out.println(">>> Bắt đầu quét và xóa TOC cũ (Deep Scan)...");
        // Duyệt ngược từ dưới lên
        for (int i = doc.getBodyElements().size() - 1; i >= 0; i--) {
            IBodyElement elem = doc.getBodyElements().get(i);
            
            boolean shouldDelete = false;

            // 1. NẾU LÀ CONTENT CONTROL (SDT)
            // [MẤU CHỐT]: Phải mở hộp ra kiểm tra từng đoạn văn bên trong
            if (elem instanceof XWPFSDT) {
                XWPFSDT sdt = (XWPFSDT) elem;
                if (isTOCSDT(sdt)) {
                    shouldDelete = true;
                }
            }

            // 2. NẾU LÀ ĐOẠN VĂN (PARAGRAPH)
            else if (elem instanceof XWPFParagraph) {
                XWPFParagraph p = (XWPFParagraph) elem;
                if (isTOCParagraph(p)) {
                    shouldDelete = true;
                }
            }
            
            // 3. NẾU LÀ BẢNG (TABLE) - (Giữ nguyên logic cũ nếu cần)
            else if (elem instanceof XWPFTable) {
                // Logic check bảng đã hoạt động tốt ở các bước trước nên tôi không nhắc lại ở đây
                // Bạn có thể giữ lại hàm isStrictTOCTable nếu muốn
            }

            if (shouldDelete) {
                doc.removeBodyElement(i);
                System.out.println(">>> [DELETED] Đã xóa thành phần TOC tại index " + i);
            }
        }
    }

    // --- HÀM MỚI: KIỂM TRA BÊN TRONG HỘP SDT ---
    private static boolean isTOCSDT(XWPFSDT sdt) {
        // Cách 1: Check Metadata (Vỏ hộp) - Đôi khi file lỗi sẽ bị mất cái này
        String tag = (sdt.getTag() != null) ? sdt.getTag().toUpperCase() : "";
        String title = (sdt.getTitle() != null) ? sdt.getTitle().toUpperCase() : "";
        if (tag.contains("TOC") || title.contains("TABLE OF CONTENTS")) {
            return true;
        }

        // Cách 2: Check Nội dung bên trong (Ruột hộp) -> [QUAN TRỌNG NHẤT]
        // Quét tất cả các element con bên trong hộp
        ISDTContent content = sdt.getContent();
        // Lấy text thô để check nhanh tiêu đề
        String allText = content.getText().toUpperCase();
        if (allText.contains("MỤC LỤC") || allText.contains("CONTENTS")) {
            // Chưa vội return true, check kỹ hơn style để tránh xóa nhầm đoạn văn thường
        }

        // Duyệt qua từng paragraph con bên trong hộp SDT
        // Lưu ý: Content của SDT có thể chứa Table hoặc Paragraph
        // Do thư viện POI hạn chế việc truy cập children của SDT, ta dùng toString() để quét XML của ruột
        // Đây là cách duy nhất hoạt động ổn định trên mọi phiên bản POI
        String internalXML = content.toString(); // Lấy XML của nội dung bên trong
        
        // Tìm dấu hiệu Style TOC bên trong XML của hộp
        // <w:pStyle w:val="TOC1"/> hoặc "TOCHeading"
        if (internalXML.contains("w:val=\"TOC1\"") || 
            internalXML.contains("w:val=\"TOC2\"") ||
            internalXML.contains("w:val=\"TOCHeading\"") ||
            internalXML.contains("w:val=\"TableofFigures\"")) {
            return true;
        }

        return false;
    }

    // --- HÀM KIỂM TRA ĐOẠN VĂN LẺ (Đã cập nhật TableofFigures) ---
    private static boolean isTOCParagraph(XWPFParagraph p) {
        String style = (p.getStyleID() != null) ? p.getStyleID().toUpperCase() : "";
        String text = p.getText().trim().toUpperCase();

        // 1. Check Style (Bao gồm cả TableofFigures bạn vừa phát hiện)
        if (style.startsWith("TOC") || 
            style.equals("TABLEOFFIGURES") || 
            style.equals("TABLEOFAUTHORITIES")) {
            return true;
        }

        // 2. Check Tiêu đề (Check mềm dẻo hơn)
        if (text.equals("MỤC LỤC") || text.equals("CONTENTS") ||
            text.startsWith("DANH MỤC CÁC BẢNG") ||
            text.startsWith("DANH MỤC BẢNG") ||
            text.startsWith("DANH MỤC CÁC HÌNH") ||
            text.startsWith("DANH MỤC HÌNH") ||
            text.startsWith("DANH MỤC CÁC HÌNH")) {
            return true;
        }

        // 3. Check Field Code (PAGEREF)
        for (XWPFRun r : p.getRuns()) {
            for (CTText instr : r.getCTR().getInstrTextList()) {
                String code = instr.getStringValue().trim().toUpperCase();
                if (code.contains("TOC ") || code.contains("PAGEREF _TOC")) {
                    return true;
                }
            }
        }
        return false;
    }


    // ==========================================
    // 2. CHUẨN HÓA CẤU TRÚC TOC (THUẬT TOÁN NEO CHẶN)
    // ==========================================
    public static void standardizeTOCStructure(XWPFDocument doc, int startIdx) {
        System.out.println(">>> Bắt đầu chuẩn hóa cấu trúc TOC theo thuật toán Neo chặn (Anchor)...");

        // 1. DỌN SẠCH TOC CŨ (Quét bằng máy hút bụi Block Deletion đã làm trước đó)
        removeOldTOCs(doc); 

        // 2. TÌM VỊ TRÍ CHÈN TỐI ƯU THEO LOGIC MỚI
        int insertIndex = findOptimalTOCInsertPosition(doc, startIdx);
        System.out.println(">>> Đã chốt vị trí chèn các Danh mục tại đoạn văn thứ: " + insertIndex);

        // 3. LẦN LƯỢT CHÈN CÁC KHỐI (Khối chèn sau sẽ đẩy xuống dưới)
        // Kết quả sẽ ra: Mục lục -> Danh mục Hình -> Danh mục Bảng
        insertIndex = insertTOCBlock(doc, insertIndex, "MỤC LỤC", "TOC \\o \"1-3\" \\h \\z \\u");
        insertIndex = insertTOCBlock(doc, insertIndex, "DANH MỤC HÌNH ẢNH", "TOC \\h \\z \\c \"Hình\"");
        insertIndex = insertTOCBlock(doc, insertIndex, "DANH MỤC BẢNG BIỂU", "TOC \\h \\z \\c \"Bảng\"");

        // 4. XỬ LÝ DANH MỤC TỪ VIẾT TẮT
        int vietTatIdx = findHeadingIndex(doc, "DANH MỤC TỪ VIẾT TẮT", 0);
        if (vietTatIdx != -1) {
            System.out.println("   - Đã có Danh mục từ viết tắt. Bỏ qua không chèn thêm.");
        } else {
            System.out.println("   - Chưa có Danh mục từ viết tắt. Đang chèn khung giữ chỗ...");
            // Chèn từ viết tắt xuống dưới cùng của cụm Mục lục
            insertIndex = insertPlaceholderBlock(doc, insertIndex, "DANH MỤC TỪ VIẾT TẮT");
        }
        
        System.out.println(">>> [DONE] Hoàn tất chuẩn hóa cấu trúc thành công.");
    }

    // ==========================================
    // THUẬT TOÁN TÌM VỊ TRÍ CHÈN MỤC LỤC (Logic Tối Giản)
    // ==========================================
    private static int findOptimalTOCInsertPosition(XWPFDocument doc, int startIdx) {
        List<XWPFParagraph> paras = doc.getParagraphs();

        // 1. Ưu tiên 1: Tìm "LỜI MỞ ĐẦU" hoặc "MỞ ĐẦU"
        for (int i = startIdx; i < paras.size(); i++) {
            String text = paras.get(i).getText().trim().toUpperCase();
            if (text.equals("LỜI MỞ ĐẦU") || text.equals("MỞ ĐẦU") || text.equals("ĐẶT VẤN ĐỀ")) {
                System.out.println("   -> Đã tìm thấy '" + text + "' tại index: " + i + ". Sẽ chèn MỤC LỤC lên TRÊN nó.");
                return i;
            }
        }

        // 2. Ưu tiên 2: Nếu không có Mở đầu, tìm Chương 1 / Phần 1
        for (int i = startIdx; i < paras.size(); i++) {
            String text = paras.get(i).getText().trim().toUpperCase();
            // Sử dụng hàm isChapterStart đã có sẵn của bạn hoặc check tay
            if (isChapterStart(text) || text.startsWith("CHƯƠNG 1") || text.startsWith("PHẦN 1")) {
                System.out.println("   -> Không có Lời mở đầu. Tìm thấy Chương 1 tại index: " + i + ". Sẽ chèn MỤC LỤC lên TRÊN nó.");
                return i;
            }
        }

        // 3. Dự phòng: Nếu không tìm thấy gì, chèn ngay sau trang bìa
        System.out.println("   -> Không tìm thấy Mốc. Chèn Mục lục mặc định tại index: " + startIdx);
        return startIdx;
    }

    /**
     * Hàm tìm vị trí của một dòng Heading dựa trên từ khóa (tương đối)
     */
    private static int findHeadingIndex(XWPFDocument doc, String keyword, int startIndex) {
        List<XWPFParagraph> paras = doc.getParagraphs();
        if (startIndex >= paras.size()) return -1;

        for (int i = startIndex; i < paras.size(); i++) {
            String text = paras.get(i).getText().trim().toUpperCase();
            if (text.equals(keyword) || text.startsWith(keyword + " ") || text.startsWith(keyword + ":") || text.endsWith(keyword)) {
                return i;
            }
        }
        return -1;
    }

    /**
     * LOGIC QUÉT LEVEL 1 ĐỂ TÌM ĐIỂM CUỐI
     * Quét từ dòng startHeadingIndex + 1 xuống dưới.
     * Dừng lại khi gặp bất kỳ dấu hiệu nào của một "Heading lớn" (Chương, Phần, Mục lục khác...).
     * Trả về index của dòng đó (để ta chèn nội dung vào TRƯỚC nó).
     */
    private static int findSectionEndIndex(XWPFDocument doc, int startHeadingIndex) {
        List<XWPFParagraph> paras = doc.getParagraphs();
        
        for (int i = startHeadingIndex + 1; i < paras.size(); i++) {
            XWPFParagraph p = paras.get(i);
            String text = p.getText().trim();
            String upperText = text.toUpperCase();
            String style = (p.getStyleID() != null) ? p.getStyleID() : "";

            // Bỏ qua dòng trống
            if (text.isEmpty()) continue;

            // 1. DẤU HIỆU HEADING QUA STYLE
            if (style.startsWith("Heading 1") || style.startsWith("Heading1")) {
                return i; 
            }

            // 2. DẤU HIỆU HEADING QUA TỪ KHÓA (Nếu style chưa chuẩn)
            if (upperText.startsWith("CHƯƠNG ") || 
                upperText.startsWith("PHẦN ") || 
                upperText.equals("LỜI MỞ ĐẦU") ||
                upperText.equals("MỞ ĐẦU") ||
                upperText.startsWith("DANH MỤC ") || // Gặp danh mục khác cũng dừng
                upperText.equals("MỤC LỤC")) {
                
                // Kiểm tra thêm điều kiện in đậm để chắc chắn đó là tiêu đề
                if (isBold(p)) return i;
            }
            
            // 3. DẤU HIỆU SỐ THỨ TỰ (VD: "1. ", "I. ")
            // Regex: Bắt đầu bằng số hoặc I/V/X, theo sau là dấu chấm và khoảng trắng
            if (text.matches("^(\\d+|[IVX]+)\\.\\s+.*") && isBold(p)) {
                return i;
            }
        }
        
        // Nếu không tìm thấy gì (hết file), chèn vào cuối cùng
        return paras.size();
    }

    // Helper check in đậm
    private static boolean isBold(XWPFParagraph p) {
        for (XWPFRun r : p.getRuns()) {
            if (r.isBold()) return true;
        }
        return false; // Mặc định
    }

    /**     
     * HÀM CHÈN KHỐI TOC (Đã sửa lỗi STXmlSpace và setValue)
     */
    private static int insertTOCBlock(XWPFDocument doc, int index, String title, String fieldCode) {
        // Đảm bảo index hợp lệ
        if (index > doc.getParagraphs().size()) index = doc.getParagraphs().size();
        
        XmlCursor cursor = (index < doc.getParagraphs().size()) 
                ? doc.getParagraphs().get(index).getCTP().newCursor() 
                : doc.getDocument().getBody().newCursor(); // Fallback nếu ở cuối file
        
        // 1. Chèn Tiêu đề
        XWPFParagraph pTitle = doc.insertNewParagraph(cursor);
        pTitle.setStyle("Heading1");
        pTitle.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun rTitle = pTitle.createRun();
        rTitle.setText(title);
        rTitle.setBold(true);
        rTitle.setFontSize(14);
        rTitle.setFontFamily("Times New Roman");
        rTitle.setColor("000000");
        rTitle.addBreak(BreakType.TEXT_WRAPPING);

        // 2. Chèn TOC Field
        XWPFParagraph pTOC = doc.insertNewParagraph(cursor);
        pTOC.setStyle("TOC1"); 
        
        // BEGIN
        XWPFRun r1 = pTOC.createRun();
        CTFldChar fldChar = r1.getCTR().addNewFldChar();
        fldChar.setFldCharType(STFldCharType.BEGIN);
        fldChar.setDirty(true); // Ép Word cập nhật lại số trang

        // INSTR TEXT
        XWPFRun r2 = pTOC.createRun();
        CTText ctText = r2.getCTR().addNewInstrText();
        // [ĐÃ SỬA] Bỏ setSpace(STXmlSpace)
        // [ĐÃ SỬA] Dùng setStringValue thay vì setValue
        ctText.setStringValue(" " + fieldCode + " "); 

        // SEPARATE
        XWPFRun r3 = pTOC.createRun();
        r3.getCTR().addNewFldChar().setFldCharType(STFldCharType.SEPARATE);
        
        // END
        XWPFRun r4 = pTOC.createRun();
        r4.getCTR().addNewFldChar().setFldCharType(STFldCharType.END);

        // 3. Chèn Ngắt trang
        XWPFParagraph pBreak = doc.insertNewParagraph(cursor);
        pBreak.setPageBreak(true);

        return index + 3; // Tăng index lên 3 đơn vị (Title, TOC, Break)
    }

    /**
     * HÀM CHÈN KHỐI GIỮ CHỖ (TỪ VIẾT TẮT)
     */
    private static int insertPlaceholderBlock(XWPFDocument doc, int index, String title) {
        if (index > doc.getParagraphs().size()) index = doc.getParagraphs().size();
        XmlCursor cursor = (index < doc.getParagraphs().size()) 
                ? doc.getParagraphs().get(index).getCTP().newCursor() 
                : doc.getDocument().getBody().newCursor();

        XWPFParagraph pTitle = doc.insertNewParagraph(cursor);
        pTitle.setStyle("Heading1");
        pTitle.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun rTitle = pTitle.createRun();
        rTitle.setText(title);
        rTitle.setBold(true);
        rTitle.setFontSize(14);
        rTitle.setFontFamily("Times New Roman");
        rTitle.setColor("000000");

        XWPFParagraph pContent = doc.insertNewParagraph(cursor);
        XWPFRun rContent = pContent.createRun();
        rContent.setText("(Chèn danh mục từ viết tắt tại đây)");
        rContent.setItalic(true);
        rContent.setFontSize(13);
        rContent.setFontFamily("Times New Roman");

        XWPFParagraph pBreak = doc.insertNewParagraph(cursor);
        pBreak.setPageBreak(true);

        return index + 3;
    }

    
    // --- CÁC HÀM PHỤ TRỢ MỚI ---

    // Kiểm tra xem dòng này có phải là bắt đầu Chương 1/Phần 1 không
    private static boolean isChapterStart(String text) {
        if (text == null) return false;
        return CHAPTER_START_REGEX.matcher(text.toUpperCase()).matches();
    }

   
    

    // --- HÀM KIỂM TRA BULLET / LIST ---
    private static boolean isBulletOrList(XWPFParagraph p) {
        // 1. Check list tự động của Word (XML numPr)
        if (p.getCTP().getPPr() != null && p.getCTP().getPPr().isSetNumPr()) {
            return true;
        }
        // 2. Check gõ tay thủ công
        String text = p.getText().trim();
        if (text.startsWith("- ") || text.startsWith("• ") || text.startsWith("+ ")) {
            return true;
        }
        return false;
    }
    
    
}