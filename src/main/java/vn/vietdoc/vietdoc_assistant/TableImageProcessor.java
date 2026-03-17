package vn.vietdoc.vietdoc_assistant;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;

import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.apache.xmlbeans.XmlCursor;
import vn.pipeline.Annotation;
import vn.pipeline.VnCoreNLP;
import vn.pipeline.Word;

import javax.xml.namespace.QName;
import java.io.IOException;
import java.math.BigInteger;
import java.text.Normalizer;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class TableImageProcessor {

    // ==========================================
    // 1. CONSTANTS
    // ==========================================
    private static final long EMU_PER_CM = 360000L;
    private static final long TWIPS_PER_CM = 567L;
    private static final double SAFE_WIDTH_CM = 15.5;
    private static final long MAX_WIDTH_EMU = (long) (SAFE_WIDTH_CM * EMU_PER_CM);
    private static final BigInteger TABLE_WIDTH_TWIPS = BigInteger.valueOf((long) (SAFE_WIDTH_CM * TWIPS_PER_CM));

    private static VnCoreNLP pipeline;
    private static final int MAX_CAPTION_WORDS = 30;
    
    private static final Pattern EXISTING_CAPTION_PATTERN = Pattern.compile(
            "^\\s*(Hình|Bảng|Sơ đồ|Biểu đồ|Mô hình|Danh sách|Figure|Table|Chart|Image)\\s*\\d+.*",
            Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE
    );

    // Regex nhận diện chương
    private static final Pattern CHAPTER_NUMBER_REGEX = Pattern.compile("^\\s*(\\d+)\\.\\s+.*"); 
    // [FIX] Thêm \\u00A0 để bắt space đặc biệt, dùng (?i) để không phân biệt hoa thường
    
    private static final Pattern CHAPTER_KEYWORD_REGEX = Pattern.compile(
        "^[\\s\\u00A0]*(?:CHƯƠNG|PHẦN)[\\s\\u00A0.:\\-]*([0-9IVX]+).*", 
        Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE
    );

    private static final String[] LVL1_KEYWORDS = {
            "MỞ ĐẦU", "KẾT LUẬN", "LỜI CẢM ƠN", "LỜI MỞ ĐẦU",
            "DANH MỤC", "TÀI LIỆU THAM KHẢO", "PHỤ LỤC", "TÓM TẮT", "LỜI CAM ĐOAN"
    };

    // NLP Stopwords
    private static final Set<String> LEAD_WORDS = new HashSet<>(Arrays.asList(
            "sau", "duoi", "day", "nhu", "hinh", "bang", "bieu do", "ket qua", "ben", "phia", "minh hoa", "so do"
    ));
    private static final Set<String> STOP_WORDS = new HashSet<>(Arrays.asList(
            ".", ",", ";", "!", "?", "(", ")", "[", "]", "\"", "'", ":", "-", "/", "–"
    ));
    private static final String[] HARD_BREAKERS_NO_ACCENT = {
            "la", "thi", "ma", "boi", "vi", "do", "tai", "trong", "cua", "cho", "thay", "duoc", "bi", "gom", "co", "thuoc", "nam"
    };
    private static final Set<String> ALLOWED_POS = new HashSet<>(Arrays.asList(
            "N", "Np", "Ny", "Nc", "Nu", "M", "A", "V", "E", "L", "P", "X"
    ));

    static {
        try {
            pipeline = new VnCoreNLP(new String[]{"wseg", "pos"});
        } catch (IOException e) { e.printStackTrace(); }
    }

    // ==========================================
    // 2. MAIN PROCESSING
    // ==========================================

    public static void process(XWPFDocument doc, int startIndex, FormattingParameters params) {
        System.out.println(">>> Bắt đầu xử lý Bảng/Ảnh (Bỏ qua bảng 1x1)...");
        setupPageMargins(doc);

        List<IBodyElement> elements = doc.getBodyElements();
        
        int currentChapter = 0; 
        int tableCount = 0;
        int imageCount = 0;

        for (int i = 0; i < elements.size(); i++) {
            if (i < startIndex) continue;

            IBodyElement element = elements.get(i);

            // 1. CHECK HEADING (Để cập nhật số chương)
            if (element instanceof XWPFParagraph) {
                XWPFParagraph p = (XWPFParagraph) element;
                int detectedChapter = detectChapterNumber(p);
                
                if (detectedChapter != -1) {
                    if (detectedChapter != currentChapter) {
                        currentChapter = detectedChapter;
                        tableCount = 0;
                        imageCount = 0;
                        System.out.println(">>> Phát hiện Chương " + currentChapter + ": Reset bộ đếm.");
                    }
                }
            }

            // 2. XỬ LÝ BẢNG
            if (element instanceof XWPFTable) {
                XWPFTable table = (XWPFTable) element;

                // [BƯỚC 1] Dọn dẹp hàng/cột thừa
                cleanupTable(table);

                // [BƯỚC 2] Kiểm tra loại bảng
                boolean isLayoutTable = isIgnoredTable(table); 

                // [BƯỚC 3] Format bảng (Luôn chạy để chuẩn hóa khung ngoài & font)
                // Truyền isLayoutTable vào để quyết định có kẻ lưới bên trong hay không
                fixTableTotally(table, isLayoutTable, params);

                // [BƯỚC 4] Nếu là bảng Layout -> Dừng xử lý tại đây (Không đánh số Caption)
                if (isLayoutTable) {
                    System.out.println(">>> Đã chuẩn hóa bảng Layout (Giữ khung ngoài, bỏ khung trong, không caption).");
                    continue; 
                }
                
                // --- PHẦN DƯỚI NÀY CHỈ CHẠY CHO BẢNG DỮ LIỆU THƯỜNG ---
                
                // Tăng số thứ tự
                tableCount++;
                
                // ... (Logic xử lý caption giữ nguyên) ...
                XWPFParagraph paraBefore = (i > 0 && elements.get(i - 1) instanceof XWPFParagraph) 
                                           ? (XWPFParagraph) elements.get(i - 1) : null;
                XWPFParagraph paraAfter = (i < elements.size() - 1 && elements.get(i + 1) instanceof XWPFParagraph) 
                                          ? (XWPFParagraph) elements.get(i + 1) : null;

                if (paraBefore != null) {
                    handleTableCaption(paraBefore, paraAfter, currentChapter, tableCount, params);
                }
            }

            // 3. XỬ LÝ ẢNH
            else if (element instanceof XWPFParagraph) {
                XWPFParagraph paraCurrent = (XWPFParagraph) element;
                
                // Sử dụng hasImage MỚI
                if (hasImage(paraCurrent)) {
                    imageCount++;
                    standardizeParagraph(paraCurrent, params);
                    XWPFParagraph paraBefore = findCandidateCaptionBefore(elements, i);
                    XWPFParagraph paraAfter = findCandidateCaptionAfter(elements, i);

                    // Xử lý ảnh
                    handleImageCaption(doc, paraCurrent, paraBefore, paraAfter, currentChapter, imageCount, params);
                    // (Ảnh tạo dòng mới ở DƯỚI nên không ảnh hưởng index, không cần i++)
                }
            }
        }
        System.out.println(">>> Hoàn tất xử lý.");
    }

    

    // ==========================================
    // 3. CAPTION HANDLING
    // ==========================================

    // ... Trong Class TableImageProcessor ...

    private static void handleTableCaption(XWPFParagraph paraBefore, XWPFParagraph paraAfter, int chapter, int index, FormattingParameters params) {
        String finalContent = "";
        boolean isCaptionFound = false;

        // 1. KIỂM TRA ĐÚNG VỊ TRÍ (Ở TRÊN - paraBefore)
        if (isExistingCaption(paraBefore.getText())) {
            // Đã có caption đúng chỗ -> Lấy nội dung
            finalContent = extractCaptionContent(paraBefore.getText());
            isCaptionFound = true;
        } 
        // 2. KIỂM TRA SAI VỊ TRÍ (Ở DƯỚI - paraAfter)
        else if (paraAfter != null && isExistingCaption(paraAfter.getText())) {
            // Có caption nhưng nằm sai chỗ (ở dưới bảng) -> Lấy nội dung & Xóa ở dưới đi
            finalContent = extractCaptionContent(paraAfter.getText());
            clearParagraph(paraAfter); // Xóa dòng sai vị trí
            isCaptionFound = true;
            System.out.println(">>> [FIX] Dời tên Bảng từ dưới lên trên.");
        }

        // 3. NẾU CHƯA CÓ GÌ -> DÙNG AI/NLP ĐỂ ĐOÁN
        if (!isCaptionFound || finalContent.isEmpty()) {
            finalContent = predictTableCaption(paraBefore.getText());
        }

        // 4. ÁP DỤNG STYLE & ĐÁNH SỐ LẠI (Ghi đè vào paraBefore)
        paraBefore.setStyle("BangStyle");
        removeListProperties(paraBefore);
        
        paraBefore.setAlignment(ParagraphAlignment.CENTER);
        
        clearParagraph(paraBefore); // Xóa text cũ để ghi mới chuẩn format

        int displayChapter = (chapter == 0) ? 1 : chapter;
        String label = "Bảng " + displayChapter + "." + index + ": ";

        XWPFRun r = paraBefore.createRun();
        r.setText(label + finalContent);
        r.setBold(true); 
        r.setFontFamily("Times New Roman"); 
        r.setFontSize(params.getFontSizeBody());
        
        System.out.println("[TABLE] " + label + finalContent);
    }
    
    // Đổi kiểu trả về thành boolean (true = đã chèn thêm dòng mới)
    private static boolean handleImageCaption(XWPFDocument doc, XWPFParagraph paraCurrent, XWPFParagraph paraBefore, XWPFParagraph paraAfter, int chapter, int index, FormattingParameters params) {
        String finalContent = "";
        boolean isCaptionFound = false;
        XWPFParagraph targetPara = null;

        if (paraAfter != null && isExistingCaption(paraAfter.getText())) {
            finalContent = extractCaptionContent(paraAfter.getText());
            targetPara = paraAfter;
            isCaptionFound = true;
        } else if (paraBefore != null && isExistingCaption(paraBefore.getText())) {
            finalContent = extractCaptionContent(paraBefore.getText());
            clearParagraph(paraBefore);
            isCaptionFound = true;
        }

        if (!isCaptionFound || finalContent.isEmpty()) {
            finalContent = predictImageCaption("", ""); // Giữ logic cũ của bạn ở đây
        }

        boolean isInserted = false;
        if (targetPara == null) {
            // TẠO DÒNG MỚI ĐỂ BẢO VỆ VĂN BẢN CỦA USER
            org.apache.xmlbeans.XmlCursor cursor = paraCurrent.getCTP().newCursor();
            cursor.toNextSibling(); 
            targetPara = doc.insertNewParagraph(cursor);
            isInserted = true;
        } else {
            clearParagraph(targetPara);
        }

        targetPara.setStyle("HinhStyle");
        removeListProperties(targetPara);
        targetPara.setAlignment(ParagraphAlignment.CENTER);

        int displayChapter = (chapter == 0) ? 1 : chapter;
        String label = "Hình " + displayChapter + "." + index + ": ";

        XWPFRun r = targetPara.createRun();
        r.setText(label + finalContent);
        r.setItalic(params.isItalicCaption());
        r.setFontFamily("Times New Roman"); 
        r.setFontSize(params.getFontSizeBody());

        System.out.println("[IMAGE] " + label + finalContent);
        return isInserted;
    }

    // ==========================================
    // 4. HELPER: CHAPTER DETECTION
    // ==========================================

    private static int detectChapterNumber(String text) {
        if (text == null || text.trim().isEmpty()) return -1;
        String raw = text.trim();
        Matcher m1 = CHAPTER_NUMBER_REGEX.matcher(raw);
        if (m1.matches()) {
            try { return Integer.parseInt(m1.group(1)); } catch (NumberFormatException e) { return -1; }
        }
        Matcher m2 = CHAPTER_KEYWORD_REGEX.matcher(raw);
        if (m2.matches()) {
            try { return Integer.parseInt(m2.group(1)); } catch (NumberFormatException e) { return -1; }
        }
        if (isLevel1Keyword(raw)) return -1; 
        return -1;
    }

    private static boolean isLevel1Keyword(String text) {
        String up = text.toUpperCase();
        for (String k : LVL1_KEYWORDS) {
            if (up.startsWith(k)) return true;
        }
        return false;
    }

    // ==========================================
    // 5. REFORMAT LOGIC
    // ==========================================

    // [SỬA] Thêm tham số boolean isLayoutTable
    private static void fixTableTotally(XWPFTable table, boolean isLayoutTable, FormattingParameters params) {
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        if (tblPr == null) tblPr = table.getCTTbl().addNewTblPr();

        // 1. Reset Spacing & Width
        CTTblWidth spacing = tblPr.isSetTblCellSpacing() ? tblPr.getTblCellSpacing() : tblPr.addNewTblCellSpacing();
        spacing.setType(STTblWidth.DXA);
        spacing.setW(BigInteger.ZERO);

        CTTblWidth width = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
        width.setType(STTblWidth.DXA);
        width.setW(TABLE_WIDTH_TWIPS);

        // 2. KẺ KHUNG (BORDER)
        CTTblBorders borders = tblPr.isSetTblBorders() ? tblPr.getTblBorders() : tblPr.addNewTblBorders();
        
        // Luôn kẻ khung ngoài (Theo yêu cầu: "không xóa outside border")
        createBorder(borders.isSetBottom() ? borders.getBottom() : borders.addNewBottom());
        createBorder(borders.isSetTop() ? borders.getTop() : borders.addNewTop());
        createBorder(borders.isSetLeft() ? borders.getLeft() : borders.addNewLeft());
        createBorder(borders.isSetRight() ? borders.getRight() : borders.addNewRight());

        // CHỈ KẺ KHUNG TRONG NẾU KHÔNG PHẢI LÀ BẢNG DÀN TRANG (LAYOUT)
        if (!isLayoutTable) {
            createBorder(borders.isSetInsideH() ? borders.getInsideH() : borders.addNewInsideH());
            createBorder(borders.isSetInsideV() ? borders.getInsideV() : borders.addNewInsideV());
        } else {
            // Nếu là bảng layout -> Xóa khung trong (để đảm bảo đúng tính chất bảng 1x1/layout)
            if (borders.isSetInsideH()) borders.unsetInsideH();
            if (borders.isSetInsideV()) borders.unsetInsideV();
        }

        // 3. Căn chỉnh cột (Giữ nguyên logic cũ)
        int colCount = (!table.getRows().isEmpty()) ? table.getRow(0).getTableCells().size() : 0;
        BigInteger colWidth = (colCount > 0) ? TABLE_WIDTH_TWIPS.divide(BigInteger.valueOf(colCount)) : BigInteger.ZERO;

        List<XWPFTableRow> rows = table.getRows();
        for (int i = 0; i < rows.size(); i++) {
            XWPFTableRow row = rows.get(i);

            // [MỚI THÊM] 1. Lặp lại dòng tiêu đề (Header) khi bảng rớt sang trang mới
            if (i == 0) {
                row.setRepeatHeader(true);
            }

            // [MỚI THÊM] 2. Ngăn không cho nội dung trong 1 dòng bị cắt ngang qua 2 trang
            row.setCantSplitRow(true);

            // Xóa spacing thừa của dòng (Code cũ của bạn)
            if (row.getCtRow().isSetTrPr()) {
                CTTrPr trPr = row.getCtRow().getTrPr();
                XmlCursor cursor = trPr.newCursor();
                if (cursor.toChild(new QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "tblCellSpacing"))) cursor.removeXml();
                cursor.dispose();
            }

            for (XWPFTableCell cell : row.getTableCells()) {
                // Set width cột
                if (colCount > 0) {
                    CTTcPr tcPr = cell.getCTTc().getTcPr();
                    if (tcPr == null) tcPr = cell.getCTTc().addNewTcPr();
                    CTTblWidth cellW = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
                    cellW.setType(STTblWidth.DXA);
                    cellW.setW(colWidth);
                }
                cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                // Format nội dung bên trong
                for (XWPFParagraph p : cell.getParagraphs()) {
                    standardizeParagraph(p, params);
                    // Nếu là bảng dữ liệu: Dòng đầu in đậm, canh giữa. Các dòng sau canh trái.
                    if (!isLayoutTable) {
                        if (i == 0) {
                            for (XWPFRun r : p.getRuns()) r.setBold(true);
                            p.setAlignment(ParagraphAlignment.CENTER);
                        } else {
                            p.setAlignment(ParagraphAlignment.LEFT);
                        }
                    } else {
                        // Nếu là bảng Layout: Chỉ chuẩn hóa font, không ép in đậm hay canh lề cứng nhắc
                        // (Hoặc tùy bạn, ở đây tôi để mặc định canh trái cho gọn)
                        p.setAlignment(ParagraphAlignment.LEFT);
                    }
                }
            }
        }
    }

    private static void standardizeParagraph(XWPFParagraph p, FormattingParameters params) {
        p.setSpacingBetween(params.getLineSpacing());
        boolean hasImage = false;

        for (XWPFRun r : p.getRuns()) {
            r.setFontFamily("Times New Roman");
            r.setFontSize(params.getFontSizeBody());
            
            // Logic tìm và xử lý ảnh (Deep copy bằng XmlBeans)
            List<org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing> drawings = r.getCTR().getDrawingList();
            if (drawings != null && !drawings.isEmpty()) {
                hasImage = true;
                for (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing drawing : drawings) {
                    
                    List<CTAnchor> anchorList = drawing.getAnchorList();
                    if (anchorList != null && !anchorList.isEmpty()) {
                        // Duyệt ngược từ cuối lên để xóa an toàn
                        for (int j = anchorList.size() - 1; j >= 0; j--) {
                            CTAnchor anchor = anchorList.get(j);
                            CTInline inline = drawing.addNewInline();
                            
                            // Ép khoảng cách viền = 0
                            inline.setDistT(0L);
                            inline.setDistB(0L);
                            inline.setDistL(0L);
                            inline.setDistR(0L);
                            
                            // Clone toàn bộ node XML an toàn
                            inline.addNewExtent().set(anchor.getExtent());
                            inline.addNewDocPr().set(anchor.getDocPr());
                            inline.addNewGraphic().set(anchor.getGraphic());
                            
                            if (anchor.isSetEffectExtent()) {
                                inline.addNewEffectExtent().set(anchor.getEffectExtent());
                            } else {
                                inline.addNewEffectExtent();
                            }
                            
                            if (anchor.isSetCNvGraphicFramePr()) {
                                inline.addNewCNvGraphicFramePr().set(anchor.getCNvGraphicFramePr());
                            }
                            
                            // Xóa thẻ Anchor cũ
                            drawing.removeAnchor(j);
                        }
                    }

                    // Resize ảnh
                    for (CTInline inline : drawing.getInlineList()) {
                        resizeInlineImage(inline);
                    }
                }
            }
        }

        // Căn giữa ảnh và xóa khoảng thụt lề
        if (hasImage) {
            p.setAlignment(ParagraphAlignment.CENTER);
            if (p.getCTP().getPPr() != null && p.getCTP().getPPr().isSetInd()) {
                p.getCTP().getPPr().unsetInd(); 
            }
        }
    }

    private static void resizeInlineImage(CTInline inline) {
        CTPositiveSize2D extent = inline.getExtent();
        if (extent != null) {
            long currentW = extent.getCx();
            long currentH = extent.getCy();
            if (currentW > MAX_WIDTH_EMU) {
                double scale = (double) MAX_WIDTH_EMU / currentW;
                long newH = (long) (currentH * scale);
                extent.setCx(MAX_WIDTH_EMU);
                extent.setCy(newH);
            }
        }
    }

    private static void createBorder(CTBorder border) {
        border.setVal(STBorder.SINGLE);
        border.setSz(BigInteger.valueOf(4));
        border.setSpace(BigInteger.ZERO);
        border.setColor("000000");
    }

    private static void setupPageMargins(XWPFDocument doc) {
        try {
            CTSectPr sectPr = doc.getDocument().getBody().getSectPr();
            if (sectPr == null) sectPr = doc.getDocument().getBody().addNewSectPr();
            CTPageMar pageMar = sectPr.getPgMar();
            if (pageMar == null) pageMar = sectPr.addNewPgMar();
            pageMar.setTop(BigInteger.valueOf(1417));
            pageMar.setBottom(BigInteger.valueOf(1134));
            pageMar.setLeft(BigInteger.valueOf(1984));
            pageMar.setRight(BigInteger.valueOf(1134));
        } catch (Exception e) {}
    }

    // ==========================================
    // 6. NLP HELPERS
    // ==========================================

    private static boolean isExistingCaption(String text) {
        if (text == null || text.trim().isEmpty()) return false;
        return EXISTING_CAPTION_PATTERN.matcher(text.trim()).matches();
    }

    // --- CẬP NHẬT HÀM KIỂM TRA ẢNH ĐỂ KHÔNG BỎ SÓT ---
    private static boolean hasImage(XWPFParagraph p) {
        if (p == null) return false;
        
        // 1. Cách cũ an toàn cho ảnh chìm
        for (XWPFRun r : p.getRuns()) {
            if (r.getEmbeddedPictures() != null && !r.getEmbeddedPictures().isEmpty()) {
                return true;
            }
        }
        
        // 2. [QUAN TRỌNG] Quét trực tiếp XML gốc của đoạn văn
        // Bắt gọn các loại ảnh nổi (wp:anchor), ảnh VML cũ (v:imagedata), SmartArt, OLE...
        String xml = p.getCTP().toString();
        
        // Chỉ cần chứa một trong các thẻ vẽ đồ họa này là chắc chắn có hình
        if (xml.contains("<w:drawing>") || 
            xml.contains("<wp:anchor") || 
            xml.contains("<wp:inline") || 
            xml.contains("<v:imagedata") || 
            xml.contains("<w:pict>")) {
            return true;
        }

        return false;
    }

    private static String predictTableCaption(String textBefore) {
        try {
            if (pipeline == null) return "Bảng số liệu";
            String cleanBefore = textBefore.trim();
            String tableRegex = "^\\s*(Bảng|Table)\\s*[\\d.]+\\s*[:.–-]?\\s*([^\\n]+)";
            Matcher mBefore = Pattern.compile(tableRegex, Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE).matcher(getLastLine(cleanBefore));
            if (mBefore.find()) return mBefore.group(2).trim(); 
            String caption = extractBackward(cleanBefore);
            return !caption.isEmpty() ? caption : "Bảng số liệu";
        } catch (Exception e) { return "Bảng số liệu"; }
    }

    private static String predictImageCaption(String textBefore, String textAfter) {
        try {
            if (pipeline == null) return "Ảnh minh họa";
            String cleanBefore = textBefore.trim();
            String cleanAfter = textAfter.trim();
            String regex = "^\\s*(Hình|Sơ đồ|Biểu đồ|Bảng|Figure|Chart)\\s*\\d*[:.]?\\s*([^\\n\\.]+)";
            Matcher m = Pattern.compile(regex, Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE).matcher(cleanAfter);
            if (m.find()) return m.group(2).trim();
            String captionBefore = extractBackward(cleanBefore);
            if (isCaptionGood(captionBefore)) return captionBefore;
            if (!cleanAfter.isEmpty()) {
                String captionAfter = extractForward(cleanAfter);
                if (isCaptionGood(captionAfter)) return captionAfter;
            }
            return !captionBefore.isEmpty() ? captionBefore : "Ảnh minh họa";
        } catch (Exception e) { return "Ảnh minh họa"; }
    }

    private static String extractBackward(String text) {
        if (text.isEmpty()) return "";
        try {
            String lastSentence = text.length() > 250 ? text.substring(text.length() - 250) : text;
            Annotation annotation = new Annotation(lastSentence);
            pipeline.annotate(annotation);
            if (annotation.getSentences().isEmpty()) return "";
            List<Word> words = annotation.getSentences().get(annotation.getSentences().size() - 1).getWords();
            List<String> captionWords = new ArrayList<>();
            boolean hasNoun = false, foundWord = false;
            for (int i = words.size() - 1; i >= 0; i--) {
                Word w = words.get(i);
                String form = w.getForm().replace("_", " "), noAccent = removeAccent(form).toLowerCase(), pos = w.getPosTag();
                if (STOP_WORDS.contains(w.getForm()) || w.getForm().equals(":")) { if (!foundWord) continue; else break; }
                if (LEAD_WORDS.contains(noAccent)) continue;
                boolean isBreaker = false;
                for (String b : HARD_BREAKERS_NO_ACCENT) if (noAccent.equals(b) || noAccent.startsWith(b + " ")) { isBreaker = true; break; }
                if (isBreaker) break;
                foundWord = true;
                if (pos.startsWith("N") || pos.equals("X") || pos.equals("M") || pos.equals("Ny") || Character.isUpperCase(form.charAt(0))) hasNoun = true;
                captionWords.add(form);
                if (captionWords.size() >= MAX_CAPTION_WORDS) break;
            }
            if (!hasNoun) return "";
            Collections.reverse(captionWords);
            return String.join(" ", captionWords).trim();
        } catch (Exception e) { return ""; }
    }

    private static String extractForward(String text) {
        if (text.isEmpty()) return "";
        try {
            String firstSentence = text.length() > 250 ? text.substring(0, 250) : text;
            int dotIndex = firstSentence.indexOf('.');
            if (dotIndex != -1) firstSentence = firstSentence.substring(0, dotIndex);
            Annotation annotation = new Annotation(firstSentence);
            pipeline.annotate(annotation);
            if (annotation.getSentences().isEmpty()) return "";
            List<Word> words = annotation.getSentences().get(0).getWords();
            List<String> captionWords = new ArrayList<>();
            boolean foundNoun = false;
            for (Word w : words) {
                String form = w.getForm().replace("_", " "), noAccent = removeAccent(form).toLowerCase(), pos = w.getPosTag();
                boolean isBreaker = false;
                if (STOP_WORDS.contains(form)) isBreaker = true;
                else for (String b : HARD_BREAKERS_NO_ACCENT) if (noAccent.equals(b) || noAccent.startsWith(b + " ")) { isBreaker = true; break; }
                if (isBreaker) { if (foundNoun) break; else continue; }
                if (pos.startsWith("N") || pos.equals("X") || pos.equals("M") || pos.equals("Ny") || Character.isUpperCase(form.charAt(0))) foundNoun = true;
                if (foundNoun) { if (ALLOWED_POS.contains(pos) || Character.isUpperCase(form.charAt(0))) captionWords.add(form); else break; }
                if (captionWords.size() >= MAX_CAPTION_WORDS) break;
            }
            return String.join(" ", captionWords).trim();
        } catch (Exception e) { return ""; }
    }

    private static boolean isCaptionGood(String caption) {
        if (caption == null || caption.isEmpty()) return false;
        while (caption.endsWith(":") || caption.endsWith(".")) caption = caption.substring(0, caption.length() - 1).trim();
        if (caption.split("\\s+").length < 2) return false;
        String noAccent = removeAccent(caption).toLowerCase();
        return !noAccent.endsWith("duoi day") && !noAccent.endsWith("sau day");
    }

    private static String removeAccent(String s) {
        if (s == null) return "";
        String temp = Normalizer.normalize(s, Normalizer.Form.NFD);
        Pattern pattern = Pattern.compile("\\p{InCombiningDiacriticalMarks}+");
        return pattern.matcher(temp).replaceAll("").replace('đ', 'd').replace('Đ', 'D');
    }

    private static String getLastLine(String text) {
        if (text == null) return "";
        String[] lines = text.split("\n");
        for (int i = lines.length - 1; i >= 0; i--) { if (!lines[i].trim().isEmpty()) return lines[i]; }
        return text;
    }

    // ==========================================
    // SỬA: Cập nhật hàm này để nhận XWPFParagraph thay vì String
    // ==========================================
    // ==========================================
    // SỬA: Hàm detectChapterNumber nhận XWPFParagraph và xử lý La Mã
    // ==========================================
    private static int detectChapterNumber(XWPFParagraph p) {
        String text = p.getText();
        if (text == null || text.trim().isEmpty()) return -1;
        
        // 1. Chuẩn hóa chuỗi: Viết hoa + Thay thế các loại khoảng trắng đặc biệt thành dấu cách thường
        String raw = text.toUpperCase().trim();
        raw = raw.replace('\u00A0', ' ').replace((char)160, ' ').replaceAll("\\s+", " ");
        
        String style = (p.getStyleID() != null) ? p.getStyleID() : "";

        // --- CÁCH 1: BẮT TỪ KHÓA "CHƯƠNG" HOẶC "PHẦN" (MẠNH MẼ NHẤT) ---
        // Logic: Tìm vị trí chữ "CHƯƠNG", lấy phần text phía sau nó để phân tích
        int keywordIdx = raw.indexOf("CHƯƠNG");
        if (keywordIdx == -1) keywordIdx = raw.indexOf("PHẦN");

        if (keywordIdx != -1) {
            // Cắt chuỗi từ sau chữ "CHƯƠNG". Ví dụ: "CHƯƠNG II: TỔNG QUAN" -> " II: TỔNG QUAN"
            String suffix = raw.substring(keywordIdx + 6).trim(); // 6 là độ dài chữ CHƯƠNG/PHẦN (tương đối)
            if (raw.contains("PHẦN")) suffix = raw.substring(raw.indexOf("PHẦN") + 4).trim();
            
            // Lấy từ đầu tiên (được tách bởi dấu cách, dấu chấm hoặc dấu hai chấm)
            // Ví dụ: "II: TỔNG QUAN" -> Lấy "II"
            String[] parts = suffix.split("[ .:-]");
            if (parts.length > 0) {
                int val = romanToDecimal(parts[0]);
                if (val != -1) return val;
            }
        }

        // --- CÁCH 2: BẮT SỐ ĐẦU DÒNG (VÍ DỤ "1. ", "2. ") ---
        // Bắt buộc phải là Style Heading 1
        if (style.startsWith("Heading 1") || style.startsWith("Heading1")) {
            // Regex đơn giản chỉ bắt số ở đầu
            Matcher m = Pattern.compile("^(\\d+)\\..*").matcher(raw);
            if (m.matches()) {
                try { return Integer.parseInt(m.group(1)); } catch (Exception e) { return -1; }
            }
        }

        // --- CÁCH 3: CÁC TỪ KHÓA LEVEL 1 KHÁC (MỞ ĐẦU/KẾT LUẬN) ---
        if (isLevel1Keyword(raw)) return 0; // Reset về 0

        return -1;
    }

    // --- HÀM CHUYỂN ĐỔI SỐ LA MÃ (Hỗ trợ I đến XX) ---
    private static int romanToDecimal(String roman) {
        if (roman == null || roman.isEmpty()) return -1;
        // Xóa sạch dấu chấm, hai chấm nếu còn sót
        String r = roman.toUpperCase().replace(".", "").replace(":", "").trim();
        
        // Nếu là số thường (1, 2, 3...)
        if (r.matches("\\d+")) {
            try { return Integer.parseInt(r); } catch (Exception e) { return -1; }
        }

        // Nếu là La Mã
        switch (r) {
            case "I": return 1;
            case "II": return 2;
            case "III": return 3;
            case "IV": return 4;
            case "V": return 5;
            case "VI": return 6;
            case "VII": return 7;
            case "VIII": return 8;
            case "IX": return 9;
            case "X": return 10;
            case "XI": return 11;
            case "XII": return 12;
            case "XIII": return 13;
            case "XIV": return 14;
            case "XV": return 15;
            case "XVI": return 16;
            case "XVII": return 17;
            case "XVIII": return 18;
            case "XIX": return 19;
            case "XX": return 20;
            default: return -1;
        }
    }


    // --- CẬP NHẬT REGEX ĐỂ BÓC TÁCH TRIỆT ĐỂ SỐ THỨ TỰ ---
    // Loại bỏ mọi tiền tố rác kiểu "1.1.", "1.1. 1: ", "1: ", "Bảng 1.", "Hình 1: "
    private static final Pattern CAPTION_CONTENT_REGEX = Pattern.compile(
        "^\\s*(?:(?:Hình|Bảng|Sơ đồ|Biểu đồ|Mô hình|Danh sách|Figure|Table|Chart|Image)\\s*)?([\\dIVX.\\s]+)[:.\\-–]?\\s*(.*)", 
        Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE
    );

    // --- HÀM BÓC TÁCH NỘI DUNG CAPTION CŨ ---
    private static String extractCaptionContent(String text) {
        if (text == null) return "";
        String raw = text.trim();
        Matcher m = CAPTION_CONTENT_REGEX.matcher(raw);
        if (m.matches()) {
            return m.group(2).trim(); // Trích xuất group(2) chứa nội dung thật sự
        }
        
        // Fallback: Nếu người dùng viết "Bảng: Nội dung" (Chỉ có chữ, không gõ số)
        String lower = raw.toLowerCase();
        if (lower.startsWith("bảng") || lower.startsWith("hình") || lower.startsWith("sơ đồ") || lower.startsWith("biểu đồ")) {
             int colonIdx = raw.indexOf(":");
             if (colonIdx == -1) colonIdx = raw.indexOf(".");
             if (colonIdx != -1) return raw.substring(colonIdx + 1).trim();
             // Chặt bỏ từ khóa nếu không có dấu
             return raw.replaceFirst("(?i)^(Hình|Bảng|Sơ đồ|Biểu đồ|Mô hình)\\s*", "").trim();
        }

        return raw; 
    }
    
    // Hàm xóa trắng đoạn văn (để viết lại hoặc xóa caption sai vị trí)
    private static void clearParagraph(XWPFParagraph p) {
        for (int i = p.getRuns().size() - 1; i >= 0; i--) {
            p.removeRun(i);
        }
    }

    // ==========================================
    // 7. CLEANUP HELPERS (XÓA HÀNG/CỘT THỪA)
    // ==========================================

    private static void cleanupTable(XWPFTable table) {
        // 1. Xóa hàng trống trước
        removeEmptyRows(table);

        // 2. Xóa cột trống sau (Logic phức tạp hơn vì Word quản lý theo dòng)
        removeEmptyColumns(table);
    }

    private static void removeEmptyRows(XWPFTable table) {
        List<XWPFTableRow> rows = table.getRows();
        // Duyệt ngược để xóa không bị lệch index
        for (int i = rows.size() - 1; i >= 0; i--) {
            XWPFTableRow row = rows.get(i);
            if (isRowEmpty(row)) {
                table.removeRow(i);
            }
        }
    }

    private static void removeEmptyColumns(XWPFTable table) {
        List<XWPFTableRow> rows = table.getRows();
        if (rows.isEmpty()) return;

        // Tìm số lượng cột tối đa
        int maxCol = 0;
        for (XWPFTableRow row : rows) {
            maxCol = Math.max(maxCol, row.getTableCells().size());
        }

        // Đánh dấu các cột trống
        boolean[] isEmptyCol = new boolean[maxCol];
        Arrays.fill(isEmptyCol, true); // Giả sử ban đầu trống hết

        for (int c = 0; c < maxCol; c++) {
            for (XWPFTableRow row : rows) {
                // Nếu cột c có dữ liệu ở bất kỳ dòng nào -> Không trống
                if (c < row.getTableCells().size() && !isCellEmpty(row.getCell(c))) {
                    isEmptyCol[c] = false;
                    break;
                }
            }
        }

        // Xóa các cột được đánh dấu là trống (Duyệt ngược cột)
        for (int c = maxCol - 1; c >= 0; c--) {
            if (isEmptyCol[c]) {
                for (XWPFTableRow row : rows) {
                    if (c < row.getTableCells().size()) {
                        row.removeCell(c); // Xóa cell tại vị trí c của mọi dòng
                    }
                }
            }
        }
    }

    private static boolean isRowEmpty(XWPFTableRow row) {
        for (XWPFTableCell cell : row.getTableCells()) {
            if (!isCellEmpty(cell)) return false;
        }
        return true;
    }

    private static boolean isCellEmpty(XWPFTableCell cell) {
        if (cell == null) return true;
        // 1. Check text
        String text = cell.getText().trim();
        if (!text.isEmpty()) return false;

        // 2. Check ảnh (Nếu có ảnh thì không coi là trống)
        for (XWPFParagraph p : cell.getParagraphs()) {
            for (XWPFRun r : p.getRuns()) {
                if (!r.getEmbeddedPictures().isEmpty()) return false;
            }
        }
        return true;
    }


    // ==========================================
    // CHECK BẢNG CẦN BỎ QUA (LAYOUT/BORDER/SIGNATURE)
    // ==========================================
    private static boolean isIgnoredTable(XWPFTable table) {
        List<XWPFTableRow> rows = table.getRows();
        
        // 1. Check kích thước: Nếu không có dòng nào hoặc chỉ có 1 ô (1x1)
        if (rows.isEmpty()) return true;
        if (rows.size() == 1 && rows.get(0).getTableCells().size() <= 1) return true;

        // 2. Check Inside Borders (Đường kẻ lưới bên trong)
        // Nếu không có đường kẻ dọc và ngang bên trong -> Coi là bảng dàn trang -> Bỏ qua
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        if (tblPr != null && tblPr.isSetTblBorders()) {
            CTTblBorders borders = tblPr.getTblBorders();
            
            boolean hasInsideH = hasBorder(borders.getInsideH());
            boolean hasInsideV = hasBorder(borders.getInsideV());

            // Nếu cả 2 đường viền trong đều KHÔNG CÓ -> Bỏ qua
            if (!hasInsideH && !hasInsideV) {
                return true; 
            }
        } else {
            // Trường hợp không set TblBorders -> Mặc định Word có thể hiển thị hoặc không.
            // Nhưng thường bảng dữ liệu chuẩn sẽ có border.
            // Tùy chọn: Nếu không có border setting nào -> Bỏ qua luôn cho an toàn?
            // Hoặc return false để giữ lại. Ở đây tôi chọn giữ lại nếu không chắc chắn.
            return false;
        }

        return false;
    }

    // Helper kiểm tra một đường viền có tồn tại và hiển thị không
    private static boolean hasBorder(CTBorder border) {
        if (border == null) return false;
        // Nếu val là NONE hoặc NIL -> Coi như không có
        if (border.getVal() == STBorder.NONE || border.getVal() == STBorder.NIL) return false;
        return true;
    }

    // --- HÀM MỚI: XÓA THUỘC TÍNH DANH SÁCH/BULLET ---
    private static void removeListProperties(XWPFParagraph p) {
        if (p.getCTP().getPPr() == null) return;
        
        // 1. Xóa Numbering (Dấu chấm, gạch đầu dòng, bullet)
        if (p.getCTP().getPPr().isSetNumPr()) {
            p.getCTP().getPPr().unsetNumPr();
        }
        
        // 2. Xóa thụt lề (Indentation) thường đi kèm với danh sách
        if (p.getCTP().getPPr().isSetInd()) {
            p.getCTP().getPPr().unsetInd();
        }
    }


    // --- HÀM TÌM CAPTION Ở DƯỚI AN TOÀN ---
    private static XWPFParagraph findCandidateCaptionAfter(List<IBodyElement> elements, int currentIndex) {
        for (int i = currentIndex + 1; i < elements.size() && i <= currentIndex + 3; i++) { // Chỉ nhìn xa 3 dòng
            IBodyElement elem = elements.get(i);
            if (elem instanceof XWPFParagraph) {
                XWPFParagraph p = (XWPFParagraph) elem;
                String text = p.getText().trim();
                if (text.isEmpty()) continue; // Bỏ qua dòng trống
                
                // Nếu dòng chứa chữ, ta lấy luôn dòng này làm caption
                return p; 
            } else {
                return null; // Gặp bảng -> Dừng
            }
        }
        return null;
    }

    // --- HÀM TÌM CAPTION Ở TRÊN AN TOÀN ---
    private static XWPFParagraph findCandidateCaptionBefore(List<IBodyElement> elements, int currentIndex) {
        for (int i = currentIndex - 1; i >= 0 && i >= currentIndex - 3; i--) { // Chỉ nhìn xa 3 dòng
            IBodyElement elem = elements.get(i);
            if (elem instanceof XWPFParagraph) {
                XWPFParagraph p = (XWPFParagraph) elem;
                String text = p.getText().trim();
                if (text.isEmpty()) continue; // Bỏ qua dòng trống
                
                // Trả về dòng text đầu tiên tìm được
                return p;
            } else {
                return null; // Gặp bảng -> Dừng
            }
        }
        return null;
    }



    
}