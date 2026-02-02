package vn.vietdoc.vietdoc_assistant;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
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
    // 1. CONSTANTS FOR FORMATTING (REFORMAT)
    // ==========================================
    private static final long EMU_PER_CM = 360000L;
    private static final long TWIPS_PER_CM = 567L;
    private static final double SAFE_WIDTH_CM = 15.5;
    private static final long MAX_WIDTH_EMU = (long) (SAFE_WIDTH_CM * EMU_PER_CM);
    private static final BigInteger TABLE_WIDTH_TWIPS = BigInteger.valueOf((long) (SAFE_WIDTH_CM * TWIPS_PER_CM));

    // ==========================================
    // 2. CONSTANTS FOR NLP & CAPTIONING
    // ==========================================
    private static VnCoreNLP pipeline;
    private static final int MAX_CAPTION_WORDS = 30;

    // Regex nhận diện caption đã tồn tại
    private static final Pattern EXISTING_CAPTION_PATTERN = Pattern.compile(
            "^\\s*(Hình|Bảng|Sơ đồ|Biểu đồ|Mô hình|Danh sách|Figure|Table|Chart|Image)\\s*\\d+.*",
            Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE
    );

    // Stopwords & Keywords
    private static final Set<String> LEAD_WORDS = new HashSet<>(Arrays.asList(
            "sau", "duoi", "day", "nhu", "hinh", "bang", "bieu do", "ket qua", "ben", "phia", "minh hoa", "so do"
    ));
    private static final Set<String> STOP_WORDS = new HashSet<>(Arrays.asList(
            ".", ",", ";", "!", "?", "(", ")", "[", "]", "\"", "'", ":", "-", "/", "–"
    ));
    private static final String[] HARD_BREAKERS_NO_ACCENT = {
            "la", "thi", "ma", "boi", "vi", "do", "tai", "trong", "cua",
            "cho", "thay", "duoc", "bi", "gom", "co", "thuoc", "nam",
            "thong qua", "dua tren", "bang cach", "su dung", "ap dung",
            "mo ta", "the hien", "bieu dien", "bao gom", "chi tiet"
    };
    private static final Set<String> ALLOWED_POS = new HashSet<>(Arrays.asList(
            "N", "Np", "Ny", "Nc", "Nu", "M", "A", "V", "E", "L", "P", "X"
    ));

    // Khởi tạo VnCoreNLP
    static {
        try {
            String[] annotators = {"wseg", "pos"};
            pipeline = new VnCoreNLP(annotators);
            System.out.println(">>> [TableImageProcessor] VnCoreNLP Loaded!");
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println(">>> [TableImageProcessor] Lỗi khởi tạo VnCoreNLP: " + e.getMessage());
        }
    }

    // ==========================================
    // 3. MAIN PROCESSING METHOD
    // ==========================================

    /**
     * Hàm xử lý chính: Duyệt qua document để làm đẹp Bảng/Ảnh và tự điền tên (Caption)
     *
     * @param doc        Tài liệu DOCX đang xử lý
     * @param startIndex Vị trí index bắt đầu duyệt (để bỏ qua phần Mục lục hoặc bìa nếu cần)
     */
    public static void process(XWPFDocument doc, int startIndex) {
        System.out.println(">>> Bắt đầu xử lý Bảng và Ảnh từ index: " + startIndex);
        
        // Setup Page Margins cho toàn bộ document trước
        setupPageMargins(doc);

        List<IBodyElement> elements = doc.getBodyElements();

        for (int i = 0; i < elements.size(); i++) {
            // Bỏ qua các phần tử nằm trước startIndex
            if (i < startIndex) continue;

            IBodyElement element = elements.get(i);

            // --- XỬ LÝ BẢNG (TABLE) ---
            if (element instanceof XWPFTable) {
                XWPFTable table = (XWPFTable) element;

                // A. Reformat: Làm đẹp bảng (Kẻ khung, width, spacing...)
                fixTableTotally(table);

                // B. Caption: Xử lý tên bảng (Dựa vào đoạn văn ngay phía trên)
                if (i > 0 && elements.get(i - 1) instanceof XWPFParagraph) {
                    XWPFParagraph paraBefore = (XWPFParagraph) elements.get(i - 1);
                    handleTableCaption(paraBefore);
                }
            }

            // --- XỬ LÝ ĐOẠN VĂN (PARAGRAPH) CHỨA ẢNH ---
            else if (element instanceof XWPFParagraph) {
                XWPFParagraph paraCurrent = (XWPFParagraph) element;

                // A. Reformat: Chuẩn hóa font và Resize ảnh (nếu có)
                // Lưu ý: Hàm này sẽ check nếu có ảnh thì resize, không thì chỉ chỉnh font
                standardizeParagraph(paraCurrent);

                // B. Caption: Xử lý tên ảnh (Nếu đoạn văn này thực sự chứa ảnh)
                if (hasImage(paraCurrent)) {
                    // Cần check đoạn văn ngay phía sau để chèn tên
                    if (i < elements.size() - 1 && elements.get(i + 1) instanceof XWPFParagraph) {
                        XWPFParagraph paraAfter = (XWPFParagraph) elements.get(i + 1);
                        
                        // Lấy text đoạn trước đó để hỗ trợ dự đoán (Context)
                        String textBefore = (i > 0 && elements.get(i - 1) instanceof XWPFParagraph)
                                ? ((XWPFParagraph) elements.get(i - 1)).getText() : "";
                        
                        handleImageCaption(textBefore, paraCurrent, paraAfter);
                    }
                }
            }
        }
        System.out.println(">>> Hoàn tất xử lý Bảng và Ảnh.");
    }

    // ==========================================
    // 4. LOGIC REFORMAT (LÀM ĐẸP)
    // ==========================================

    private static void fixTableTotally(XWPFTable table) {
        // --- 1. TABLE LEVEL SETUP ---
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        if (tblPr == null) tblPr = table.getCTTbl().addNewTblPr();

        // Ép Spacing = 0
        CTTblWidth spacing = tblPr.isSetTblCellSpacing() ? tblPr.getTblCellSpacing() : tblPr.addNewTblCellSpacing();
        spacing.setType(STTblWidth.DXA);
        spacing.setW(BigInteger.ZERO);

        // Ép Width bảng = 15.5cm
        CTTblWidth width = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
        width.setType(STTblWidth.DXA);
        width.setW(TABLE_WIDTH_TWIPS);

        // Kẻ khung
        CTTblBorders borders = tblPr.isSetTblBorders() ? tblPr.getTblBorders() : tblPr.addNewTblBorders();
        createBorder(borders.isSetBottom() ? borders.getBottom() : borders.addNewBottom());
        createBorder(borders.isSetTop() ? borders.getTop() : borders.addNewTop());
        createBorder(borders.isSetLeft() ? borders.getLeft() : borders.addNewLeft());
        createBorder(borders.isSetRight() ? borders.getRight() : borders.addNewRight());
        createBorder(borders.isSetInsideH() ? borders.getInsideH() : borders.addNewInsideH());
        createBorder(borders.isSetInsideV() ? borders.getInsideV() : borders.addNewInsideV());

        // --- TÍNH TOÁN COLUMN WIDTH ---
        int colCount = 0;
        if (!table.getRows().isEmpty()) {
            colCount = table.getRow(0).getTableCells().size();
        }
        BigInteger colWidth = (colCount > 0)
                ? TABLE_WIDTH_TWIPS.divide(BigInteger.valueOf(colCount))
                : BigInteger.ZERO;

        // --- 2. ROW & CELL LEVEL ---
        List<XWPFTableRow> rows = table.getRows();
        for (int i = 0; i < rows.size(); i++) {
            XWPFTableRow row = rows.get(i);
            CTRow ctRow = row.getCtRow();

            // Fix lỗi Spacing dòng
            if (ctRow.isSetTrPr()) {
                CTTrPr trPr = ctRow.getTrPr();
                XmlCursor cursor = trPr.newCursor();
                if (cursor.toChild(new QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "tblCellSpacing"))) {
                    cursor.removeXml();
                }
                cursor.dispose();
            }
            if (ctRow.isSetTblPrEx()) {
                CTTblPrEx tblPrEx = ctRow.getTblPrEx();
                if (tblPrEx.isSetTblCellSpacing()) tblPrEx.unsetTblCellSpacing();
                if (tblPrEx.isSetTblBorders()) tblPrEx.unsetTblBorders();
            }

            // Cell styling
            for (XWPFTableCell cell : row.getTableCells()) {
                // Ép chiều rộng
                if (colCount > 0) {
                    CTTcPr tcPr = cell.getCTTc().getTcPr();
                    if (tcPr == null) tcPr = cell.getCTTc().addNewTcPr();
                    CTTblWidth cellW = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
                    cellW.setType(STTblWidth.DXA);
                    cellW.setW(colWidth);
                }
                cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                // Paragraph bên trong cell
                for (XWPFParagraph p : cell.getParagraphs()) {
                    standardizeParagraph(p); // Font chuẩn
                    if (i == 0) { // Header
                        for (XWPFRun r : p.getRuns()) r.setBold(true);
                        p.setAlignment(ParagraphAlignment.CENTER);
                    } else {
                        p.setAlignment(ParagraphAlignment.BOTH);
                    }
                }
                // Màu nền Header
                if (i == 0) {
                    cell.setColor("E7E6E6");
                }
            }
        }
    }

    private static void standardizeParagraph(XWPFParagraph p) {
        p.setSpacingBetween(1.5);
        boolean foundImage = false;

        for (XWPFRun r : p.getRuns()) {
            r.setFontFamily("Times New Roman");
            r.setFontSize(14); // Có thể parametrise nếu muốn linh động

            // Tìm và resize ảnh
            List<org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing> drawings = r.getCTR().getDrawingList();
            if (drawings != null && !drawings.isEmpty()) {
                foundImage = true;
                for (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing drawing : drawings) {
                    for (CTInline inline : drawing.getInlineList()) {
                        resizeInlineImage(inline);
                    }
                }
            }
        }

        // Nếu có ảnh thì căn giữa đoạn văn
        if (foundImage) {
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
        CTSectPr sectPr = doc.getDocument().getBody().getSectPr();
        if (sectPr == null) sectPr = doc.getDocument().getBody().addNewSectPr();
        CTPageMar pageMar = sectPr.getPgMar();
        if (pageMar == null) pageMar = sectPr.addNewPgMar();

        pageMar.setTop(BigInteger.valueOf(1417));
        pageMar.setBottom(BigInteger.valueOf(1134));
        pageMar.setLeft(BigInteger.valueOf(1984));
        pageMar.setRight(BigInteger.valueOf(1134));
    }

    // ==========================================
    // 5. LOGIC AUTO-CAPTION (ĐIỀN TÊN)
    // ==========================================

    private static void handleTableCaption(XWPFParagraph paraBefore) {
        String text = paraBefore.getText();
        
        // 1. Check: Đã có tên chưa?
        if (isExistingCaption(text)) {
            System.out.println("[SKIP TABLE] Đã có tên: " + text);
            return;
        }

        // 2. Dự đoán tên
        String predicted = predictTableCaption(text);

        // 3. Điền vào
        if (text.trim().isEmpty()) {
            // Nếu dòng trống -> điền trực tiếp
            if (paraBefore.getRuns().isEmpty()) paraBefore.createRun();
            paraBefore.getRuns().get(0).setText("Bảng [AUTO]: " + predicted);
            paraBefore.getRuns().get(0).setBold(true);
            paraBefore.getRuns().get(0).setFontFamily("Times New Roman");
            paraBefore.getRuns().get(0).setFontSize(13); // Caption thường nhỏ hơn 1 chút hoặc nghiêng
            paraBefore.setAlignment(ParagraphAlignment.CENTER);
        } else {
            // Nếu có text dẫn dắt -> append dòng mới
            XWPFRun run = paraBefore.createRun();
            run.addBreak();
            run.setText("Bảng [AUTO]: " + predicted);
            run.setBold(true);
            run.setColor("0000FF"); // Đánh dấu màu xanh để người dùng biết là auto
        }
        System.out.println("[FILL TABLE] " + predicted);
    }

    private static void handleImageCaption(String textBefore, XWPFParagraph imagePara, XWPFParagraph paraAfter) {
        String textAfter = paraAfter.getText();

        // 1. Check: Đã có tên chưa?
        if (isExistingCaption(textAfter)) {
            System.out.println("[SKIP IMAGE] Đã có tên: " + textAfter);
            return;
        }

        // 2. Dự đoán tên (kết hợp text trước ảnh và nội dung trong chính paragraph ảnh)
        String combinedContext = textBefore + " " + imagePara.getText();
        String predicted = predictImageCaption(combinedContext, textAfter);

        // 3. Điền vào
        if (textAfter.trim().isEmpty()) {
            if (paraAfter.getRuns().isEmpty()) paraAfter.createRun();
            paraAfter.getRuns().get(0).setText("Hình [AUTO]: " + predicted);
            paraAfter.getRuns().get(0).setBold(true);
            paraAfter.getRuns().get(0).setFontFamily("Times New Roman");
            paraAfter.getRuns().get(0).setFontSize(13);
            paraAfter.setAlignment(ParagraphAlignment.CENTER);
        } else {
            XWPFRun run = paraAfter.createRun();
            run.addBreak();
            run.setText("Hình [AUTO]: " + predicted);
            run.setBold(true);
            run.setColor("FF0000"); // Đỏ để phân biệt
        }
        System.out.println("[FILL IMAGE] " + predicted);
    }

    // --- NLP UTILITIES ---

    private static String predictTableCaption(String textBefore) {
        try {
            if (pipeline == null) return "Bảng số liệu";
            String cleanBefore = textBefore.trim();
            // Regex Bảng
            String tableRegex = "^\\s*(Bảng|Table)\\s*[\\d.]+\\s*[:.–-]?\\s*([^\\n]+)";
            Matcher mBefore = Pattern.compile(tableRegex, Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE).matcher(getLastLine(cleanBefore));
            if (mBefore.find()) return mBefore.group(0).trim();
            // NLP Backward
            String caption = extractBackward(cleanBefore);
            if (isCaptionGood(caption)) return caption;
            return !caption.isEmpty() ? caption : "Bảng số liệu";
        } catch (Exception e) { return "Bảng số liệu"; }
    }

    private static String predictImageCaption(String textBefore, String textAfter) {
        try {
            if (pipeline == null) return "Ảnh minh họa";
            String cleanBefore = textBefore.trim();
            String cleanAfter = textAfter.trim();
            // Regex Ảnh
            String regex = "^\\s*(Hình|Sơ đồ|Biểu đồ|Bảng|Figure|Chart)\\s*\\d*[:.]?\\s*([^\\n\\.]+)";
            Matcher m = Pattern.compile(regex, Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE).matcher(cleanAfter);
            if (m.find()) return m.group(0).trim();
            // Backward
            String captionBefore = extractBackward(cleanBefore);
            if (isCaptionGood(captionBefore)) return captionBefore;
            // Forward
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

    // --- HELPER FUNCTIONS ---

    private static boolean isExistingCaption(String text) {
        if (text == null || text.trim().isEmpty()) return false;
        return EXISTING_CAPTION_PATTERN.matcher(text.trim()).matches();
    }

    private static boolean hasImage(XWPFParagraph paragraph) {
        for (XWPFRun run : paragraph.getRuns()) {
            if (!run.getEmbeddedPictures().isEmpty()) return true;
        }
        return false;
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
}