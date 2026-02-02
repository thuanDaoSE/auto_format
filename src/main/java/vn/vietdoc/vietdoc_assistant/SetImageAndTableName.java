package vn.vietdoc.vietdoc_assistant;

import org.apache.poi.xwpf.usermodel.*;
import vn.pipeline.Annotation;
import vn.pipeline.VnCoreNLP;
import vn.pipeline.Word;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.Normalizer;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class SetImageAndTableName {

    // --- ĐƯỜNG DẪN FILE ---
    // Bạn dùng file gốc (chưa xóa tên) để chạy cái này nhé
    private static final String INPUT_FILE = "D:\\Workspaces\\vietdoc-assistant\\PTDACNTT_Nhom14 (1).docx";
    private static final String OUTPUT_FILE = "D:\\Workspaces\\vietdoc-assistant\\PTDACNTT_Nhom14_FILLED.docx";

    // --- REGEX NHẬN DIỆN CAPTION CŨ ---
    // Nếu dòng bắt đầu bằng các từ này -> Coi là đã có tên -> Bỏ qua
    private static final Pattern EXISTING_CAPTION_PATTERN = Pattern.compile(
        "^\\s*(Hình|Bảng|Sơ đồ|Biểu đồ|Mô hình|Danh sách|Figure|Table|Chart|Image)\\s*\\d+.*",
        Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE
    );

    // --- CẤU HÌNH NLP ---
    private static VnCoreNLP pipeline;
    private static final int MAX_CAPTION_WORDS = 30;

    // Các biến cấu hình NLP (Lead words, Stop words...) giữ nguyên như cũ
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

    static {
        try {
            String[] annotators = {"wseg", "pos"}; 
            pipeline = new VnCoreNLP(annotators); 
            System.out.println(">>> VnCoreNLP Loaded!");
        } catch (IOException e) { e.printStackTrace(); }
    }

    public static void main(String[] args) {
        System.out.println("=== CHẾ ĐỘ AUTO-FILL: CHỈ ĐIỀN KHI THIẾU ===");
        try {
            processDocx(INPUT_FILE, OUTPUT_FILE);
            System.out.println("=== XONG! ===");
        } catch (Exception e) { e.printStackTrace(); }
    }

    public static void processDocx(String inputPath, String outputPath) throws IOException {
        FileInputStream fis = new FileInputStream(inputPath);
        XWPFDocument document = new XWPFDocument(fis);
        List<IBodyElement> elements = document.getBodyElements();

        for (int i = 0; i < elements.size(); i++) {
            IBodyElement element = elements.get(i);

            // ---------------- XỬ LÝ BẢNG ----------------
            if (element instanceof XWPFTable) {
                // Kiểm tra đoạn văn TRÊN bảng
                if (i > 0 && elements.get(i - 1) instanceof XWPFParagraph) {
                    XWPFParagraph paraBefore = (XWPFParagraph) elements.get(i - 1);
                    String text = paraBefore.getText();

                    // 1. CHECK: Đã có tên chưa?
                    if (isExistingCaption(text)) {
                        System.out.println("[SKIP TABLE] Đã có tên: " + text);
                        continue; // Bỏ qua
                    }

                    // 2. Nếu chưa có -> Dự đoán
                    String predicted = predictTableCaption(text); // Logic cũ: chỉ dùng textBefore
                    
                    // 3. Điền vào (Nếu dòng trên trống thì điền, nếu có text rác thì chèn thêm dòng mới)
                    if (text.trim().isEmpty()) {
                        paraBefore.createRun().setText("Bảng [AUTO]: " + predicted);
                        // Format cho đẹp (Đậm, Canh giữa)
                        paraBefore.setAlignment(ParagraphAlignment.CENTER);
                        paraBefore.getRuns().get(0).setBold(true);
                    } else {
                        // Trường hợp dòng trên là câu dẫn ("Dưới đây là bảng..."), ta tạo dòng mới
                        // (Phần này POI xử lý chèn dòng hơi phức tạp, tạm thời ta append vào dòng đó để an toàn)
                        XWPFRun run = paraBefore.createRun();
                        run.addBreak(); // Xuống dòng
                        run.setText("Bảng [AUTO]: " + predicted);
                        run.setBold(true);
                        run.setColor("0000FF"); // Xanh
                    }
                    System.out.println("[FILL TABLE] " + predicted);
                }
            }

            // ---------------- XỬ LÝ ẢNH ----------------
            else if (element instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) element;
                if (hasImage(paragraph)) {
                    // Kiểm tra đoạn văn SAU ảnh
                    if (i < elements.size() - 1 && elements.get(i + 1) instanceof XWPFParagraph) {
                        XWPFParagraph paraAfter = (XWPFParagraph) elements.get(i + 1);
                        String text = paraAfter.getText();

                        // 1. CHECK: Đã có tên chưa?
                        if (isExistingCaption(text)) {
                            System.out.println("[SKIP IMAGE] Đã có tên: " + text);
                            continue; // Bỏ qua
                        }

                        // 2. Nếu chưa có -> Dự đoán
                        String textBefore = (i > 0 && elements.get(i - 1) instanceof XWPFParagraph) 
                                            ? ((XWPFParagraph) elements.get(i - 1)).getText() : "";
                        if (!paragraph.getText().trim().isEmpty()) textBefore += " " + paragraph.getText();
                        
                        String predicted = predictImageCaption(textBefore, text); // TextAfter lúc này là rác hoặc trống

                        // 3. Điền vào
                        if (text.trim().isEmpty()) {
                            paraAfter.createRun().setText("Hình [AUTO]: " + predicted);
                            paraAfter.setAlignment(ParagraphAlignment.CENTER);
                            paraAfter.getRuns().get(0).setBold(true);
                        } else {
                            // Nếu dòng sau là text thường, chèn thêm
                            XWPFRun run = paraAfter.createRun();
                            run.addBreak();
                            run.setText("Hình [AUTO]: " + predicted);
                            run.setBold(true);
                            run.setColor("FF0000"); // Đỏ
                        }
                        System.out.println("[FILL IMAGE] " + predicted);
                    }
                }
            }
        }

        FileOutputStream fos = new FileOutputStream(outputPath);
        document.write(fos);
        fos.close();
        document.close();
        fis.close();
    }

    // --- HÀM KIỂM TRA CAPTION CÓ SẴN ---
    private static boolean isExistingCaption(String text) {
        if (text == null || text.trim().isEmpty()) return false;
        return EXISTING_CAPTION_PATTERN.matcher(text.trim()).matches();
    }

    // --- LOGIC DỰ ĐOÁN (COPY TỪ BẢN TRƯỚC) ---
    public static String predictTableCaption(String textBefore) {
        // Logic y hệt bản DocxAutoLabeler
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

    public static String predictImageCaption(String textBefore, String textAfter) {
        // Logic y hệt bản DocxAutoLabeler
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

    // --- CÁC HÀM NLP SUPPORT (GIỮ NGUYÊN) ---
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

    private static boolean hasImage(XWPFParagraph paragraph) {
        for (XWPFRun run : paragraph.getRuns()) if (!run.getEmbeddedPictures().isEmpty()) return true;
        return false;
    }

    private static boolean isCaptionGood(String caption) {
        if (caption == null || caption.isEmpty()) return false;
        while (caption.endsWith(":") || caption.endsWith(".")) caption = caption.substring(0, caption.length() - 1).trim();
        if (caption.split("\\s+").length < 2) return false;
        String noAccent = removeAccent(caption).toLowerCase();
        return !noAccent.endsWith("duoi day") && !noAccent.endsWith("sau day");
    }

    public static String removeAccent(String s) {
        if (s == null) return "";
        String temp = Normalizer.normalize(s, Normalizer.Form.NFD);
        Pattern pattern = Pattern.compile("\\p{InCombiningDiacriticalMarks}+");
        return pattern.matcher(temp).replaceAll("").replace('đ','d').replace('Đ','D');
    }
    
    private static String getLastLine(String text) {
        if (text == null) return "";
        String[] lines = text.split("\n");
        for (int i = lines.length - 1; i >= 0; i--) { if (!lines[i].trim().isEmpty()) return lines[i]; }
        return text;
    }
}