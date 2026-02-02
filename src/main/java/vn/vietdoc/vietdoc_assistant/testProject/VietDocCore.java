package vn.vietdoc.vietdoc_assistant.testProject;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.apache.xmlbeans.XmlCursor;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

public class VietDocCore {

    // ================= CẤU HÌNH AI =================
    private static final int CURRENT_PROVIDER = 2; // 1 = DeepSeek, 2 = Google Gemma
    private static final int BATCH_SIZE = 5; // Số lượng ảnh gửi đi trong 1 request (Tối ưu tốc độ)

    // --- DEEPSEEK ---
    private static final String DEEPSEEK_API_KEY = "sk-xxxxxxxxxxxx";
    private static final String DEEPSEEK_URL = "https://api.deepseek.com/chat/completions";
    private static final String DEEPSEEK_MODEL = "deepseek-chat";

    // --- GOOGLE GEMMA ---
    private static final String[] GOOGLE_API_KEYS = {
        "AIzaSyDvM-p9dId2FmbORHwNW4EtFhGg81Q6CHo", // Key mẫu
        "AIzaSyAMtoHdeizjVhBWnWdOe8yFPqGon6RMfcQ"
    };
    private static final String GOOGLE_MODEL = "gemma-3-27b-it";
    private static final String GOOGLE_URL_TEMPLATE = "https://generativelanguage.googleapis.com/v1beta/models/%s:generateContent?key=%s";

    // --- FORMAT WORD ---
    private static final BigInteger MARGIN_LEFT = BigInteger.valueOf((long) (3.5 * 567));
    private static final BigInteger MARGIN_RIGHT = BigInteger.valueOf((long) (2.0 * 567));
    private static final BigInteger MARGIN_TOP = BigInteger.valueOf((long) (2.0 * 567));
    private static final BigInteger MARGIN_BOTTOM = BigInteger.valueOf((long) (2.0 * 567));
    private static final String FONT_FAMILY = "Times New Roman";

    // Class nội bộ để lưu thông tin ảnh chờ xử lý
    private static class ImageRequest {
        XWPFParagraph paragraph; // Vị trí đoạn văn chứa ảnh
        String context;          // Ngữ cảnh văn bản xung quanh
        
        public ImageRequest(XWPFParagraph p, String c) {
            this.paragraph = p;
            this.context = c;
        }
    }

    public static void main(String[] args) {
        String inputPath = "D:\\Workspaces\\vietdoc-assistant\\input_doan.docx";
        String outputPath = "D:\\Workspaces\\vietdoc-assistant\\output_vietdoc_pro.docx";

        System.out.println(">>> VIETDOC CORE: Khởi động...");
        System.out.println(">>> CHẾ ĐỘ: Xử lý theo lô (Batch Processing)");
        
        processDocument(inputPath, outputPath);
    }

    public static void processDocument(String inputPath, String outputPath) {
        try (FileInputStream fis = new FileInputStream(inputPath);
             XWPFDocument document = new XWPFDocument(fis)) {

            // 1. Format & Gom ảnh
            setupPageMargins(document);
            setupPageNumber(document);
            
            // List này sẽ chứa toàn bộ yêu cầu caption cần xử lý
            List<ImageRequest> pendingImages = new ArrayList<>();
            processParagraphsAndCollectImages(document, pendingImages);
            
            // 2. Xử lý AI theo lô (Batch Processing)
            if (!pendingImages.isEmpty()) {
                System.out.println(">>> Tìm thấy " + pendingImages.size() + " ảnh cần đặt tên. Bắt đầu gọi AI...");
                processBatches(pendingImages);
            } else {
                System.out.println(">>> Không tìm thấy ảnh nào cần xử lý.");
            }

            processTables(document);
            generateTOC(document);

            // 3. Xuất file
            try (FileOutputStream out = new FileOutputStream(outputPath)) {
                document.write(out);
                System.out.println("🚀 [Done] File đã xong: " + outputPath);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // ================= 1. GOM ẢNH (KHÔNG GỌI AI NGAY) =================
    
    private static void processParagraphsAndCollectImages(XWPFDocument document, List<ImageRequest> pendingImages) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (int i = 0; i < paragraphs.size(); i++) {
            XWPFParagraph p = paragraphs.get(i);
            
            // Format cơ bản
            p.setSpacingAfter(120);
            for (XWPFRun r : p.getRuns()) {
                r.setFontFamily(FONT_FAMILY);
                r.setFontSize(13);
            }

            // Detect ảnh & Đưa vào hàng đợi
            if (containsImage(p)) {
                // Kiểm tra xem đã có caption chưa (dòng tiếp theo có phải bắt đầu bằng "Hình" không)
                boolean hasCaption = (i + 1 < paragraphs.size()) && 
                                     paragraphs.get(i+1).getText().trim().toLowerCase().startsWith("hình");
                
                if (!hasCaption) {
                    // Lấy ngữ cảnh rộng hơn (trước sau 3 dòng) để AI hiểu rõ hơn
                    String context = getContext(paragraphs, i - 3, i + 3);
                    // Thêm vào hàng đợi xử lý
                    pendingImages.add(new ImageRequest(p, context));
                }
            }
        }
    }

    // ================= 2. XỬ LÝ BATCH (LOGIC MỚI) =================

    private static void processBatches(List<ImageRequest> allRequests) {
        int total = allRequests.size();
        
        for (int i = 0; i < total; i += BATCH_SIZE) {
            int end = Math.min(i + BATCH_SIZE, total);
            List<ImageRequest> batch = allRequests.subList(i, end);
            
            System.out.println("\n--- Đang xử lý Batch từ ảnh " + (i+1) + " đến " + end + " ---");
            
            // 1. Tạo Prompt gộp
            StringBuilder userPromptBuilder = new StringBuilder();
            userPromptBuilder.append("Danh sách các ngữ cảnh ảnh cần đặt tên:\n");
            for (int k = 0; k < batch.size(); k++) {
                // Đánh dấu số thứ tự để AI không bị nhầm
                userPromptBuilder.append("ITEM_").append(k).append(": ").append(batch.get(k).context).append("\n\n");
            }
            
            String systemPrompt = "Bạn là trợ lý soạn thảo văn bản học thuật." +
                    "Nhiệm vụ: Đặt tên (caption) ngắn gọn cho từng ngữ cảnh ảnh được cung cấp.\n" +
                    "Quy tắc phản hồi QUAN TRỌNG:\n" +
                    "1. Chỉ trả về danh sách các caption.\n" +
                    "2. Các caption ngăn cách nhau CHÍNH XÁC bởi chuỗi ký tự '|||'.\n" +
                    "3. Định dạng: Hình [số]: [Tên ngắn gọn].\n" +
                    "4. Không thêm lời dẫn, không thêm số thứ tự ở đầu dòng, chỉ nội dung text ngăn cách bởi |||.";

            String userPrompt = userPromptBuilder.toString();

            // LOG REQUEST (Theo yêu cầu của bạn)
            System.out.println(">>> REQUEST SENT TO AI:\n" + userPrompt);

            // 2. Gọi AI
            String aiResponse = callAI(systemPrompt, userPrompt);
            
            // LOG RESPONSE
            System.out.println(">>> RESPONSE FROM AI:\n" + aiResponse);

            // 3. Parse kết quả và điền vào Word
            if (aiResponse != null && !aiResponse.isEmpty()) {
                // Tách chuỗi dựa trên dấu phân cách |||
                String[] captions = aiResponse.split("\\|\\|\\|");
                
                // Gán ngược lại vào văn bản
                for (int k = 0; k < batch.size(); k++) {
                    if (k < captions.length) {
                        String caption = captions[k].trim();
                        // Xử lý sạch caption nếu AI lỡ thêm xuống dòng
                        caption = caption.replace("\n", "").replace("ITEM_" + k + ":", "").trim();
                        
                        insertCaption(batch.get(k).paragraph, caption);
                    } else {
                        System.err.println("⚠️ Cảnh báo: AI trả về ít caption hơn số lượng ảnh gửi đi ở batch này.");
                    }
                }
            }
        }
    }

    // ================= LOGIC GỌI AI & PARSING JSON =================
    
    private static String callAI(String systemPrompt, String userPrompt) {
        if (CURRENT_PROVIDER == 1) return callDeepSeek(systemPrompt, userPrompt);
        else return callGoogleGemma(systemPrompt, userPrompt);
    }

    private static String callGoogleGemma(String systemPrompt, String userPrompt) {
        try {
            String randomKey = GOOGLE_API_KEYS[new Random().nextInt(GOOGLE_API_KEYS.length)];
            String url = String.format(GOOGLE_URL_TEMPLATE, GOOGLE_MODEL, randomKey);
            String fullPrompt = systemPrompt + "\n\nUser Request: " + userPrompt;
            
            // Escape JSON cẩn thận
            String jsonBody = String.format("{\"contents\": [{\"parts\": [{\"text\": \"%s\"}]}]}", escapeJson(fullPrompt));

            return sendPostRequest(url, null, jsonBody, "candidates");
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private static String callDeepSeek(String systemPrompt, String userPrompt) {
        try {
            String jsonBody = String.format(
                "{\"model\": \"%s\", \"messages\": [{\"role\": \"system\", \"content\": \"%s\"}, {\"role\": \"user\", \"content\": \"%s\"}], \"stream\": false}", 
                DEEPSEEK_MODEL, escapeJson(systemPrompt), escapeJson(userPrompt)
            );
            return sendPostRequest(DEEPSEEK_URL, DEEPSEEK_API_KEY, jsonBody, "choices");
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private static String sendPostRequest(String url, String apiKey, String jsonBody, String responseType) {
        try {
            HttpClient client = HttpClient.newHttpClient();
            HttpRequest.Builder builder = HttpRequest.newBuilder()
                    .uri(URI.create(url))
                    .header("Content-Type", "application/json");
            if (apiKey != null) builder.header("Authorization", "Bearer " + apiKey);

            HttpRequest request = builder.POST(HttpRequest.BodyPublishers.ofString(jsonBody, StandardCharsets.UTF_8)).build();
            HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());

            if (response.statusCode() == 200) {
                if (responseType.equals("choices")) return extractContentDeepSeek(response.body());
                else return extractContentGoogle(response.body());
            } else {
                System.err.println("❌ API Error (" + response.statusCode() + "): " + response.body());
                return null;
            }
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // --- Parser JSON Thủ Công (Giữ nguyên logic của bạn nhưng tinh chỉnh) ---
    private static String extractContentDeepSeek(String json) { return parseJsonValue(json, "\"content\":"); }
    private static String extractContentGoogle(String json) { return parseJsonValue(json, "\"text\":"); }

    private static String parseJsonValue(String json, String key) {
        try {
            int start = json.indexOf(key);
            if (start == -1) return null;
            start = json.indexOf("\"", start + key.length()) + 1;
            StringBuilder sb = new StringBuilder();
            boolean escaped = false;
            for (int i = start; i < json.length(); i++) {
                char c = json.charAt(i);
                if (escaped) {
                    if (c == 'n') sb.append('\n'); else if (c == '"') sb.append('"'); else sb.append(c);
                    escaped = false;
                } else {
                    if (c == '\\') escaped = true;
                    else if (c == '"') break;
                    else sb.append(c);
                }
            }
            return sb.toString();
        } catch (Exception e) { return ""; }
    }
    
    private static String escapeJson(String text) {
        if (text == null) return "";
        return text.replace("\\", "\\\\").replace("\"", "\\\"").replace("\n", "\\n").replace("\r", "");
    }

    // ================= CÁC HÀM HỖ TRỢ POI =================
    private static void insertCaption(XWPFParagraph imgP, String text) {
        XmlCursor cursor = imgP.getCTP().newCursor();
        XWPFParagraph cP = imgP.getBody().insertNewParagraph(cursor);
        cP.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun r = cP.createRun();
        r.setText(text);
        r.setItalic(true);
        r.setFontSize(12);
        System.out.println("✅ Inserted: " + text);
    }
    
    private static boolean containsImage(XWPFParagraph p) {
        for (XWPFRun r : p.getRuns()) if (r.getEmbeddedPictures().size() > 0) return true;
        return false;
    }
    
    private static String getContext(List<XWPFParagraph> list, int start, int end) {
        StringBuilder sb = new StringBuilder();
        for (int i = Math.max(0, start); i < Math.min(list.size(), end); i++) sb.append(list.get(i).getText()).append(" ");
        String s = sb.toString().trim();
        return s.length() > 300 ? s.substring(0, 300) : s; // Tăng context lên 300 ký tự cho AI dễ hiểu
    }
    
    private static void setupPageMargins(XWPFDocument document) { /* Giữ nguyên code cũ */
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        if (sectPr == null) sectPr = document.getDocument().getBody().addNewSectPr();
        CTPageMar pageMar = sectPr.getPgMar();
        if (pageMar == null) pageMar = sectPr.addNewPgMar();
        pageMar.setLeft(MARGIN_LEFT); pageMar.setRight(MARGIN_RIGHT);
        pageMar.setTop(MARGIN_TOP); pageMar.setBottom(MARGIN_BOTTOM);
    }
    private static void setupPageNumber(XWPFDocument document) { /* Giữ nguyên code cũ */
        XWPFHeaderFooterPolicy policy = document.createHeaderFooterPolicy();
        XWPFFooter footer = policy.createFooter(STHdrFtr.DEFAULT);
        XWPFParagraph p = footer.createParagraph();
        p.setAlignment(ParagraphAlignment.CENTER);
        p.getCTP().addNewFldSimple().setInstr("PAGE \\* MERGEFORMAT");
    }
    private static void processTables(XWPFDocument document) { /* Giữ nguyên code cũ */
        for (XWPFTable table : document.getTables()) {
            CTTblWidth tblW = table.getCTTbl().getTblPr().addNewTblW();
            tblW.setType(STTblWidth.DXA);
            tblW.setW(BigInteger.valueOf(9000));
        }
    }
    private static void generateTOC(XWPFDocument document) { /* Giữ nguyên code cũ */
        XWPFParagraph p = document.createParagraph();
        p.getCTP().addNewFldSimple().setInstr("TOC \\o \"1-3\" \\h \\z \\u");
    }
}