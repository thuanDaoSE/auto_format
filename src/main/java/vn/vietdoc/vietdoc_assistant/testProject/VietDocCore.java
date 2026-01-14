package vn.vietdoc.vietdoc_assistant.testProject;

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
import java.util.List;
import java.util.Random;

public class VietDocCore {

    // ================= CẤU HÌNH AI (LINH HOẠT) =================
    
    // 1 = DeepSeek, 2 = Google Gemma (Free Tier)
    private static final int CURRENT_PROVIDER = 2; 

    // --- CẤU HÌNH DEEPSEEK ---
    private static final String DEEPSEEK_API_KEY = "sk-xxxxxxxxxxxx"; // Điền key DeepSeek
    private static final String DEEPSEEK_URL = "https://api.deepseek.com/chat/completions";
    private static final String DEEPSEEK_MODEL = "deepseek-chat";

    // --- CẤU HÌNH GOOGLE GEMMA (Dùng mảng để xoay vòng 5 acc) ---
    private static final String[] GOOGLE_API_KEYS = {
        "AIzaSy-Account1-xxxxxx",
        "AIzaSy-Account2-xxxxxx",
        "AIzaSy-Account3-xxxxxx",
        "AIzaSy-Account4-xxxxxx",
        "AIzaSy-Account5-xxxxxx"
    };
    // Lưu ý: Tên model có thể thay đổi tùy thời điểm Google release (gemma-2-27b-it, gemma-3-27b-it, v.v.)
    private static final String GOOGLE_MODEL = "gemma-2-27b-it"; 
    private static final String GOOGLE_URL_TEMPLATE = "https://generativelanguage.googleapis.com/v1beta/models/%s:generateContent?key=%s";

    // --- CẤU HÌNH FORMAT (HARD-CODE) ---
    private static final BigInteger MARGIN_LEFT = BigInteger.valueOf((long) (3.5 * 567)); 
    private static final BigInteger MARGIN_RIGHT = BigInteger.valueOf((long) (2.0 * 567));
    private static final BigInteger MARGIN_TOP = BigInteger.valueOf((long) (2.0 * 567));
    private static final BigInteger MARGIN_BOTTOM = BigInteger.valueOf((long) (2.0 * 567));
    private static final String FONT_FAMILY = "Times New Roman";
    private static final BigInteger SPACING_AFTER = BigInteger.valueOf(120); 

    public static void main(String[] args) {
        String inputPath = "D:\\Workspaces\\vietdoc-assistant\\input_doan.docx"; 
        String outputPath = "D:\\Workspaces\\vietdoc-assistant\\output_vietdoc_pro.docx";

        System.out.println(">>> VIETDOC CORE: Khởi động...");
        System.out.println(">>> AI PROVIDER: " + (CURRENT_PROVIDER == 1 ? "DeepSeek V3" : "Google Gemma 27B"));
        
        processDocument(inputPath, outputPath);
    }

    public static void processDocument(String inputPath, String outputPath) {
        try (FileInputStream fis = new FileInputStream(inputPath);
             XWPFDocument document = new XWPFDocument(fis)) {

            // 1. Format Cứng (Apache POI)
            setupPageMargins(document);
            setupPageNumber(document);
            processParagraphs(document); // Xử lý font, lề, ảnh
            processTables(document);
            generateTOC(document);

            // 2. Xuất file
            try (FileOutputStream out = new FileOutputStream(outputPath)) {
                document.write(out);
                System.out.println("🚀 [Done] File đã xong: " + outputPath);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // ================= LOGIC GỌI AI LINH HOẠT =================

    /**
     * Hàm điều phối gọi AI dựa trên Provider đã chọn
     */
    private static String callAI(String systemPrompt, String userPrompt) {
        if (CURRENT_PROVIDER == 1) {
            return callDeepSeek(systemPrompt, userPrompt);
        } else {
            return callGoogleGemma(systemPrompt, userPrompt);
        }
    }

    // --- LOGIC GỌI DEEPSEEK (OpenAI Standard) ---
    private static String callDeepSeek(String systemPrompt, String userPrompt) {
        try {
            String jsonBody = String.format(
                "{" +
                "  \"model\": \"%s\"," +
                "  \"messages\": [" +
                "    {\"role\": \"system\", \"content\": \"%s\"}," +
                "    {\"role\": \"user\", \"content\": \"%s\"}" +
                "  ]," +
                "  \"stream\": false" +
                "}", 
                DEEPSEEK_MODEL, escapeJson(systemPrompt), escapeJson(userPrompt)
            );

            return sendPostRequest(DEEPSEEK_URL, DEEPSEEK_API_KEY, jsonBody, "choices");
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // --- LOGIC GỌI GOOGLE GEMMA (Google Native Standard) ---
    private static String callGoogleGemma(String systemPrompt, String userPrompt) {
        try {
            // 1. Chọn ngẫu nhiên 1 Key để tránh Rate Limit (Load Balancing đơn giản)
            String randomKey = GOOGLE_API_KEYS[new Random().nextInt(GOOGLE_API_KEYS.length)];
            String url = String.format(GOOGLE_URL_TEMPLATE, GOOGLE_MODEL, randomKey);

            // 2. Cấu trúc JSON của Google khác DeepSeek (dùng "contents" -> "parts")
            // Google không có role "system" rõ ràng trong free tier cũ, ta gộp vào user prompt hoặc dùng cấu trúc mới
            // Ở đây gộp System prompt vào User prompt để an toàn nhất cho mọi model
            String fullPrompt = systemPrompt + "\n\nUser Request: " + userPrompt;

            String jsonBody = String.format(
                "{" +
                "  \"contents\": [{" +
                "    \"parts\": [{\"text\": \"%s\"}]" +
                "  }]" +
                "}", 
                escapeJson(fullPrompt)
            );

            // Google API dùng key trên URL, không cần Header Authorization
            return sendPostRequest(url, null, jsonBody, "candidates");

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // --- HÀM GỬI HTTP REQUEST CHUNG ---
    private static String sendPostRequest(String url, String apiKey, String jsonBody, String responseType) {
        try {
            HttpClient client = HttpClient.newHttpClient();
            HttpRequest.Builder builder = HttpRequest.newBuilder()
                    .uri(URI.create(url))
                    .header("Content-Type", "application/json");
            
            if (apiKey != null) {
                builder.header("Authorization", "Bearer " + apiKey);
            }

            HttpRequest request = builder.POST(HttpRequest.BodyPublishers.ofString(jsonBody, StandardCharsets.UTF_8)).build();
            HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());

            if (response.statusCode() == 200) {
                // Parse kết quả dựa trên loại API
                if (responseType.equals("choices")) {
                    return extractContentDeepSeek(response.body());
                } else {
                    return extractContentGoogle(response.body());
                }
            } else {
                System.err.println("❌ API Error (" + response.statusCode() + "): " + response.body());
                return null;
            }
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // ================= PARSER THỦ CÔNG (ĐỂ KHÔNG CẦN THƯ VIỆN JSON) =================

    private static String extractContentDeepSeek(String json) {
        // Tìm "content": "..."
        return parseJsonValue(json, "\"content\":");
    }

    private static String extractContentGoogle(String json) {
        // Tìm "text": "..." (Google trả về trong parts)
        return parseJsonValue(json, "\"text\":");
    }

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
                    if (c == 'n') sb.append('\n');
                    else if (c == '\"') sb.append('\"');
                    else if (c == 't') sb.append('\t');
                    else sb.append(c);
                    escaped = false;
                } else {
                    if (c == '\\') {
                        escaped = true;
                    } else if (c == '\"') {
                        break; // Hết chuỗi
                    } else {
                        sb.append(c);
                    }
                }
            }
            return sb.toString();
        } catch (Exception e) {
            return "Lỗi parse JSON";
        }
    }
    
    // Hàm escape ký tự đặc biệt để tạo JSON hợp lệ
    private static String escapeJson(String text) {
        if (text == null) return "";
        return text.replace("\\", "\\\\")
                   .replace("\"", "\\\"")
                   .replace("\n", "\\n")
                   .replace("\r", "");
    }

    // ================= LOGIC POI XỬ LÝ WORD (NHƯ CŨ) =================
    // (Đã rút gọn để tập trung vào phần AI, bạn giữ nguyên logic processParagraphs như file trước)

    private static void processParagraphs(XWPFDocument document) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (int i = 0; i < paragraphs.size(); i++) {
            XWPFParagraph p = paragraphs.get(i);
            
            // Format cơ bản
            p.setSpacingAfter(120);
            for (XWPFRun r : p.getRuns()) {
                r.setFontFamily(FONT_FAMILY);
                r.setFontSize(13);
            }

            // DETECT ẢNH & GỌI AI
            if (containsImage(p)) {
                boolean hasCaption = (i + 1 < paragraphs.size()) && 
                                     paragraphs.get(i+1).getText().trim().toLowerCase().startsWith("hình");
                if (!hasCaption) {
                    String context = getContext(paragraphs, i - 2, i + 2);
                    System.out.println("🔍 [AI] Đang sinh caption cho ảnh...");
                    
                    // GỌI HÀM AI MỚI TẠI ĐÂY
                    String aiCaption = callAI(
                        "Bạn là trợ lý soạn thảo. Chỉ trả về 1 câu caption ngắn gọn bắt đầu bằng 'Hình [số]:'.", 
                        "Hãy đặt tên cho hình ảnh dựa trên ngữ cảnh này: " + context
                    );
                    
                    if (aiCaption != null) insertCaption(p, aiCaption);
                }
            }
        }
    }
    
    private static void insertCaption(XWPFParagraph imgP, String text) {
        XmlCursor cursor = imgP.getCTP().newCursor();
        XWPFParagraph cP = imgP.getBody().insertNewParagraph(cursor);
        cP.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun r = cP.createRun();
        r.setText(text);
        r.setItalic(true);
        r.setFontSize(12);
        System.out.println("✅ Caption: " + text);
    }

    // --- Helper Functions (Giữ nguyên) ---
    private static void setupPageMargins(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        if (sectPr == null) sectPr = document.getDocument().getBody().addNewSectPr();
        CTPageMar pageMar = sectPr.getPgMar();
        if (pageMar == null) pageMar = sectPr.addNewPgMar();
        pageMar.setLeft(MARGIN_LEFT); pageMar.setRight(MARGIN_RIGHT);
        pageMar.setTop(MARGIN_TOP); pageMar.setBottom(MARGIN_BOTTOM);
    }
    private static void setupPageNumber(XWPFDocument document) {
        XWPFHeaderFooterPolicy policy = document.createHeaderFooterPolicy();
        XWPFFooter footer = policy.createFooter(STHdrFtr.DEFAULT);
        XWPFParagraph p = footer.createParagraph();
        p.setAlignment(ParagraphAlignment.CENTER);
        p.getCTP().addNewFldSimple().setInstr("PAGE \\* MERGEFORMAT");
    }
    private static void processTables(XWPFDocument document) {
        for (XWPFTable table : document.getTables()) {
            CTTblWidth tblW = table.getCTTbl().getTblPr().addNewTblW();
            tblW.setType(STTblWidth.DXA);
            tblW.setW(BigInteger.valueOf(9000));
        }
    }
    private static void generateTOC(XWPFDocument document) {
        XWPFParagraph p = document.createParagraph();
        p.getCTP().addNewFldSimple().setInstr("TOC \\o \"1-3\" \\h \\z \\u");
    }
    private static boolean containsImage(XWPFParagraph p) {
        for (XWPFRun r : p.getRuns()) if (r.getEmbeddedPictures().size() > 0) return true;
        return false;
    }
    private static String getContext(List<XWPFParagraph> list, int start, int end) {
        StringBuilder sb = new StringBuilder();
        for (int i = Math.max(0, start); i < Math.min(list.size(), end); i++) sb.append(list.get(i).getText()).append(" ");
        return sb.length() > 200 ? sb.substring(0, 200) : sb.toString();
    }
}