package vn.vietdoc.vietdoc_assistant;

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
import java.util.regex.Pattern;

/**
 * VIETDOC CORE - Hybrid Document Processing System
 * 
 * LAYER 1: Hard Logic (Apache POI) - Precision formatting
 * LAYER 2: Soft Logic (AI APIs) - Content generation & understanding
 * 
 * Architecture: All functionality consolidated in one file as per requirements
 */
public class VietDocCore {

    // ================= CONFIGURATION =================
    
    /**
     * DEBUG MODE: Set to true to skip AI calls and use hard-coded placeholders
     */
    private static final boolean DEBUG_MODE = false;
    
    /**
     * AI Provider: 1 = DeepSeek, 2 = Google Gemma
     */
    private static final int CURRENT_PROVIDER = 2;
    
    /**
     * Batch size for processing images (optimize API calls)
     */
    private static final int BATCH_SIZE = 5;
    
    /**
     * Context window: chars above and below image for captioning
     */
    private static final int CONTEXT_CHARS_ABOVE = 150;
    private static final int CONTEXT_CHARS_BELOW = 100;
    
    // ================= API CONFIGURATION =================
    
    // DeepSeek Configuration
    private static final String DEEPSEEK_API_KEY = "sk-xxxxxxxxxxxx"; // Replace with actual key
    private static final String DEEPSEEK_URL = "https://api.deepseek.com/chat/completions";
    private static final String DEEPSEEK_MODEL = "deepseek-chat";
    
    // Google Gemma Configuration
    private static final String[] GOOGLE_API_KEYS = {
        "AIzaSyDvM-p9dId2FmbORHwNW4EtFhGg81Q6CHo", // Replace with actual keys
        "AIzaSyAMtoHdeizjVhBWnWdOe8yFPqGon6RMfcQ"
    };
    private static final String GOOGLE_MODEL = "gemma-3-27b-it";
    private static final String GOOGLE_URL_TEMPLATE = "https://generativelanguage.googleapis.com/v1beta/models/%s:generateContent?key=%s";
    
    // Rate limiting for Google Free Tier
    private static final long RATE_LIMIT_DELAY_MS = 1000; // 1 second delay between requests
    private static long lastApiCallTime = 0;
    
    // ================= FORMATTING CONSTANTS (Hard-coded per Vietnam Academic Standards) =================
    
    // Margins (in twips: 1cm = 567 twips)
    private static final BigInteger MARGIN_LEFT = BigInteger.valueOf((long) (3.5 * 567));   // 3.5cm
    private static final BigInteger MARGIN_RIGHT = BigInteger.valueOf((long) (2.0 * 567));  // 2.0cm
    private static final BigInteger MARGIN_TOP = BigInteger.valueOf((long) (2.0 * 567));    // 2.0cm
    private static final BigInteger MARGIN_BOTTOM = BigInteger.valueOf((long) (2.0 * 567)); // 2.0cm
    
    // Font settings
    private static final String FONT_FAMILY = "Times New Roman";
    private static final int FONT_SIZE_BODY = 13;
    private static final int FONT_SIZE_HEADING1 = 14;
    private static final int FONT_SIZE_CAPTION = 12;
    
    // Paragraph settings
    private static final int SPACING_BEFORE = 0;      // 0pt
    private static final int SPACING_AFTER = 120;     // 6pt = 120 twips
    private static final int LINE_SPACING_MULTIPLIER = 360; // 1.5 lines = 150%
    private static final int INDENTATION_FIRST_LINE = 567; // decree 30
    
    // Table settings
    private static final BigInteger TABLE_WIDTH = BigInteger.valueOf(9000); // ~16cm in twips
    
    // Missing section patterns (Vietnamese academic document structure)
    private static final String[] REQUIRED_SECTIONS = {
        "Lời cam đoan",
        "Lời cảm ơn",
        "Mục lục",
        "Danh mục",
        "Tóm tắt"
    };
    
    // ================= INNER CLASSES =================
    
    /**
     * Represents an image that needs caption generation
     */
    private static class ImageRequest {
        XWPFParagraph paragraph;
        String context;
        int imageIndex;
        
        public ImageRequest(XWPFParagraph p, String c, int idx) {
            this.paragraph = p;
            this.context = c;
            this.imageIndex = idx;
        }
    }
    
    // ================= MAIN ENTRY POINT =================
    
    public static void main(String[] args) {
        String inputPath = "D:\\Worksaces\\vietdoc-assistant\\input_doan.docx";
        String outputPath = "D:\\Worksppaces\\vietdoc-assistant\\output_vietdoc.docx";
        
        if (args.length >= 2) {
            inputPath = args[0];
            outputPath = args[1];
        }
        
        System.out.println("========================================");
        System.out.println("VIETDOC ASSISTANT - Document Processor");
        System.out.println("========================================");
        System.out.println("Debug Mode: " + DEBUG_MODE);
        System.out.println("AI Provider: " + (CURRENT_PROVIDER == 1 ? "DeepSeek" : "Google Gemma"));
        System.out.println("Input: " + inputPath);
        System.out.println("Output: " + outputPath);
        System.out.println("========================================\n");
        
        processDocument(inputPath, outputPath);
    }
    
    /**
     * Main document processing orchestrator
     */
    public static void processDocument(String inputPath, String outputPath) {
        try (FileInputStream fis = new FileInputStream(inputPath);
             XWPFDocument document = new XWPFDocument(fis)) {
            
            System.out.println("[1/6] Setting up page margins and formatting...");
            setupPageMargins(document);
            setupPageNumber(document);
            
            System.out.println("[2/6] Processing paragraphs and collecting images...");
            List<ImageRequest> pendingImages = new ArrayList<>();
            int[] imageCounter = {1}; // Use array to pass by reference
            processParagraphsAndCollectImages(document, pendingImages, imageCounter);
            
            System.out.println("[3/6] Checking for missing sections...");
            checkAndInsertMissingSections(document);
            
            System.out.println("[4/6] Processing tables...");
            processTables(document);
            
            System.out.println("[5/6] Generating captions for images...");
            if (!pendingImages.isEmpty()) {
                System.out.println("Found " + pendingImages.size() + " image(s) needing captions.");
                processBatches(pendingImages);
            } else {
                System.out.println("No images found requiring caption generation.");
            }
            
            System.out.println("[6/6] Saving document...");
            try (FileOutputStream out = new FileOutputStream(outputPath)) {
                document.write(out);
                System.out.println("\n✅ SUCCESS: Document processed and saved to: " + outputPath);
            }
            
        } catch (IOException e) {
            System.err.println("❌ ERROR: Failed to process document: " + e.getMessage());
            e.printStackTrace();
        } catch (Exception e) {
            System.err.println("❌ ERROR: Unexpected error: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    // ================= LAYER 1: HARD LOGIC - FORMATTING =================
    
    /**
     * Setup page margins according to Vietnam Academic Standards
     */
    private static void setupPageMargins(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        if (sectPr == null) {
            sectPr = document.getDocument().getBody().addNewSectPr();
        }
        
        CTPageMar pageMar = sectPr.getPgMar();
        if (pageMar == null) {
            pageMar = sectPr.addNewPgMar();
        }
        
        pageMar.setLeft(MARGIN_LEFT);
        pageMar.setRight(MARGIN_RIGHT);
        pageMar.setTop(MARGIN_TOP);
        pageMar.setBottom(MARGIN_BOTTOM);
        
        System.out.println("  ✓ Margins set: Left=3.5cm, Others=2.0cm");
    }
    
    /**
     * Setup page numbering in footer (centered)
     */
    private static void setupPageNumber(XWPFDocument document) {
        try {
            XWPFHeaderFooterPolicy policy = document.createHeaderFooterPolicy();
            XWPFFooter footer = policy.createFooter(STHdrFtr.DEFAULT);
            XWPFParagraph footerPara = footer.createParagraph();
            footerPara.setAlignment(ParagraphAlignment.CENTER);
            footerPara.getCTP().addNewFldSimple().setInstr("PAGE \\* MERGEFORMAT");
            System.out.println("  ✓ Page numbering configured");
        } catch (Exception e) {
            System.err.println("  ⚠ Warning: Could not set up page numbering: " + e.getMessage());
        }
    }
    
    /**
     * Process all paragraphs: format and collect images needing captions
     */
    private static void processParagraphsAndCollectImages(XWPFDocument document, 
                                                          List<ImageRequest> pendingImages, 
                                                          int[] imageCounter) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        
        for (int i = 0; i < paragraphs.size(); i++) {
            XWPFParagraph para = paragraphs.get(i);
            
            // Apply formatting to paragraph
            formatParagraph(para);
            
            // Check if paragraph contains image
            if (containsImage(para)) {
                // Check if caption already exists (next paragraph starts with "Hình")
                boolean hasCaption = false;
                if (i + 1 < paragraphs.size()) {
                    String nextText = paragraphs.get(i + 1).getText().trim().toLowerCase();
                    hasCaption = nextText.startsWith("hình") || nextText.startsWith("bảng");
                }
                
                if (!hasCaption) {
                    // Extract context window (150 chars above, 100 chars below)
                    String context = extractContext(paragraphs, i, CONTEXT_CHARS_ABOVE, CONTEXT_CHARS_BELOW);
                    pendingImages.add(new ImageRequest(para, context, imageCounter[0]++));
                }
            }
        }
        
        System.out.println("  ✓ Processed " + paragraphs.size() + " paragraph(s)");
    }
    
    /**
     * Format paragraph according to Vietnam Academic Standards
     */
    private static void formatParagraph(XWPFParagraph para) {
        if (para.getStyle() != null && para.getStyle().equals("Normal")) {
            para.setAlignment(ParagraphAlignment.BOTH); // Chuẩn Nghị định 30 cho nội dung
        }else{
            // Set alignment: CENTER (CENTER)
            para.setAlignment(ParagraphAlignment.CENTER);
        }
        
        // Set spacing: Before=0pt, After=6pt
        para.setSpacingBefore(SPACING_BEFORE);
        para.setSpacingAfter(SPACING_AFTER);
        
        // Set line spacing: 1.5 lines
        CTPPr ppr = para.getCTP().getPPr();
        if (ppr == null) {
            ppr = para.getCTP().addNewPPr();
        }
        CTSpacing spacing = ppr.getSpacing();
        if (spacing == null) {
            spacing = ppr.addNewSpacing();
        }
        spacing.setLineRule(STLineSpacingRule.AUTO);
        spacing.setLine(BigInteger.valueOf(LINE_SPACING_MULTIPLIER));
        
        boolean isHeading1 = isHeading1(para);
        // Set indentation for first line
        if (!isHeading1 && para.getAlignment() != ParagraphAlignment.CENTER) {
            para.setIndentationFirstLine(INDENTATION_FIRST_LINE);
        } else {
            // If it is a heading or centered, reset the first line indentation to avoid being shifted
            para.setIndentationFirstLine(0);
        }
        
        // Format runs: Font family and size
        int fontSize = isHeading1 ? FONT_SIZE_HEADING1 : FONT_SIZE_BODY;
        
        for (XWPFRun run : para.getRuns()) {
            run.setFontFamily(FONT_FAMILY);
            run.setFontSize(fontSize);
            if (isHeading1) {
                run.setBold(true);
            }

           // --- XỬ LÝ MẬT ĐỘ CHỮ (FIX DÍNH/THƯA) ---
        // SỬ DỤNG HÀM CÓ SẴN CỦA POI (Thay vì chọc vào CTRPr gây lỗi)
        
        // Fix lỗi Character Spacing (Khoảng cách giữa các ký tự)
        // Giá trị 0 là Normal.
        run.setCharacterSpacing(0); 

        // Lưu ý: Hàm setTextScale (w:w) thường không có sẵn trong bản POI Lite.
        // Tuy nhiên lỗi dính chữ chủ yếu do Spacing, nên setCharacterSpacing(0) là đủ xử lý 90% trường hợp.
        // Nếu bắt buộc phải xử lý Scale mà không có poi-ooxml-full, ta bỏ qua để tránh lỗi biên dịch.
        }
    }
    
    /**
     * Check if paragraph is a Heading 1 (size 14, bold)
     */
    private static boolean isHeading1(XWPFParagraph para) {
        // Simple heuristic: check if first run is bold and size 14
        if (para.getRuns().isEmpty()) return false;
        XWPFRun firstRun = para.getRuns().get(0);
        return firstRun.isBold() && firstRun.getFontSize() == FONT_SIZE_HEADING1;
    }
    
    /**
     * Process all tables: set width to prevent overflow
     */
    private static void processTables(XWPFDocument document) {
        int tableCount = 0;
        for (XWPFTable table : document.getTables()) {
            try {
                CTTblPr tblPr = table.getCTTbl().getTblPr();
                if (tblPr == null) {
                    tblPr = table.getCTTbl().addNewTblPr();
                }
                
                CTTblWidth tblW = tblPr.getTblW();
                if (tblW == null) {
                    tblW = tblPr.addNewTblW();
                }
                tblW.setType(STTblWidth.DXA);
                tblW.setW(TABLE_WIDTH);
                
                tableCount++;
            } catch (Exception e) {
                System.err.println("  ⚠ Warning: Could not format table: " + e.getMessage());
            }
        }
        System.out.println("  ✓ Processed " + tableCount + " table(s)");
    }
    
    /**
     * Check for missing required sections and insert placeholders
     */
    private static void checkAndInsertMissingSections(XWPFDocument document) {
        String documentText = extractFullText(document).toLowerCase();
        List<String> missingSections = new ArrayList<>();
        
        for (String section : REQUIRED_SECTIONS) {
            if (!documentText.contains(section.toLowerCase())) {
                missingSections.add(section);
            }
        }
        
        if (!missingSections.isEmpty()) {
            System.out.println("  ⚠ Found " + missingSections.size() + " missing section(s): " + missingSections);
            
            // Insert missing sections at the beginning (after title if exists)
            XWPFParagraph firstPara = document.getParagraphs().isEmpty() ? 
                document.createParagraph() : document.getParagraphs().get(0);
            
            for (String section : missingSections) {
                if (DEBUG_MODE) {
                    // In debug mode, insert placeholder
                    insertMissingSection(document, firstPara, section, "[PLACEHOLDER: " + section + "]");
                } else {
                    // In production, call AI to generate content
                    String content = generateMissingSection(section, documentText);
                    insertMissingSection(document, firstPara, section, content);
                }
            }
        } else {
            System.out.println("  ✓ All required sections found");
        }
    }
    
    /**
     * Insert a missing section into document
     */
    private static void insertMissingSection(XWPFDocument document, XWPFParagraph afterPara, 
                                            String sectionName, String content) {
        try {
            XmlCursor cursor = afterPara.getCTP().newCursor();
            XWPFParagraph newPara = document.insertNewParagraph(cursor);
            
            // Insert section title
            XWPFRun titleRun = newPara.createRun();
            titleRun.setText(sectionName);
            titleRun.setBold(true);
            titleRun.setFontSize(FONT_SIZE_HEADING1);
            titleRun.setFontFamily(FONT_FAMILY);
            
            // Insert content paragraph
            XWPFParagraph contentPara = document.insertNewParagraph(cursor);
            XWPFRun contentRun = contentPara.createRun();
            contentRun.setText(content);
            contentRun.setFontSize(FONT_SIZE_BODY);
            contentRun.setFontFamily(FONT_FAMILY);
            
            formatParagraph(newPara);
            formatParagraph(contentPara);
            
            System.out.println("  ✓ Inserted section: " + sectionName);
        } catch (Exception e) {
            System.err.println("  ⚠ Warning: Could not insert section " + sectionName + ": " + e.getMessage());
        }
    }
    
    // ================= LAYER 2: SOFT LOGIC - AI INTEGRATION =================
    
    /**
     * Process images in batches to optimize API calls
     */
    private static void processBatches(List<ImageRequest> allRequests) {
        int total = allRequests.size();
        
        for (int i = 0; i < total; i += BATCH_SIZE) {
            int end = Math.min(i + BATCH_SIZE, total);
            List<ImageRequest> batch = allRequests.subList(i, end);
            
            System.out.println("\n  Processing batch " + (i/BATCH_SIZE + 1) + " (images " + (i+1) + "-" + end + ")...");
            
            if (DEBUG_MODE) {
                // Debug mode: insert test captions
                processBatchDebug(batch);
            } else {
                // Production mode: call AI
                processBatchAI(batch);
            }
        }
    }
    
    /**
     * Process batch in debug mode (no AI calls)
     */
    private static void processBatchDebug(List<ImageRequest> batch) {
        for (ImageRequest req : batch) {
            String caption = "Hình " + req.imageIndex + ": [TEST CAPTION - Debug Mode]";
            insertCaption(req.paragraph, caption);
        }
        System.out.println("  ✓ Debug mode: Inserted " + batch.size() + " test caption(s)");
    }
    
    /**
     * Process batch using AI
     */
    private static void processBatchAI(List<ImageRequest> batch) {
        try {
            // Build prompt
            StringBuilder userPromptBuilder = new StringBuilder();
            userPromptBuilder.append("Danh sách các ngữ cảnh ảnh cần đặt tên:\n\n");
            for (int k = 0; k < batch.size(); k++) {
                userPromptBuilder.append("ITEM_").append(k).append(": ")
                    .append(batch.get(k).context).append("\n\n");
            }
            
            String systemPrompt = "Bạn là trợ lý soạn thảo văn bản học thuật nghiêm ngặt. " +
                    "Nhiệm vụ: Đặt tên (caption) ngắn gọn cho từng ngữ cảnh ảnh được cung cấp.\n" +
                    "Quy tắc phản hồi QUAN TRỌNG:\n" +
                    "1. Chỉ trả về danh sách các caption.\n" +
                    "2. Các caption ngăn cách nhau CHÍNH XÁC bởi chuỗi ký tự '|||'.\n" +
                    "3. Định dạng: Hình [số]: [Tên ngắn gọn].\n" +
                    "4. Không thêm lời dẫn, không thêm số thứ tự ở đầu dòng, chỉ nội dung text ngăn cách bởi |||.\n" +
                    "5. Output JSON only.";
            
            String userPrompt = userPromptBuilder.toString();
            
            // Call AI
            String aiResponse = callAI(systemPrompt, userPrompt);
            
            if (aiResponse != null && !aiResponse.isEmpty()) {
                // Parse response
                String[] captions = aiResponse.split("\\|\\|\\|");
                
                // Insert captions
                for (int k = 0; k < batch.size(); k++) {
                    if (k < captions.length) {
                        String caption = cleanCaption(captions[k], k);
                        insertCaption(batch.get(k).paragraph, caption);
                    } else {
                        // Fallback: use placeholder
                        String fallbackCaption = "Hình " + batch.get(k).imageIndex + ": [Caption không khả dụng]";
                        insertCaption(batch.get(k).paragraph, fallbackCaption);
                        System.err.println("  ⚠ Warning: AI returned fewer captions than requested. Using fallback.");
                    }
                }
                System.out.println("  ✓ Generated " + batch.size() + " caption(s)");
            } else {
                // Fallback: insert placeholders
                System.err.println("  ⚠ Warning: AI call failed. Using fallback captions.");
                for (ImageRequest req : batch) {
                    String fallbackCaption = "Hình " + req.imageIndex + ": [Caption không khả dụng]";
                    insertCaption(req.paragraph, fallbackCaption);
                }
            }
            
        } catch (Exception e) {
            System.err.println("  ❌ Error processing batch: " + e.getMessage());
            // Fallback: insert placeholders
            for (ImageRequest req : batch) {
                String fallbackCaption = "Hình " + req.imageIndex + ": [Caption không khả dụng]";
                insertCaption(req.paragraph, fallbackCaption);
            }
        }
    }
    
    /**
     * Generate content for missing section using AI
     */
    private static String generateMissingSection(String sectionName, String documentContext) {
        try {
            String systemPrompt = "Bạn là trợ lý soạn thảo văn bản học thuật nghiêm ngặt. " +
                    "Nhiệm vụ: Tạo nội dung cho phần '" + sectionName + "' trong một văn bản học thuật Việt Nam. " +
                    "Output JSON only.";
            
            String userPrompt = "Tạo nội dung ngắn gọn (khoảng 100-200 từ) cho phần '" + sectionName + "'. " +
                    "Ngữ cảnh tài liệu: " + documentContext.substring(0, Math.min(500, documentContext.length()));
            
            String response = callAI(systemPrompt, userPrompt);
            
            if (response != null && !response.isEmpty()) {
                return response;
            }
        } catch (Exception e) {
            System.err.println("  ⚠ Warning: Could not generate content for " + sectionName + ": " + e.getMessage());
        }
        
        // Fallback
        return "[Nội dung cho phần " + sectionName + " sẽ được bổ sung sau]";
    }
    
    /**
     * Main AI call dispatcher
     */
    private static String callAI(String systemPrompt, String userPrompt) {
        if (CURRENT_PROVIDER == 1) {
            return callDeepSeek(systemPrompt, userPrompt);
        } else {
            return callGoogleGemma(systemPrompt, userPrompt);
        }
    }
    
    /**
     * Call DeepSeek API (OpenAI-compatible)
     */
    private static String callDeepSeek(String systemPrompt, String userPrompt) {
        try {
            enforceRateLimit();
            
            String jsonBody = String.format(
                "{\"model\": \"%s\", \"messages\": [{\"role\": \"system\", \"content\": \"%s\"}, {\"role\": \"user\", \"content\": \"%s\"}], \"stream\": false}",
                DEEPSEEK_MODEL, escapeJson(systemPrompt), escapeJson(userPrompt)
            );
            
            return sendPostRequest(DEEPSEEK_URL, DEEPSEEK_API_KEY, jsonBody, "choices");
            
        } catch (Exception e) {
            System.err.println("  ❌ DeepSeek API error: " + e.getMessage());
            return null;
        }
    }
    
    /**
     * Call Google Gemma API
     */
    private static String callGoogleGemma(String systemPrompt, String userPrompt) {
        try {
            enforceRateLimit();
            
            // Rotate API keys for rate limit handling
            String randomKey = GOOGLE_API_KEYS[new Random().nextInt(GOOGLE_API_KEYS.length)];
            String url = String.format(GOOGLE_URL_TEMPLATE, GOOGLE_MODEL, randomKey);
            
            String fullPrompt = systemPrompt + "\n\nUser Request: " + userPrompt;
            String jsonBody = String.format(
                "{\"contents\": [{\"parts\": [{\"text\": \"%s\"}]}]}",
                escapeJson(fullPrompt)
            );
            
            return sendPostRequest(url, null, jsonBody, "candidates");
            
        } catch (Exception e) {
            System.err.println("  ❌ Google Gemma API error: " + e.getMessage());
            return null;
        }
    }
    
    /**
     * Send HTTP POST request
     */
    private static String sendPostRequest(String url, String apiKey, String jsonBody, String responseType) {
        try {
            HttpClient client = HttpClient.newHttpClient();
            HttpRequest.Builder builder = HttpRequest.newBuilder()
                    .uri(URI.create(url))
                    .header("Content-Type", "application/json");
            
            if (apiKey != null && !apiKey.isEmpty() && !apiKey.contains("xxxx")) {
                builder.header("Authorization", "Bearer " + apiKey);
            }
            
            HttpRequest request = builder
                    .POST(HttpRequest.BodyPublishers.ofString(jsonBody, StandardCharsets.UTF_8))
                    .build();
            
            HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());
            
            if (response.statusCode() == 200) {
                if (responseType.equals("choices")) {
                    return extractContentDeepSeek(response.body());
                } else {
                    return extractContentGoogle(response.body());
                }
            } else {
                System.err.println("  ❌ API Error (" + response.statusCode() + "): " + response.body());
                return null;
            }
            
        } catch (Exception e) {
            System.err.println("  ❌ HTTP request error: " + e.getMessage());
            return null;
        }
    }
    
    /**
     * Extract content from DeepSeek response (choices.message.content)
     */
    private static String extractContentDeepSeek(String json) {
        return parseJsonValue(json, "\"content\":");
    }
    
    /**
     * Extract content from Google response (candidates.content.parts.text)
     */
    private static String extractContentGoogle(String json) {
        return parseJsonValue(json, "\"text\":");
    }
    
    /**
     * Simple JSON parser (manual parsing to avoid dependencies)
     */
    private static String parseJsonValue(String json, String key) {
        try {
            int start = json.indexOf(key);
            if (start == -1) return null;
            
            start = json.indexOf("\"", start + key.length()) + 1;
            if (start == 0) return null;
            
            StringBuilder sb = new StringBuilder();
            boolean escaped = false;
            
            for (int i = start; i < json.length(); i++) {
                char c = json.charAt(i);
                
                if (escaped) {
                    if (c == 'n') sb.append('\n');
                    else if (c == '"') sb.append('"');
                    else if (c == '\\') sb.append('\\');
                    else sb.append(c);
                    escaped = false;
                } else {
                    if (c == '\\') {
                        escaped = true;
                    } else if (c == '"') {
                        break;
                    } else {
                        sb.append(c);
                    }
                }
            }
            
            return sb.toString().trim();
            
        } catch (Exception e) {
            System.err.println("  ⚠ JSON parsing error: " + e.getMessage());
            return "";
        }
    }
    
    /**
     * Enforce rate limiting (especially for Google Free Tier)
     */
    private static void enforceRateLimit() {
        long currentTime = System.currentTimeMillis();
        long timeSinceLastCall = currentTime - lastApiCallTime;
        
        if (timeSinceLastCall < RATE_LIMIT_DELAY_MS) {
            try {
                Thread.sleep(RATE_LIMIT_DELAY_MS - timeSinceLastCall);
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
            }
        }
        
        lastApiCallTime = System.currentTimeMillis();
    }
    
    // ================= UTILITY FUNCTIONS =================
    
    /**
     * Check if paragraph contains an image
     */
    private static boolean containsImage(XWPFParagraph para) {
        for (XWPFRun run : para.getRuns()) {
            if (run.getEmbeddedPictures().size() > 0) {
                return true;
            }
        }
        return false;
    }
    
    /**
     * Extract context around a paragraph (limited chars above and below)
     */
    private static String extractContext(List<XWPFParagraph> paragraphs, int index, 
                                        int charsAbove, int charsBelow) {
        StringBuilder sb = new StringBuilder();
        
        // Extract text above
        int charsCollected = 0;
        for (int i = index - 1; i >= 0 && charsCollected < charsAbove; i--) {
            String text = paragraphs.get(i).getText();
            if (charsCollected + text.length() <= charsAbove) {
                sb.insert(0, text + " ");
                charsCollected += text.length();
            } else {
                int remaining = charsAbove - charsCollected;
                sb.insert(0, text.substring(Math.max(0, text.length() - remaining)) + " ");
                break;
            }
        }
        
        // Extract text below
        charsCollected = 0;
        for (int i = index + 1; i < paragraphs.size() && charsCollected < charsBelow; i++) {
            String text = paragraphs.get(i).getText();
            if (charsCollected + text.length() <= charsBelow) {
                sb.append(text).append(" ");
                charsCollected += text.length();
            } else {
                int remaining = charsBelow - charsCollected;
                sb.append(text.substring(0, Math.min(remaining, text.length())));
                break;
            }
        }
        
        return sb.toString().trim();
    }
    
    /**
     * Extract full text from document for section checking
     */
    private static String extractFullText(XWPFDocument document) {
        StringBuilder sb = new StringBuilder();
        for (XWPFParagraph para : document.getParagraphs()) {
            sb.append(para.getText()).append(" ");
        }
        return sb.toString();
    }
    
    /**
     * Insert caption after image paragraph
     */
    private static void insertCaption(XWPFParagraph imgPara, String captionText) {
        try {
            XmlCursor cursor = imgPara.getCTP().newCursor();
            XWPFParagraph captionPara = imgPara.getBody().insertNewParagraph(cursor);
            
            captionPara.setAlignment(ParagraphAlignment.CENTER);
            
            XWPFRun run = captionPara.createRun();
            run.setText(captionText);
            run.setItalic(true);
            run.setFontSize(FONT_SIZE_CAPTION);
            run.setFontFamily(FONT_FAMILY);
            
            formatParagraph(captionPara);
            
            System.out.println("    ✓ Inserted caption: " + captionText);
            
        } catch (Exception e) {
            System.err.println("    ⚠ Warning: Could not insert caption: " + e.getMessage());
        }
    }
    
    /**
     * Clean caption text from AI response
     */
    private static String cleanCaption(String caption, int index) {
        if (caption == null) return "";
        
        // Remove ITEM_X: prefix if present
        caption = caption.replaceAll("ITEM_" + index + ":\\s*", "");
        
        // Remove newlines and extra spaces
        caption = caption.replace("\n", " ").replace("\r", "");
        caption = caption.replaceAll("\\s+", " ").trim();
        
        // Ensure it starts with "Hình" if not already
        if (!caption.toLowerCase().startsWith("hình")) {
            caption = "Hình " + (index + 1) + ": " + caption;
        }
        
        return caption;
    }
    
    /**
     * Escape JSON special characters
     */
    private static String escapeJson(String text) {
        if (text == null) return "";
        return text.replace("\\", "\\\\")
                   .replace("\"", "\\\"")
                   .replace("\n", "\\n")
                   .replace("\r", "\\r")
                   .replace("\t", "\\t");
    }
}

