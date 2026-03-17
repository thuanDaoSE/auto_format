package vn.vietdoc.vietdoc_assistant.controller; // Đổi package cho đúng dự án của bạn

import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import vn.vietdoc.vietdoc_assistant.VietDocEngine; // Import class xử lý của bạn
import vn.vietdoc.vietdoc_assistant.FormattingParameters; // Import DTO
import vn.vietdoc.vietdoc_assistant.VietDocStyleConfig; // Import config

@RestController
@RequestMapping("/api/format")
@CrossOrigin(origins = "*") // Tạm thời cho phép mọi trang web gọi API này để test cho dễ
public class FormatController {

    private static final String DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    private static final long MAX_FILE_SIZE = 15L * 1024L * 1024L; // 15MB

    @PostMapping("/upload")
    public ResponseEntity<byte[]> uploadAndFormatFile(
            @RequestParam("file") MultipartFile file,
            @RequestParam(value = "marginLeft", required = false) Double marginLeft,
            @RequestParam(value = "marginRight", required = false) Double marginRight,
            @RequestParam(value = "marginTop", required = false) Double marginTop,
            @RequestParam(value = "marginBottom", required = false) Double marginBottom,
            @RequestParam(value = "fontSizeBody", required = false) Integer fontSizeBody,
            @RequestParam(value = "fontSizeHeading1", required = false) Integer fontSizeHeading1,
            @RequestParam(value = "fontSizeHeading2", required = false) Integer fontSizeHeading2,
            @RequestParam(value = "fontSizeHeading3", required = false) Integer fontSizeHeading3,
            @RequestParam(value = "fontSizeCaption", required = false) Integer fontSizeCaption,
            @RequestParam(value = "lineSpacing", required = false) Double lineSpacing,
            @RequestParam(value = "indentationFirstLine", required = false) Double indentationFirstLine,
            @RequestParam(value = "affectTableSize", required = false) Boolean affectTableSize,
            @RequestParam(value = "italicCaption", required = false) Boolean italicCaption) {

        // 1. KIỂM TRA BẢO MẬT ĐẦU VÀO (Chặn virus, file quá to, file sai đuôi)
        try {
            validateFile(file);
        } catch (IllegalArgumentException e) {
            // Nếu file có vấn đề -> Đá ra ngoài ngay lập tức
            return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(null);
        }

        // 2. TẠO FORMATTING PARAMETERS VỚI DEFAULT VALUE
        FormattingParameters params = VietDocStyleConfig.createDefaultParameters();

        // Override các giá trị nếu user truyền vào
        if (marginLeft != null)
            params.setMarginLeft(marginLeft);
        if (marginRight != null)
            params.setMarginRight(marginRight);
        if (marginTop != null)
            params.setMarginTop(marginTop);
        if (marginBottom != null)
            params.setMarginBottom(marginBottom);
        if (fontSizeBody != null)
            params.setFontSizeBody(fontSizeBody);
        if (fontSizeHeading1 != null)
            params.setFontSizeHeading1(fontSizeHeading1);
        if (fontSizeHeading2 != null)
            params.setFontSizeHeading2(fontSizeHeading2);
        if (fontSizeHeading3 != null)
            params.setFontSizeHeading3(fontSizeHeading3);
        if (fontSizeCaption != null)
            params.setFontSizeCaption(fontSizeCaption);
        if (lineSpacing != null)
            params.setLineSpacing(lineSpacing);
        if (indentationFirstLine != null)
            params.setIndentationFirstLine(indentationFirstLine);
        if (affectTableSize != null)
            params.setAffectTableSize(affectTableSize);
        if (italicCaption != null)
            params.setItalicCaption(italicCaption);

        // 3. VALIDATE CÁC THAM SỐ
        try {
            params.validate();
        } catch (IllegalArgumentException e) {
            System.err.println("Lỗi validation tham số: " + e.getMessage());
            return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(null);
        }

        System.out.println("Formatting Parameters: " + params.toString());

        // 4. ĐƯA VÀO LÒ XỬ LÝ (VietDocEngine)
        try {
            // Đọc file tải lên thành mảng byte
            byte[] fileBytes = file.getBytes();

            // Đưa vào lò xử lý với formatting parameters
            byte[] formattedFileBytes = VietDocEngine.processForWeb(fileBytes, params);

            // 5. ĐÓNG GÓI VÀ GỬI TRẢ FILE ĐÃ LÀM ĐẸP CHO NGƯỜI DÙNG TẢI XUỐNG
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType(DOCX_MIME_TYPE));
            // Đặt tên file khi người dùng tải về
            headers.setContentDispositionFormData("attachment", "FormatPro_" + file.getOriginalFilename());

            return new ResponseEntity<>(formattedFileBytes, headers, HttpStatus.OK);

        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }

    // --- HÀM KIỂM TRA BẢO MẬT ---
    private void validateFile(MultipartFile file) {
        if (file == null || file.isEmpty())
            throw new IllegalArgumentException("File trống.");
        if (file.getSize() > MAX_FILE_SIZE)
            throw new IllegalArgumentException("Dung lượng vượt quá 15MB.");

        String filename = file.getOriginalFilename();
        if (filename == null || !filename.toLowerCase().endsWith(".docx")) {
            throw new IllegalArgumentException("Chỉ hỗ trợ .docx");
        }

        String contentType = file.getContentType();
        if (contentType == null
                || (!contentType.equals(DOCX_MIME_TYPE) && !contentType.equals("application/octet-stream"))) {
            throw new IllegalArgumentException("Định dạng file không hợp lệ.");
        }
        if (contentType.contains("macroEnabled")) {
            throw new IllegalArgumentException("Không hỗ trợ file chứa Macro.");
        }
    }
}