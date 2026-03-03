package vn.vietdoc.vietdoc_assistant.controller; // Đổi package cho đúng dự án của bạn

import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import vn.vietdoc.vietdoc_assistant.VietDocEngine; // Import class xử lý của bạn

@RestController
@RequestMapping("/api/format")
@CrossOrigin(origins = "*") // Tạm thời cho phép mọi trang web gọi API này để test cho dễ
public class FormatController {

    private static final String DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    private static final long MAX_FILE_SIZE = 15L * 1024L * 1024L; // 15MB

    @PostMapping("/upload")
    public ResponseEntity<byte[]> uploadAndFormatFile(@RequestParam("file") MultipartFile file) {
        
        // 1. KIỂM TRA BẢO MẬT ĐẦU VÀO (Chặn virus, file quá to, file sai đuôi)
        try {
            validateFile(file);
        } catch (IllegalArgumentException e) {
            // Nếu file có vấn đề -> Đá ra ngoài ngay lập tức
            return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(null); 
        }

        // 2. ĐƯA VÀO LÒ XỬ LÝ (VietDocEngine)
        try {
            // Gọi hàm mới tạo ở Bước 1
            // Đọc file tải lên thành mảng byte
            byte[] fileBytes = file.getBytes(); 

            // Đưa vào lò xử lý
            byte[] formattedFileBytes = VietDocEngine.processForWeb(fileBytes);

            // 3. ĐÓNG GÓI VÀ GỬI TRẢ FILE ĐÃ LÀM ĐẸP CHO NGƯỜI DÙNG TẢI XUỐNG
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
        if (file == null || file.isEmpty()) throw new IllegalArgumentException("File trống.");
        if (file.getSize() > MAX_FILE_SIZE) throw new IllegalArgumentException("Dung lượng vượt quá 15MB.");
        
        String filename = file.getOriginalFilename();
        if (filename == null || !filename.toLowerCase().endsWith(".docx")) {
            throw new IllegalArgumentException("Chỉ hỗ trợ .docx");
        }
        
        String contentType = file.getContentType();
        if (contentType == null || (!contentType.equals(DOCX_MIME_TYPE) && !contentType.equals("application/octet-stream"))) {
            throw new IllegalArgumentException("Định dạng file không hợp lệ.");
        }
        if (contentType.contains("macroEnabled")) {
            throw new IllegalArgumentException("Không hỗ trợ file chứa Macro.");
        }
    }
}