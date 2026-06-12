package vn.vietdoc.vietdoc_assistant;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.nio.file.Files;
import java.util.List;

public class VietDocEngine {

    // Đường dẫn tới file style (Bạn nhớ chép file này vào VPS nhé)
    // Nếu dùng Spring Boot, tốt nhất nên đặt file này vào thư mục resources hoặc
    // map volume bên ngoài.
    private static final String STYLE_XML_PATH_DOCKER = "/app/data/styles_config.xml";
    private static final String STYLE_XML_PATH_LOCAL = "../data/styles_config.xml";
    private static final String STYLE_XML_PATH_FALLBACK = "src/main/java/vn/vietdoc/vietdoc_assistant/styles_config.xml";

    private static String getStyleXmlPath() {
        if (new File(STYLE_XML_PATH_DOCKER).exists()) {
            return STYLE_XML_PATH_DOCKER;
        }
        if (new File(STYLE_XML_PATH_LOCAL).exists()) {
            return STYLE_XML_PATH_LOCAL;
        }
        return STYLE_XML_PATH_FALLBACK;
    }

    /**
     * Hàm chính chuyên dùng cho Controller (Web) gọi vào
     * 
     * @param uploadedFileBytes Bytes của file DOCX từ user tải lên
     * @param params            FormattingParameters chứa các tham số tùy chỉnh
     * @return Bytes của file DOCX đã format
     */
    public static byte[] processForWeb(byte[] uploadedFileBytes, FormattingParameters params) throws Exception {
        // 1. Tạo file tạm ngẫu nhiên trên ổ cứng VPS để tránh đụng độ khi có nhiều
        // người dùng cùng lúc
        File tempInputFile = File.createTempFile("formatpro_input_", ".docx");
        File tempOutputFile = File.createTempFile("formatpro_output_", ".docx");

        try {
            // Lưu mảng byte người dùng gửi lên thành file tạm
            Files.write(tempInputFile.toPath(), uploadedFileBytes);

            // ========================================================
            // BƯỚC 1: TẠO FILE OUTPUT TỪ INPUT (COPY) & XÓA TOC CŨ
            // ========================================================
            try {
                // Class này của bạn đang nhận vào String path, nên ta truyền đường dẫn file tạm
                // vào
                Docx4jTOCRemover.removeTOCSafe(tempInputFile.getAbsolutePath(), tempOutputFile.getAbsolutePath());
            } catch (Exception e) {
                System.out.println("Lỗi ở Docx4jTOCRemover, chuyển sang copy thuần.");
                Files.copy(tempInputFile.toPath(), tempOutputFile.toPath(),
                        java.nio.file.StandardCopyOption.REPLACE_EXISTING);
            }

            // ========================================================
            // BƯỚC 2: CHUẨN HÓA LOGIC BẰNG POI
            // ========================================================
            System.out.println(">>> Giai đoạn 1: Chuẩn hóa Heading & TOC (POI)...");
            try (FileInputStream fis = new FileInputStream(tempOutputFile);
                    XWPFDocument document = new XWPFDocument(fis)) {

                int startIdx = 0;
                try {
                    startIdx = CoverPageDetector.detect(document) + 1;
                } catch (Exception e) {
                }

                try {
                    System.out.println("... Đang gỡ bỏ thuộc tính TOC khỏi Trang bìa ...");
                    CoverPageDetector.removeHeadingsFromCover(document);
                } catch (Exception e) {
                }

                try {
                    System.out.println("... Đang dọn dẹp Page Break thủ công ...");
                    PageCleaner.cleanAllBreaksAndEmptyLines(document, startIdx);
                } catch (Exception e) {
                }

                try {
                    TableImageProcessor.process(document, startIdx, params);
                } catch (Exception e) {
                }
                try {
                    GeneralFormatter.formatNormalText(document, startIdx, params);
                } catch (Exception e) {
                }

                HeadingProcessor.process(document, startIdx);

                try {
                    PageCleaner.clean(document, startIdx);
                } catch (Exception e) {
                }
                try {
                    BulletConverter.process(document);
                } catch (Exception e) {
                }

                HeadingProcessor.addOrUpdateTOC(document, startIdx);
                PageNumberFormatter.formatPageNumbers(document, params);
                forceAllSectionsToA4(document, params);

                // Ghi đè lại vào file tạm output
                try (FileOutputStream fos = new FileOutputStream(tempOutputFile)) {
                    document.write(fos);
                    System.out.println(">>> [POI] Đã xử lý nội dung xong.");
                }
            }

            // ========================================================
            // BƯỚC 3: THAY THẾ FILE STYLES.XML BÊN TRONG (ZIP SWAP)
            // ========================================================
            String styleXmlPath = getStyleXmlPath();
            File styleFile = new File(styleXmlPath);
            if (styleFile.exists()) {
                StyleSwapper.swapStyles(tempOutputFile.getAbsolutePath(), styleXmlPath, params);
            } else {
                System.err.println("CẢNH BÁO: Không tìm thấy file styles_config.xml tại " + styleXmlPath);
            }

            // 4. Đọc file tạm đã hoàn thiện thành mảng byte để gửi trả cho Controller
            return Files.readAllBytes(tempOutputFile.toPath());

        } finally {
            // [QUAN TRỌNG NHẤT] Dọn dẹp chiến trường
            // Dù thành công hay lỗi, luôn phải xóa file tạm để chống tràn ổ cứng VPS
            if (tempInputFile.exists())
                tempInputFile.delete();
            if (tempOutputFile.exists())
                tempOutputFile.delete();
        }
    }

    // --- HÀM ÉP KHỔ GIẤY A4 + MARGIN --- (CẬP NHẬT)
    private static void forceAllSectionsToA4(XWPFDocument document, FormattingParameters params) {
        BigInteger width = BigInteger.valueOf(11906);
        BigInteger height = BigInteger.valueOf(16838);

        CTSectPr bodySectPr = document.getDocument().getBody().getSectPr();
        if (bodySectPr != null) {
            setPageSize(bodySectPr, width, height, params);
        }

        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (XWPFParagraph p : paragraphs) {
            if (p.getCTP().getPPr() != null && p.getCTP().getPPr().getSectPr() != null) {
                CTSectPr sectionSectPr = p.getCTP().getPPr().getSectPr();
                setPageSize(sectionSectPr, width, height, params);
            }
        }
    }

    private static void setPageSize(CTSectPr sectPr, BigInteger width, BigInteger height, FormattingParameters params) {
        CTPageSz pgSz = sectPr.getPgSz();
        if (pgSz == null) {
            pgSz = sectPr.addNewPgSz();
        }
        pgSz.setW(width);
        pgSz.setH(height);

        // 2. Set Căn lề (THÊM BƯỚC CHECK NULL Ở ĐÂY)
        CTPageMar pgMar = sectPr.getPgMar();
        if (pgMar == null) {
            pgMar = sectPr.addNewPgMar(); // Nếu file Word chưa có thẻ margin thì tạo mới
        }
        // Set margins từ FormattingParameters
        sectPr.getPgMar().setLeft(params.getMarginLeftTwips());
        sectPr.getPgMar().setRight(params.getMarginRightTwips());
        sectPr.getPgMar().setTop(params.getMarginTopTwips());
        sectPr.getPgMar().setBottom(params.getMarginBottomTwips());
    }
}