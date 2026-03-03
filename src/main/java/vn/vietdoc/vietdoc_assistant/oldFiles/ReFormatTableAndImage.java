package vn.vietdoc.vietdoc_assistant.oldFiles;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.apache.xmlbeans.XmlCursor;
import javax.xml.namespace.QName;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.List;

public class ReFormatTableAndImage {

    private static final long EMU_PER_CM = 360000L;
    private static final long TWIPS_PER_CM = 567L;
    private static final double SAFE_WIDTH_CM = 15.5; 
    
    private static final long MAX_WIDTH_EMU = (long) (SAFE_WIDTH_CM * EMU_PER_CM); 
    private static final BigInteger TABLE_WIDTH_TWIPS = BigInteger.valueOf((long) (SAFE_WIDTH_CM * TWIPS_PER_CM)); 

    public static void main(String[] args) {
        String inputPath = "D:\\Workspaces\\vietdoc-assistant\\image_table_eg.docx";
        String outputPath = "D:\\Workspaces\\vietdoc-assistant\\fixed_beautiful.docx"; // Output file đẹp hơn

        try (FileInputStream fis = new FileInputStream(inputPath);
             XWPFDocument doc = new XWPFDocument(fis)) {

            System.out.println("Đang làm đẹp bảng và ảnh...");
            setupPageMargins(doc);

            for (XWPFTable table : doc.getTables()) {
                fixTableTotally(table);
            }

            for (XWPFParagraph p : doc.getParagraphs()) {
                standardizeParagraph(p);
            }

            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                doc.write(fos);
                System.out.println("Xong! File đẹp đã lưu tại: " + outputPath);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

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

        // --- TÍNH TOÁN "DISTRIBUTE COLUMNS EVENLY" ---
        // Lấy số lượng cột dựa trên dòng đầu tiên
        int colCount = 0;
        if (!table.getRows().isEmpty()) {
            colCount = table.getRow(0).getTableCells().size();
        }
        
        // Tính chiều rộng mỗi cột: Tổng width / Số cột
        BigInteger colWidth = (colCount > 0) 
                ? TABLE_WIDTH_TWIPS.divide(BigInteger.valueOf(colCount)) 
                : BigInteger.ZERO;

        // --- 2. ROW & CELL LEVEL PROCESSING ---
        List<XWPFTableRow> rows = table.getRows();
        for (int i = 0; i < rows.size(); i++) {
            XWPFTableRow row = rows.get(i);
            CTRow ctRow = row.getCtRow();

            // A. Fix lỗi Spacing dòng (Logic cũ vẫn giữ để đảm bảo kỹ thuật)
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

            // B. "DISTRIBUTE ROWS EVENLY" (Mô phỏng)
            // Set chiều cao tối thiểu cho dòng để chữ không bị bí. ~0.8cm (450 twips) là đẹp.
            // Nếu dòng đã có height rule thì thôi, nếu chưa thì set.
            // (Bạn có thể bỏ qua nếu muốn dòng tự co giãn theo nội dung)
            // row.setHeight(450); 

            // C. CELL LEVEL BEAUTIFICATION
            for (XWPFTableCell cell : row.getTableCells()) {
                // 1. Ép chiều rộng cột đều nhau (Distribute Columns)
                if (colCount > 0) {
                    CTTcPr tcPr = cell.getCTTc().getTcPr();
                    if (tcPr == null) tcPr = cell.getCTTc().addNewTcPr();
                    CTTblWidth cellW = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
                    cellW.setType(STTblWidth.DXA);
                    cellW.setW(colWidth);
                }

                // 2. Căn giữa dọc (Vertical Alignment) -> Giúp bảng rất thoáng
                cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                // 3. Xử lý Paragraph bên trong
                for (XWPFParagraph p : cell.getParagraphs()) {
                    standardizeParagraph(p); // Font chuẩn
                    
                    // Nếu là Header (Dòng 0), ta In đậm + Căn giữa ngang
                    if (i == 0) {
                        for (XWPFRun r : p.getRuns()) {
                            r.setBold(true); // Header in đậm
                        }
                        p.setAlignment(ParagraphAlignment.CENTER); // Header căn giữa
                    } else {
                        // Dữ liệu bình thường căn đều hoặc trái tùy ý (thường Justify hoặc Left)
                        p.setAlignment(ParagraphAlignment.BOTH); 
                    }
                }

                // 4. Tô màu nền cho Header (Row 0)
                if (i == 0) {
                    cell.setColor("E7E6E6"); // Màu xám nhạt (Light Gray) chuẩn công sở
                }
            }
        }
    }

    private static void standardizeParagraph(XWPFParagraph p) {
        p.setSpacingBetween(1.5);
        boolean hasImage = false;

        for (XWPFRun r : p.getRuns()) {
            r.setFontFamily("Times New Roman");
            r.setFontSize(14);
            
            // Logic tìm và xử lý ảnh
            List<org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing> drawings = r.getCTR().getDrawingList();
            if (drawings != null && !drawings.isEmpty()) {
                hasImage = true;
                for (org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing drawing : drawings) {
                    
                    // --- CÁCH 2: THAO TÁC TRỰC TIẾP VỚI CẤU TRÚC XML ---
                    // Kiểm tra xem có thẻ Anchor (ảnh đang bị Wrap) không
                    List<CTAnchor> anchorList = drawing.getAnchorList();
                    if (anchorList != null && !anchorList.isEmpty()) {
                        // Duyệt ngược từ cuối lên để xóa an toàn mà không bị lỗi index
                        for (int j = anchorList.size() - 1; j >= 0; j--) {
                            CTAnchor anchor = anchorList.get(j);
                            
                            // Tạo một thẻ Inline mới trong cùng cấu trúc Drawing
                            CTInline inline = drawing.addNewInline();
                            
                            // Bắt buộc phải cấu hình lề (dist) cho thẻ Inline (Word yêu cầu)
                            inline.setDistT(0L);
                            inline.setDistB(0L);
                            inline.setDistL(0L);
                            inline.setDistR(0L);
                            
                            // Sao chép toàn bộ các thuộc tính quan trọng từ Anchor sang Inline
                            // .set() là cách clone node XML an toàn nhất trong Apache XmlBeans
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
                            
                            // Xóa thẻ Anchor cũ sau khi đã copy xong
                            drawing.removeAnchor(j);
                        }
                    }

                    // --- XỬ LÝ RESIZE KÍCH THƯỚC ---
                    // Vòng lặp này giờ sẽ xử lý cả ảnh Inline cũ và ảnh Wrap vừa được ép sang Inline
                    for (CTInline inline : drawing.getInlineList()) {
                        resizeInlineImage(inline);
                    }
                }
            }
        }

        if (hasImage) {
            p.setAlignment(ParagraphAlignment.CENTER);
            if (p.getCTP().getPPr() != null && p.getCTP().getPPr().isSetInd()) {
                p.getCTP().getPPr().unsetInd(); // Xóa thụt đầu dòng (indent) để ảnh căn giữa chuẩn xác
            }
        }
    }

    private static String getBlipId(CTAnchor anchor) {
        XmlCursor cursor = anchor.newCursor();
        cursor.push();
        // Quét tìm thẻ <a:blip> chứa ID ảnh
        QName blipQName = new QName("http://schemas.openxmlformats.org/drawingml/2006/main", "blip");
        while (cursor.hasNextToken()) {
            if (cursor.isStart() && cursor.getName().equals(blipQName)) {
                String rId = cursor.getAttributeText(new QName("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed"));
                cursor.pop();
                return rId;
            }
            cursor.toNextToken();
        }
        cursor.pop();
        return null; // Không tìm thấy
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
}