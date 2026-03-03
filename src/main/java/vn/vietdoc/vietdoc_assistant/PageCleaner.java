package vn.vietdoc.vietdoc_assistant;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import java.util.List;

public class PageCleaner {

    /**
     * HÀM QUÉT DỌN TOÀN DIỆN (HÚT BỤI)
     * Xóa sạch mọi Section Break, Page Break và Dòng trống thừa từ một vị trí chỉ định.
     */
    public static void cleanAllBreaksAndEmptyLines(XWPFDocument doc, int startIndex) {
        System.out.println(">>> [PageCleaner] Đang dọn dẹp toàn bộ Break và Dòng rỗng từ index: " + startIndex);
        
        List<XWPFParagraph> paragraphs = doc.getParagraphs();
        if (startIndex <= 0 || startIndex >= paragraphs.size()) return;

        // Quét từ trên xuống dưới (Bắt đầu từ sau trang bìa)
        boolean restart = true;
        while (restart) {
            restart = false;
            paragraphs = doc.getParagraphs(); // Cập nhật lại list sau khi xóa
            
            for (int i = startIndex; i < paragraphs.size(); i++) {
                XWPFParagraph p = paragraphs.get(i);
                
                // 1. DỌN DẸP SECTION BREAK VÀ PAGE BREAK TRONG PROPERTIES
                if (p.getCTP().isSetPPr()) {
                    CTPPr ppr = p.getCTP().getPPr();
                    if (ppr.isSetSectPr()) ppr.unsetSectPr(); // Xóa ngắt phần
                    if (ppr.isSetPageBreakBefore()) ppr.unsetPageBreakBefore(); // Xóa ngắt trang cứng
                }

                // 2. DỌN DẸP PAGE BREAK THỦ CÔNG (Ctrl + Enter) TRONG RUNS
                for (XWPFRun r : p.getRuns()) {
                    if (r.getCTR() != null) {
                        List<CTBr> brList = r.getCTR().getBrList();
                        if (brList != null) {
                            for (int j = brList.size() - 1; j >= 0; j--) {
                                CTBr br = brList.get(j);
                                if (br.getType() == STBrType.PAGE || br.getType() == STBrType.COLUMN) {
                                    r.getCTR().removeBr(j); // Xóa ngắt trang/cột
                                }
                            }
                        }
                    }
                }

                // 3. HÚT BỤI DÒNG TRỐNG (Dòng không có chữ, không có ảnh)
                if (isEmptyParagraph(p)) {
                    doc.removeBodyElement(doc.getPosOfParagraph(p));
                    restart = true; // Reset vòng lặp vì list paragraph đã bị thay đổi kích thước
                    break;
                }
            }
        }
        System.out.println(">>> [PageCleaner] Hoàn tất dọn dẹp rác định dạng.");
    }

    // --- HELPER: Kiểm tra xem dòng có rỗng hoàn toàn không ---
    private static boolean isEmptyParagraph(XWPFParagraph p) {
        // Nếu có chữ -> Không rỗng
        if (p.getText().trim().length() > 0) return false;
        
        // Kiểm tra xem có ẩn chứa hình ảnh, hình vẽ hay biểu đồ không
        for (XWPFRun r : p.getRuns()) {
            if (!r.getEmbeddedPictures().isEmpty()) return false;
            if (r.getCTR() != null) {
                if (r.getCTR().sizeOfDrawingArray() > 0) return false;
                if (r.getCTR().sizeOfPictArray() > 0) return false;
                if (r.getCTR().sizeOfObjectArray() > 0) return false;
            }
        }
        return true;
    }

    public static void clean(XWPFDocument doc, int startIndex) {
        List<XWPFParagraph> paragraphs = doc.getParagraphs();
        if (startIndex <= 0 || startIndex >= paragraphs.size()) return;

        // --- BƯỚC 1: BẢO VỆ BÌA (GIỮ NGUYÊN SECTION BREAK CỦA USER) ---
        // Tìm Section Break gần nhất ở biên giới Bìa - Thân
        int protectedIndex = -1;
        if (hasSectionBreak(paragraphs.get(startIndex - 1))) {
            protectedIndex = startIndex - 1;
        } else if (hasSectionBreak(paragraphs.get(startIndex))) {
            protectedIndex = startIndex;
        }

        // Nếu không tìm thấy cái nào -> Tạo mới tại cuối bìa
        if (protectedIndex == -1) {
            XWPFParagraph lastCoverP = paragraphs.get(startIndex - 1);
            if (hasPageBreak(lastCoverP)) removePageBreak(lastCoverP); // Xóa PageBreak thường
            insertSectionBreak(lastCoverP);
            protectedIndex = startIndex - 1;
        }

        // --- BƯỚC 2: HẠ CẤP SECTION THỪA TRONG NỘI DUNG ---
        // Duyệt từ sau bìa đến áp chót
        for (int i = startIndex; i < paragraphs.size() - 1; i++) {
            if (i == protectedIndex) continue; // Không đụng vào biên giới bìa

            XWPFParagraph curr = paragraphs.get(i);
            XWPFParagraph next = paragraphs.get(i + 1);
            
            if (hasSectionBreak(curr)) {
                // 1. Xóa Section Break (để gộp định dạng, sửa lỗi mất khung/vỡ lề)
                removeSectionBreak(curr);
                
                // 2. Chuyển quyền ngắt trang cho dòng kế tiếp (next)
                // Thay vì chèn dòng mới (gây dấu lạ), ta bảo dòng sau: "Hãy nhảy sang trang mới trước khi bắt đầu"
                if (!hasPageBreak(next)) { // Nếu dòng sau chưa có ngắt trang
                    next.setPageBreak(true); // Set PageBreakBefore = true
                }
            }
        }

        // --- BƯỚC 3: DỌN RÁC (ENTER TRỐNG & PAGE BREAK KÉP) ---
        boolean restart = true;
        while (restart) {
            restart = false;
            for (int i = paragraphs.size() - 2; i >= startIndex; i--) {
                if (i == protectedIndex) continue;

                XWPFParagraph curr = paragraphs.get(i);
                XWPFParagraph next = paragraphs.get(i + 1);

                // A. XÓA "ENTER THỪA" (Dòng trống vô nghĩa)
                // Nếu dòng hiện tại trống VÀ dòng sau đã có lệnh ngắt trang (PageBreakBefore)
                // Thì dòng hiện tại là thừa -> Xóa
                boolean currEmpty = isEmptyParagraph(curr);
                boolean nextHasBreak = hasPageBreak(next) || (next.getCTP().getPPr() != null && next.getCTP().getPPr().isSetPageBreakBefore());
                
                if (currEmpty) {
                    // Nếu dòng này không có break gì cả -> Xóa
                    if (!hasBreak(curr)) {
                        doc.removeBodyElement(doc.getPosOfParagraph(curr));
                        restart = true; break;
                    }
                    // Nếu dòng này có PageBreak, nhưng dòng sau cũng có PageBreak -> Xóa dòng này
                    if (hasPageBreak(curr) && nextHasBreak) {
                        doc.removeBodyElement(doc.getPosOfParagraph(curr));
                        restart = true; break;
                    }
                }
            }
        }
    }

    // --- HELPER METHODS ---

    private static void insertSectionBreak(XWPFParagraph p) {
        if (p.getCTP().getPPr() == null) p.getCTP().addNewPPr();
        if (p.getCTP().getPPr().getSectPr() == null) {
            CTSectPr sectPr = p.getCTP().getPPr().addNewSectPr();
            sectPr.addNewType().setVal(STSectionMark.NEXT_PAGE);
        }
    }

    private static void removeSectionBreak(XWPFParagraph p) {
        if (p.getCTP().getPPr() != null && p.getCTP().getPPr().isSetSectPr()) {
            p.getCTP().getPPr().unsetSectPr();
        }
    }

    private static void removePageBreak(XWPFParagraph p) {
        // Xóa inline break
        for (XWPFRun r : p.getRuns()) {
            CTR ctr = r.getCTR();
            for (int i = ctr.sizeOfBrArray() - 1; i >= 0; i--) {
                if (ctr.getBrArray(i).getType() == STBrType.PAGE) ctr.removeBr(i);
            }
        }
        // Xóa property break
        if (p.getCTP().getPPr() != null && p.getCTP().getPPr().isSetPageBreakBefore()) {
            p.getCTP().getPPr().unsetPageBreakBefore();
        }
    }

    private static boolean hasSectionBreak(XWPFParagraph p) {
        return p.getCTP().getPPr() != null && p.getCTP().getPPr().getSectPr() != null;
    }

    private static boolean hasPageBreak(XWPFParagraph p) {
        for (XWPFRun r : p.getRuns()) {
            for (CTBr br : r.getCTR().getBrList()) {
                if (br.getType() == STBrType.PAGE) return true;
            }
        }
        if (p.getCTP().getPPr() != null && p.getCTP().getPPr().isSetPageBreakBefore()) return true;
        return false;
    }
    
    private static boolean hasBreak(XWPFParagraph p) {
        return hasPageBreak(p) || hasSectionBreak(p);
    }
}