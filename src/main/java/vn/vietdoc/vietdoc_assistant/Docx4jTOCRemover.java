package vn.vietdoc.vietdoc_assistant;

import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.docx4j.jaxb.Context;

import java.io.File;
import java.math.BigInteger;
import java.util.List;
import jakarta.xml.bind.JAXBElement;

public class Docx4jTOCRemover {

    public static void removeTOCSafe(String inputPath, String outputPath) throws Exception {
        System.out.println(">>> [Docx4j] Chế độ: Grafting (Ghép Section vào đoạn văn trước)...");
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(inputPath));
        MainDocumentPart mainPart = wordMLPackage.getMainDocumentPart();
        ObjectFactory factory = Context.getWmlObjectFactory();

        List<Object> bodyContent = mainPart.getContent();
        boolean found = false;

        // Duyệt ngược
        for (int i = bodyContent.size() - 1; i >= 0; i--) {
            Object obj = bodyContent.get(i);
            Object unwrappedObj = XmlUtils.unwrap(obj);

            if (unwrappedObj instanceof SdtBlock) {
                SdtBlock sdtBlock = (SdtBlock) unwrappedObj;
                SdtPr sdtPr = sdtBlock.getSdtPr();

                if (isTOCGallery(sdtPr)) {
                    System.out.println(">>> Phát hiện TOC tại index: " + i);

                    // 1. TÌM SECTION BREAK CŨ
                    SectPr oldSectPr = findSectPrInside(sdtBlock);

                    if (oldSectPr != null) {
                        // 2. TẠO BẢN SAO SẠCH CỦA SECTION BREAK
                        // Copy để giữ Header/Footer, nhưng xóa các ID rác
                        SectPr cleanSectPr = XmlUtils.deepCopy(oldSectPr);
                        cleanSectPr.setRsidR(null);
                        cleanSectPr.setRsidDel(null);
                        cleanSectPr.setRsidRPr(null);
                        cleanSectPr.setRsidSect(null);
                        
                        // Đảm bảo khổ giấy A4 (theo yêu cầu giữ lại để chắc chắn)
                        SectPr.PgSz pgSz = cleanSectPr.getPgSz();
                        if (pgSz == null) {
                            pgSz = factory.createSectPrPgSz();
                            cleanSectPr.setPgSz(pgSz);
                        }
                        pgSz.setW(BigInteger.valueOf(11906));
                        pgSz.setH(BigInteger.valueOf(16838));

                        // 3. CHIẾN THUẬT "GHÉP CÀNH" (GRAFTING)
                        // Tìm đoạn văn ngay phía trước TOC (dòng cuối trang bìa)
                        boolean grafted = false;
                        if (i > 0) {
                            Object prevObj = XmlUtils.unwrap(bodyContent.get(i - 1));
                            if (prevObj instanceof P) {
                                System.out.println(">>> [INFO] Tìm thấy đoạn văn phía trước. Đang ghép Section Break...");
                                P prevP = (P) prevObj;
                                if (prevP.getPPr() == null) prevP.setPPr(factory.createPPr());
                                
                                // Gán Section Break vào đoạn văn trước đó
                                // Điều này biến dòng cuối trang bìa thành điểm ngắt section
                                prevP.getPPr().setSectPr(cleanSectPr);
                                grafted = true;
                            }
                        }

                        // 4. NẾU KHÔNG GHÉP ĐƯỢC (TOC nằm đầu file hoặc trước Table)
                        // Buộc phải tạo đoạn văn mới nhưng dùng ID sạch
                        if (!grafted) {
                            System.out.println(">>> [INFO] Không ghép được. Tạo đoạn văn đệm mới.");
                            P newP = factory.createP();
                            PPr pPr = factory.createPPr();
                            pPr.setSectPr(cleanSectPr);
                            newP.setPPr(pPr);
                            
                            // Chèn vào vị trí hiện tại của TOC
                            bodyContent.add(i, newP);
                            // Vì chèn vào i, nên TOC bị đẩy xuống i+1, ta cần xóa i+1
                            // Nhưng logic ở dưới là remove(i), nên ta tăng i lên 1 ở vòng lặp sau? 
                            // Cách đơn giản: add vào i, sau đó remove cái i+1 (là cái TOC cũ)
                             bodyContent.remove(i + 1);
                             // Đã remove xong, set flag để không remove lần nữa ở dưới
                             found = true; 
                             continue; // Sang vòng lặp tiếp (đã xử lý xong node này)
                        }
                    } 

                    // 5. XÓA TOC (Nếu đã graft thành công hoặc không có SectPr)
                    bodyContent.remove(i);
                    System.out.println(">>> [SUCCESS] Đã xóa TOC và xử lý Section Break.");
                    found = true;
                }
            }
        }

        if (!found) {
            System.out.println(">>> Không tìm thấy TOC.");
        }

        sanitizeRunProperties(wordMLPackage);

        // FIX LỖI NAMESPACE (Chèn đoạn này trước khi save)
        Document wDoc = wordMLPackage.getMainDocumentPart().getJaxbElement();
        String ignorable = wDoc.getIgnorable();

        if (ignorable != null) {
            // Word 365 thường thêm w16du và w16sdtfl vào Ignorable
            // nhưng docx4j cũ lại không biết để thêm xmlns tương ứng -> Gây lỗi.
            // Giải pháp: Xóa chúng khỏi chuỗi Ignorable.
            
            String newIgnorable = ignorable.replace("w16du", "").replace("w16sdtfl", "").trim();
            
            // Xử lý thừa dấu cách do replace
            newIgnorable = newIgnorable.replaceAll(" +", " "); 
            
            wDoc.setIgnorable(newIgnorable);
            
            System.out.println(">>> [NAMESPACE FIX] Đã xóa w16du/w16sdtfl khỏi Ignorable.");
        }

        wordMLPackage.save(new File(outputPath));
        System.out.println(">>> Đã lưu file: " + outputPath);
    }

    // --- HELPER FUNCTIONS ---

    private static boolean isTOCGallery(SdtPr sdtPr) {
        if (sdtPr == null) return false;
        for (Object item : sdtPr.getRPrOrAliasOrLock()) {
            Object val = XmlUtils.unwrap(item);
            if (val instanceof CTSdtDocPart) {
                CTSdtDocPart docPart = (CTSdtDocPart) val;
                if (docPart.getDocPartGallery() != null && 
                    ("Table of Contents".equals(docPart.getDocPartGallery().getVal()) || 
                     "TOC".equals(docPart.getDocPartGallery().getVal()))) {
                    return true;
                }
            }
        }
        return false;
    }

    private static SectPr findSectPrInside(Object obj) {
        Object unwrapped = XmlUtils.unwrap(obj);
        if (unwrapped instanceof SectPr) return (SectPr) unwrapped;
        if (unwrapped instanceof P) { 
            PPr pPr = ((P) unwrapped).getPPr();
            if (pPr != null && pPr.getSectPr() != null) return pPr.getSectPr();
        }
        else if (unwrapped instanceof ContentAccessor) { 
            List<Object> children = ((ContentAccessor) unwrapped).getContent();
            for (Object child : children) {
                SectPr found = findSectPrInside(child);
                if (found != null) return found;
            }
        }
        return null;
    }

    private static void sanitizeRunProperties(WordprocessingMLPackage wordMLPackage) {
    MainDocumentPart mainPart = wordMLPackage.getMainDocumentPart();
    ObjectFactory factory = Context.getWmlObjectFactory();

    // Lấy tất cả các Run (R) trong văn bản
    List<Object> runs = getAllElementFromObject(mainPart.getContent(), R.class);
    int fixCount = 0;

    for (Object obj : runs) {
        R run = (R) obj;
        RPr oldRPr = run.getRPr();
        
        // Nếu không có định dạng gì thì bỏ qua
        if (oldRPr == null) continue;

        // --- BƯỚC QUAN TRỌNG: TẠO RPr MỚI TINH ---
        RPr newRPr = factory.createRPr();

        // --- CHÉP THUỘC TÍNH THEO ĐÚNG THỨ TỰ SCHEMA CỦA OPENXML ---
        // Thứ tự bắt buộc: rFonts -> b -> i -> ... -> color -> sz ...
        // Việc gọi setter theo thứ tự này giúp đảm bảo code tường minh, 
        // dù JAXB sẽ tự lo việc sắp xếp khi save, nhưng ta làm kỹ cho chắc.

        // 1. Run Fonts (rFonts) - Thường gây lỗi nhất
        if (oldRPr.getRFonts() != null) newRPr.setRFonts(oldRPr.getRFonts());

        // 2. Bold (b)
        if (oldRPr.getB() != null) newRPr.setB(oldRPr.getB());
        if (oldRPr.getBCs() != null) newRPr.setBCs(oldRPr.getBCs());

        // 3. Italic (i)
        if (oldRPr.getI() != null) newRPr.setI(oldRPr.getI());
        if (oldRPr.getICs() != null) newRPr.setICs(oldRPr.getICs());

        // 4. Strike / Caps (Các định dạng chữ khác)
        if (oldRPr.getCaps() != null) newRPr.setCaps(oldRPr.getCaps());
        if (oldRPr.getSmallCaps() != null) newRPr.setSmallCaps(oldRPr.getSmallCaps());
        if (oldRPr.getStrike() != null) newRPr.setStrike(oldRPr.getStrike());

        // 5. Color (color) - Thường bị đặt sai chỗ
        if (oldRPr.getColor() != null) newRPr.setColor(oldRPr.getColor());

        // 6. Spacing / Position (Giãn dòng, chỉ số trên/dưới)
        if (oldRPr.getSpacing() != null) newRPr.setSpacing(oldRPr.getSpacing());
        if (oldRPr.getVertAlign() != null) newRPr.setVertAlign(oldRPr.getVertAlign());

        // 7. Size (sz)
        if (oldRPr.getSz() != null) newRPr.setSz(oldRPr.getSz());
        if (oldRPr.getSzCs() != null) newRPr.setSzCs(oldRPr.getSzCs());

        // 8. Underline (u)
        if (oldRPr.getU() != null) newRPr.setU(oldRPr.getU());
        
        // 9. Highlight / Shading (Màu nền)
        if (oldRPr.getHighlight() != null) newRPr.setHighlight(oldRPr.getHighlight());
        if (oldRPr.getShd() != null) newRPr.setShd(oldRPr.getShd());

        // --- GÁN LẠI VÀO RUN ---
        // Thay thế hoàn toàn cái cũ bị lỗi bằng cái mới sạch sẽ
        run.setRPr(newRPr);
        fixCount++;
    }
    
    System.out.println(">>> [SANITIZER] Đã làm sạch định dạng cho " + fixCount + " run(s).");
}

    // Hàm bổ trợ tìm kiếm (Helper)
    private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new java.util.ArrayList<Object>();
        if (obj instanceof JAXBElement) obj = ((JAXBElement<?>) obj).getValue();

        if (obj.getClass().equals(toSearch)) {
            result.add(obj);
        } else if (obj instanceof org.docx4j.wml.ContentAccessor) {
            List<?> children = ((org.docx4j.wml.ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }
        }
        return result;
    }
}