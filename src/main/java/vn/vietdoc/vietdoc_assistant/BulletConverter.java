package vn.vietdoc.vietdoc_assistant;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigInteger;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class BulletConverter {

    private static final Pattern MANUAL_BULLET_PATTERN = Pattern.compile("^\\s*([-+*•▪])\\s+(.*)");

    public static void process(XWPFDocument doc) {
        System.out.println(">>> Đang chuẩn hóa Bullet (Reset lề thủ công)...");

        // 1. Khởi tạo Style chuẩn
        BigInteger standardNumID = setupCustomBulletStyles(doc);

        if (standardNumID == null) {
            System.err.println("Lỗi: Không thể khởi tạo hệ thống Numbering.");
            return;
        }

        List<XWPFParagraph> paragraphs = doc.getParagraphs();
        for (XWPFParagraph p : paragraphs) {
            String text = p.getText();
            
            if (text == null || text.trim().isEmpty() || (p.getStyleID() != null && p.getStyleID().startsWith("Heading"))) {
                continue;
            }

            // --- TRƯỜNG HỢP 1: ĐÃ LÀ BULLET TỰ ĐỘNG (CÓ SẴN) ---
            if (p.getNumID() != null) {
                if (isBulletList(doc, p.getNumID())) {
                    int currentLevel = p.getNumIlvl() != null ? p.getNumIlvl().intValue() : 0;
                    
                    // Gán Style chuẩn
                    applyNumbering(p, standardNumID, currentLevel);
                    
                    // [FIX QUAN TRỌNG] Xóa Indent thủ công cũ để các dòng thẳng tắp theo Style mới
                    resetParagraphIndent(p);
                }
                continue; 
            }

            // --- TRƯỜNG HỢP 2: BULLET THỦ CÔNG (KÝ TỰ) ---
            Matcher m = MANUAL_BULLET_PATTERN.matcher(text);
            if (m.matches()) {
                String marker = m.group(1);
                String content = m.group(2);
                int level = mapMarkerToLevel(marker);

                // 1. Xóa nội dung cũ
                while (!p.getRuns().isEmpty()) p.removeRun(0);
                
                // 2. Ghi nội dung mới
                XWPFRun run = p.createRun();
                run.setText(content);
                run.setFontFamily("Times New Roman");
                run.setFontSize(13);

                // 3. Gán List chuẩn
                applyNumbering(p, standardNumID, level);

                // 4. Xóa Indent tay
                resetParagraphIndent(p);

                System.out.println("Converted manual: [" + marker + "] -> Level " + level);
            }
        }
    }

    // --- CÁC HÀM BỔ TRỢ ---

    /**
     * Xóa sạch các thiết lập thụt lề (Left, Right, Hanging) thủ công trên đoạn văn
     * để đoạn văn ăn theo thiết lập chuẩn của Bullet Style.
     */
    private static void resetParagraphIndent(XWPFParagraph p) {
        if (p.getCTP().isSetPPr()) {
            CTPPr ppr = p.getCTP().getPPr();
            
            // Xóa Indent (Left/Right/Hanging)
            if (ppr.isSetInd()) {
                ppr.unsetInd();
            }
            
            // Xóa Tabs (để tránh dấu Tab cũ làm nhảy chữ lung tung)
            if (ppr.isSetTabs()) {
                ppr.unsetTabs();
            }
            
            // Xóa Spacing Line (nếu muốn Bullet khít nhau hơn thì có thể chỉnh ở đây, hiện tại giữ nguyên)
            // if (ppr.isSetSpacing()) ppr.unsetSpacing();
        }
    }

    private static boolean isBulletList(XWPFDocument doc, BigInteger numID) {
        try {
            XWPFNumbering numbering = doc.getNumbering();
            if (numbering == null) return false;
            XWPFNum num = numbering.getNum(numID);
            if (num == null) return false;
            XWPFAbstractNum abstractNum = numbering.getAbstractNum(num.getCTNum().getAbstractNumId().getVal());
            if (abstractNum == null) return false;
            
            CTAbstractNum ctAbstractNum = abstractNum.getAbstractNum();
            if (ctAbstractNum.sizeOfLvlArray() > 0) {
                CTLvl lvl0 = ctAbstractNum.getLvlArray(0);
                if (lvl0.isSetNumFmt()) {
                    STNumberFormat.Enum fmt = lvl0.getNumFmt().getVal();
                    return fmt == STNumberFormat.BULLET;
                }
            }
            return false;
        } catch (Exception e) {
            return false;
        }
    }

    private static int mapMarkerToLevel(String marker) {
        switch (marker) {
            case "-": return 0;
            case "+": return 1;
            case "*": return 2;
            default:  return 0;
        }
    }

    private static void applyNumbering(XWPFParagraph p, BigInteger numID, int level) {
        CTPPr ppr = p.getCTP().isSetPPr() ? p.getCTP().getPPr() : p.getCTP().addNewPPr();
        CTNumPr numPr = ppr.isSetNumPr() ? ppr.getNumPr() : ppr.addNewNumPr();
        
        CTDecimalNumber numIdNode = numPr.isSetNumId() ? numPr.getNumId() : numPr.addNewNumId();
        numIdNode.setVal(numID);

        CTDecimalNumber ilvlNode = numPr.isSetIlvl() ? numPr.getIlvl() : numPr.addNewIlvl();
        ilvlNode.setVal(BigInteger.valueOf(level));
    }

    private static BigInteger setupCustomBulletStyles(XWPFDocument doc) {
        try {
            XWPFNumbering numbering = doc.getNumbering();
            if (numbering == null) numbering = doc.createNumbering();

            CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
            cTAbstractNum.setAbstractNumId(BigInteger.valueOf(9999)); 

            // Cấu hình Indent chuẩn cho từng Level
            // Level 0: Left 720 (1.27cm), Hanging 360
            setupLevel(cTAbstractNum, 0, "-", "Symbol", 720, 360); 
            // Level 1: Left 1440 (2.54cm), Hanging 360
            setupLevel(cTAbstractNum, 1, "+", "Symbol", 1440, 360);
            // Level 2: Left 2160 (3.81cm), Hanging 360
            setupLevel(cTAbstractNum, 2, "•", "Times New Roman", 2160, 360);
            
            XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
            BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
            return numbering.addNum(abstractNumID);

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private static void setupLevel(CTAbstractNum abstractNum, int levelId, String bulletChar, String fontName, int left, int hanging) {
        CTLvl lvl = abstractNum.addNewLvl();
        lvl.setIlvl(BigInteger.valueOf(levelId));

        lvl.addNewNumFmt().setVal(STNumberFormat.BULLET);
        lvl.addNewLvlText().setVal(bulletChar);

        // Thiết lập Indent trong Style gốc
        CTInd ind = lvl.addNewPPr().addNewInd(); 
        ind.setLeft(BigInteger.valueOf(left));
        ind.setHanging(BigInteger.valueOf(hanging));

        CTFonts fonts = lvl.addNewRPr().addNewRFonts();
        fonts.setAscii(fontName);
        fonts.setHAnsi(fontName);
        fonts.setHint(STHint.DEFAULT);
    }
}