package vn.vietdoc.vietdoc_assistant;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import java.util.List;

public class GeneralFormatter {

    public static void setupPageMargins(XWPFDocument doc) {
        CTSectPr sectPr = doc.getDocument().getBody().getSectPr();
        if (sectPr == null) sectPr = doc.getDocument().getBody().addNewSectPr();
        CTPageMar pageMar = sectPr.getPgMar();
        if (pageMar == null) pageMar = sectPr.addNewPgMar();

        pageMar.setLeft(VietDocStyleConfig.MARGIN_LEFT);
        pageMar.setRight(VietDocStyleConfig.MARGIN_RIGHT);
        pageMar.setTop(VietDocStyleConfig.MARGIN_TOP);
        pageMar.setBottom(VietDocStyleConfig.MARGIN_BOTTOM);
    }

    public static void formatNormalText(XWPFDocument doc, int startIndex) {
        List<XWPFParagraph> paragraphs = doc.getParagraphs();
        
        for (int i = startIndex; i < paragraphs.size(); i++) {
            XWPFParagraph p = paragraphs.get(i);
            String styleID = p.getStyleID();

            // Format các đoạn văn thường (Không phải Heading/Title)
            if (styleID == null || (!styleID.startsWith("Heading") && !styleID.startsWith("Title"))) {
                
                // 1. Set Spacing (BẮT BUỘC PHẢI CÓ LineSpacingRule)
                p.setSpacingBefore(VietDocStyleConfig.SPACING_BEFORE);
                p.setSpacingAfter(VietDocStyleConfig.SPACING_AFTER);
                p.setSpacingBetween(VietDocStyleConfig.LINE_SPACING, LineSpacingRule.AUTO); 

                // 2. Thụt đầu dòng & Căn đều
                p.setFirstLineIndent(VietDocStyleConfig.INDENTATION_FIRST_LINE);
                p.setAlignment(ParagraphAlignment.BOTH); 

                // 3. Set Font
                for (XWPFRun r : p.getRuns()) {
                    r.setFontFamily(VietDocStyleConfig.FONT_FAMILY);
                    r.setFontSize(VietDocStyleConfig.FONT_SIZE_BODY);
                }
            }
        }
    }
}