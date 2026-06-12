package vn.vietdoc.vietdoc_assistant;

import java.math.BigInteger;

public class VietDocStyleConfig {
    // Page Margins (Twips)
    public static final BigInteger MARGIN_LEFT = BigInteger.valueOf((long) (3.5 * 567)); // 3.5cm
    public static final BigInteger MARGIN_RIGHT = BigInteger.valueOf((long) (2.0 * 567)); // 2.0cm
    public static final BigInteger MARGIN_TOP = BigInteger.valueOf((long) (2.5 * 567)); // 2.5cm
    public static final BigInteger MARGIN_BOTTOM = BigInteger.valueOf((long) (2.0 * 567)); // 2.0cm

    // Fonts & Sizes
    public static final String FONT_FAMILY = "Times New Roman";
    public static final int FONT_SIZE_BODY = 14;
    public static final int FONT_SIZE_HEADING1 = 14; // Các heading nhỏ hơn tự giảm
    public static final int FONT_SIZE_CAPTION = 12;

    // Spacing & Indentation
    public static final int SPACING_BEFORE = 120; // 0pt
    public static final int SPACING_AFTER = 0; // 6pt = 120 twips
    public static final int LINE_SPACING_MULTIPLIER = 360; // 1.5 lines = 1.5 * 240
    public static final int INDENTATION_FIRST_LINE = 567; // 1cm

    // Table
    public static final BigInteger TABLE_WIDTH = BigInteger.valueOf(9000); // ~16cm

    // Cỡ chữ chuẩn cho từng Level (Theo Nghị định 30 hoặc chuẩn học thuật)
    public static final int FONT_SIZE_HEADING2 = 14; // Đậm
    public static final int FONT_SIZE_HEADING3 = 14; // Đậm, nghiêng hoặc thường
    public static final int FONT_SIZE_HEADING4 = 14; // Bằng văn bản thường
    public static final int FONT_SIZE_NORMAL = 14;

    // Spacing & Indentation (Chuẩn Nghị định 30)
    // Cách đoạn trước 0pt, Cách đoạn sau 6pt

    // Line Spacing: 360 = 1.5 lines. (240 = Single)
    public static final int LINE_SPACING = 360;

    // ===== FACTORY & VALIDATION METHODS =====

    /**
     * Tạo FormattingParameters mặc định từ VietDocStyleConfig
     */
    public static FormattingParameters createDefaultParameters() {
        FormattingParameters params = new FormattingParameters();
        params.setMarginLeft(3.5);
        params.setMarginRight(2.0);
        params.setMarginTop(2.5);
        params.setMarginBottom(2.0);
        params.setFontSizeBody(FONT_SIZE_BODY);
        params.setFontSizeHeading1(FONT_SIZE_HEADING1);
        params.setFontSizeHeading2(FONT_SIZE_HEADING2);
        params.setFontSizeHeading3(FONT_SIZE_HEADING3);
        params.setFontSizeCaption(FONT_SIZE_CAPTION);
        params.setLineSpacing(1.5); // LINE_SPACING 360 = 1.5
        params.setSpacingBefore(SPACING_BEFORE);
        params.setSpacingAfter(SPACING_AFTER);
        params.setIndentationFirstLine(1.0); // 1cm = 567 twips
        params.setAffectTableSize(false);
        params.setItalicCaption(false);
        return params;
    }

}