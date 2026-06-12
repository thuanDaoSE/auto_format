package vn.vietdoc.vietdoc_assistant;

import java.math.BigInteger;

/**
 * DTO để wrap tất cả các tham số tùy chỉnh formatting
 * Các giá trị mặc định được lấy từ VietDocStyleConfig
 */
public class FormattingParameters {

    // ===== MARGINS (cm) - Range: 1.0 - 4.0 =====
    private Double marginLeft;
    private Double marginRight;
    private Double marginTop;
    private Double marginBottom;

    // ===== FONT SIZES (pt) - Range: 10 - 20 =====
    private Integer fontSizeBody;
    private Integer fontSizeHeading1;
    private Integer fontSizeHeading2;
    private Integer fontSizeHeading3;
    private Integer fontSizeCaption;

    // ===== SPACING & INDENTATION =====
    private Double lineSpacing; // Range: 1.0 - 2.0
    private Integer spacingBefore;
    private Integer spacingAfter;
    private Double indentationFirstLine; // Range: 1.0 - 2.0 cm

    // ===== TABLE/IMAGE SETTINGS =====
    private Boolean affectTableSize; // Có áp dụng lên table/image?
    private Boolean italicCaption; // Làm nghiêng caption?

    // ===== Getters & Setters =====

    public Double getMarginLeft() {
        return marginLeft != null ? marginLeft : 3.5;
    }

    public void setMarginLeft(Double marginLeft) {
        this.marginLeft = marginLeft;
    }

    public Double getMarginRight() {
        return marginRight != null ? marginRight : 2.0;
    }

    public void setMarginRight(Double marginRight) {
        this.marginRight = marginRight;
    }

    public Double getMarginTop() {
        return marginTop != null ? marginTop : 2.5;
    }

    public void setMarginTop(Double marginTop) {
        this.marginTop = marginTop;
    }

    public Double getMarginBottom() {
        return marginBottom != null ? marginBottom : 2.0;
    }

    public void setMarginBottom(Double marginBottom) {
        this.marginBottom = marginBottom;
    }

    public Integer getFontSizeBody() {
        return fontSizeBody != null ? fontSizeBody : 14;
    }

    public void setFontSizeBody(Integer fontSizeBody) {
        this.fontSizeBody = fontSizeBody;
    }

    public Integer getFontSizeHeading1() {
        return fontSizeHeading1 != null ? fontSizeHeading1 : 14;
    }

    public void setFontSizeHeading1(Integer fontSizeHeading1) {
        this.fontSizeHeading1 = fontSizeHeading1;
    }

    public Integer getFontSizeHeading2() {
        return fontSizeHeading2 != null ? fontSizeHeading2 : 14;
    }

    public void setFontSizeHeading2(Integer fontSizeHeading2) {
        this.fontSizeHeading2 = fontSizeHeading2;
    }

    public Integer getFontSizeHeading3() {
        return fontSizeHeading3 != null ? fontSizeHeading3 : 14;
    }

    public void setFontSizeHeading3(Integer fontSizeHeading3) {
        this.fontSizeHeading3 = fontSizeHeading3;
    }

    public Integer getFontSizeCaption() {
        return fontSizeCaption != null ? fontSizeCaption : 12;
    }

    public void setFontSizeCaption(Integer fontSizeCaption) {
        this.fontSizeCaption = fontSizeCaption;
    }

    public Double getLineSpacing() {
        return lineSpacing != null ? lineSpacing : 1.5;
    }

    public void setLineSpacing(Double lineSpacing) {
        this.lineSpacing = lineSpacing;
    }

    public Integer getSpacingBefore() {
        return spacingBefore != null ? spacingBefore : 120;
    }

    public void setSpacingBefore(Integer spacingBefore) {
        this.spacingBefore = spacingBefore;
    }

    public Integer getSpacingAfter() {
        return spacingAfter != null ? spacingAfter : 0;
    }

    public void setSpacingAfter(Integer spacingAfter) {
        this.spacingAfter = spacingAfter;
    }

    public Double getIndentationFirstLine() {
        return indentationFirstLine != null ? indentationFirstLine : 1.0;
    }

    public void setIndentationFirstLine(Double indentationFirstLine) {
        this.indentationFirstLine = indentationFirstLine;
    }

    public Boolean isAffectTableSize() {
        return affectTableSize != null ? affectTableSize : false;
    }

    public void setAffectTableSize(Boolean affectTableSize) {
        this.affectTableSize = affectTableSize;
    }

    public Boolean isItalicCaption() {
        return italicCaption != null ? italicCaption : false;
    }

    public void setItalicCaption(Boolean italicCaption) {
        this.italicCaption = italicCaption;
    }

    // ===== CONVERSION HELPERS =====

    /**
     * Chuyển đổi margin từ cm sang Twips (1 cm = 567 Twips)
     */
    public BigInteger getMarginLeftTwips() {
        return BigInteger.valueOf((long) (getMarginLeft() * 567));
    }

    public BigInteger getMarginRightTwips() {
        return BigInteger.valueOf((long) (getMarginRight() * 567));
    }

    public BigInteger getMarginTopTwips() {
        return BigInteger.valueOf((long) (getMarginTop() * 567));
    }

    public BigInteger getMarginBottomTwips() {
        return BigInteger.valueOf((long) (getMarginBottom() * 567));
    }

    /**
     * Chuyển đổi font size từ pt sang half-point (Word XML dùng half-point)
     * 1 pt = 2 half-point
     */
    public int getFontSizeBodyHalfPoint() {
        return getFontSizeBody() * 2;
    }

    public int getFontSizeHeading1HalfPoint() {
        return getFontSizeHeading1() * 2;
    }

    public int getFontSizeHeading2HalfPoint() {
        return getFontSizeHeading2() * 2;
    }

    public int getFontSizeHeading3HalfPoint() {
        return getFontSizeHeading3() * 2;
    }

    public int getFontSizeCaptionHalfPoint() {
        return getFontSizeCaption() * 2;
    }

    /**
     * Chuyển đổi line spacing từ ratio sang XML units
     * 1.0 = 240, 1.5 = 360, 2.0 = 480
     */
    public int getLineSpacingXmlUnits() {
        return (int) (getLineSpacing() * 240);
    }

    /**
     * Chuyển đổi indentation từ cm sang Twips
     */
    public BigInteger getIndentationFirstLineTwips() {
        return BigInteger.valueOf((long) (getIndentationFirstLine() * 567));
    }

    // ===== VALIDATION =====

    /**
     * Validate tất cả các tham số
     * 
     * @throws IllegalArgumentException nếu giá trị ngoài range
     */
    public void validate() {
        // Validate margins
        validateMargin(getMarginLeft(), "marginLeft");
        validateMargin(getMarginRight(), "marginRight");
        validateMargin(getMarginTop(), "marginTop");
        validateMargin(getMarginBottom(), "marginBottom");

        // Validate font sizes
        validateFontSize(getFontSizeBody(), "fontSizeBody");
        validateFontSize(getFontSizeHeading1(), "fontSizeHeading1");
        validateFontSize(getFontSizeHeading2(), "fontSizeHeading2");
        validateFontSize(getFontSizeHeading3(), "fontSizeHeading3");
        validateFontSize(getFontSizeCaption(), "fontSizeCaption");

        // Validate line spacing
        double lineSpace = getLineSpacing();
        if (lineSpace < 1.0 || lineSpace > 2.0) {
            throw new IllegalArgumentException("Line spacing phải từ 1.0 đến 2.0, nhận được: " + lineSpace);
        }

        // Validate indentation
        validateIndentation(getIndentationFirstLine(), "indentationFirstLine");
    }

    private void validateMargin(Double margin, String fieldName) {
        if (margin < 1.0 || margin > 4.0) {
            throw new IllegalArgumentException(fieldName + " phải từ 1.0 - 4.0 cm, nhận được: " + margin);
        }
    }

    private void validateFontSize(Integer fontSize, String fieldName) {
        if (fontSize < 10 || fontSize > 20) {
            throw new IllegalArgumentException(fieldName + " phải từ 10 - 20 pt, nhận được: " + fontSize);
        }
    }

    private void validateIndentation(Double indent, String fieldName) {
        if (indent < 1.0 || indent > 2.0) {
            throw new IllegalArgumentException(fieldName + " phải từ 1.0 - 2.0 cm, nhận được: " + indent);
        }
    }

    @Override
    public String toString() {
        return "FormattingParameters{" +
                "marginLeft=" + getMarginLeft() +
                ", marginRight=" + getMarginRight() +
                ", marginTop=" + getMarginTop() +
                ", marginBottom=" + getMarginBottom() +
                ", fontSizeBody=" + getFontSizeBody() +
                ", fontSizeHeading1=" + getFontSizeHeading1() +
                ", fontSizeHeading2=" + getFontSizeHeading2() +
                ", fontSizeHeading3=" + getFontSizeHeading3() +
                ", fontSizeCaption=" + getFontSizeCaption() +
                ", lineSpacing=" + getLineSpacing() +
                ", indentationFirstLine=" + getIndentationFirstLine() +
                ", affectTableSize=" + isAffectTableSize() +
                ", italicCaption=" + isItalicCaption() +
                '}';
    }
}
