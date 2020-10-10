package by.tolkun.school.entity;

import java.util.Objects;

/**
 * Class to represent font of cell in spreed sheet.
 */
public class SpreadsheetFont implements Cloneable {

    /**
     * Font name of text in cell.
     */
    private String fontName;

    /**
     * Font offset of text in cell.
     */
    private Short fontOffset;

    /**
     * Is text bold in cell.
     */
    private Boolean isBold;

    /**
     * Is text italic in cell.
     */
    private Boolean isItalic;

    /**
     * Is text underlined in cell.
     */
    private Boolean isUnderlined;

    /**
     * Is test is double underlined in cell.
     */
    private Boolean isDoubleUnderlined;

    /**
     * Is text is strikeout in cell.
     */
    private Boolean isStrikeout;

    /**
     * Text size in point of cell.
     */
    private Short sizeInPoints;

    /**
     * Class to build object of class {@link SpreadsheetFont}.
     */
    public static class Builder {

        /**
         * The instance of outer class for building.
         */
        private final SpreadsheetFont font = new SpreadsheetFont();

        /**
         * Set font name.
         *
         * @param fontName the name of font
         * @return builder
         */
        public Builder fontName(String fontName) {
            font.fontName = fontName;
            return this;
        }

        /**
         * Set font offset.
         *
         * @param fontOffset the offset of font
         * @return builder
         */
        public Builder fontOffset(Short fontOffset) {
            font.fontOffset = fontOffset;
            return this;
        }

        /**
         * Set font bold.
         *
         * @param isBold is {@code true} if font is bold, {@code false}
         *               otherwise
         * @return builder
         */
        public Builder isBold(Boolean isBold) {
            font.isBold = isBold;
            return this;
        }

        /**
         * Set font italic.
         *
         * @param isItalic is {@code true} if font is italic, {@code false}
         *                 otherwise
         * @return builder
         */
        public Builder isItalic(Boolean isItalic) {
            font.isItalic = isItalic;
            return this;
        }

        /**
         * Set font underline.
         *
         * @param isUnderlined is {@code true} if font is underlined,
         *                     {@code false} otherwise
         * @return builder
         */
        public Builder isUnderlined(Boolean isUnderlined) {
            font.isUnderlined = isUnderlined;
            return this;
        }

        /**
         * Set font double underline.
         *
         * @param isDoubleUnderlined is {@code true} if font is double
         *                           underlined, {@code false} otherwise
         * @return builder
         */
        public Builder isDoubleUnderlined(Boolean isDoubleUnderlined) {
            font.isDoubleUnderlined = isDoubleUnderlined;
            return this;
        }

        /**
         * Set font strikeout.
         *
         * @param isStrikeout the {@code true} if font is strikeout,
         *                    {@code false} otherwise
         * @return builder
         */
        public Builder isStrikeout(Boolean isStrikeout) {
            font.isStrikeout = isStrikeout;
            return this;
        }

        /**
         * Set font size in points.
         *
         * @param sizeInPoints the size in points of font
         * @return builder
         */
        public Builder sizeInPoints(Short sizeInPoints) {
            font.sizeInPoints = sizeInPoints;
            return this;
        }

        /**
         * Build font object.
         *
         * @return built font
         */
        public SpreadsheetFont build() {
            return font;
        }
    }

    /**
     * Get font name.
     *
     * @return font name
     */
    public String getFontName() {
        return fontName;
    }

    /**
     * Get font offset.
     *
     * @return font offset
     */
    public Short getFontOffset() {
        return fontOffset;
    }

    /**
     * Check is font bold.
     *
     * @return {@code true} if font is bold, {@code false} otherwise
     */
    public Boolean isBold() {
        return isBold;
    }

    /**
     * Check is font italic.
     *
     * @return {@code true} if font is italic, {@code false} otherwise
     */
    public Boolean isItalic() {
        return isItalic;
    }

    /**
     * Check is font underlined.
     *
     * @return {@code true} if font is underlined, {@code false} otherwise
     */
    public Boolean isUnderlined() {
        return isUnderlined;
    }

    /**
     * Check is font double underlined.
     *
     * @return {@code true} if font is double underlined,
     * {@code false} otherwise
     */
    public Boolean isDoubleUnderlined() {
        return isDoubleUnderlined;
    }

    /**
     * Check is font strikeout.
     *
     * @return {@code true} if font is strikeout, {@code false} otherwise
     */
    public Boolean isStrikeout() {
        return isStrikeout;
    }

    /**
     * Get font size in points.
     *
     * @return size of font in points
     */
    public Short getSizeInPoints() {
        return sizeInPoints;
    }

    /**
     * Get a new font that applies the given font to this one, ignoring
     * all null fields.
     *
     * @param font the font to apply
     * @return new font with applied properties
     */
    public SpreadsheetFont applyFont(SpreadsheetFont font) {
        SpreadsheetFont newFont = clone();
        applyFont(font, newFont);
        return newFont;
    }

    /**
     * Apply font: copy properties of source font to destination font.
     *
     * @param source      the source font
     * @param destination the destination font
     */
    private void applyFont(SpreadsheetFont source,
                           SpreadsheetFont destination) {
        if (source.fontName != null) {
            destination.fontName = source.fontName;
        }
        if (source.fontOffset != null) {
            destination.fontOffset = source.fontOffset;
        }
        if (source.isBold != null) {
            destination.isBold = source.isBold;
        }
        if (source.isItalic != null) {
            destination.isItalic = source.isItalic;
        }
        if (source.isUnderlined != null) {
            destination.isUnderlined = source.isUnderlined;
        }
        if (source.isDoubleUnderlined != null) {
            destination.isDoubleUnderlined = source.isDoubleUnderlined;
        }
        if (source.isStrikeout != null) {
            destination.isStrikeout = source.isStrikeout;
        }
        if (source.sizeInPoints != null)
            destination.sizeInPoints = source.sizeInPoints;
    }

    /**
     * Compares this font to the specified object. The result is {@code true}
     * if and only if the argument is not null and is a {@code SpreadsheetFont}.
     *
     * @param o the object to compare this {@code SpreadsheetFont} against
     * @return {@code true} if the given object represents
     * a {code SpreadsheetFont} equivalent to this string, {@code false}
     * otherwise
     */
    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof SpreadsheetFont)) return false;
        SpreadsheetFont that = (SpreadsheetFont) o;
        return Objects.equals(fontName, that.fontName) &&
                Objects.equals(fontOffset, that.fontOffset) &&
                Objects.equals(isBold, that.isBold) &&
                Objects.equals(isItalic, that.isItalic) &&
                Objects.equals(isUnderlined, that.isUnderlined) &&
                Objects.equals(isDoubleUnderlined, that.isDoubleUnderlined) &&
                Objects.equals(isStrikeout, that.isStrikeout) &&
                Objects.equals(sizeInPoints, that.sizeInPoints);
    }

    /**
     * Compute hash code of {@code SpreadsheetFont}.
     *
     * @return a hash code for this font
     */
    @Override
    public int hashCode() {
        return Objects.hash(fontName, fontOffset, isBold, isItalic,
                isUnderlined, isDoubleUnderlined, isStrikeout, sizeInPoints);
    }

    /**
     * Get clone of the object.
     *
     * @return font
     */
    public SpreadsheetFont clone() {
        SpreadsheetFont font = new SpreadsheetFont();
        applyFont(this, font);
        return font;
    }
}
