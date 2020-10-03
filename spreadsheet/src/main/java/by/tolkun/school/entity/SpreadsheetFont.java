package by.tolkun.school.entity;

public class SpreadsheetFont {
    private String fontName;
    private Integer fontOffset;
    private Boolean isBold;
    private Boolean isItalic;
    private Boolean isUnderlined;
    private Boolean isDoubleUnderlined;
    private Boolean isStrikeout;
    private Integer sizeInPoints;

    /**
     * Class to build object of class {@link SpreadsheetFont}.
     */
    public static class Builder {
        private SpreadsheetFont font;

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
        public Builder fontOffset(Integer fontOffset) {
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
        public Builder sizeInPoints(Integer sizeInPoints) {
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
}
