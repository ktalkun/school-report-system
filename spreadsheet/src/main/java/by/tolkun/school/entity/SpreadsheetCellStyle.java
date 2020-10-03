package by.tolkun.school.entity;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class SpreadsheetCellStyle {
    private SpreadsheetFont font;
    private HorizontalAlignment horizontalAlignment;
    private VerticalAlignment verticalAlignment;

    private BorderStyle topBorderStyle;
    private BorderStyle rightBorderStyle;
    private BorderStyle bottomBorderStyle;
    private BorderStyle leftBorderStyle;

    private Color topBorderColor;
    private Color rightBorderColor;
    private Color bottomBorderColor;
    private Color leftBorderColor;

    private String dataFormatString;
    private Color backgroundColor;

    private Boolean isLocked;
    private Boolean isHidden;
    private Boolean isTextWrapped;

    private Integer indention;
    private Integer rotation;

    /**
     * Class to build object of class {@link SpreadsheetCellStyle}.
     */
    public static class Builder {
        private SpreadsheetCellStyle style;

        /**
         * Set font.
         *
         * @param font the font of cell
         * @return builder
         */
        public Builder font(SpreadsheetFont font) {
            style.font = font;
            return this;
        }

        /**
         * Set horizontal alignment of cell.
         *
         * @param alignment the horizontal alignment of cell
         * @return builder
         */
        public Builder horizontalAlignment(HorizontalAlignment alignment) {
            style.horizontalAlignment = alignment;
            return this;
        }

        /**
         * Set vertical alignment of cell.
         *
         * @param alignment the vertical alignment of cell
         * @return builder
         */
        public Builder verticalAlignment(VerticalAlignment alignment) {
            style.verticalAlignment = alignment;
            return this;
        }

        /**
         * Set top border style of cell.
         *
         * @param borderStyle the top border style of cell
         * @return builder
         */
        public Builder topBorderStyle(BorderStyle borderStyle) {
            style.topBorderStyle = borderStyle;
            return this;
        }

        /**
         * Set right border style of cell.
         *
         * @param borderStyle the right border style of cell
         * @return builder
         */
        public Builder rightBorderStyle(BorderStyle borderStyle) {
            style.rightBorderStyle = borderStyle;
            return this;
        }

        /**
         * Set bottom border style of cell.
         *
         * @param borderStyle the bottom border style of cell
         * @return builder
         */
        public Builder bottomBorderStyle(BorderStyle borderStyle) {
            style.bottomBorderStyle = borderStyle;
            return this;
        }

        /**
         * Set left border style of cell.
         *
         * @param borderStyle the left border style of cell
         * @return builder
         */
        public Builder leftBorderStyle(BorderStyle borderStyle) {
            style.leftBorderStyle = borderStyle;
            return this;
        }

        /**
         * Set top border color of cell.
         *
         * @param borderColor the top border color of cell
         * @return builder
         */
        public Builder topBorderColor(Color borderColor) {
            style.topBorderColor = borderColor;
            return this;
        }

        /**
         * Set right border color of cell.
         *
         * @param borderColor the right border color of cell
         * @return builder
         */
        public Builder rightBorderColor(Color borderColor) {
            style.rightBorderColor = borderColor;
            return this;
        }

        /**
         * Set bottom border color of cell.
         *
         * @param borderColor the bottom border color of cell
         * @return builder
         */
        public Builder bottomBorderColor(Color borderColor) {
            style.bottomBorderColor = borderColor;
            return this;
        }

        /**
         * Set left border color of cell.
         *
         * @param borderColor the left border color of cell
         * @return builder
         */
        public Builder leftBorderColor(Color borderColor) {
            style.leftBorderColor = borderColor;
            return this;
        }

        /**
         * The data format string of cell.
         *
         * @param dataFormatString the data format string of cell
         * @return builder
         */
        public Builder dataFormatString(String dataFormatString) {
            style.dataFormatString = dataFormatString;
            return this;
        }

        /**
         * Set background color of cell.
         *
         * @param color the background color of cell
         * @return builder
         */
        public Builder backgroundColor(Color color) {
            style.backgroundColor = color;
            return this;
        }

        /**
         * Set lock of cell.
         *
         * @param isLocked {@code true} if cell locked, {@code false}
         *                 otherwise
         * @return builder
         */
        public Builder isLocked(Boolean isLocked) {
            style.isLocked = isLocked;
            return this;
        }

        /**
         * Set cell hidden.
         *
         * @param isHidden {@code true} if cell hidden, {@code false}
         *                 otherwise
         * @return builder
         */
        public Builder isHidden(Boolean isHidden) {
            style.isHidden = isHidden;
            return this;
        }

        /**
         * Set cell text wrap.
         *
         * @param isTextWrapped {@code true} if cell hidden, {@code false}
         *                      otherwise
         * @return builder
         */
        public Builder isTextWrapped(Boolean isTextWrapped) {
            style.isTextWrapped = isTextWrapped;
            return this;
        }

        /**
         * Set indentation of cell.
         *
         * @param indention the indentation of cell
         * @return builder
         */
        public Builder indention(Integer indention) {
            style.indention = indention;
            return this;
        }

        /**
         * Set text rotation of cell.
         *
         * @param rotation the text rotation of cell
         * @return builder
         */
        public Builder rotation(Integer rotation) {
            style.rotation = rotation;
            return this;
        }

        /**
         * Build style object.
         *
         * @return built style
         */
        public SpreadsheetCellStyle build() {
            return style;
        }
    }
}
