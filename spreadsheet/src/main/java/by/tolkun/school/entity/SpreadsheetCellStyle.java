package by.tolkun.school.entity;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.Objects;

/**
 * Class to represent style of cell in spreed sheet.
 */
public class SpreadsheetCellStyle implements Cloneable {

    /**
     * Font of cell.
     */
    private SpreadsheetFont font;

    /**
     * Horizontal alignment of cell.
     */
    private HorizontalAlignment horizontalAlignment;

    /**
     * Vertical alignment of cell.
     */
    private VerticalAlignment verticalAlignment;

    /**
     * Top border style of cell.
     */
    private BorderStyle topBorderStyle;

    /**
     * Right border style of cell.
     */
    private BorderStyle rightBorderStyle;

    /**
     * Bottom border style of cell.
     */
    private BorderStyle bottomBorderStyle;

    /**
     * Left border style of cell.
     */
    private BorderStyle leftBorderStyle;

    /**
     * Top border color of cell.
     */
    private Color topBorderColor;

    /**
     * Right border color of cell.
     */
    private Color rightBorderColor;

    /**
     * Bottom border color of cell.
     */
    private Color bottomBorderColor;

    /**
     * Left border color of cell.
     */
    private Color leftBorderColor;

    /**
     * Data format string of cell.
     */
    private String dataFormatString;

    /**
     * Background color of cell.
     */
    private Color backgroundColor;

    /**
     * Is locked cell.
     */
    private Boolean isLocked;

    /**
     * Is hidden cell.
     */
    private Boolean isHidden;

    /**
     * Is text wrapped in cell.
     */
    private Boolean isTextWrapped;

    /**
     * Indentation of cell.
     */
    private Short indention;

    /**
     * Text rotation int cell.
     */
    private Short rotation;

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
        public Builder indention(Short indention) {
            style.indention = indention;
            return this;
        }

        /**
         * Set text rotation of cell.
         *
         * @param rotation the text rotation of cell
         * @return builder
         */
        public Builder rotation(Short rotation) {
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

    /**
     * Get font.
     *
     * @return font of cell
     */
    public SpreadsheetFont getFont() {
        return font;
    }

    /**
     * Get horizontal alignment.
     *
     * @return horizontal alignment
     */
    public HorizontalAlignment getHorizontalAlignment() {
        return horizontalAlignment;
    }

    /**
     * Get vertical alignment.
     *
     * @return vertical alignment
     */
    public VerticalAlignment getVerticalAlignment() {
        return verticalAlignment;
    }

    /**
     * Get border top style.
     *
     * @return border top style
     */
    public BorderStyle getTopBorderStyle() {
        return topBorderStyle;
    }

    /**
     * Get border right style.
     *
     * @return border right style
     */
    public BorderStyle getRightBorderStyle() {
        return rightBorderStyle;
    }

    /**
     * Get border bottom style.
     *
     * @return border bottom style
     */
    public BorderStyle getBottomBorderStyle() {
        return bottomBorderStyle;
    }

    /**
     * Get border left style.
     *
     * @return left bottom style
     */
    public BorderStyle getLeftBorderStyle() {
        return leftBorderStyle;
    }

    /**
     * Get border top color.
     *
     * @return top bottom color
     */
    public Color getTopBorderColor() {
        return topBorderColor;
    }

    /**
     * Get border right color.
     *
     * @return right bottom color
     */
    public Color getRightBorderColor() {
        return rightBorderColor;
    }

    /**
     * Get border bottom color.
     *
     * @return bottom bottom color
     */
    public Color getBottomBorderColor() {
        return bottomBorderColor;
    }

    /**
     * Get border left color.
     *
     * @return left bottom color
     */
    public Color getLeftBorderColor() {
        return leftBorderColor;
    }

    /**
     * Get data formatting string.
     *
     * @return data formatting string
     */
    public String getDataFormatString() {
        return dataFormatString;
    }

    /**
     * Get background color.
     *
     * @return background color
     */
    public Color getBackgroundColor() {
        return backgroundColor;
    }

    /**
     * Check is cell locked.
     *
     * @return {@code true} if cell is locked, {@code false} otherwise
     */
    public Boolean isLocked() {
        return isLocked;
    }

    /**
     * Check cell is hidden.
     *
     * @return {@code true} if cell is hidden, {@code false} otherwise
     */
    public Boolean isHidden() {
        return isHidden;
    }

    /**
     * Check cell text is wrapped.
     *
     * @return {@code true} if cell text is wrapped, {@code false} otherwise
     */
    public Boolean isTextWrapped() {
        return isTextWrapped;
    }

    /**
     * Get cell indentation.
     *
     * @return indentation
     */
    public Short getIndention() {
        return indention;
    }

    /**
     * Get text rotation of cell in degrees.
     *
     * @return text rotation in degrees
     */
    public Short getRotation() {
        return rotation;
    }

    /**
     * Apply style: copy properties of source style to destination style.
     *
     * @param source      the source style
     * @param destination the destination style
     */
    private void applyStyle(SpreadsheetCellStyle source,
                            SpreadsheetCellStyle destination) {
        if (destination.font == null) {
            destination.font = source.font;
        } else {
            destination.font = destination.font.applyFont(source.font);
        }
        if (source.horizontalAlignment != null) {
            destination.horizontalAlignment = source.horizontalAlignment;
        }
        if (source.verticalAlignment != null) {
            destination.verticalAlignment = source.verticalAlignment;
        }
        if (source.topBorderStyle != null) {
            destination.topBorderStyle = source.topBorderStyle;
        }
        if (source.rightBorderStyle != null) {
            destination.rightBorderStyle = source.rightBorderStyle;
        }
        if (source.bottomBorderStyle != null) {
            destination.bottomBorderStyle = source.bottomBorderStyle;
        }
        if (source.leftBorderStyle != null) {
            destination.leftBorderStyle = source.leftBorderStyle;
        }
        if (source.topBorderColor != null) {
            destination.topBorderColor = source.topBorderColor;
        }
        if (source.rightBorderColor != null) {
            destination.rightBorderColor = source.rightBorderColor;
        }
        if (source.bottomBorderColor != null) {
            destination.bottomBorderColor = source.bottomBorderColor;
        }
        if (source.leftBorderColor != null) {
            destination.leftBorderColor = source.leftBorderColor;
        }
        if (source.dataFormatString != null) {
            destination.dataFormatString = source.dataFormatString;
        }
        if (source.backgroundColor != null) {
            destination.backgroundColor = source.backgroundColor;
        }
        if (source.isLocked != null) {
            destination.isLocked = source.isLocked;
        }
        if (source.isHidden != null) {
            destination.isHidden = source.isHidden;
        }
        if (source.isTextWrapped != null) {
            destination.isTextWrapped = source.isTextWrapped;
        }
        if (source.indention != null) {
            destination.indention = source.indention;
        }
        if (source.rotation != null) {
            destination.rotation = source.rotation;
        }
    }

    /**
     * Compares this cell style to the specified object. The result is
     * {@code true} if and only if the argument is not null and
     * is a {@code SpreadsheetCellStyle}.
     *
     * @param o the object to compare this {@code SpreadsheetCellStyle} against
     * @return {@code true} if the given object represents
     * a {code SpreadsheetCellStyle} equivalent to this cell style, {@code false}
     * otherwise
     */
    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof SpreadsheetCellStyle)) return false;
        SpreadsheetCellStyle that = (SpreadsheetCellStyle) o;
        return Objects.equals(font, that.font) &&
                horizontalAlignment == that.horizontalAlignment &&
                verticalAlignment == that.verticalAlignment &&
                topBorderStyle == that.topBorderStyle &&
                rightBorderStyle == that.rightBorderStyle &&
                bottomBorderStyle == that.bottomBorderStyle &&
                leftBorderStyle == that.leftBorderStyle &&
                Objects.equals(topBorderColor, that.topBorderColor) &&
                Objects.equals(rightBorderColor, that.rightBorderColor) &&
                Objects.equals(bottomBorderColor, that.bottomBorderColor) &&
                Objects.equals(leftBorderColor, that.leftBorderColor) &&
                Objects.equals(dataFormatString, that.dataFormatString) &&
                Objects.equals(backgroundColor, that.backgroundColor) &&
                Objects.equals(isLocked, that.isLocked) &&
                Objects.equals(isHidden, that.isHidden) &&
                Objects.equals(isTextWrapped, that.isTextWrapped) &&
                Objects.equals(indention, that.indention) &&
                Objects.equals(rotation, that.rotation);
    }

    /**
     * Compute hash code of {@code SpreadsheetCellStyle}.
     *
     * @return a hash code for this cell style
     */
    @Override
    public int hashCode() {
        return Objects.hash(font, horizontalAlignment, verticalAlignment,
                topBorderStyle, rightBorderStyle, bottomBorderStyle,
                leftBorderStyle, topBorderColor, rightBorderColor,
                bottomBorderColor, leftBorderColor, dataFormatString,
                backgroundColor, isLocked, isHidden, isTextWrapped, indention,
                rotation);
    }

    /**
     * Get clone of the object.
     *
     * @return style
     */
    public SpreadsheetCellStyle clone() {
        SpreadsheetCellStyle style = new SpreadsheetCellStyle();
        applyStyle(this, style);
        return style;
    }
}
