package by.tolkun.school.entity;

import by.tolkun.school.exception.SpreadsheetException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.util.Calendar;
import java.util.Date;

/**
 * Class to represent cell of spreadsheet cell.
 */
public class SpreadsheetCell {

    /**
     * POI cell of sheet.
     */
    private XSSFCell cell;

    /**
     * Style of cell.
     */
    private SpreadsheetCellStyle style;

    /**
     * Data formatter.
     */
    private final DataFormatter dataFormatter = new DataFormatter();

    /**
     * Constructor with parameters.
     *
     * @param cell the cell of sheet
     */
    public SpreadsheetCell(XSSFCell cell) {
        this.cell = cell;
    }

    /**
     * Get poi cell.
     *
     * @return poi cell
     */
    public XSSFCell getPoiCell() {
        return cell;
    }

    /**
     * Get cell style.
     *
     * @return style of cell
     */
    public SpreadsheetCellStyle getStyle() {
        return style;
    }

    /**
     * Get value of cell. Returns the formatted value of a cell as a String
     * regardless of the cell type. If the Excel format pattern cannot be
     * parsed then the cell value will be formatted using a default format.
     * When passed a null or blank cell, this method will return an empty
     * String (""). Formulas in formula type cells will not be evaluated.
     *
     * @return the formatted cell value as a String
     */
    public String getValue() {
        return dataFormatter.formatCellValue(cell);
    }

    /**
     * Set value of cell.
     *
     * @param value the value is one from {@link String}, {@link Number},
     *              {@link Date}, {@link Calendar}, {@link Boolean},
     *              {@link RichTextString}
     */
    public void setValue(Object value) {
        try {
            if (value == null) {
                cell.setCellValue((String) null);
            } else if (value instanceof String) {
                if (((String) value).startsWith("=")) {
                    cell.setCellFormula(((String) value).substring(1));
                } else {
                    cell.setCellValue((String) value);
                }
            } else if (value instanceof Number) {
                double num = ((Number) value).doubleValue();
                if (Double.isNaN(num) || Double.isInfinite(num)) {
                    cell.setCellValue("");
                } else {
                    cell.setCellValue(num);
                }
            } else if (value instanceof Date) {
                cell.setCellValue((Date) value);
            } else if (value instanceof Calendar) {
                cell.setCellValue((Calendar) value);
            } else if (value instanceof Boolean) {
                cell.setCellValue((Boolean) value);
            } else if (value instanceof RichTextString) {
                cell.setCellValue((RichTextString) value);
            } else {
                throw new SpreadsheetException(
                        String.format("Cannot set a %s [%s] as the" +
                                        " spreadsheet cell content.",
                                value.getClass().getSimpleName(),
                                value.toString()));
            }
        } catch (SpreadsheetException e) {
            e.printStackTrace();
        }
    }
}
