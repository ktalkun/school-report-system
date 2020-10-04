package by.tolkun.school.entity;

import org.apache.poi.xssf.usermodel.XSSFCell;

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
}
