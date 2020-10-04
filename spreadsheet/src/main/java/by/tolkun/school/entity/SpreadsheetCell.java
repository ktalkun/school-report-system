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
}
