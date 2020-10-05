package by.tolkun.school.entity;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * Class to represent tab (sheet) of workbook.
 */
public class SpreadsheetTab {

    /**
     * Workbook of excel document.
     */
    private final SpreadsheetWorkbook spreadsheetWorkbook;

    /**
     * Sheet (tab) of workbook.
     */
    private final XSSFSheet sheet;

    /**
     * Constructor with parameters.
     *
     * @param spreadsheetWorkbook the workbook
     * @param title               the title of new sheet
     */
    SpreadsheetTab(SpreadsheetWorkbook spreadsheetWorkbook, String title) {
        this.spreadsheetWorkbook = spreadsheetWorkbook;
        this.sheet = spreadsheetWorkbook.getPoiWorkbook().createSheet(title);
    }

    /**
     * Constructor with parameters.
     *
     * @param spreadsheetWorkbook the workbook
     * @param sheet               the sheet
     */
    SpreadsheetTab(SpreadsheetWorkbook spreadsheetWorkbook, XSSFSheet sheet) {
        this.spreadsheetWorkbook = spreadsheetWorkbook;
        this.sheet = sheet;
    }

    /**
     * Get Poi sheet.
     *
     * @return Poi sheet
     */
    public XSSFSheet getPoiSheet() {
        return sheet;
    }

    /**
     * Register style: return registered style if it exists,
     * create Poi style {@link CellStyle} from {@link SpreadsheetCellStyle}
     * and add it to style map otherwise.
     *
     * @param style the style {@link SpreadsheetCellStyle}
     * @return Poi style {@link CellStyle}
     */
    public CellStyle registerStyle(SpreadsheetCellStyle style) {
        return spreadsheetWorkbook.registerStyle(style);
    }
}
