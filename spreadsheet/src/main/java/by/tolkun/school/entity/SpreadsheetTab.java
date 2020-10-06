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
    private final SpreadsheetWorkbook workbook;

    /**
     * Sheet (tab) of workbook.
     */
    private final XSSFSheet sheet;

    /**
     * Constructor with parameters.
     *
     * @param workbook the workbook
     * @param title               the title of new sheet
     */
    SpreadsheetTab(SpreadsheetWorkbook workbook, String title) {
        this.workbook = workbook;
        this.sheet = workbook.getPoiWorkbook().createSheet(title);
    }

    /**
     * Constructor with parameters.
     *
     * @param workbook the workbook
     * @param sheet               the sheet
     */
    SpreadsheetTab(SpreadsheetWorkbook workbook, XSSFSheet sheet) {
        this.workbook = workbook;
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
        return workbook.registerStyle(style);
    }
}
