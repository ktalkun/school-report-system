package by.tolkun.school.entity;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.HashMap;
import java.util.Map;

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
     * Map of address and cell.
     */
    private final Map<String, SpreadsheetCell> cells = new HashMap<>();

    /**
     * Max number of existing row.
     */
    private int highestModifiedRow = -1;

    /**
     * Max number of existing column.
     */
    private int highestModifiedCol = -1;

    /**
     * Constructor with parameters.
     *
     * @param workbook the workbook
     * @param title    the title of new sheet
     */
    SpreadsheetTab(SpreadsheetWorkbook workbook, String title) {
        this.workbook = workbook;
        this.sheet = workbook.getPoiWorkbook().createSheet(title);
    }

    /**
     * Constructor with parameters.
     *
     * @param workbook the workbook
     * @param sheet    the sheet
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

    /**
     * Get cell by cell address.
     *
     * @param cellAddress the cell address
     * @return cell
     */
    public SpreadsheetCell getCell(String cellAddress) {
        CellReference cellReference = new CellReference(cellAddress);
        return cells.get(cellAddress);
    }

    /**
     * Get cell address.
     *
     * @param rowNum    the number of row
     * @param columnNum the number of column
     * @return cell address as string
     */
    public static String getCellAddress(int rowNum, int columnNum) {
        return CellReference.convertNumToColString(columnNum) + (rowNum + 1);
    }
}
