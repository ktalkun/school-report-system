package by.tolkun.school.entity;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
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
     * Get cell by row number and column number.
     *
     * @param rowNum    the number of row
     * @param columnNum the number of column
     * @return cell
     */
    public SpreadsheetCell getCell(int rowNum, int columnNum) {
        String address = getCellAddress(rowNum, columnNum);
        return getCell(address);
    }

    /**
     * Get cell by cell address if cell exists or create and return new cell
     * otherwise.
     *
     * @param cellAddress the cell address
     * @return cell if it exists or create and return new cell otherwise
     */
    public SpreadsheetCell getOrCreateCell(String cellAddress) {
        SpreadsheetCell cell = getCell(cellAddress);
        if (cell == null) {
            CellReference cellReference = new CellReference(cellAddress);
            int rowNum = cellReference.getRow();
            int columnNum = cellReference.getCol();
            cell = new SpreadsheetCell(this, getOrCreatePoiCell(rowNum,
                    columnNum));
            cells.put(cellAddress, cell);
            recordCellModified(rowNum, columnNum);
        }
        return cell;
    }

    /**
     * Get cell by row number and column number if cell exists or create
     * and return new cell otherwise.
     *
     * @param rowNum    the number of row
     * @param columnNum the number of column
     * @return cell if it exists or create and return new cell otherwise
     */
    public SpreadsheetCell getOrCreateCell(int rowNum, int columnNum) {
        return getOrCreateCell(getCellAddress(rowNum, columnNum));
    }

    /**
     * Get Poi row {@link XSSFRow} if it exists or create and return new row
     * otherwise.
     *
     * @param rowNum the number of row
     * @return Poi row {@link XSSFRow}
     */
    private XSSFRow getOrCreatePoiRow(int rowNum) {
        XSSFRow row = sheet.getRow(rowNum);
        if (row == null) {
            row = sheet.createRow(rowNum);
        }
        return row;
    }

    /**
     * Get Poi cell {@link XSSFCell} if it exists or create and return new cell
     * otherwise.
     *
     * @param rowNum    the number of row
     * @param columnNum the number of column
     * @return Poi cell {@link XSSFCell}
     */
    private XSSFCell getOrCreatePoiCell(int rowNum, int columnNum) {
        XSSFRow row = getOrCreatePoiRow(rowNum);
        XSSFCell cell = row.getCell(columnNum);
        if (cell == null) {
            cell = row.createCell(columnNum);
        }
        return cell;
    }

    /**
     * Record cell. If new cell is created it's necessary to update fields
     * {@code highestModifiedCol} and {@code highestModifiedRow} (update
     * dimensions of tab (sheet)).
     *
     * @param rowNum    the number of row
     * @param columnNum the number of column
     */
    private void recordCellModified(int rowNum, int columnNum) {
        if (columnNum > highestModifiedCol) {
            highestModifiedCol = columnNum;
        }
        if (rowNum > highestModifiedRow) {
            highestModifiedRow = rowNum;
        }
    }

    /**
     * Set value and style of cell by cell address.
     *
     * @param cellAddress the cell address
     * @param content     the content of cell
     * @param style       the style of cell
     */
    public void setValue(String cellAddress, Object content,
                         SpreadsheetCellStyle style) {
        CellReference cellReference = new CellReference(cellAddress);
        SpreadsheetCell cell = getOrCreateCell(cellReference.getRow(),
                cellReference.getCol());
        cell.setValue(content);
        if (style != null) {
            cell.setStyle(style);
        }
    }

    /**
     * Set value of cell by cell address.
     *
     * @param cellAddress the cell address
     * @param content     the content of cell
     */
    public void setValue(String cellAddress, Object content) {
        setValue(cellAddress, content, null);
    }

    /**
     * Set value and style of cell by row number and column number.
     *
     * @param rowNum    the number of row
     * @param columnNum the number of column
     * @param content   the content of cell
     * @param style     the style of cell
     */
    public void setValue(int rowNum, int columnNum, Object content,
                         SpreadsheetCellStyle style) {
        setValue(getCellAddress(rowNum, columnNum), content, style);
    }

    /**
     * Set value of cell by row number and column number.
     *
     * @param rowNum    the number of row
     * @param columnNum the number of column
     * @param content   the content of cell
     */
    public void setValue(int rowNum, int columnNum, Object content) {
        setValue(rowNum, columnNum, content, null);
    }

    /**
     * Set style of cell by cell address.
     *
     * @param cellAddress the cell address
     * @param style       the style of cell
     */
    public void setStyle(String cellAddress, SpreadsheetCellStyle style) {
        getOrCreateCell(cellAddress).setStyle(style);
    }

    /**
     * Set style for diapason of cells by cells' addresses.
     *
     * @param firstCellAddress the cell address of first cell
     * @param lastCellAddress  the cell address of last cell
     * @param style            the style of cell
     */
    public void setStyle(String firstCellAddress, String lastCellAddress,
                         SpreadsheetCellStyle style) {
        CellReference firstReference = new CellReference(firstCellAddress);
        CellReference lastReference = new CellReference(lastCellAddress);
        for (int row = firstReference.getRow();
             row <= lastReference.getRow();
             row++) {
            for (int col = firstReference.getCol();
                 col <= lastReference.getCol();
                 col++) {
                getOrCreateCell(row, col).setStyle(style);
            }
        }
    }

    /**
     * Set style of cell by row number and column number.
     *
     * @param rowNum    the number of cell
     * @param columnNum the number of column
     * @param style     the style of cell
     */
    public void setStyle(int rowNum, int columnNum,
                         SpreadsheetCellStyle style) {
        setStyle(getCellAddress(rowNum, columnNum), style);
    }

    /**
     * Set style for diapason of cells by cells' row numbers and column numbers.
     *
     * @param firstRowNum    the row number of first cell
     * @param firstColumnNum the column number of first cell
     * @param lastRowNum     the row number of last cell
     * @param lastColumnNum  the column number of last cell
     * @param style          the style of cell
     */
    public void setStyle(int firstRowNum, int firstColumnNum,
                         int lastRowNum, int lastColumnNum,
                         SpreadsheetCellStyle style) {
        setStyle(getCellAddress(firstRowNum, firstColumnNum),
                getCellAddress(lastRowNum, lastColumnNum), style);
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
