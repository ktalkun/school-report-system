package by.tolkun.school.entity;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

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
    private int rowCount = -1;

    /**
     * Max number of existing column.
     */
    private int columnCount = -1;

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
     * Get count of rows.
     *
     * @return count of rows
     */
    public int getRowCount() {
        return rowCount;
    }

    /**
     * Get count of columns.
     *
     * @return count of columns
     */
    public int getColumnCount() {
        return columnCount;
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
        if (columnNum > columnCount) {
            columnCount = columnNum;
        }
        if (rowNum > rowCount) {
            rowCount = rowNum;
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
     * Print values column down cell by cell using number of first cell
     * row and column.
     *
     * @param rowNum    the number of first cell row
     * @param columnNum the number of first cell column
     * @param style     the style of cell
     * @param values    the values of cells
     * @return index of the next row after the last one written
     */
    public int printDown(int rowNum, int columnNum,
                         SpreadsheetCellStyle style, Object... values) {
        for (int i = 0; i < values.length; i++) {
            setValue(rowNum + i, columnNum, values[i], style);
        }
        return rowNum + values.length;
    }

    /**
     * Print values column down cell by cell using first cell address.
     *
     * @param cellAddress the first cell address
     * @param style       the style of cell
     * @param values      the values of cells
     */
    public void printDown(String cellAddress, SpreadsheetCellStyle style,
                          Object... values) {
        CellReference cellReference = new CellReference(cellAddress);
        printDown(cellReference.getRow(), cellReference.getCol(),
                style, values);
    }

    /**
     * Print values row across cell by cell using number of first cell
     * row and column.
     *
     * @param rowNum    the number of first cell row
     * @param columnNum the number of first cell column
     * @param style     the style of cell
     * @param values    the values of cells
     * @return index of the next row after the last one written
     */
    public int printAcross(int rowNum, int columnNum,
                           SpreadsheetCellStyle style, Object... values) {
        for (int i = 0; i < values.length; i++) {
            setValue(rowNum, columnNum + i, values[i], style);
        }
        return columnNum + values.length;
    }

    /**
     * Print values row across cell by cell using first cell address.
     *
     * @param cellAddress the first cell address
     * @param style       the style of cell
     * @param values      the values of cells
     */
    public void printAcross(String cellAddress, SpreadsheetCellStyle style,
                            Object... values) {
        CellReference cellReference = new CellReference(cellAddress);
        printAcross(cellReference.getRow(), cellReference.getCol(),
                style, values);
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
     * Set top border for cells in {@code rowNum} row from
     * {@code firstColumnNum} to {@code lastColumnNum} column.
     *
     * @param rowNum         the number of row to set top border
     * @param firstColumnNum the number of first column to set top border
     * @param lastColumnNum  the number of last column to set top border
     * @param borderStyle    the border style of cell
     */
    public void setTopBorder(int rowNum, int firstColumnNum, int lastColumnNum,
                             BorderStyle borderStyle) {
        for (int columnNum = firstColumnNum;
             columnNum <= lastColumnNum; columnNum++) {
            getOrCreateCell(rowNum, columnNum)
                    .applyStyle(new SpreadsheetCellStyle.Builder()
                            .topBorderStyle(borderStyle)
                            .build()
                    );
        }
    }

    /**
     * Set top border for cells in {@code rowNum} row.
     *
     * @param rowNum      the number of row to set top border
     * @param borderStyle the border style of cell
     */
    public void setTopBorder(int rowNum, BorderStyle borderStyle) {
        setTopBorder(rowNum, 0, columnCount, borderStyle);
    }

    /**
     * Set bottom border for cells in {@code rowNum} row from
     * {@code firstColumnNum} to {@code lastColumnNum} column.
     *
     * @param rowNum         the number of row to set bottom border
     * @param firstColumnNum the number of first column to set bottom border
     * @param lastColumnNum  the number of last column to set bottom border
     * @param borderStyle    the border style of cell
     */
    public void setBottomBorder(int rowNum, int firstColumnNum,
                                int lastColumnNum, BorderStyle borderStyle) {
        for (int columnNum = firstColumnNum;
             columnNum <= lastColumnNum; columnNum++) {
            getOrCreateCell(rowNum, columnNum)
                    .applyStyle(new SpreadsheetCellStyle.Builder()
                            .bottomBorderStyle(borderStyle)
                            .build()
                    );
        }
    }

    /**
     * Set bottom border for cells in {@code rowNum} row.
     *
     * @param rowNum      the number of row to set bottom border
     * @param borderStyle the border style of cell
     */
    public void setBottomBorder(int rowNum, BorderStyle borderStyle) {
        setBottomBorder(rowNum, 0, columnCount,
                borderStyle);
    }

    /**
     * Set right border for cells in {@code columnNum} column from
     * {@code firstRowNum} to {@code lastRowNum} row.
     *
     * @param columnNum   the number of column to set right border
     * @param firstRowNum the number of first row to set right border
     * @param lastRowNum  the number of last row to set right border
     * @param borderStyle the border style of cell
     */
    public void setRightBorder(int columnNum, int firstRowNum, int lastRowNum,
                               BorderStyle borderStyle) {
        for (int rowNum = firstRowNum; rowNum <= lastRowNum; rowNum++) {
            getOrCreateCell(rowNum, columnNum)
                    .applyStyle(new SpreadsheetCellStyle.Builder()
                            .rightBorderStyle(borderStyle)
                            .build()
                    );
        }
    }

    /**
     * Set right border for cells in {@code columnNum} column.
     *
     * @param columnNum   the number of column to set right border
     * @param borderStyle the border style of cell
     */
    public void setRightBorder(int columnNum, BorderStyle borderStyle) {
        setRightBorder(columnNum, 0, rowCount,
                borderStyle);
    }

    /**
     * Set left border for cells in {@code columnNum} column from
     * {@code firstRowNum} to {@code lastRowNum} row.
     *
     * @param columnNum   the number of column to set left border
     * @param firstRowNum the number of first row to set left border
     * @param lastRowNum  the number of last row to set left border
     * @param borderStyle the border style of cell
     */
    public void setLeftBorder(int columnNum, int firstRowNum, int lastRowNum,
                              BorderStyle borderStyle) {
        for (int rowNum = firstRowNum; rowNum <= lastRowNum; rowNum++) {
            getOrCreateCell(rowNum, columnNum)
                    .applyStyle(new SpreadsheetCellStyle.Builder()
                            .leftBorderStyle(borderStyle)
                            .build()
                    );
        }
    }

    /**
     * Set left border for cells in {@code columnNum} column.
     *
     * @param columnNum   the number of column to set left border
     * @param borderStyle the border style of cell
     */
    public void setLeftBorder(int columnNum, BorderStyle borderStyle) {
        setLeftBorder(columnNum, 0, rowCount, borderStyle);
    }

    /**
     * Set surrounded border for group of cell.
     *
     * @param firstRowNum    the number of first row
     * @param lastRowNum     the number of last row
     * @param firstColumnNum the number of first column
     * @param lastColumnNum  the number of last column
     * @param borderStyle    the border style of cell
     */
    public void setSurroundBorder(int firstRowNum, int firstColumnNum,
                                  int lastRowNum, int lastColumnNum,
                                  BorderStyle borderStyle) {
        setTopBorder(firstRowNum, firstColumnNum, lastColumnNum, borderStyle);
        setBottomBorder(lastRowNum, firstColumnNum, lastColumnNum, borderStyle);
        setLeftBorder(firstRowNum, lastRowNum, firstColumnNum, borderStyle);
        setRightBorder(firstRowNum, lastRowNum, lastColumnNum, borderStyle);
    }

    /**
     * Set surrounded border for group of cell.
     *
     * @param firstCellAddress the address of first cell
     * @param lastCellAddress  the address of last cell
     * @param borderStyle      the border style of cell
     */
    public void setSurroundBorder(String firstCellAddress,
                                  String lastCellAddress,
                                  BorderStyle borderStyle) {
        CellReference firstCellReference = new CellReference(firstCellAddress);
        CellReference lastCellReference = new CellReference(lastCellAddress);
        setSurroundBorder(
                firstCellReference.getRow(), firstCellReference.getCol(),
                lastCellReference.getRow(), lastCellReference.getCol(),
                borderStyle
        );
    }

    /**
     * Insert rows.
     *
     * @param rowNum           the number of row to insert new rows from
     * @param insertedRowCount the count of rows to insert
     */
    public void insertRows(int rowNum, int insertedRowCount) {
        sheet.shiftRows(rowNum, insertedRowCount - 1, insertedRowCount);
    }

    /**
     * Insert row.
     *
     * @param rowNum the number of row to insert new row from
     */
    public void insertRow(int rowNum) {
        insertRows(rowNum, 1);
    }

    /**
     * Insert columns.
     *
     * @param columnNum           the number of column to insert new columns
     *                            from
     * @param insertedColumnCount the count of columns to insert
     */
    public void insertColumns(int columnNum, int insertedColumnCount) {
        sheet.shiftColumns(columnNum, columnCount - 1,
                insertedColumnCount);
    }

    /**
     * Remove diapason of rows by first row and last row number.
     *
     * @param firstRowNum the number of first row
     * @param lastRowNum  the number of last row
     */
    public void removeRows(int firstRowNum, int lastRowNum) {
        int delta = lastRowNum - firstRowNum + 1;
        sheet.shiftRows(lastRowNum + 1, rowCount, delta);
        rowCount -= delta;
    }

    /**
     * Remove row by row number.
     *
     * @param rowNum the number of row
     */
    public void removeRow(int rowNum) {
        removeRows(rowNum, rowNum);
    }

    /**
     * Remove diapason of columns by first column and last column number.
     *
     * @param firstColumnNum the number of first column
     * @param lastColumnNum  the number of last column
     */
    public void removeColumns(int firstColumnNum, int lastColumnNum) {
        int delta = lastColumnNum - firstColumnNum + 1;
        sheet.shiftColumns(lastColumnNum + 1, columnCount,
                delta);
        columnCount -= delta;
    }

    /**
     * Remove column by column number.
     *
     * @param columnNum the number of column
     */
    public void removeColumn(int columnNum) {
        removeColumns(columnNum, columnNum);
    }

    /**
     * Get the row's height measured in twips (1/20th of a point). If the
     * height is not set, the default worksheet value is returned,
     * See {@link XSSFSheet#getDefaultRowHeightInPoints()}.
     *
     * @param rowNum the number of row
     * @return height of row
     */
    public int getRowHeight(int rowNum) {
        return sheet.getRow(rowNum).getHeight();
    }

    /**
     * Set the height in "twips" or 1/20th of a point.
     *
     * @param rowNum the number of row
     * @param height the height in "twips" or 1/20th of a point.
     *               -1 resets to the default height
     */
    public void setRowHeight(int rowNum, short height) {
        sheet.getRow(rowNum).setHeight(height);
    }

    /**
     * Set the height for diapason of rows in "twips" or 1/20th of a point.
     *
     * @param firstRowNum the number of first row
     * @param lastRowNum  the number of last row
     * @param height      the height of rows
     */
    public void setRowsHeight(int firstRowNum, int lastRowNum, short height) {
        for (int rowNum = firstRowNum; rowNum < lastRowNum; rowNum++) {
            setRowHeight(rowNum, height);
        }
    }

    /**
     * Get the actual column width (in units of 1/256th of a character width )
     * Note, the returned value is always greater that getDefaultColumnWidth()
     * because the latter does not include margins. Actual column width
     * measured as the number of characters of the maximum digit width
     * of the numbers 0, 1, 2, ..., 9 as rendered in the normal style's font.
     * There are 4 pixels of margin padding (two on each side), plus 1 pixel
     * padding for the gridlines.
     *
     * @param columnNum the column number
     * @return width - the width in units of 1/256th of a character width
     */
    public int getColumnWidth(int columnNum) {
        return sheet.getColumnWidth(columnNum);
    }

    /**
     * Set the width (in units of 1/256th of a character width).
     * <p>
     * The maximum column width for an individual cell is 255 characters.
     * This value represents the number of characters that can be displayed
     * in a cell that is formatted with the standard font (first font in the
     * workbook).
     * <p>
     * Character width is defined as the maximum digit width of the numbers
     * {@code 0, 1, 2, ... 9} as rendered using the default font (first font
     * in the workbook). Unless you are using a very special font, the default
     * character is '0' (zero), this is true for Arial (default font font
     * in HSSF) and Calibri (default font in XSSF).
     * <p>
     * Please note, that the width set by this method includes 4 pixels
     * of margin padding (two on each side), plus 1 pixel padding
     * for the gridlines (Section 3.3.1.12 of the OOXML spec).
     * This results is a slightly less value of visible characters than passed
     * to this method (approx. 1/2 of a character).
     * <p>
     * To compute the actual number of visible characters, Excel uses
     * the following formula (Section 3.3.1.12 of the OOXML spec):
     * {@code width = Truncate([{Number of Visible Characters} * {Maximum Digit
     * Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256 }.
     * <p>
     * Using the Calibri font as an example, the maximum digit width of 11
     * point font size is 7 pixels (at 96 dpi). If you set a column width to be
     * eight characters wide, e.g. setColumnWidth(columnIndex, 8*256), then the
     * actual value of visible characters (the value shown in Excel) is derived
     * from the following equation: Truncate([numChars*7+5]/7*256)/256 = 8;
     * which gives 7.29.
     *
     * @param columnNum the number of column
     * @param width     the width in units of 1/256th of a character width
     */
    public void setColumnWidth(int columnNum, int width) {
        sheet.setColumnWidth(columnNum, width);
    }

    /**
     * Set the width for diapason of column (in units of 1/256th of a character
     * width).
     *
     * @param firstColumnNum the number of first column
     * @param lastColumnNum  the number of last column
     * @param width          the width of columns
     */
    public void setColumnsWidth(int firstColumnNum, int lastColumnNum,
                                int width) {
        for (int colNum = firstColumnNum; colNum < lastColumnNum; colNum++) {
            setColumnWidth(colNum, width);
        }
    }

    /**
     * Adjusts the row height to fit the contents.
     *
     * @param rowNum the number of row
     */
    public void autoSizeRow(int rowNum) {
        float maxCellHeight = -1;
        for (int columnNum = 0; columnNum <= columnCount; columnNum++) {
            SpreadsheetCell cell = getOrCreateCell(rowNum, columnNum);
            int fontSize = cell.getFontSizeInPoints();
            XSSFCell poiCell = cell.getPoiCell();
            if (poiCell.getCellType() == CellType.STRING) {
                String value = poiCell.getStringCellValue();
                int numLines = 1;
                for (int i = 0; i < value.length(); i++) {
                    if (value.charAt(i) == '\n') numLines++;
                }
                float cellHeight = computeRowHeightInPoints(fontSize, numLines);
                if (cellHeight > maxCellHeight) {
                    maxCellHeight = cellHeight;
                }
            }
        }

        float defaultRowHeightInPoints = sheet.getDefaultRowHeightInPoints();
        float rowHeight = maxCellHeight;
        if (rowHeight < defaultRowHeightInPoints + 1) {
            rowHeight = -1; // resets to the default.
        }

        sheet.getRow(rowNum).setHeightInPoints(rowHeight);
    }

    /**
     * Compute row height in points.
     *
     * @param fontSizeInPoints the font size in points
     * @param numLines         the number of lines
     * @return row height in points
     */
    public float computeRowHeightInPoints(int fontSizeInPoints, int numLines) {
        // A crude approximation of what excel does.
        float lineHeightInPoints = 1.4f * fontSizeInPoints;
        float rowHeightInPoints = lineHeightInPoints * numLines;
        // Round to the nearest 0.25.
        rowHeightInPoints = Math.round(rowHeightInPoints * 4) / 4f;

        // Don't shrink rows to fit the font, only grow them.
        float defaultRowHeightInPoints = sheet.getDefaultRowHeightInPoints();
        if (rowHeightInPoints < defaultRowHeightInPoints + 1) {
            rowHeightInPoints = defaultRowHeightInPoints;
        }
        return rowHeightInPoints;
    }

    /**
     * Adjusts the column width to fit the contents.
     *
     * @param columnNum the number of column
     */
    public void autoSizeColumn(int columnNum) {
        sheet.autoSizeColumn(columnNum, true);
    }

    /**
     * Adjusts the all rows' heights to fit the contents.
     */
    public void autosizeRows() {
        for (int rowNum = 0; rowNum <= rowCount; rowNum++) {
            autoSizeRow(rowNum);
        }
    }

    /**
     * Adjusts the all columns' widths to fit the contents.
     */
    public void autosizeCols() {
        for (int colNum = 0; colNum <= columnCount; colNum++) {
            autoSizeColumn(colNum);
        }
    }

    /**
     * Adjusts the  all rows' heights and all columns' widths to fit
     * the contents.
     */
    public void autosizeRowsAndCols() {
        autosizeCols();
        autosizeRows();
    }

    /**
     * Merge cells by numbers of rows and columns.
     *
     * @param firstRowNum    the number of first row
     * @param firstColumnNum the number of first column
     * @param lastRowNum     the number of last row
     * @param lastColumnNum  the number of last column
     * @param content        the content of merged cell
     * @param style          the style of cell
     */
    public void mergeCells(int firstRowNum, int firstColumnNum,
                           int lastRowNum, int lastColumnNum,
                           Object content, SpreadsheetCellStyle style) {
        setValue(firstRowNum, firstColumnNum, content);
        for (int col = firstColumnNum; col <= lastColumnNum; col++) {
            for (int row = firstRowNum; row <= lastRowNum; row++) {
                setStyle(row, col, style);
            }
        }
        sheet.addMergedRegion(new CellRangeAddress(firstRowNum, lastRowNum,
                firstColumnNum, lastColumnNum));
    }

    /**
     * Merge cells by cells' addresses.
     *
     * @param firstCellAddress the address of first cell
     * @param lastCellAddress  the address of last cell
     * @param content          the content of merged cell
     * @param style            the style of cell
     */
    public void mergeCells(String firstCellAddress, String lastCellAddress,
                           Object content, SpreadsheetCellStyle style) {
        CellReference firstReference = new CellReference(firstCellAddress);
        CellReference lastReference = new CellReference(lastCellAddress);
        mergeCells(firstReference.getRow(), lastReference.getRow(),
                firstReference.getCol(), lastReference.getCol(),
                content, style);
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

    /**
     * Clear all rows.
     */
    public void clearAll() {
        sheet.shiftRows(rowCount, rowCount * 2,
                -rowCount);
    }

    /**
     * Compares this tab to the specified object. The result is {@code true}
     * if and only if the argument is not null and is a {@code SpreadsheetTab}.
     *
     * @param o the object to compare this {@code SpreadsheetTab} against
     * @return {@code true} if the given object represents
     * a {@code SpreadsheetTab} equivalent to this tab, {@code false}
     * otherwise
     */
    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof SpreadsheetTab)) return false;
        SpreadsheetTab that = (SpreadsheetTab) o;
        return rowCount == that.rowCount &&
                columnCount == that.columnCount &&
                Objects.equals(workbook, that.workbook) &&
                Objects.equals(sheet, that.sheet) &&
                Objects.equals(cells, that.cells);
    }

    /**
     * Compute hash code of {@code SpreadsheetTab}.
     *
     * @return a hash code for this tab
     */
    @Override
    public int hashCode() {
        return Objects.hash(workbook, sheet, cells, rowCount,
                columnCount);
    }

    /**
     * Returns the string representation of the {@code SpreadsheetTab}.
     *
     * @return the string representation of the {@code SpreadsheetTab}
     */
    @Override
    public String toString() {
        return "SpreadsheetTab{" +
                "workbook=" + workbook +
                ", sheet=" + sheet +
                ", cells=" + cells +
                ", highestModifiedRow=" + rowCount +
                ", highestModifiedCol=" + columnCount +
                '}';
    }
}
