package by.tolkun.school.entity;

import by.tolkun.school.config.StudentTimetableConfig;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * Class to represent student timetable like a excel sheet.
 */
public class StudentTimetableSheet {
    /**
     * Sheet of Excel document.
     */
    private Sheet sheet;

    /**
     * Constructor with parameters.
     *
     * @param sheet of the student timetable
     */
    public StudentTimetableSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    /**
     * Get sheet of student timetable.
     *
     * @return sheet of the student timetable
     */
    public Sheet getSheet() {
        return sheet;
    }

    /**
     * Set sheet of student timetable.
     *
     * @param sheet of the student timetable in the excel workbook
     */
    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    /**
     * Create and get new font.
     *
     * @return new object of class {@code Font}
     */
    public Font createFont() {
        return sheet.getWorkbook().createFont();

    }

    /**
     * Get physical number of the rows in a sheet.
     *
     * @return physical number of the rows in a sheet
     */
    public int getPhysicalNumberOfRows() {
        return sheet.getPhysicalNumberOfRows();
    }

    /**
     * Get physical number of the columns in a sheet.
     *
     * @return physical number of the columns in a sheet
     */
    public int getPhysicalNumberOfColumns() {
        int maxNumCells = 0;
        for (Row row : sheet) {
            if (maxNumCells < row.getLastCellNum()) {
                maxNumCells = row.getLastCellNum();
            }
        }
        return maxNumCells;
    }

    /**
     * Get row by number. Return existing row, create new and return otherwise.
     *
     * @param rowNum the row number
     * @return row of {@code rowNum} number
     */
    public Row getRow(int rowNum) {
        if (sheet.getRow(rowNum) == null) {
            return sheet.createRow(rowNum);
        }
        return sheet.getRow(rowNum);
    }

    /**
     * Get cell by row and column number. Return existing cell, create new and
     * return otherwise.
     *
     * @param rowNum the row number
     * @param colNum the column number
     * @return cell of from {@code rowNum} row and {@code colNum} column
     */
    public Cell getCell(int rowNum, int colNum) {
        Row row = getRow(rowNum);
        return row.getCell(colNum, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
    }

    /**
     * Remove row by number.
     *
     * @param rowNum the number of the row to remove
     */
    public void removeRow(int rowNum) {
        sheet.shiftRows(rowNum + 1,
                getPhysicalNumberOfRows() - 1, -1);
    }

    /**
     * Remove column by number.
     *
     * @param columnNum the number of the column to remove
     */
    public void removeColumn(int columnNum) {
        sheet.shiftColumns(columnNum + 1,
                getPhysicalNumberOfColumns() - 1, -1);
    }

    /**
     * Insert row by number.
     *
     * @param rowNum the number of the row to insert
     */
    public void insertRow(int rowNum) {
        sheet.shiftRows(rowNum, getPhysicalNumberOfRows() - 1, 1);
    }

    /**
     * Insert column by number.
     *
     * @param columnNum the number of the column to insert
     */
    public void insertColumn(int columnNum) {
        sheet.shiftColumns(columnNum,
                getPhysicalNumberOfColumns() - 1, 1);
    }

    /**
     * Merge cells (create merge region).
     *
     * @param firstRowNum    the first row number of merged region
     * @param firstColumnNum the first column number of merged region
     * @param lastRowNum     the last row number of merged region
     * @param lastColumnNum  the last column number of merged region
     * @return index of this region
     */
    public int mergeCells(int firstRowNum, int firstColumnNum,
                          int lastRowNum, int lastColumnNum) {
        return sheet.addMergedRegion(new CellRangeAddress(firstRowNum, lastRowNum,
                firstColumnNum, lastColumnNum));
    }

    /**
     * Set row style by number of the row.
     *
     * @param rowNum    the number of the row to set style
     * @param cellStyle the style of cells in row with number {@code rowNum}
     */
    public void setRowStyle(int rowNum, CellStyle cellStyle) {
        for (int i = 0; i < getPhysicalNumberOfColumns(); i++) {
            getCell(rowNum, i).setCellStyle(cellStyle);
        }
    }

    /**
     * Set column style by number of the column.
     *
     * @param columnNum the number of the column to set style
     * @param cellStyle the style of cells in column with number
     *                  {@code columnNum}
     */
    public void setColumnStyle(int columnNum, CellStyle cellStyle) {
        for (int i = 0; i < getPhysicalNumberOfRows(); i++) {
            getCell(i, columnNum).setCellStyle(cellStyle);
        }
    }

    /**
     * Set row horizontal alignment for all cells in row.
     *
     * @param rowNum    the number of the row to set horizontal alignment
     * @param alignment the horizontal alignment
     */
    public void setRowHorizontalAlignment(int rowNum,
                                          HorizontalAlignment alignment) {
        for (int i = 0; i < getPhysicalNumberOfColumns(); i++) {
            getUniqueCellStyle(getCell(rowNum, i)).setAlignment(alignment);
        }
    }

    /**
     * Set column horizontal alignment for all cells in column.
     *
     * @param columnNum the number of the column to set horizontal alignment
     * @param alignment the horizontal alignment
     */
    public void setColumnHorizontalAlignment(int columnNum,
                                             HorizontalAlignment alignment) {
        for (int i = 0; i < getPhysicalNumberOfRows(); i++) {
            getUniqueCellStyle(getCell(i, columnNum)).setAlignment(alignment);
        }
    }

    /**
     * Set cell horizontal alignment.
     *
     * @param rowNum    the number of the cell row to set horizontal alignment
     * @param columnNum the number of the cell column to set horizontal alignment
     * @param alignment the horizontal alignment
     */
    public void setCellHorizontalAlignment(int rowNum, int columnNum,
                                           HorizontalAlignment alignment) {
        getUniqueCellStyle(getCell(rowNum, columnNum))
                .setAlignment(alignment);
    }

    /**
     * Set row vertical alignment for all cells in row.
     *
     * @param rowNum    the number of the row to set vertical alignment
     * @param alignment the vertical alignment
     */
    public void setRowVerticalAlignment(int rowNum,
                                        VerticalAlignment alignment) {
        for (int i = 0; i < getPhysicalNumberOfColumns(); i++) {
            getUniqueCellStyle(getCell(rowNum, i))
                    .setVerticalAlignment(alignment);
        }
    }

    /**
     * Set vertical alignment for all cells in column.
     *
     * @param columnNum the number fo the column to set vertical alignment
     * @param alignment the vertical alignment
     */
    public void setColumnVerticalAlignment(int columnNum,
                                           VerticalAlignment alignment) {
        for (int i = 0; i < getPhysicalNumberOfRows(); i++) {
            getUniqueCellStyle(getCell(i, columnNum))
                    .setVerticalAlignment(alignment);
        }
    }

    /**
     * Set cell vertical alignment.
     *
     * @param rowNum    the number of the cell row to set vertical alignment
     * @param columnNum the number of the cell column to vertical alignment
     * @param alignment the vertical alignment
     */
    public void setCellVerticalAlignment(int rowNum, int columnNum,
                                         VerticalAlignment alignment) {
        getUniqueCellStyle(getCell(rowNum, columnNum))
                .setVerticalAlignment(alignment);
    }

    /**
     * Set rotation for text in all cells in the row.
     *
     * @param rowNum   the number of the row to set rotation text
     * @param rotation the rotation of text in degrees
     */
    public void setRowRotation(int rowNum, short rotation) {
        for (int i = 0; i < getPhysicalNumberOfColumns(); i++) {
            getUniqueCellStyle(getCell(rowNum, i)).setRotation(rotation);
        }
    }

    /**
     * Set rotation for text in all cells in the column.
     *
     * @param columnNum the number of the column to set rotation text
     * @param rotation  the rotation of text in degrees
     */
    public void setColumnRotation(int columnNum, short rotation) {
        for (int i = 0; i < getPhysicalNumberOfRows(); i++) {
            getUniqueCellStyle(getCell(i, columnNum)).setRotation(rotation);
        }
    }

    /**
     * Set the top border for the row.
     *
     * @param rowNum      the number of the row to set top border
     * @param borderStyle the style of the top border
     */
    public void setRowBorderTop(int rowNum, BorderStyle borderStyle) {
        for (int i = 0; i < getPhysicalNumberOfColumns(); i++) {
            getUniqueCellStyle(getCell(rowNum, i))
                    .setBorderTop(borderStyle);
        }
    }

    /**
     * Set the bottom border for the row.
     *
     * @param rowNum      the number of the row to set bottom border
     * @param borderStyle the style of the bottom border
     */
    public void setRowBorderBottom(int rowNum, BorderStyle borderStyle) {
        for (int i = 0; i < getPhysicalNumberOfColumns(); i++) {
            getUniqueCellStyle(getCell(rowNum, i))
                    .setBorderBottom(borderStyle);
        }
    }

    /**
     * Set the right border for the column.
     *
     * @param columnNum   the number of the column to set right border
     * @param borderStyle the style of the right border
     */
    public void setColumnBorderRight(int columnNum, BorderStyle borderStyle) {
        for (int i = 0; i < getPhysicalNumberOfRows(); i++) {
            getUniqueCellStyle(getCell(i, columnNum))
                    .setBorderRight(borderStyle);
        }
    }

    /**
     * Set the left border for the column.
     *
     * @param columnNum   the number of the column to set left border
     * @param borderStyle the style of the left border
     */
    public void setColumnBorderLeft(int columnNum, BorderStyle borderStyle) {
        for (int i = 0; i < getPhysicalNumberOfRows(); i++) {
            getUniqueCellStyle(getCell(i, columnNum))
                    .setBorderLeft(borderStyle);
        }
    }

    /**
     * Set the font for the row.
     *
     * @param rowNum the number of the row to set font
     * @param font   the font
     */
    public void setRowFont(int rowNum, Font font) {
        for (int i = 0; i < getPhysicalNumberOfColumns(); i++) {
            getUniqueCellStyle(getCell(rowNum, i)).setFont(font);
        }
    }

    /**
     * Set the font for the column.
     *
     * @param columnNum the number of the column to set font
     * @param font      the font
     */
    public void setColumnFont(int columnNum, Font font) {
        for (int i = 0; i < getPhysicalNumberOfRows(); i++) {
            getUniqueCellStyle(getCell(i, columnNum)).setFont(font);
        }
    }

    /**
     * Set the font of the cell.
     *
     * @param rowNum    the number of the row of cell
     * @param columnNum the number of the column of cell
     * @param font      the font
     */
    public void setCellFont(int rowNum, int columnNum, Font font) {
        getUniqueCellStyle(getCell(rowNum, columnNum)).setFont(font);
    }

    /**
     * Set whether the text should be wrapped. Setting this flag to
     * {@code true} make all content visible within a cell by displaying
     * it on multiple lines.
     *
     * @param rowNum    the number of the row of cell
     * @param columnNum the number of the column of cell
     * @param wrapped   {@code true} if make cell wrapped, {@code false}
     *                  otherwise
     */
    public void setCellWrapText(int rowNum, int columnNum, boolean wrapped) {
        getUniqueCellStyle(getCell(rowNum, columnNum)).setWrapText(wrapped);
    }

    /**
     * Get unique cell style. Style that's not used by others cell.
     *
     * @param cell the cell to get unique style
     * @return existing cell style if it's not used by others cells otherwise
     * create and return new cell style
     */
    private CellStyle getUniqueCellStyle(Cell cell) {
        if (cell.getCellStyle().getIndex() == 0) {
            cell.setCellStyle(sheet.getWorkbook().createCellStyle());
        }
        return cell.getCellStyle();
    }

    /**
     * Get quantity of the lessons per day according to shift and class.
     *
     * @param shift       of the day of a class
     * @param schoolDay   the school day
     * @param schoolClass the school class
     * @return quantity of the lessons per day according to shift and class
     */
    public int getQtyLessonsByShiftAndDayAndClass(int shift, int schoolDay,
                                                  int schoolClass) {
        int currentShiftBeginRow
                = StudentTimetableConfig.NUM_OF_FIRST_ROW_WITH_LESSON
                + schoolDay * StudentTimetableConfig.QTY_LESSONS_PER_DAY;
        int currentShiftEndRow = currentShiftBeginRow
                + StudentTimetableConfig.MAX_QTY_LESSONS_PER_FIRST_SHIFT;

        if (shift == 2) {
            currentShiftBeginRow
                    = StudentTimetableConfig.NUM_OF_FIRST_ROW_WITH_LESSON
                    + schoolDay * StudentTimetableConfig.QTY_LESSONS_PER_DAY
                    + StudentTimetableConfig.QTY_LESSONS_PER_FIRST_SHIFT;
            currentShiftEndRow = currentShiftBeginRow
                    + StudentTimetableConfig.MAX_QTY_LESSONS_PER_SECOND_SHIFT;
        }

        int qtyLessonPerCurrentShift = 0;
        for (int i = currentShiftBeginRow; i < currentShiftEndRow; i++) {
            if (!sheet
                    .getRow(i)
                    .getCell(schoolClass)
                    .getStringCellValue()
                    .isEmpty()) {
                qtyLessonPerCurrentShift++;
            }
        }

        return qtyLessonPerCurrentShift;
    }

    /**
     * Get shift by day and class if can determine by quantities of lessons
     * by the shifts, otherwise return first shift.
     *
     * @param schoolDay   the school day
     * @param schoolClass the school class
     * @return shift according to day and class
     */
    public int getShiftByDayAndClass(int schoolDay, int schoolClass) {
        int qtyLessonsFirstShift = getQtyLessonsByShiftAndDayAndClass(1,
                schoolDay, schoolClass);
        int qtyLessonsSecondShift = getQtyLessonsByShiftAndDayAndClass(2,
                schoolDay, schoolClass);
        if (schoolDay < 0) {
            return 1;
        }

//        If cannot determine shift by quantity of lessons by the shifts.
        if (qtyLessonsFirstShift == qtyLessonsSecondShift) {
            return getShiftByDayAndClass(schoolDay - 1, schoolClass);
        }

        if (qtyLessonsFirstShift > qtyLessonsSecondShift) {
            return 1;
        }

        return 2;
    }

    /**
     * Get list of the lessons by day and class.
     *
     * @param schoolDay   the school day
     * @param schoolClass the school class
     * @return list of the lessons according to day and class
     */
    public List<String> getLessonsByDayAndClass(int schoolDay, int schoolClass) {
        int shift = getShiftByDayAndClass(schoolDay, schoolClass);
        int qtyLessonsPerCurrentShift
                = StudentTimetableConfig.MAX_QTY_LESSONS_PER_FIRST_SHIFT;
        if (shift == 2) {
            qtyLessonsPerCurrentShift
                    = StudentTimetableConfig.MAX_QTY_LESSONS_PER_SECOND_SHIFT;
        }
        int numOfFirstLessonForCurrentDayAndShift
                = StudentTimetableConfig.NUM_OF_FIRST_ROW_WITH_LESSON
                + schoolDay * StudentTimetableConfig.QTY_LESSONS_PER_DAY
                + (shift - 1) * StudentTimetableConfig.QTY_LESSONS_PER_FIRST_SHIFT;

        List<String> lessons = new ArrayList<>();

//        Read all lessons with tilings the window.
        for (int i = numOfFirstLessonForCurrentDayAndShift;
             i < numOfFirstLessonForCurrentDayAndShift
                     + qtyLessonsPerCurrentShift; i++) {
            lessons.add(sheet
                    .getRow(i)
                    .getCell(schoolClass)
                    .getStringCellValue()
                    .trim()
            );
        }

//        Remove empty string from the end of the list
//        tilings the windows will remain.
        int currentLesson = lessons.size() - 1;
        while (currentLesson >= 0 && lessons.get(currentLesson).isEmpty()) {
            currentLesson--;
        }
        lessons = lessons.subList(0, currentLesson + 1);

//        Add lessons that are before begin of second shift.
        if (shift == 2) {
            for (int i = numOfFirstLessonForCurrentDayAndShift - 1;
                 i > numOfFirstLessonForCurrentDayAndShift - 1
                         - StudentTimetableConfig.QTY_LESSONS_PER_FIRST_SHIFT;
                 i--) {
                String lesson = getCell(i, schoolClass)
                        .getStringCellValue()
                        .trim();
                if (!lesson.isEmpty()) {
                    lessons.add("!" + lesson);
                }
            }
        }

        return lessons;
    }

    /**
     * Get the list of a school classes.
     *
     * @return get the list of a school classes
     */
    public List<SchoolClass> getSchoolClasses() {
        List<SchoolClass> schoolClasses = new ArrayList<>();

//        Loop by classes to get list of SchoolClasses.
        for (int numOfCurrentClass = StudentTimetableConfig
                .NUM_OF_FIRST_COLUMN_WITH_LESSON;
             numOfCurrentClass < getPhysicalNumberOfColumns();
             numOfCurrentClass++) {

            List<SchoolDay> schoolDays = new ArrayList<>();

//            Loop by days to get list of SchoolDays.
            for (int numOfCurrentDay = 0;
                 numOfCurrentDay < StudentTimetableConfig.QTY_SCHOOL_DAYS_PER_WEEK;
                 numOfCurrentDay++) {

                int shift = getShiftByDayAndClass(numOfCurrentDay,
                        numOfCurrentClass);
                schoolDays.add(new SchoolDay(getLessonsByDayAndClass(
                        numOfCurrentDay, numOfCurrentClass), shift)
                );
            }

            schoolClasses.add(new SchoolClass(sheet
                    .getRow(StudentTimetableConfig
                            .NUM_OF_FIRST_ROW_WITH_LESSON - 1)
                    .getCell(numOfCurrentClass)
                    .getStringCellValue(), schoolDays)
            );
        }

        return schoolClasses;
    }

    /**
     * Set the width (in units of 1/256th of a character width) for range of
     * columns. The maximum column width for an individual cell is
     * 255 characters.
     *
     * @param firstColumnNum the number of first column
     * @param lastColumnNum  the number of last column
     * @param columnWidth    the width in units of 1/256th of a character width
     */
    public void setColumnsWidth(int firstColumnNum, int lastColumnNum,
                                int columnWidth) {
        for (int i = firstColumnNum; i < lastColumnNum; i++) {
            setColumnWidth(i, columnWidth);
        }
    }

    /**
     * Set the width (in units of 1/256th of a character width)
     * The maximum column width for an individual cell is 255 characters.
     *
     * @param columnNum   the number of a column
     * @param columnWidth the width in units of 1/256th of a character width
     */
    public void setColumnWidth(int columnNum, int columnWidth) {
        sheet.setColumnWidth(columnNum, columnWidth);
    }

    /**
     * Get the width (in units of 1/256th of a character width ) of a column.
     *
     * @param columnNum the number of a column
     * @return width - the width in units of 1/256th of a character width
     */
    public int getColumnWidth(int columnNum) {
        return sheet.getColumnWidth(columnNum);
    }

    /**
     * Set autosize of a row by the number.
     *
     * @param rowNum the number of a row
     */
    public void autoSizeRow(int rowNum) {
        float maxCellHeight = -1;
        for (int columnNum = 0;
             columnNum < getPhysicalNumberOfColumns();
             columnNum++) {
            Cell cell = getCell(rowNum, columnNum);
            if (cell.getCellType() == CellType.STRING) {
                int fontSize = getCellFontSizeInPoints(rowNum, columnNum);
                String value = cell.getStringCellValue();
                int numLines = 1;
                for (int i = 0; i < value.length(); i++) {
                    if (value.charAt(i) == '\n') {
                        numLines++;
                    }
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
            rowHeight = -1;    // resets to the default
        }

        getRow(rowNum).setHeightInPoints(rowHeight);
    }

    /**
     * Set autosize of column by the number.
     *
     * @param columnNum the number of a column
     */
    public void autoSizeColumn(int columnNum) {
        sheet.autoSizeColumn(columnNum);
    }

    /**
     * Set autosize of all rows.
     */
    public void autoSizeAllRows() {
        for (int i = 0; i < getPhysicalNumberOfRows(); i++) {
            autoSizeRow(i);
        }
    }

    /**
     * Set autosize of all columns.
     */
    public void autoSizeAllColumns() {
        for (int i = 0; i < getPhysicalNumberOfColumns(); i++) {
            sheet.autoSizeColumn(i);
        }
    }

    /**
     * Get font size in points of cell.
     *
     * @param rowNum    the number of a row
     * @param columnNum the number of a column
     * @return font size in points of cell
     */
    public int getCellFontSizeInPoints(int rowNum, int columnNum) {
        Cell cell = getCell(rowNum, columnNum);
        int fontIndex = cell.getCellStyle().getFontIndexAsInt();
        Font font = sheet.getWorkbook().getFontAt(fontIndex);
        return font.getFontHeightInPoints();
    }

    /**
     * Compute row height in points by font size and number of lines.
     *
     * @param fontSizeInPoints the font size in points
     * @param numLines         the number of lines
     * @return row height in points by font size and number of lines
     */
    public float computeRowHeightInPoints(int fontSizeInPoints, int numLines) {
        // A crude approximation of what excel does.
        float lineHeightInPoints = 1.4f * fontSizeInPoints;
        float rowHeightInPoints = lineHeightInPoints * numLines;
        // Round to the nearest 0.25.
        rowHeightInPoints = Math.round(rowHeightInPoints * 4) / 4f;

        // Don't shrink rows to fit the font, only grow them
        float defaultRowHeightInPoints = sheet.getDefaultRowHeightInPoints();
        if (rowHeightInPoints < defaultRowHeightInPoints + 1) {
            rowHeightInPoints = defaultRowHeightInPoints;
        }
        return rowHeightInPoints;
    }

    /**
     * Clear all rows.
     */
    public void clearAll() {
        sheet.shiftRows(getPhysicalNumberOfRows(), getPhysicalNumberOfRows() * 2,
                -(getPhysicalNumberOfRows()));
    }

    /**
     * Compares this StudentTimeTableSheet to the specified object. The result
     * is true if and only if the argument is not null and is
     * a StudentTimeTableSheet object and objects are equal.
     *
     * @param o the object to compare this StudentTimeTableSheet against
     * @return true if the given object of StudentTimeTableSheet class
     * equivalent to this student timetable sheet, false otherwise
     */
    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof StudentTimetableSheet)) return false;
        StudentTimetableSheet that = (StudentTimetableSheet) o;
        return Objects.equals(sheet, that.sheet);
    }

    /**
     * Returns a hash code for this StudentTimetableSheet.
     *
     * @return a hash code value for this object
     */
    @Override
    public int hashCode() {
        return Objects.hash(sheet);
    }

    /**
     * Returns the string representation of the StudentTimetableSheet.
     *
     * @return the string representation of the StudentTimetableSheet
     */
    @Override
    public String toString() {
        return "TimetableExcelWorkbook{" +
                "sheet=" + sheet +
                '}';
    }
}
