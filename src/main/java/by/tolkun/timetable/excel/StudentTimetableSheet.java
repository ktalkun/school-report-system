package by.tolkun.timetable.excel;

import by.tolkun.timetable.config.StudentTimetableConfig;
import by.tolkun.timetable.entity.SchoolClass;
import by.tolkun.timetable.entity.SchoolDay;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

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
     * Insert row by number.
     *
     * @param rowNum the number of the row to insert
     */
    public void insertRow(int rowNum) {
        sheet.shiftRows(rowNum, getPhysicalNumberOfRows() - 1, 1);
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
     * Set autosize of all columns.
     */
    public void autoSizeAllColumns() {
        for (int i = 0; i < getPhysicalNumberOfColumns(); i++) {
            sheet.autoSizeColumn(i);
        }
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
