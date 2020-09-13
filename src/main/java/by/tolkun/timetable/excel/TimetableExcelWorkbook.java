package by.tolkun.timetable.excel;

import config.StudentTimetableConfig;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Objects;

public class TimetableExcelWorkbook {
    private Workbook workbook;


    public TimetableExcelWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public int getPhysicalNumberOfRows(int numSheet) {
        return workbook.getSheetAt(numSheet).getPhysicalNumberOfRows();
    }

    public int getPhysicalNumberOfColumns(int numSheet) {
        int maxNumCells = 0;
        for (Row row : workbook.getSheetAt(numSheet)) {
            if (maxNumCells < row.getLastCellNum()) {
                maxNumCells = row.getLastCellNum();
            }
        }
        return maxNumCells;
    }

    public int getShiftByDayAndClass(int numSheet, int schoolDay, int schoolClass) {
        int firstShiftBeginRow = StudentTimetableConfig.NUM_OF_FIRST_ROW_WITH_LESSON
                + schoolDay * StudentTimetableConfig.LESSONS_PER_DAY;
        int firstShiftEndRow = firstShiftBeginRow
                + StudentTimetableConfig.QTY_LESSONS_PER_FIRST_SHIFT;
        int numLessonsInFirstShift = 0;
        for (int i = firstShiftBeginRow; i < firstShiftEndRow; i++) {
            if (!workbook
                    .getSheetAt(numSheet)
                    .getRow(i)
                    .getCell(schoolClass)
                    .getStringCellValue()
                    .isEmpty()) {
                numLessonsInFirstShift++;
            }
        }

        int secondShiftBeginRow = firstShiftEndRow;
        int secondShiftEndRow = secondShiftBeginRow
                + StudentTimetableConfig.QTY_LESSONS_PER_SECOND_SHIFT;
        int numLessonsInSecondShift = 0;
        for (int i = secondShiftBeginRow; i < secondShiftEndRow; i++) {
            if (!workbook.getSheetAt(numSheet)
                    .getRow(i)
                    .getCell(schoolClass)
                    .getStringCellValue()
                    .isEmpty()) {
                numLessonsInSecondShift++;
            }
        }

        if (numLessonsInFirstShift > numLessonsInSecondShift) {
            return 1;
        }
        return 2;
    }

    public void autoSizeAllColumns(int numSheet) {
        for (int i = 0; i < getPhysicalNumberOfColumns(numSheet); i++) {
            workbook.getSheetAt(numSheet).autoSizeColumn(i);
        }
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof TimetableExcelWorkbook)) return false;
        TimetableExcelWorkbook that = (TimetableExcelWorkbook) o;
        return workbook.equals(that.workbook);
    }

    @Override
    public int hashCode() {
        return Objects.hash(workbook);
    }

    @Override
    public String toString() {
        return "TimetableExcelWorkbook{" +
                "workbook=" + workbook +
                '}';
    }
}
