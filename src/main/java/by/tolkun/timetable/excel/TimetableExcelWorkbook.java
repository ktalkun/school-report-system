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
        if (getQtyLessonsByShiftAndDayAndClass(numSheet, 1, schoolDay,
                schoolClass) > getQtyLessonsByShiftAndDayAndClass(numSheet,
                2, schoolDay, schoolClass)) {
            return 1;
        }
        return 2;
    }

    public int getQtyLessonsByShiftAndDayAndClass(int numSheet, int shift,
                                                  int schoolDay, int schoolClass) {
        int currentShiftBeginRow = StudentTimetableConfig.NUM_OF_FIRST_ROW_WITH_LESSON
                + schoolDay * StudentTimetableConfig.LESSONS_PER_DAY;
        int currentShiftEndRow = currentShiftBeginRow
                + StudentTimetableConfig.QTY_LESSONS_PER_FIRST_SHIFT;

        if (shift == 2) {
            currentShiftBeginRow = currentShiftEndRow;
            currentShiftEndRow = currentShiftBeginRow
                    + StudentTimetableConfig.QTY_LESSONS_PER_SECOND_SHIFT;
        }

        int qtyLessonPerCurrentShift = 0;
        for (int i = currentShiftBeginRow; i < currentShiftEndRow; i++) {
            if (!workbook.getSheetAt(numSheet)
                    .getRow(i)
                    .getCell(schoolClass)
                    .getStringCellValue()
                    .isEmpty()) {
                qtyLessonPerCurrentShift++;
            }
        }

        return qtyLessonPerCurrentShift;
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
