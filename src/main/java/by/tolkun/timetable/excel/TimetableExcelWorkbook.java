package by.tolkun.timetable.excel;

import by.tolkun.timetable.config.StudentTimetableConfig;
import by.tolkun.timetable.entity.SchoolClass;
import by.tolkun.timetable.entity.SchoolDay;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;
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

    public int getShiftByDayAndClass(int numSheet, int schoolDay, int schoolClass) {
        if (getQtyLessonsByShiftAndDayAndClass(numSheet, 1, schoolDay,
                schoolClass) > getQtyLessonsByShiftAndDayAndClass(numSheet,
                2, schoolDay, schoolClass)) {
            return 1;
        }
        return 2;
    }

    public List<String> getLessonsByDayAndClass(int numSheet, int schoolDay,
                                                int schoolClass) {
        int shift = getShiftByDayAndClass(numSheet, schoolDay, schoolClass);
        int qtyLessonsPerCurrentShift
                = StudentTimetableConfig.QTY_LESSONS_PER_FIRST_SHIFT;
        if (shift == 2) {
            qtyLessonsPerCurrentShift
                    = StudentTimetableConfig.QTY_LESSONS_PER_SECOND_SHIFT;
        }
        int numOfFirstLessonForCurrentDayAndShift
                = StudentTimetableConfig.NUM_OF_FIRST_ROW_WITH_LESSON
                + schoolDay * StudentTimetableConfig.LESSONS_PER_DAY
                + (shift - 1) * StudentTimetableConfig.QTY_LESSONS_PER_FIRST_SHIFT;

        List<String> lessons = new ArrayList<>();
        for (int i = numOfFirstLessonForCurrentDayAndShift;
             i < numOfFirstLessonForCurrentDayAndShift
                     + qtyLessonsPerCurrentShift; i++) {
            lessons.add(workbook.getSheetAt(numSheet)
                    .getRow(i)
                    .getCell(schoolClass)
                    .getStringCellValue()
                    .trim()
            );
        }

//        Add lessons that are before begin of second shift.
        if (shift == 2) {
            for (int i = numOfFirstLessonForCurrentDayAndShift - 1;
                 i > numOfFirstLessonForCurrentDayAndShift - 1
                         - StudentTimetableConfig.QTY_LESSONS_PER_FIRST_SHIFT;
                 i--) {
                String lesson = workbook.getSheetAt(numSheet)
                        .getRow(i)
                        .getCell(schoolClass)
                        .getStringCellValue()
                        .trim();
                if (!lesson.isEmpty()) {
                    lessons.add("!" + lesson);
                }
            }
        }

        return lessons;
    }

    public List<SchoolClass> getSchoolClasses(int numSheet) {
        List<SchoolClass> schoolClasses = new ArrayList<>();
        for (int numOfCurrentClass = StudentTimetableConfig.NUM_OF_FIRST_COLUMN_WITH_LESSON;
             numOfCurrentClass < getPhysicalNumberOfColumns(numSheet);
             numOfCurrentClass++) {

            List<SchoolDay> schoolDays = new ArrayList<>();
            for (int numOfCurrentDay = 0;
                 numOfCurrentDay < StudentTimetableConfig.QTY_SCHOOL_DAYS_PER_WEEK;
                 numOfCurrentDay++) {

                int shift = getShiftByDayAndClass(numSheet, numOfCurrentDay,
                        numOfCurrentClass);
                schoolDays.add(new SchoolDay(getLessonsByDayAndClass(numSheet,
                        numOfCurrentDay, numOfCurrentClass), shift));
            }

            schoolClasses.add(new SchoolClass(workbook.getSheetAt(numSheet)
                    .getRow(StudentTimetableConfig.NUM_OF_FIRST_ROW_WITH_LESSON - 1)
                    .getCell(numOfCurrentClass)
                    .getStringCellValue(), schoolDays)
            );
        }
        return schoolClasses;
    }

    public void autoSizeAllColumns(int numSheet) {
        for (int i = 0; i < getPhysicalNumberOfColumns(numSheet); i++) {
            workbook.getSheetAt(numSheet).autoSizeColumn(i);
        }
    }

    public void clearAll(int numSheet) {
        workbook.getSheetAt(numSheet).shiftRows(getPhysicalNumberOfRows(numSheet),
                getPhysicalNumberOfRows(numSheet) * 2,
                -(getPhysicalNumberOfRows(numSheet)));
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
