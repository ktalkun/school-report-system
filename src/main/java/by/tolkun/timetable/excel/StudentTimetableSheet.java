package by.tolkun.timetable.excel;

import by.tolkun.timetable.config.StudentTimetableConfig;
import by.tolkun.timetable.entity.SchoolClass;
import by.tolkun.timetable.entity.SchoolDay;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class StudentTimetableSheet {
    private Sheet sheet;

    public StudentTimetableSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public Sheet getSheet() {
        return sheet;
    }

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
     * Get quantity of the lessons per day according to shift and class.
     *
     * @param shift of the day of a class
     * @param schoolDay the school day
     * @param schoolClass the school class
     * @return quantity of the lessons per day according to shift and class
     */
    public int getQtyLessonsByShiftAndDayAndClass(int shift, int schoolDay,
                                                  int schoolClass) {
        int currentShiftBeginRow
                = StudentTimetableConfig.NUM_OF_FIRST_ROW_WITH_LESSON
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
     * Get shift by day and class.
     *
     * @param schoolDay the school day
     * @param schoolClass the school class
     * @return shift according to day and class
     */
    public int getShiftByDayAndClass(int schoolDay, int schoolClass) {
        if (getQtyLessonsByShiftAndDayAndClass(1, schoolDay, schoolClass)
                > getQtyLessonsByShiftAndDayAndClass(2, schoolDay,
                schoolClass)) {
            return 1;
        }
        return 2;
    }

    /**
     * Get list of the lessons by day and class.
     *
     * @param schoolDay the school day
     * @param schoolClass the school class
     * @return list of the lessons according to day and class
     */
    public List<String> getLessonsByDayAndClass(int schoolDay, int schoolClass) {
        int shift = getShiftByDayAndClass(schoolDay, schoolClass);
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
                String lesson = sheet
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
//        Loop by classes.
        for (int numOfCurrentClass = StudentTimetableConfig.NUM_OF_FIRST_COLUMN_WITH_LESSON;
             numOfCurrentClass < getPhysicalNumberOfColumns(numSheet);
             numOfCurrentClass++) {

            List<SchoolDay> schoolDays = new ArrayList<>();
//            Loop by days.
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
        if (!(o instanceof StudentTimetableSheet)) return false;
        StudentTimetableSheet that = (StudentTimetableSheet) o;
        return Objects.equals(sheet, that.sheet);
    }

    @Override
    public int hashCode() {
        return Objects.hash(sheet);
    }

    @Override
    public String toString() {
        return "TimetableExcelWorkbook{" +
                "sheet=" + sheet +
                '}';
    }
}
