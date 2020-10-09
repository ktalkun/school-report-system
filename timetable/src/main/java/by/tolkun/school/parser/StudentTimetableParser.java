package by.tolkun.school.parser;

import by.tolkun.school.config.StudentTimetableConfig;
import by.tolkun.school.entity.SchoolClass;
import by.tolkun.school.entity.SchoolDay;
import by.tolkun.school.entity.SpreadsheetTab;
import by.tolkun.school.entity.StudentTimetable;

import java.util.ArrayList;
import java.util.List;

/**
 * Class to parse tab {@link by.tolkun.school.entity.SpreadsheetTab} into
 * student timetable {@link by.tolkun.school.entity.StudentTimetable}.
 */
public class StudentTimetableParser {

    /**
     * Parse tab of excel workbook to student timetable
     * {@link StudentTimetable}.
     *
     * @param tab the tab (sheet)
     * @return student timetable
     */
    public static StudentTimetable parse(SpreadsheetTab tab) {
        List<SchoolClass> schoolClasses = new ArrayList<>();
        // Loop by classes to get list of SchoolClasses.
        for (int classNum = StudentTimetableConfig
                .NUM_OF_FIRST_COLUMN_WITH_LESSON;
             classNum < tab.getColumnCount();
             classNum++) {

            List<SchoolDay> schoolDays = new ArrayList<>();
            // Loop by days to get list of SchoolDays.
            for (int dayNum = 0;
                 dayNum < StudentTimetableConfig.QTY_SCHOOL_DAYS_PER_WEEK;
                 dayNum++) {
                int shift = parseShift(tab, dayNum, classNum);
                schoolDays.add(new SchoolDay(
                        parseLessons(tab, shift, dayNum, classNum), shift
                ));
            }

            String schoolClassName = tab
                    .getCell(StudentTimetableConfig
                            .NUM_OF_FIRST_ROW_WITH_LESSON - 1, classNum)
                    .getValue();
            schoolClasses.add(new SchoolClass(schoolClassName, schoolDays));
        }

        return new StudentTimetable(schoolClasses);
    }

    /**
     * Parse shift by day and class if can determine by count of lessons
     * at shifts, otherwise return {@code 1} shift.
     *
     * @param tab         the tab (sheet)
     * @param schoolDay   the school day
     * @param schoolClass the school class
     * @return shift according to day and class
     */
    private static int parseShift(SpreadsheetTab tab, int schoolDay,
                                  int schoolClass) {
        if (schoolDay < 0) {
            return 1;
        }

        int lessonFirstShiftCount = countLessonsAtShift(tab,
                1, schoolDay, schoolClass);
        int lessonSecondShiftCount = countLessonsAtShift(tab,
                2, schoolDay, schoolClass);

        // If cannot determine shift by count of lessons by the shifts.
        if (lessonFirstShiftCount == lessonSecondShiftCount) {
            return parseShift(tab, schoolDay - 1, schoolClass);
        }

        if (lessonFirstShiftCount > lessonSecondShiftCount) {
            return 1;
        }

        return 2;
    }

    /**
     * Count lessons per day according to shift, day and class.
     *
     * @param tab         the tab (sheet)
     * @param shift       the shift of school day
     * @param schoolDay   the school day
     * @param schoolClass the school class
     * @return count of the lessons per day according to shift and class
     */
    private static int countLessonsAtShift(SpreadsheetTab tab,
                                           int shift,
                                           int schoolDay,
                                           int schoolClass) {
        int shiftBeginRow
                = StudentTimetableConfig.NUM_OF_FIRST_ROW_WITH_LESSON
                + schoolDay * StudentTimetableConfig.QTY_LESSONS_PER_DAY;
        int shiftEndRow = shiftBeginRow
                + StudentTimetableConfig.MAX_QTY_LESSONS_PER_FIRST_SHIFT;

        if (shift == 2) {
            shiftBeginRow
                    = StudentTimetableConfig.NUM_OF_FIRST_ROW_WITH_LESSON
                    + schoolDay * StudentTimetableConfig.QTY_LESSONS_PER_DAY
                    + StudentTimetableConfig.QTY_LESSONS_PER_FIRST_SHIFT;
            shiftEndRow = shiftBeginRow
                    + StudentTimetableConfig.MAX_QTY_LESSONS_PER_SECOND_SHIFT;
        }

        int lessonPerShiftCount = 0;
        for (int i = shiftBeginRow; i < shiftEndRow; i++) {
            if (!tab.getCell(i, schoolClass).getValue().isEmpty()) {
                lessonPerShiftCount++;
            }
        }

        return lessonPerShiftCount;
    }

    /**
     * Parse list of the lessons by shift, day and class.
     *
     * @param schoolDay   the school day
     * @param schoolClass the school class
     * @return list of the lessons according to day and class
     */
    private static List<String> parseLessons(SpreadsheetTab tab, int shift,
                                             int schoolDay, int schoolClass) {
        int lessonCount
                = StudentTimetableConfig.MAX_QTY_LESSONS_PER_FIRST_SHIFT;
        if (shift == 2) {
            lessonCount
                    = StudentTimetableConfig.MAX_QTY_LESSONS_PER_SECOND_SHIFT;
        }
        int numFirstLesson
                = StudentTimetableConfig.NUM_OF_FIRST_ROW_WITH_LESSON
                + schoolDay * StudentTimetableConfig.QTY_LESSONS_PER_DAY
                + (shift - 1) * StudentTimetableConfig.QTY_LESSONS_PER_FIRST_SHIFT;

        List<String> lessons = new ArrayList<>();
        // Read all lessons with tilings the window.
        for (int i = numFirstLesson; i < numFirstLesson + lessonCount; i++) {
            lessons.add(tab.getCell(i, schoolClass).getValue().trim());
        }
        return lessons;
    }
}
