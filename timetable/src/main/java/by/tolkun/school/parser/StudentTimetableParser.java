package by.tolkun.school.parser;

import by.tolkun.school.config.StudentTimetableConfig;
import by.tolkun.school.entity.SpreadsheetTab;

/**
 * Class to parse tab {@link by.tolkun.school.entity.SpreadsheetTab} into
 * student timetable {@link by.tolkun.school.entity.StudentTimetable}.
 */
public class StudentTimetableParser {

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
}
