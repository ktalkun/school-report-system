package by.tolkun.timetable.config;

public final class StudentTimetableConfig {
    public static final int QTY_LESSONS_PER_FIRST_SHIFT = 6;
    public static final int MAX_QTY_LESSONS_PER_FIRST_SHIFT = 8;
    public static final int QTY_LESSONS_PER_SECOND_SHIFT = 6;
    public static final int MAX_QTY_LESSONS_PER_SECOND_SHIFT = 6;
    public static final int QTY_LESSONS_PER_DAY = QTY_LESSONS_PER_FIRST_SHIFT
            + QTY_LESSONS_PER_SECOND_SHIFT;
    public static final int QTY_SCHOOL_DAYS_PER_WEEK = 5;
    public static final int NUM_OF_FIRST_ROW_WITH_LESSON = 2;
    public static final int NUM_OF_FIRST_COLUMN_WITH_LESSON = 1;
}
