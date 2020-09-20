package by.tolkun.timetable.config;

/**
 * Class to represent properties of student timetable.
 */
public final class StudentTimetableConfig {
    /**
     * The fixed quantity of lessons per first shift.
     */
    public static final int QTY_LESSONS_PER_FIRST_SHIFT = 6;

    /**
     * The max quantity of lessons per first. It includes the fixed
     * quantity of lessons per first shift and lessons that may be at the
     * second shift.
     */
    public static final int MAX_QTY_LESSONS_PER_FIRST_SHIFT = 8;

    /**
     * The fixed quantity of lessons per second shift.
     */
    public static final int QTY_LESSONS_PER_SECOND_SHIFT = 6;

    /**
     * The max quantity of lessons per second shift. It includes the fixed
     * quantity of lessons per second shift and lessons that may be later
     * of the second shift.
     */
    public static final int MAX_QTY_LESSONS_PER_SECOND_SHIFT = 6;

    /**
     * The quantity of lessons per day.
     */
    public static final int QTY_LESSONS_PER_DAY = QTY_LESSONS_PER_FIRST_SHIFT
            + QTY_LESSONS_PER_SECOND_SHIFT;

    /**
     * The quantity of days per week.
     */
    public static final int QTY_SCHOOL_DAYS_PER_WEEK = 5;

    /**
     * The number of first row with lessons.
     */
    public static final int NUM_OF_FIRST_ROW_WITH_LESSON = 2;

    /**
     * The number of first column with lessons.
     */
    public static final int NUM_OF_FIRST_COLUMN_WITH_LESSON = 1;
}
