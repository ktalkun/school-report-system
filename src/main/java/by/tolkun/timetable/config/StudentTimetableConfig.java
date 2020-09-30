package by.tolkun.timetable.config;

import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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
     * The list of week days' names.
     */
    public static final List<String> WEEK_DAYS = Arrays.asList(
            "Понедльник",
            "Вторник",
            "Среда",
            "Четверг",
            "Пятница",
            "Суббота",
            "Воскресенье"
    );

    /**
     * Map of subject to replace.
     */
    public static final Map<String, String> SUBJECTS_TO_REPLACE
            = new HashMap<String, String>() {{
        put("а/а/а", "английский язык");
        put("ин/ин", "информатика");
        put("тр/тр", "трудовое обуч.");
    }};

    public static final int CURRENT_YEAR = 2020;

    public static final String DIRECTOR_NAME = "Коресташова  Е.Н.";

    public static final String CURRENT_HALF_YEAR = "I";

    public static final String TIMETABLE_NAME
            = "Расписание уроков 5-11 классов\n" +
            CURRENT_HALF_YEAR + " полугодие " + CURRENT_YEAR + " / "
            + (CURRENT_YEAR + 1) + " учебного года";

    public static final String TIMETABLE_SIGN = "         УТВЕРЖДАЮ \n" +
            "Директор государственного учреждения образования\n" +
            " «Средняя школа № 49 г.Минска»\n" +
            "\n" +
            "________________" + DIRECTOR_NAME + "\n" +
            "________________" + CURRENT_YEAR + " г.";


    /**
     * The number of first row with lessons.
     */
    public static final int NUM_OF_FIRST_ROW_WITH_LESSON = 2;

    /**
     * The number of first column with lessons.
     */
    public static final int NUM_OF_FIRST_COLUMN_WITH_LESSON = 1;
}
