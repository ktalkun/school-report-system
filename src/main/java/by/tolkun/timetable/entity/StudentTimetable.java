package by.tolkun.timetable.entity;

import java.util.Comparator;
import java.util.List;
import java.util.Objects;

/**
 * Class to represent student timetable like list of school classes. Every
 * school class has list of week days.
 */
public class StudentTimetable {
    /**
     * The list of school classes.
     */
    private List<SchoolClass> schoolClasses;


    /**
     * Constructor with parameters.
     *
     * @param schoolClasses the list of school classes
     */
    public StudentTimetable(List<SchoolClass> schoolClasses) {
        this.schoolClasses = schoolClasses;
    }

    /**
     * Get list of school classes.
     *
     * @return list of school classes
     */
    public List<SchoolClass> getSchoolClasses() {
        return schoolClasses;
    }

    /**
     * Set list of school classes.
     *
     * @param schoolClasses the list of school classes
     */
    public void setSchoolClasses(List<SchoolClass> schoolClasses) {
        this.schoolClasses = schoolClasses;
    }

    /**
     * Get {@code SchoolClass} by the number.
     *
     * @param schoolClassNum the number for school class
     * @return {@code SchooClass} by the number
     */
    public SchoolClass getSchoolClass(int schoolClassNum) {
        return schoolClasses.get(schoolClassNum);
    }

    /**
     * Get max quantity lessons per day among all classes.
     *
     * @param dayNum the number of day
     * @return max quantity lessons per {@code numDay} or -1 if timetable is empty
     */
    public int getQtyMaxLessonsPerDay(int dayNum) {
        return schoolClasses
                .stream()
                .max(Comparator.comparingInt(o -> o.getWeek().get(dayNum)
                        .getLessons().size()))
                .map(schoolClass -> schoolClass.getWeek().get(dayNum)
                        .getLessons().size())
                .orElse(-1);
    }

    /**
     * Compares this StudentTimetable to the specified object. The result
     * is true if and only if the argument is not null and is a StudentTimetable
     * object that represents the same sequence of characters as this object.
     *
     * @param o the object to compare this StudentTimetable against
     * @return true if the given object represents a StudentTimetable
     * equivalent to this StudentTimetable, false otherwise
     */
    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof StudentTimetable)) return false;
        StudentTimetable that = (StudentTimetable) o;
        return schoolClasses.equals(that.schoolClasses);
    }

    /**
     * Returns a hash code for this StudentTimetable.
     *
     * @return a hash code value for this object
     */
    @Override
    public int hashCode() {
        return Objects.hash(schoolClasses);
    }

    /**
     * Returns the string representation of the StudentTimetable.
     *
     * @return the string representation of the StudentTimetable
     */
    @Override
    public String toString() {
        return "StudentTimetable{" +
                "schoolClasses=" + schoolClasses +
                '}';
    }
}
