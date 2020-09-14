package by.tolkun.timetable.entity;

import java.util.Comparator;
import java.util.List;
import java.util.Objects;

public class StudentTimetable {
    private List<SchoolClass> schoolClasses;

    public StudentTimetable(List<SchoolClass> schoolClasses) {
        this.schoolClasses = schoolClasses;
    }

    public List<SchoolClass> getSchoolClasses() {
        return schoolClasses;
    }

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
     * @param numDay the number of day
     * @return max quantity lessons per {@code numDay} or -1 if timetable is empty
     */
    public int getQtyMaxLessonsPerDay(int numDay) {
        return schoolClasses
                .stream()
                .max(Comparator.comparingInt(o -> o.getWeek().get(numDay)
                        .getLessons().size()))
                .map(schoolClass -> schoolClass.getWeek().get(numDay)
                        .getLessons().size())
                .orElse(-1);
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof StudentTimetable)) return false;
        StudentTimetable that = (StudentTimetable) o;
        return schoolClasses.equals(that.schoolClasses);
    }

    @Override
    public int hashCode() {
        return Objects.hash(schoolClasses);
    }

    @Override
    public String toString() {
        return "StudentTimetable{" +
                "schoolClasses=" + schoolClasses +
                '}';
    }
}
