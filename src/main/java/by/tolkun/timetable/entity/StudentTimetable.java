package by.tolkun.timetable.entity;

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
