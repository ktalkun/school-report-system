package by.tolkun.timetable.entity;

import java.util.List;
import java.util.Objects;

public class SchoolDay {
    private List<String> lessons;
    private int shift;

    public SchoolDay(List<String> lessons, int shift) {
        this.lessons = lessons;
        this.shift = shift;
    }

    public List<String> getLessons() {
        return lessons;
    }

    public void setLessons(List<String> lessons) {
        this.lessons = lessons;
    }

    public String getLesson(int i) {
        return lessons.get(i);
    }

    public int getShift() {
        return shift;
    }

    public void setShift(int shift) {
        this.shift = shift;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        SchoolDay schoolDay = (SchoolDay) o;
        return shift == schoolDay.shift &&
                Objects.equals(lessons, schoolDay.lessons);
    }

    @Override
    public int hashCode() {
        return Objects.hash(lessons, shift);
    }

    @Override
    public String toString() {
        return "SchoolDay{" +
                "lessons=" + lessons +
                ", shift=" + shift +
                '}';
    }
}
