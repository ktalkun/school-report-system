package by.tolkun.timetable.entity;

import java.util.List;
import java.util.Objects;

public class SchoolClass {
    private String name;
    private List<SchoolDay> week;

    public SchoolClass(String name, List<SchoolDay> week) {
        this.name = name;
        this.week = week;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<SchoolDay> getWeek() {
        return week;
    }

    public void setWeek(List<SchoolDay> week) {
        this.week = week;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof SchoolClass)) return false;
        SchoolClass that = (SchoolClass) o;
        return Objects.equals(name, that.name) &&
                Objects.equals(week, that.week);
    }

    @Override
    public int hashCode() {
        return Objects.hash(name, week);
    }

    @Override
    public String toString() {
        return "SchoolClass{" +
                "name='" + name + '\'' +
                ", week=" + week +
                '}';
    }
}
