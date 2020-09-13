package by.tolkun.timetable.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Objects;

public class TimetableExcelWorkbook {
    private Workbook workbook;

    public TimetableExcelWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public int getPhysicalNumberOfRows(int numSheet) {
        return workbook.getSheetAt(numSheet).getPhysicalNumberOfRows();
    }

    public int getPhysicalNumberOfColumns(int numSheet) {
        int maxNumCells = 0;
        for (Row row : workbook.getSheetAt(numSheet)) {
            if (maxNumCells < row.getLastCellNum()) {
                maxNumCells = row.getLastCellNum();
            }
        }
        return maxNumCells;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof TimetableExcelWorkbook)) return false;
        TimetableExcelWorkbook that = (TimetableExcelWorkbook) o;
        return workbook.equals(that.workbook);
    }

    @Override
    public int hashCode() {
        return Objects.hash(workbook);
    }

    @Override
    public String toString() {
        return "TimetableExcelWorkbook{" +
                "workbook=" + workbook +
                '}';
    }
}
