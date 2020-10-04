package by.tolkun.school.entity;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Class to represent workbook.
 */
public class SpreadsheetWorkbook {

    /**
     * Poi workbook.
     */
    private XSSFWorkbook workbook;

    /**
     * Default constructor.
     */
    public SpreadsheetWorkbook() {
        this(new XSSFWorkbook());
    }

    /**
     * Constructor with parameters.
     *
     * @param workbook the workbook
     */
    public SpreadsheetWorkbook(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }
}
