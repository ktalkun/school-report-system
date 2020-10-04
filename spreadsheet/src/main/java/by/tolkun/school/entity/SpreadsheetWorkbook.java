package by.tolkun.school.entity;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.Map;

/**
 * Class to represent workbook.
 */
public class SpreadsheetWorkbook {

    /**
     * Poi workbook.
     */
    private final XSSFWorkbook workbook;

    /**
     * Map of tabs by index.
     */
    private final Map<Integer, SpreadsheetTab> tabsByIndex = new HashMap<>();

    /**
     * Map of tabs by title.
     */
    private final Map<String, SpreadsheetTab> tabsByTitle = new HashMap<>();

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
        if (workbook.getNumberOfSheets() > 0) {
            for (int index = 0; index < workbook.getNumberOfSheets(); index++) {
                XSSFSheet sheet = workbook.getSheetAt(index);
                createExistingTab(sheet);
            }
        }
    }

    /**
     * Create existing sheet.
     *
     * @param sheet the sheet of workbook
     */
    private void createExistingTab(XSSFSheet sheet) {
        SpreadsheetTab tab = new SpreadsheetTab(this, sheet);
        tabsByTitle.put(sheet.getSheetName(), tab);
        tabsByIndex.put(getPoiWorkbook().getSheetIndex(tab.getPoiSheet()), tab);
    }

    /**
     * Create tab.
     *
     * @param title the title of new tab (sheet)
     * @return tab (sheet)
     */
    public SpreadsheetTab createTab(String title) {
        if (getTab(title) != null) {
            throw new IllegalArgumentException("Workbook already has a sheet"
                    + " with title: " + title);
        }
        // Create tab by creating sheet in poi workbook.
        SpreadsheetTab tab = new SpreadsheetTab(this, title);
        tabsByTitle.put(title, tab);
        tabsByIndex.put(getPoiWorkbook().getSheetIndex(tab.getPoiSheet()), tab);
        return tab;
    }

    /**
     * Get tab by index.
     *
     * @param index the index of tab (sheet)
     * @return tab (sheet)
     */
    public SpreadsheetTab getTab(int index) {
        return tabsByIndex.get(index);
    }

    /**
     * Get tab by title.
     *
     * @param title the title of tab (sheet)
     * @return tab (sheet)
     */
    public SpreadsheetTab getTab(String title) {
        return tabsByTitle.get(title);
    }

    /**
     * Get Poi workbook.
     *
     * @return Poi workbook
     */
    public XSSFWorkbook getPoiWorkbook() {
        return workbook;
    }
}
