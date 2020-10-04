package by.tolkun.school.entity;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
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
     * Map of spreadsheet font {@link SpreadsheetFont}
     * and Poi font {@link Font}.
     */
    private final Map<SpreadsheetFont, Font> fontMap = new HashMap<>();

    /**
     * Map of spreadsheet style {@link SpreadsheetCellStyle}
     * and Poi style {@link CellStyle}.
     */
    private final Map<SpreadsheetCellStyle, CellStyle> styleMap
            = new HashMap<>();

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

    /**
     * Register style: create Poi style {@link CellStyle} from
     * {@link SpreadsheetCellStyle} and add it to style map.
     *
     * @param style the style {@link SpreadsheetCellStyle}
     * @return Poi style {@link CellStyle}
     */
    public CellStyle registerStyle(SpreadsheetCellStyle style) {
        CellStyle cellStyle = styleMap.get(style);
        if (cellStyle == null) {
            cellStyle = createNewStyle(style);
            styleMap.put(style, cellStyle);
        }
        return cellStyle;
    }

    /**
     * Create Poi style {@link CellStyle} from spreadsheet style
     * {@link SpreadsheetCellStyle}
     *
     * @param style the style {@link SpreadsheetCellStyle}
     * @return Poi style {@link CellStyle}
     */
    private CellStyle createNewStyle(SpreadsheetCellStyle style) {
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        if (style.getHorizontalAlignment() != null) {
            cellStyle.setAlignment(style.getHorizontalAlignment());
        }
        if (style.getVerticalAlignment() != null) {
            cellStyle.setVerticalAlignment(style.getVerticalAlignment());
        }
        if (style.getTopBorderStyle() != null) {
            cellStyle.setBorderTop(style.getTopBorderStyle());
        }
        if (style.getRightBorderStyle() != null) {
            cellStyle.setBorderLeft(style.getRightBorderStyle());
        }
        if (style.getBottomBorderStyle() != null) {
            cellStyle.setBorderBottom(style.getBottomBorderStyle());
        }
        if (style.getLeftBorderStyle() != null) {
            cellStyle.setBorderRight(style.getLeftBorderStyle());
        }
        if (style.isLocked() != null) {
            cellStyle.setLocked(style.isLocked());
        }
        if (style.isHidden() != null) {
            cellStyle.setHidden(style.isHidden());
        }
        if (style.isTextWrapped() != null) {
            cellStyle.setWrapText(style.isTextWrapped());
        }
        if (style.getIndention() != null) {
            cellStyle.setIndention(style.getIndention());
        }
        if (style.getRotation() != null) {
            cellStyle.setRotation(style.getRotation());
        }
        if (style.getTopBorderColor() != null) {
            cellStyle.setTopBorderColor(XSSFColor
                    .toXSSFColor(style.getTopBorderColor()));
        }
        if (style.getLeftBorderColor() != null) {
            cellStyle.setLeftBorderColor(XSSFColor
                    .toXSSFColor(style.getLeftBorderColor()));
        }
        if (style.getBottomBorderColor() != null)
            cellStyle.setBottomBorderColor(XSSFColor
                    .toXSSFColor(style.getBottomBorderColor()));
        if (style.getRightBorderColor() != null) {
            cellStyle.setRightBorderColor(XSSFColor
                    .toXSSFColor(style.getRightBorderColor()));
        }
        if (style.getFont() != null) {
            cellStyle.setFont(registerFont(style.getFont()));
        }
        if (style.getBackgroundColor() != null) {
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellStyle.setFillForegroundColor(XSSFColor
                    .toXSSFColor(style.getBackgroundColor()));
        }
        if (style.getDataFormatString() != null) {
            cellStyle.setDataFormat(workbook.createDataFormat()
                    .getFormat(style.getDataFormatString()));
        }
        return cellStyle;
    }

    /**
     * Register font: create Poi font {@link Font} from {@link SpreadsheetCell}
     * and add it to font map.
     *
     * @param font the font {@link SpreadsheetFont}
     * @return Poi font {@link Font}
     */
    public Font registerFont(SpreadsheetFont font) {
        Font poiFont = fontMap.get(font);
        if (poiFont == null) {
            poiFont = createNewFont(font);
            fontMap.put(font, poiFont);
        }
        return poiFont;
    }

    /**
     * Create Poi font {@link Font} from spreadsheet font
     * {@link SpreadsheetFont}.
     *
     * @param font the font {@link SpreadsheetFont}
     * @return Poi font {@link Font}
     */
    private Font createNewFont(SpreadsheetFont font) {
        XSSFFont poiFont = workbook.createFont();
        if (font.getFontName() != null) {
            poiFont.setFontName(font.getFontName());
        }
        if (font.getFontOffset() != null) {
            poiFont.setTypeOffset(font.getFontOffset());
        }
        if (font.isBold() != null) {
            poiFont.setBold(font.isBold());
        }
        if (font.isItalic() != null) {
            poiFont.setItalic(font.isItalic());
        }
        if (font.isUnderlined() != null) {
            poiFont.setUnderline(font.isUnderlined() ? Font.U_SINGLE
                    : Font.U_NONE);
        }
        if (font.isDoubleUnderlined() != null)
            poiFont.setUnderline(font.isDoubleUnderlined() ? Font.U_DOUBLE :
                    Font.U_NONE);
        if (font.isStrikeout() != null) {
            poiFont.setStrikeout(font.isStrikeout());
        }
        if (font.getSizeInPoints() != null)
            poiFont.setFontHeightInPoints(font.getSizeInPoints());
        return poiFont;
    }
}
