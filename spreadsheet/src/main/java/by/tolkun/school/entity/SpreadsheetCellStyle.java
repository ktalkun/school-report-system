package by.tolkun.school.entity;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class SpreadsheetCellStyle {
    private SpreadsheetFont font;
    private HorizontalAlignment horizontalAlignment;
    private VerticalAlignment verticalAlignment;

    private BorderStyle topBorderStyle;
    private BorderStyle rightBorderStyle;
    private BorderStyle bottomBorderStyle;
    private BorderStyle leftBorderStyle;

    private Color topBorderColor;
    private Color rightBorderColor;
    private Color bottomBorderColor;
    private Color leftBorderColor;

    private String dataFormatString;
    private Color backgroundColor;

    private Boolean isLocked;
    private Boolean isHidden;
    private Boolean isTextWrapped;

    private Integer indention;
    private Integer rotation;
}
