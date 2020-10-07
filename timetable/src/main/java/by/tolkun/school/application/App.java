package by.tolkun.school.application;

import by.tolkun.school.config.StudentTimetableConfig;
import by.tolkun.school.entity.SchoolClass;
import by.tolkun.school.entity.StudentTimetable;
import by.tolkun.school.entity.StudentTimetableSheet;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class App {

    public static void main(String[] args) {
        File file = new File("src/main/resources/timetable.xlsx");
        try (FileInputStream fis = new FileInputStream(file)) {
            Workbook workbook = new XSSFWorkbook(fis);
            if (file.isFile() && file.exists()) {
                System.out.println("File is opened.");
            } else {
                System.out.println("File either isn't existed of can't be opened!");
            }

            StudentTimetableSheet studentTimetableSheet = new StudentTimetableSheet(workbook.getSheetAt(0));


//            Sheet sheet = workbook.getSheetAt(0);

            System.out.println("Reading and parsing timetable into the classes list...");
//            Read file and filling list of school classes.
            List<SchoolClass> schoolClasses = studentTimetableSheet.getSchoolClasses();
            StudentTimetable studentTimetable = new StudentTimetable(schoolClasses);

//            System.out.println(schoolClasses.size());
//            for (SchoolClass schoolClass : schoolClasses) {
//                System.out.println(schoolClass);
//            }

            System.out.println("Clearing all...");
            studentTimetableSheet.clearAll();

            List<Integer> sumOfMaxQtyLessonsPerDaysBeforeList = new ArrayList<>();
            int currentMaxLessonsPerDaySum = 0;
            for (int i = 0; i < StudentTimetableConfig.QTY_SCHOOL_DAYS_PER_WEEK; i++) {
                int maxLessonsPerCurrentDay = studentTimetable.getQtyMaxLessonsPerDay(i);
                if (maxLessonsPerCurrentDay == -1) {
                    throw new Exception("StudentTimetable is empty!");
                }
                sumOfMaxQtyLessonsPerDaysBeforeList.add(i, currentMaxLessonsPerDaySum);
                currentMaxLessonsPerDaySum += maxLessonsPerCurrentDay;
            }

            System.out.println("Writing lessons of classes...");
//            Write timetable.
//            Loop by classes.
            Font zeroLessonFont = studentTimetableSheet.createFont();
            zeroLessonFont.setItalic(true);
            for (int i = 0; i < schoolClasses.size(); i++) {
//                Loop by days.
                for (int j = 0; j < schoolClasses.get(i).getWeek().size(); j++) {
//                    Loop by lessons.
                    for (int k = 0;
                         k < schoolClasses.get(i).getSchoolDay(j).getLessons().size();
                         k++) {
                        int rowNum = sumOfMaxQtyLessonsPerDaysBeforeList.get(j) + k;
                        String currentLesson = schoolClasses.get(i)
                                .getSchoolDay(j).getLesson(k);
                        if (currentLesson.contains("!")) {
                            currentLesson = currentLesson.substring(1);
                            studentTimetableSheet
                                    .setCellFont(rowNum, i, zeroLessonFont);
                        }
                        if (StudentTimetableConfig.SUBJECTS_TO_REPLACE
                                .get(currentLesson) != null) {
                            currentLesson = StudentTimetableConfig
                                    .SUBJECTS_TO_REPLACE.get(currentLesson);
                        }
                        studentTimetableSheet.getCell(rowNum, i)
                                .setCellValue(currentLesson);
                    }
                }
            }

            System.out.println("Inserting class names...");
//            Insert row before first row.
            studentTimetableSheet.insertRow(0);
            Font boldFont = studentTimetableSheet.createFont();
            boldFont.setBold(true);
            studentTimetableSheet.setRowFont(0, boldFont);
//            Add classes' names.
            for (int i = 0; i < studentTimetable.getSchoolClasses().size(); i++) {
                studentTimetableSheet
                        .getCell(0, i)
                        .setCellValue(studentTimetable.getSchoolClass(i).getName());
            }

//            Insert column before first column to add numbers of lessons.
            System.out.println("Inserting numbers of lessons...");
            studentTimetableSheet.insertColumn(0);
//            Set font for first row and first column.
            studentTimetableSheet.setColumnFont(0, boldFont);
//            Set center alignment for column with numbers of lesson.
            System.out.println("Setting center alignment for first column...");
            studentTimetableSheet.setColumnHorizontalAlignment(0,
                    HorizontalAlignment.CENTER);
//            Add numbers of lessons.
//            Loop by days.
            for (int i = 0; i < StudentTimetableConfig.QTY_SCHOOL_DAYS_PER_WEEK; i++) {
//                Loop by lessons of day.
                int rowOfFirstLessonCurrentDayNum = sumOfMaxQtyLessonsPerDaysBeforeList.get(i) + 1;
                int maxQtyLessonPerCurrentDay = studentTimetable.getQtyMaxLessonsPerDay(i);
                for (int j = rowOfFirstLessonCurrentDayNum;
                     j < rowOfFirstLessonCurrentDayNum + maxQtyLessonPerCurrentDay;
                     j++) {
                    studentTimetableSheet
                            .getCell(j, 0)
                            .setCellValue(j - rowOfFirstLessonCurrentDayNum + 1);
                }
            }

//            Insert column before first column to add days' names.
            studentTimetableSheet.insertColumn(0);
            studentTimetableSheet.setColumnBorderRight(0,
                    BorderStyle.MEDIUM);
            studentTimetableSheet.setColumnVerticalAlignment(0,
                    VerticalAlignment.CENTER);
            studentTimetableSheet.setColumnHorizontalAlignment(0,
                    HorizontalAlignment.CENTER);
            Font weekDaysFont = studentTimetableSheet.createFont();
            weekDaysFont.setFontHeight((short) (16 * 20));
            studentTimetableSheet.setColumnFont(0, weekDaysFont);
            studentTimetableSheet
                    .setColumnRotation(0, (short) 90);
            for (int i = 0; i < sumOfMaxQtyLessonsPerDaysBeforeList.size() - 1; i++) {
                studentTimetableSheet.mergeCells(
                        sumOfMaxQtyLessonsPerDaysBeforeList.get(i) + 1,
                        0,
                        sumOfMaxQtyLessonsPerDaysBeforeList.get(i + 1),
                        0
                );
            }
            studentTimetableSheet.mergeCells(
                    sumOfMaxQtyLessonsPerDaysBeforeList
                            .get(sumOfMaxQtyLessonsPerDaysBeforeList.size() - 1) + 1,
                    0,
                    studentTimetableSheet.getPhysicalNumberOfRows() - 1,
                    0
            );
            for (int i = 0;
                 i < StudentTimetableConfig.QTY_SCHOOL_DAYS_PER_WEEK;
                 i++) {
                studentTimetableSheet.getCell(
                        sumOfMaxQtyLessonsPerDaysBeforeList.get(i) + 1,
                        0)
                        .setCellValue(StudentTimetableConfig.WEEK_DAYS.get(i));
            }

//            Set rows' medium borders by days.
            System.out.println("Setting rows' borders...");
            for (int rowNum : sumOfMaxQtyLessonsPerDaysBeforeList) {
                studentTimetableSheet.setRowBorderTop(rowNum + 1,
                        BorderStyle.MEDIUM);
            }

//            Set columns' thin borders by classes
            System.out.println("Setting columns' borders...");
            for (int i = 0; i < schoolClasses.size(); i++) {
                studentTimetableSheet.setColumnBorderLeft(i + 2,
                        BorderStyle.THIN);
            }
//            Get list of first column of parallels.
            List<Integer> firstColumnInParallelNum = new ArrayList<>();
            int currentParallelNum = 0;
            for (int i = 0; i < schoolClasses.size(); i++) {
                String currentSchoolClassName = schoolClasses.get(i).getName();
                int schoolClassNum = Integer.parseInt(currentSchoolClassName
                        .substring(0, currentSchoolClassName.length() - 1));
                if (currentParallelNum != schoolClassNum) {
                    currentParallelNum = schoolClassNum;
                    firstColumnInParallelNum.add(i + 2);
                }
            }
//            Set columns' medium borders by parallels.
            for (int columnNum : firstColumnInParallelNum) {
                studentTimetableSheet.setColumnBorderLeft(columnNum,
                        BorderStyle.MEDIUM);
            }


            System.out.println("Setting autosize for all columns...");
            studentTimetableSheet.autoSizeAllColumns();

//            Insert header of timetable.
            studentTimetableSheet.insertRow(0);
            studentTimetableSheet.setRowBorderBottom(0,
                    BorderStyle.MEDIUM);
            int bottomSectionQtySchoolClasses = schoolClasses.size() / 2;
            int topSectionQtySchoolClasses = schoolClasses.size() + 2
                    - bottomSectionQtySchoolClasses;

            int sidePartHeaderQtyColumns = topSectionQtySchoolClasses / 4;
            int centerPartHeaderQtyColumns = topSectionQtySchoolClasses
                    - 2 * sidePartHeaderQtyColumns;

            int beginColumnLeftPartHeaderNum = 0;
            int endColumnLeftPartHeaderNum = sidePartHeaderQtyColumns - 1;
            studentTimetableSheet
                    .mergeCells(0, beginColumnLeftPartHeaderNum,
                            0, endColumnLeftPartHeaderNum);

            int beginColumnCenterPartHeaderNum = endColumnLeftPartHeaderNum + 1;
            int endColumnCenterPartHeaderNum = beginColumnCenterPartHeaderNum
                    + centerPartHeaderQtyColumns - 1;
            studentTimetableSheet
                    .mergeCells(0, beginColumnCenterPartHeaderNum,
                            0, endColumnCenterPartHeaderNum);
            studentTimetableSheet.setCellHorizontalAlignment(0,
                    beginColumnCenterPartHeaderNum, HorizontalAlignment.CENTER);
            studentTimetableSheet.setCellVerticalAlignment(0,
                    beginColumnCenterPartHeaderNum, VerticalAlignment.CENTER);
            studentTimetableSheet.setCellWrapText(0,
                    beginColumnCenterPartHeaderNum, true);
            Font centerPartHeaderFont = studentTimetableSheet.createFont();
            centerPartHeaderFont.setFontHeight((short) (28 * 20));
            studentTimetableSheet.setCellFont(0,
                    beginColumnCenterPartHeaderNum, centerPartHeaderFont);
            studentTimetableSheet
                    .getCell(0, beginColumnCenterPartHeaderNum)
                    .setCellValue(StudentTimetableConfig.TIMETABLE_NAME);

            int beginColumnRightPartHeaderNum = endColumnCenterPartHeaderNum + 1;
            int endColumnRightPartHeaderNum = beginColumnRightPartHeaderNum
                    + sidePartHeaderQtyColumns - 1;
            studentTimetableSheet
                    .mergeCells(0, beginColumnRightPartHeaderNum,
                            0, endColumnRightPartHeaderNum);
            studentTimetableSheet.setCellHorizontalAlignment(0,
                    beginColumnRightPartHeaderNum, HorizontalAlignment.LEFT);
            studentTimetableSheet.setCellVerticalAlignment(0,
                    beginColumnRightPartHeaderNum, VerticalAlignment.CENTER);
            studentTimetableSheet.setCellWrapText(0,
                    beginColumnRightPartHeaderNum, true);
            Font leftPartFont = studentTimetableSheet.createFont();
            leftPartFont.setFontHeight((short) (12 * 20));
            studentTimetableSheet.setCellFont(0,
                    beginColumnRightPartHeaderNum, leftPartFont);
            studentTimetableSheet
                    .getCell(0, beginColumnRightPartHeaderNum)
                    .setCellValue(StudentTimetableConfig.TIMETABLE_SIGN);

            studentTimetableSheet.autoSizeRow(0);

            int maxColumnWidth = 0;
            for (int i = 0; i < studentTimetableSheet.getPhysicalNumberOfColumns(); i++) {
                int currentColumnWidth = studentTimetableSheet.getSheet()
                        .getColumnWidth(i);
                if (maxColumnWidth < currentColumnWidth) {
                    maxColumnWidth = currentColumnWidth;
                }
            }

            float width = ((float) (3 * 7 + 5) / 7 * 256) / 256;
            studentTimetableSheet.setColumnWidth(1, Math.round(width) * 256);

            studentTimetableSheet.setColumnsWidth(2,
                    topSectionQtySchoolClasses,
                    maxColumnWidth);

            int sumOfColumnsWidthsTopSection = 0;
            for (int i = 0; i < topSectionQtySchoolClasses + 2; i++) {
                sumOfColumnsWidthsTopSection += studentTimetableSheet
                        .getColumnWidth(i);
            }

//            Create if not exists out file and write result to it.
            File outFile = new File("src/main/resources/result.xlsx");
            if (!outFile.exists()) {
                outFile.createNewFile();
            }
            FileOutputStream fos = new FileOutputStream(outFile);
            System.out.println("Writing workbook...");
            studentTimetableSheet.getSheet().getWorkbook().write(fos);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }


    }

}
