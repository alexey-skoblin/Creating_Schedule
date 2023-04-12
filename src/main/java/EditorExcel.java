import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;

import java.awt.Color;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.*;

@Setter
@Getter
@Slf4j
public class EditorExcel {
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFFont standardFont;

    static class CellStyles {
        static XSSFCellStyle standard;
        static XSSFCellStyle interval;
        static XSSFCellStyle employ;
        static XSSFCellStyle title;
        static XSSFCellStyle numerator;
        static XSSFCellStyle denominator;
        static XSSFCellStyle assigmentByNumerator;
        static XSSFCellStyle assigmentByDenominator;
        static XSSFCellStyle assigmentByConstant;
        static XSSFCellStyle numberSession;
        static XSSFCellStyle timeSession;
        static XSSFCellStyle practices;
        static XSSFCellStyle lectures;
        static XSSFCellStyle lab;
        static XSSFCellStyle corp1;
        static XSSFCellStyle corp4;
        static XSSFCellStyle corp6;
        static XSSFCellStyle libraries;
    }

    Short sizeStandardRow = 400;
    Short sizeIntervalRow = 100;
    Integer sizeHorizontal = 256 * 3;

/*
    @Getter
    @Setter
    @AllArgsConstructor
    class CellEdit {
        Integer[][] arrayListSizes;
        XSSFCellStyle style;
        Boolean isUniteNumber;
        Boolean isRotation;
        Boolean isUniteRotation;
        AbstractMap enumStringMap;
        AbstractMap xssfCellStyleMap;
    }

    List<CellEdit> arrayElementsRow = Arrays.asList(
            new CellEdit(new Integer[][]{{1, 1}}, null, true, true, true,
                    new EnumMap<EducationClass.RotationWeek, String>(Map.ofEntries(
                            Map.entry(EducationClass.RotationWeek.Denominator, "З")
                    )),
                    new EnumMap<EducationClass.RotationWeek, XSSFCellStyle>(Map.ofEntries(
                            Map.entry(EducationClass.RotationWeek.Denominator, CellStyles.denominator)
                    ))
            )
    );
    CellEdit intervalCellEditRow = new CellEdit(
            new Integer[][]{},
            CellStyles.interval, false, false, false, null, null
    );
*/

    EditorExcel() {
        //employCellStyle
        CellStyles.employ = workbook.createCellStyle();
        CellStyles.employ.setAlignment(HorizontalAlignment.CENTER);
        CellStyles.employ.setVerticalAlignment(VerticalAlignment.CENTER);

        //standardCellStyle
        CellStyles.standard = workbook.createCellStyle();
        CellStyles.standard.setAlignment(HorizontalAlignment.CENTER);
        CellStyles.standard.setVerticalAlignment(VerticalAlignment.CENTER);
        CellStyles.standard.setBorderTop(BorderStyle.MEDIUM);
        CellStyles.standard.setBorderBottom(BorderStyle.MEDIUM);
        CellStyles.standard.setBorderLeft(BorderStyle.MEDIUM);
        CellStyles.standard.setBorderRight(BorderStyle.MEDIUM);
        CellStyles.standard.setFillForegroundColor(new XSSFColor(new Color(255, 255, 255), new DefaultIndexedColorMap()));
        CellStyles.standard.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        //Font
        standardFont = workbook.createFont();
        standardFont.setFontName("Bahnschrift");
        standardFont.setFontHeight(11);
        standardFont.setBold(true);
        CellStyles.employ.setFont(standardFont);
        CellStyles.standard.setFont(standardFont);

        //IntervalCellStyle
        CellStyles.interval = workbook.createCellStyle();
        CellStyles.interval.setBorderTop(BorderStyle.MEDIUM);
        CellStyles.interval.setBorderBottom(BorderStyle.MEDIUM);
        CellStyles.interval.setBorderLeft(BorderStyle.MEDIUM);
        CellStyles.interval.setBorderRight(BorderStyle.MEDIUM);

        CellStyles.title = CellStyles.standard.copy();
        CellStyles.title.setFillForegroundColor(new XSSFColor(new Color(172, 185, 202), new DefaultIndexedColorMap()));

        CellStyles.numerator = CellStyles.standard.copy();
        CellStyles.numerator.setFillForegroundColor(new XSSFColor(new Color(0, 112, 192), new DefaultIndexedColorMap()));

        CellStyles.denominator = CellStyles.standard.copy();
        CellStyles.denominator.setFillForegroundColor(new XSSFColor(new Color(255, 80, 80), new DefaultIndexedColorMap()));

        CellStyles.assigmentByNumerator = CellStyles.standard.copy();
        CellStyles.assigmentByNumerator.setFillForegroundColor(new XSSFColor(new Color(155, 194, 230), new DefaultIndexedColorMap()));

        CellStyles.assigmentByDenominator = CellStyles.standard.copy();
        CellStyles.assigmentByDenominator.setFillForegroundColor(new XSSFColor(new Color(248, 203, 173), new DefaultIndexedColorMap()));

        CellStyles.assigmentByConstant = CellStyles.standard.copy();
        CellStyles.assigmentByConstant.setFillForegroundColor(new XSSFColor(new Color(221, 235, 247), new DefaultIndexedColorMap()));

        CellStyles.numberSession = CellStyles.standard.copy();
        CellStyles.numberSession.setFillForegroundColor(new XSSFColor(new Color(201, 201, 201), new DefaultIndexedColorMap()));

        CellStyles.timeSession = CellStyles.standard.copy();
        CellStyles.timeSession.setFillForegroundColor(new XSSFColor(new Color(255, 217, 102), new DefaultIndexedColorMap()));

        CellStyles.practices = CellStyles.standard.copy();
        CellStyles.practices.setFillForegroundColor(new XSSFColor(new Color(172, 185, 202), new DefaultIndexedColorMap()));

        CellStyles.lectures = CellStyles.standard.copy();
        CellStyles.lectures.setFillForegroundColor(new XSSFColor(new Color(248, 203, 173), new DefaultIndexedColorMap()));

        CellStyles.lab = CellStyles.standard.copy();
        CellStyles.lab.setFillForegroundColor(new XSSFColor(new Color(198, 224, 180), new DefaultIndexedColorMap()));

        CellStyles.corp1 = CellStyles.standard.copy();
        CellStyles.corp1.setFillForegroundColor(new XSSFColor(new Color(0, 176, 240), new DefaultIndexedColorMap()));

        CellStyles.corp4 = CellStyles.standard.copy();
        CellStyles.corp4.setFillForegroundColor(new XSSFColor(new Color(255, 192, 0), new DefaultIndexedColorMap()));

        CellStyles.corp6 = CellStyles.standard.copy();
        CellStyles.corp6.setFillForegroundColor(new XSSFColor(new Color(0, 176, 80), new DefaultIndexedColorMap()));

        CellStyles.libraries = CellStyles.standard.copy();
        CellStyles.libraries.setFillForegroundColor(new XSSFColor(new Color(198, 89, 17), new DefaultIndexedColorMap()));
    }

    void writeFile(Schedule schedule) {
        XSSFSheet sheet = workbook.createSheet("sheet");

        Integer numberLine = 0;
        sheet.createRow(numberLine).setHeight(sizeStandardRow);
        sheet.getRow(numberLine).setRowStyle(CellStyles.employ);
        numberLine++;

        for (Schedule.Day day : schedule.getWeek()) {
            if (day.getClassesDay().size() != 0) {
                numberLine = addTitleRow(sheet, day, numberLine, 13);

                numberLine = addEducationClassRow(sheet, numberLine, day.getClassesDay().get(0));
                for (int i = 1; i < day.getClassesDay().size(); i++) {
                    numberLine = addIntervalRow(sheet, numberLine);
                    numberLine = addEducationClassRow(sheet, numberLine, day.getClassesDay().get(i));
                }

                //Отступ для следующего дня.
                sheet.createRow(numberLine).setHeight(sizeStandardRow);
                sheet.getRow(numberLine).setRowStyle(CellStyles.employ);
                numberLine++;
            }
        }

        for (int i = 0; i < 13 + 4; i++) {
            sheet.setColumnWidth(i, sizeHorizontal);
            sheet.autoSizeColumn(i);
        }

        try (OutputStream fileOut = new FileOutputStream(schedule.getName() + ".xlsx")) {
            workbook.write(fileOut);
        } catch (Exception e) {
            log.atError().log("Error write excel!");
            e.printStackTrace();
        }
    }

    private Integer addEducationClassRow(XSSFSheet sheet, Integer numberLine, EducationClass educationClass) {
        XSSFRow row = sheet.createRow(numberLine);
        row.setHeight(sizeStandardRow);
        XSSFCell copyCell;

        //cellRotation
        XSSFCell cellRotation;
        if (educationClass.getStatusRotation() == EducationClass.RotationWeek.Numerator) {
            cellRotation = row.createCell(1);
            cellRotation.setCellValue("Ч");
            assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 1, 1), sheet);
            cellRotation.setCellStyle(CellStyles.numerator);
        } else if (educationClass.getStatusRotation() == EducationClass.RotationWeek.Denominator) {
            cellRotation = row.createCell(1);
            cellRotation.setCellValue("З");
            assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 1, 1), sheet);
            cellRotation.setCellStyle(CellStyles.denominator);
        }
        cellRotation = row.getCell(1);

        //cellNumberEducationClass
        XSSFCell cellNumberEducationClass = row.createCell(2);
        cellNumberEducationClass.setCellValue(educationClass.getNumber());
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 2, 2), sheet);
        cellNumberEducationClass.setCellStyle(CellStyles.numberSession);

        //cellTimeSession
        XSSFCell cellTimeSession = row.createCell(3);
        cellTimeSession.setCellValue(EducationClass.ScheduleCalls[educationClass.getNumber() - 1]);
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 3, 3), sheet);
        cellTimeSession.setCellStyle(CellStyles.timeSession);

        //cellNumberEducationClass Copy
        copyCell = row.createCell(4);
        copyCell.copyCellFrom(cellNumberEducationClass, new CellCopyPolicy());
        copyCell.setCellStyle(cellNumberEducationClass.getCellStyle());

        XSSFCell cellStudySubject;
        //cellRotation Copy + cellStudySubject + cellRotation
        if (educationClass.getStatusRotation() != EducationClass.RotationWeek.Continuously) {
            copyCell = row.createCell(5);
            copyCell.copyCellFrom(cellRotation, new CellCopyPolicy());
            copyCell.setCellStyle(cellRotation.getCellStyle());

            cellStudySubject = row.createCell(6);
            cellStudySubject.setCellValue(educationClass.getStudySubject());
            assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 6, 6), sheet);
            if (educationClass.getStatusRotation() == EducationClass.RotationWeek.Numerator) {
                cellStudySubject.setCellStyle(CellStyles.assigmentByNumerator);
            } else {
                cellStudySubject.setCellStyle(CellStyles.assigmentByDenominator);
            }

            copyCell = row.createCell(7);
            copyCell.copyCellFrom(cellRotation, new CellCopyPolicy());
            copyCell.setCellStyle(cellRotation.getCellStyle());
        } else {
            cellStudySubject = row.createCell(5);
            cellStudySubject.setCellValue(educationClass.getStudySubject());
            assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 5, 5 + 2), sheet);
            sheet.addMergedRegion(new CellRangeAddress(numberLine, numberLine, 5, 5 + 2));
            cellStudySubject.setCellStyle(CellStyles.assigmentByConstant);
        }

        //cellNumberEducationClass Copy
        copyCell = row.createCell(8);
        copyCell.copyCellFrom(cellNumberEducationClass, new CellCopyPolicy());
        copyCell.setCellStyle(cellNumberEducationClass.getCellStyle());

        //cellNameEducator
        XSSFCell cellNameEducator = row.createCell(9);
        cellNameEducator.setCellValue(educationClass.getNameEducator());
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 9, 9), sheet);

        cellNameEducator.setCellStyle(switch (educationClass.getStatusRotation()) {
            case Continuously -> CellStyles.assigmentByConstant;
            case Numerator -> CellStyles.assigmentByNumerator;
            case Denominator -> CellStyles.assigmentByDenominator;
        });

        //cellNumberEducationClass Copy
        copyCell = row.createCell(10);
        copyCell.copyCellFrom(cellNumberEducationClass, new CellCopyPolicy());
        copyCell.setCellStyle(cellNumberEducationClass.getCellStyle());

        //cellTypeSubject
        XSSFCell cellTypeSubject = row.createCell(11);
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 11, 11), sheet);
        switch (educationClass.getTypeSubject()) {
            case Lecture -> {
                cellTypeSubject.setCellValue("Лек");
                cellTypeSubject.setCellStyle(CellStyles.lectures);
            }
            case Laboratory -> {
                cellTypeSubject.setCellValue("Лаб");
                cellTypeSubject.setCellStyle(CellStyles.lab);
            }
            case Practice -> {
                cellTypeSubject.setCellValue("Прак");
                cellTypeSubject.setCellStyle(CellStyles.practices);
            }
        }

        //cellStudyAuditorium
        XSSFCell cellStudyAuditorium = row.createCell(12);
        cellStudyAuditorium.setCellValue((String) educationClass.getStudyAuditorium());
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 12, 12), sheet);

        //cellStudyCorp
        XSSFCell cellStudyCorp = row.createCell(13);
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 13, 13), sheet);
        switch (educationClass.getStudyCorp()) {
            case Corp1 -> {
                cellStudyCorp.setCellValue("1к");
                cellStudyAuditorium.setCellStyle(CellStyles.corp1);
                cellStudyCorp.setCellStyle(CellStyles.corp1);
            }
            case Corp4 -> {
                cellStudyCorp.setCellValue("4к");
                cellStudyAuditorium.setCellStyle(CellStyles.corp4);
                cellStudyCorp.setCellStyle(CellStyles.corp4);
            }
            case Corp6 -> {
                cellStudyCorp.setCellValue("6к");
                cellStudyAuditorium.setCellStyle(CellStyles.corp6);
                cellStudyCorp.setCellStyle(CellStyles.corp6);
            }
            case Library -> {
                cellStudyCorp.setCellValue("НБ");
                cellStudyAuditorium.setCellStyle(CellStyles.libraries);
                cellStudyCorp.setCellStyle(CellStyles.libraries);
            }
        }

        //cellNumberEducationClass Copy
        copyCell = row.createCell(14);
        copyCell.copyCellFrom(cellNumberEducationClass, new CellCopyPolicy());
        copyCell.setCellStyle(cellNumberEducationClass.getCellStyle());

        //cellRotation Copy
        if (educationClass.getStatusRotation() != EducationClass.RotationWeek.Continuously) {
            copyCell = row.createCell(15);
            copyCell.copyCellFrom(cellRotation, new CellCopyPolicy());
            copyCell.setCellStyle(cellRotation.getCellStyle());
        }

        numberLine++;
        return numberLine;
    }

    private Integer addEducationClassPairRow(XSSFSheet sheet, Integer numberLine, EducationClass educationClass) {
        XSSFRow row = sheet.createRow(numberLine);
        row.setHeight(sizeStandardRow);
        XSSFCell copyCell;

        //cellRotation
        XSSFCell cellRotation;
        if (educationClass.getStatusRotation() == EducationClass.RotationWeek.Numerator) {
            cellRotation = row.createCell(1);
            cellRotation.setCellValue("Ч");
            assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 1, 1), sheet);
            cellRotation.setCellStyle(CellStyles.numerator);
        } else if (educationClass.getStatusRotation() == EducationClass.RotationWeek.Denominator) {
            cellRotation = row.createCell(1);
            cellRotation.setCellValue("З");
            assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 1, 1), sheet);
            cellRotation.setCellStyle(CellStyles.denominator);
        }
        cellRotation = row.getCell(1);

        //cellNumberEducationClass
        XSSFCell cellNumberEducationClass = row.createCell(2);
        cellNumberEducationClass.setCellValue(educationClass.getNumber());
        sheet.addMergedRegion(new CellRangeAddress(numberLine, numberLine + 1, 2, 2));
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine + 1, 2, 2), sheet);
        cellNumberEducationClass.setCellStyle(CellStyles.numberSession);

        //cellTimeSession
        XSSFCell cellTimeSession = row.createCell(3);
        cellTimeSession.setCellValue(EducationClass.ScheduleCalls[educationClass.getNumber() - 1]);
        sheet.addMergedRegion(new CellRangeAddress(numberLine, numberLine + 1, 3, 3));
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine + 1, 3, 3), sheet);
        cellTimeSession.setCellStyle(CellStyles.timeSession);

        //cellNumberEducationClass Copy
        copyCell = row.createCell(4);
        copyCell.copyCellFrom(cellNumberEducationClass, new CellCopyPolicy());
        sheet.addMergedRegion(new CellRangeAddress(numberLine, numberLine + 1, 4, 4));
        copyCell.setCellStyle(cellNumberEducationClass.getCellStyle());

        XSSFCell cellStudySubject;
        //cellRotation Copy + cellStudySubject + cellRotation
        if (educationClass.getStatusRotation() != EducationClass.RotationWeek.Continuously) {
            copyCell = row.createCell(5);
            sheet.addMergedRegion(new CellRangeAddress(numberLine, numberLine + 1, 5, 5));
            copyCell.copyCellFrom(cellRotation, new CellCopyPolicy());
            copyCell.setCellStyle(cellRotation.getCellStyle());

            cellStudySubject = row.createCell(6);
            cellStudySubject.setCellValue(educationClass.getStudySubject());
            assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 6, 6), sheet);
            if (educationClass.getStatusRotation() == EducationClass.RotationWeek.Numerator) {
                cellStudySubject.setCellStyle(CellStyles.assigmentByNumerator);
            } else {
                cellStudySubject.setCellStyle(CellStyles.assigmentByDenominator);
            }

            copyCell = row.createCell(7);
            sheet.addMergedRegion(new CellRangeAddress(numberLine, numberLine + 1, 7, 7));
            copyCell.copyCellFrom(cellRotation, new CellCopyPolicy());
            copyCell.setCellStyle(cellRotation.getCellStyle());
        } else {
            cellStudySubject = row.createCell(5);
            cellStudySubject.setCellValue(educationClass.getStudySubject());
            assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 5, 5 + 2), sheet);
            sheet.addMergedRegion(new CellRangeAddress(numberLine, numberLine, 5, 5 + 2));
            cellStudySubject.setCellStyle(CellStyles.assigmentByConstant);
        }

        //cellNumberEducationClass Copy
        copyCell = row.createCell(8);
        sheet.addMergedRegion(new CellRangeAddress(numberLine, numberLine + 1, 8, 8));
        copyCell.copyCellFrom(cellNumberEducationClass, new CellCopyPolicy());
        copyCell.setCellStyle(cellNumberEducationClass.getCellStyle());

        //cellNameEducator
        XSSFCell cellNameEducator = row.createCell(9);
        cellNameEducator.setCellValue(educationClass.getNameEducator());
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 9, 9), sheet);

        cellNameEducator.setCellStyle(switch (educationClass.getStatusRotation()) {
            case Continuously -> CellStyles.assigmentByConstant;
            case Numerator -> CellStyles.assigmentByNumerator;
            case Denominator -> CellStyles.assigmentByDenominator;
        });

        //cellNumberEducationClass Copy
        copyCell = row.createCell(10);
        sheet.addMergedRegion(new CellRangeAddress(numberLine, numberLine + 1, 10, 10));
        copyCell.copyCellFrom(cellNumberEducationClass, new CellCopyPolicy());
        copyCell.setCellStyle(cellNumberEducationClass.getCellStyle());

        //cellTypeSubject
        XSSFCell cellTypeSubject = row.createCell(11);
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 11, 11), sheet);
        switch (educationClass.getTypeSubject()) {
            case Lecture -> {
                cellTypeSubject.setCellValue("Лек");
                cellTypeSubject.setCellStyle(CellStyles.lectures);
            }
            case Laboratory -> {
                cellTypeSubject.setCellValue("Лаб");
                cellTypeSubject.setCellStyle(CellStyles.lab);
            }
            case Practice -> {
                cellTypeSubject.setCellValue("Прак");
                cellTypeSubject.setCellStyle(CellStyles.practices);
            }
        }

        //cellStudyAuditorium
        XSSFCell cellStudyAuditorium = row.createCell(12);
        cellStudyAuditorium.setCellValue((String) educationClass.getStudyAuditorium());
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 12, 12), sheet);

        //cellStudyCorp
        XSSFCell cellStudyCorp = row.createCell(13);
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 13, 13), sheet);
        switch (educationClass.getStudyCorp()) {
            case Corp1 -> {
                cellStudyCorp.setCellValue("1к");
                cellStudyAuditorium.setCellStyle(CellStyles.corp1);
                cellStudyCorp.setCellStyle(CellStyles.corp1);
            }
            case Corp4 -> {
                cellStudyCorp.setCellValue("4к");
                cellStudyAuditorium.setCellStyle(CellStyles.corp4);
                cellStudyCorp.setCellStyle(CellStyles.corp4);
            }
            case Corp6 -> {
                cellStudyCorp.setCellValue("6к");
                cellStudyAuditorium.setCellStyle(CellStyles.corp6);
                cellStudyCorp.setCellStyle(CellStyles.corp6);
            }
            case Library -> {
                cellStudyCorp.setCellValue("НБ");
                cellStudyAuditorium.setCellStyle(CellStyles.libraries);
                cellStudyCorp.setCellStyle(CellStyles.libraries);
            }
        }

        //cellNumberEducationClass Copy
        copyCell = row.createCell(14);
        sheet.addMergedRegion(new CellRangeAddress(numberLine, numberLine + 1, 14, 14));
        copyCell.copyCellFrom(cellNumberEducationClass, new CellCopyPolicy());
        copyCell.setCellStyle(cellNumberEducationClass.getCellStyle());

        //cellRotation Copy
        if (educationClass.getStatusRotation() != EducationClass.RotationWeek.Continuously) {
            copyCell = row.createCell(15);
            copyCell.copyCellFrom(cellRotation, new CellCopyPolicy());
            copyCell.setCellStyle(cellRotation.getCellStyle());
        }

        numberLine++;
        return numberLine;
    }

    Integer addTitleRow(XSSFSheet sheet, Schedule.Day day, Integer numberLine, Integer sizeTable) {
        XSSFRow row = sheet.createRow(numberLine);
        row.setHeight(sizeStandardRow);
        XSSFCell baseCell = row.createCell(3);
        baseCell.setCellValue(day.getName());
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 3, sizeTable), sheet);
        sheet.addMergedRegion(new CellRangeAddress(numberLine, numberLine, 3, sizeTable));
        baseCell.setCellStyle(CellStyles.title);
        numberLine++;
        return numberLine;
    }

    private Integer addIntervalRow(XSSFSheet sheet, Integer numberLine) {
        XSSFRow row = sheet.createRow(numberLine);
        row.setHeight(sizeIntervalRow);

        //1 cell
        XSSFCell cell1 = row.createCell(2);
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 2, 2), sheet);
        cell1.setCellStyle(CellStyles.interval);
        //2 cell
        XSSFCell cell2 = row.createCell(3);
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 3, 3), sheet);
        cell2.setCellStyle(CellStyles.interval);
        //3 cell
        XSSFCell cell3 = row.createCell(4);
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 4, 4), sheet);
        cell3.setCellStyle(CellStyles.interval);
        //4 cell += 3
        XSSFCell cell4 = row.createCell(5);
        sheet.addMergedRegion(new CellRangeAddress(numberLine, numberLine, 5, 5 + 2));
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 5, 5 + 2), sheet);
        cell4.setCellStyle(CellStyles.interval);
        //5 cell
        XSSFCell cell5 = row.createCell(8);
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 8, 8), sheet);
        cell5.setCellStyle(CellStyles.interval);
        //6 cell
        XSSFCell cell6 = row.createCell(9);
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 9, 9), sheet);
        cell6.setCellStyle(CellStyles.interval);
        //7 cell
        XSSFCell cell7 = row.createCell(10);
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 10, 10), sheet);
        cell7.setCellStyle(CellStyles.interval);
        //8 cell
        XSSFCell cell8 = row.createCell(11);
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 11, 11), sheet);
        cell8.setCellStyle(CellStyles.interval);
        //9 cell
        XSSFCell cell9 = row.createCell(12);
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 12, 12), sheet);
        cell9.setCellStyle(CellStyles.interval);
        //10 cell
        XSSFCell cell10 = row.createCell(13);
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 13, 13), sheet);
        cell10.setCellStyle(CellStyles.interval);
        //11 cell
        XSSFCell cel11 = row.createCell(14);
        assignWallsRegionMEDIUMBorders(new CellRangeAddress(numberLine, numberLine, 14, 14), sheet);
        cel11.setCellStyle(CellStyles.interval);

        numberLine++;
        return numberLine;
    }

    void assignWallsRegionMEDIUMBorders(CellRangeAddress region, XSSFSheet sheet) {
        RegionUtil.setBorderTop(BorderStyle.MEDIUM, region, sheet);
        RegionUtil.setBorderBottom(BorderStyle.MEDIUM, region, sheet);
        RegionUtil.setBorderLeft(BorderStyle.MEDIUM, region, sheet);
        RegionUtil.setBorderRight(BorderStyle.MEDIUM, region, sheet);
    }
}

