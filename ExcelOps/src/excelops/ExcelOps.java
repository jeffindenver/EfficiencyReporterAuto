package excelops;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.StringJoiner;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderExtent;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;

/**
 *
 * @author JShepherd
 */
public class ExcelOps {

    public ExcelOps() {

    }

    public XSSFWorkbook openWorkbook(String filename) throws InvalidFormatException, IOException {
        File file = new File(filename);
        XSSFWorkbook wb = new XSSFWorkbook();

        try {
            wb = XSSFWorkbookFactory.createWorkbook(file, true);
        } catch (InvalidFormatException | IOException e) {
            System.out.println(e.getMessage());
            throw e;
        } finally {
            if (wb != null) {
                try {
                    wb.close();
                } catch (IOException e) {
                    System.out.println(e.getMessage());
                    throw e;
                }
            }
        }
        return wb;
    }

    public List<String> sheetToList(XSSFWorkbook wb) {
        XSSFSheet sheet = wb.getSheetAt(0);
        return toListHelper(sheet);
    }

    public List<String> sheetToList(XSSFWorkbook wb, int sheetIndex) {
        XSSFSheet sheet = wb.getSheetAt(sheetIndex);
        return toListHelper(sheet);
    }

    private List<String> toListHelper(XSSFSheet sheet) {
        List<String> list = new ArrayList<>();
        Iterator<Row> rowIterator = sheet.iterator();

        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            Iterator<Cell> cellIterator = currentRow.iterator();
            StringJoiner joiner = new StringJoiner(",");

            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                if (currentCell.getCellType() == CellType.STRING) {
                    joiner.add(currentCell.getStringCellValue());
                } else if (currentCell.getCellType() == CellType.NUMERIC) {
                    joiner.add(String.valueOf(currentCell.getNumericCellValue()));
                }
            }
            list.add(joiner.toString());
        }
        return list;
    }

    public void writeWorkbook(XSSFWorkbook wb, String filename) {
        try (OutputStream fileOut = new FileOutputStream(filename)) {
            wb.write(fileOut);
        } catch (FileNotFoundException e) {
            System.out.println(e.getMessage());
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
        //change this to throw exceptions
    }

    public void formatSheet(XSSFWorkbook wb, int maxRow, int maxCol) {
        XSSFSheet sheet = wb.createSheet("FormattingTest");
        for (int i = 0; i < maxRow; i++) {
            XSSFRow row = wb.getSheetAt(0).createRow(i);
            for (int k = 0; k < maxCol; k++) {
                XSSFCell cell = row.createCell(k);
                System.out.println("Cell created at column " + cell.getColumnIndex());
            }
        }
        //This shouldn't be here. This should be a class in the program that 
        //is using ExcelOps 

// Create Fonts        
        Font bodyFont = wb.createFont();
        bodyFont.setFontName("Calibri");
        bodyFont.setFontHeightInPoints((short) 10);

        Font titleFont = wb.createFont();
        titleFont.setFontName("Calibri");
        titleFont.setFontHeightInPoints((short) 18);
        titleFont.setItalic(true);
        titleFont.setBold(true);

        Font headerFont = wb.createFont();
        headerFont.setFontName("Calibri");
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 10);

// Create Styles 
        XSSFCellStyle titleStyle = wb.createCellStyle();
        titleStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleStyle.setFont(titleFont);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);

        XSSFCellStyle headerStyle = wb.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setFont(headerFont);

        XSSFCellStyle agentNameStyle = wb.createCellStyle();
        agentNameStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        agentNameStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFCellStyle twoDecimalStyle = wb.createCellStyle();
        XSSFDataFormat twoDecimalFormat = wb.createDataFormat();
        twoDecimalStyle.setDataFormat(twoDecimalFormat.getFormat("0.00"));
        twoDecimalStyle.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
        twoDecimalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        twoDecimalStyle.setAlignment(HorizontalAlignment.CENTER);
        twoDecimalStyle.setFont(bodyFont);

        XSSFCellStyle percentageStyle = wb.createCellStyle();
        XSSFDataFormat percentageFormat = wb.createDataFormat();
        percentageStyle.setDataFormat(percentageFormat.getFormat("0.00%"));
        percentageStyle.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
        percentageStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        percentageStyle.setAlignment(HorizontalAlignment.CENTER);
        percentageStyle.setFont(bodyFont);

        XSSFCellStyle lightGreenStyle = wb.createCellStyle();
        lightGreenStyle.setDataFormat(percentageFormat.getFormat("0.00%"));
        lightGreenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        lightGreenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFCellStyle separatorStyle = wb.createCellStyle();
        separatorStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        separatorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

//row 0 blue background, "EMS Inc" in italics, merged
//row 1 blue background, "Performance Summary" in italics, merged 
        sheet.getRow(0).getCell(0).setCellValue("EMS Inc");
        for (Cell cell : sheet.getRow(0)) {
            cell.setCellStyle(titleStyle);
        }
        sheet.getRow(1).getCell(0).setCellValue("Performance Summary");
        for (Cell cell : sheet.getRow(1)) {
            cell.setCellStyle(titleStyle);
        }

        //merge first row
        sheet.addMergedRegion(new CellRangeAddress(
                0, //first row
                0, //last row
                0, //first col
                maxCol - 1 //last col));
        ));

        //merge second row
        sheet.addMergedRegion(new CellRangeAddress(
                1, //first row
                1, //last row
                0, //first col
                maxCol - 1 //last col
        ));

//row 2 Gray background and merged
        for (Cell cell : sheet.getRow(2)) {
            cell.setCellStyle(separatorStyle);
        }

        //merge third (2) row and set color to grey
        sheet.addMergedRegion(new CellRangeAddress(
                2, //first row
                2, //last row
                0, //first col
                maxCol - 1 //last column
        ));

//row 3 Header
        sheet.setColumnWidth(0, 28 * 256);
        int width = 16 * 256;
        for (int i = 1; i < maxCol; i++) {
            sheet.setColumnWidth(i, width);
        }

        final String[] header = {"Agent last name", "Login Time", "Working Time", "Talk Time", "ACW Time",
            "% ACW Time", "Available Time", "Handle Time", "Working Rate", "Occupancy"};

        int index = 0;
        for (Cell cell : sheet.getRow(3)) {
            cell.setCellStyle(headerStyle);
            cell.setCellValue(header[index]);
            index++;
        }

//column A style
        for (int i = 4; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            row.getCell(0).setCellStyle(agentNameStyle);
        }

//column B C D E style
        for (int i = 4; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            for (int k = 1; k < 5; k++) {
                row.getCell(k).setCellStyle(twoDecimalStyle);
            }
        }

//column F style
        for (int i = 4; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            row.getCell(5).setCellStyle(percentageStyle);
        }

//column G H Style
        for (int i = 4; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            for (int k = 6; k < 8; k++) {
                row.getCell(k).setCellStyle(twoDecimalStyle);
            }
        }

//column I J Style
        for (int i = 4; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            for (int k = 8; k < 10; k++) {
                row.getCell(k).setCellStyle(percentageStyle);
            }
        }

//set column 9 style
        for (int i = 4; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            row.getCell(9).setCellStyle(lightGreenStyle);
        }
//set column 8 style
        for (int i = 4; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            row.getCell(8).setCellStyle(lightGreenStyle);
        }

//add top and bottom border to every row below the header
        PropertyTemplate pt = new PropertyTemplate();
        pt.drawBorders(new CellRangeAddress(
                4, //first row
                maxRow - 1,//last row
                1, //first col
                maxCol - 1//last col
        ), BorderStyle.THIN, BorderExtent.HORIZONTAL);
        pt.applyBorders(sheet);

        pt = new PropertyTemplate();
        pt.drawBorders(new CellRangeAddress(
                4,//first row
                maxRow - 1,//last row
                0,//first col
                0//last col
        ), BorderStyle.THIN, BorderExtent.ALL);
        pt.applyBorders(sheet);

        pt = new PropertyTemplate();
        pt.drawBorders(new CellRangeAddress(
                4,//first row
                maxRow - 1,//last row
                8,//first col
                maxCol - 1// last col
        ), BorderStyle.THIN, BorderExtent.RIGHT);
        pt.applyBorders(sheet);

//Set conditional formatting
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();
        ConditionalFormattingRule greaterThan80 = sheetCF.createConditionalFormattingRule(ComparisonOperator.GE, ".8");
        ConditionalFormattingRule greaterThan70 = sheetCF.createConditionalFormattingRule(ComparisonOperator.GE, ".7");
        ConditionalFormattingRule lessThan70 = sheetCF.createConditionalFormattingRule(ComparisonOperator.LT, ".7");

        PatternFormatting lessThan70Format = lessThan70.createPatternFormatting();
        lessThan70Format.setFillBackgroundColor(IndexedColors.RED.getIndex());

        PatternFormatting greaterThan70Format = greaterThan70.createPatternFormatting();
        greaterThan70Format.setFillBackgroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());

        PatternFormatting greaterThan80Format = greaterThan80.createPatternFormatting();
        greaterThan80Format.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.getIndex());

        ConditionalFormattingRule[] cfRules = {greaterThan80, greaterThan70, lessThan70};

        CellRangeAddress[] regions = {
            CellRangeAddress.valueOf("I5:I30")
        };

        sheetCF.addConditionalFormatting(regions, cfRules);

//set formulas
        int column = 5;
        XSSFRow row;
        for (int i = 4; i < maxRow; i++) {
            row = sheet.getRow(i);
            int cellNum = i + 1;
            row.getCell(column).setCellFormula("IFERROR(E" + cellNum + "/B" + cellNum + ", \"-\")");
        }

        column = 6;
        for (int i = 4; i < maxRow; i++) {
            row = sheet.getRow(i);
            int cellNum = i + 1;
            row.getCell(column).setCellFormula("IFERROR(C" + cellNum + "-D" + cellNum + ", \"-\")");
        }

        column = 7;
        for (int i = 4; i < maxRow; i++) {
            row = sheet.getRow(i);
            int cellNum = i + 1;
            row.getCell(column).setCellFormula("D" + cellNum + "+E" + cellNum);
        }

        column = 8;
        for (int i = 4; i < maxRow; i++) {
            row = sheet.getRow(i);
            int cellNum = i + 1;
            row.getCell(column).setCellFormula("IFERROR(C" + cellNum + "/B" + cellNum + ", \"-\")");
        }

        column = 9;
        for (int i = 4; i < maxRow; i++) {
            row = sheet.getRow(i);
            int cellNum = i + 1;
            row.getCell(column).setCellFormula("IFERROR(H" + cellNum + "/(H" + cellNum + "+G" + cellNum + "), \"-\")");
        }

    }

    public void printList(List<String> list) {
        for (String s : list) {
            System.out.println(s);
        }
    }

}
