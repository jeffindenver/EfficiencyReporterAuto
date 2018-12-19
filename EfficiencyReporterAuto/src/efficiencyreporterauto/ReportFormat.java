package efficiencyreporterauto;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.StringJoiner;
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
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;
import org.apache.poi.xssf.usermodel.CustomIndexedColorMap;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author JShepherd
 */
public class ReportFormat {

    private final XSSFWorkbook wb;
    private final XSSFSheet sheet;
    private Font bodyFont;
    private Font titleFont;
    private Font headerFont;
    private XSSFCellStyle titleStyle;
    private XSSFCellStyle headerStyle;
    private XSSFCellStyle agentNameStyle;
    private XSSFCellStyle twoDecimalStyle;
    private XSSFCellStyle percentageStyle;
    private XSSFCellStyle lightGreenStyle;
    private XSSFCellStyle separatorStyle;
    private XSSFColor myBlue;
    private XSSFColor myYellow;
    private final int maxRow;
    private final int maxCol;

    private final int HEADER_SIZE = 4;
    private final int COLUMN_SIZE = 10;

    public ReportFormat(int max, String date) {
        this.wb = new XSSFWorkbook();
        this.sheet = wb.createSheet(composeSheetName(date));
        this.maxRow = max + HEADER_SIZE;
        this.maxCol = COLUMN_SIZE;
        buildSheet(date);
    }

    public ReportFormat(XSSFWorkbook workbook, int max, String date) {
        this.wb = workbook;
        this.sheet = wb.createSheet(composeSheetName(date));
        this.maxRow = max + HEADER_SIZE;
        this.maxCol = COLUMN_SIZE;
        buildSheet(date);
    }

    private void buildSheet(String date) {
        createColors();
        createFonts();
        createStyles();
        formatSheet(date);
        setColumnStyles();
        setBorders();
        setFormulas();
        setConditionalFormatting();
    }

    private void createColors() {
        DefaultIndexedColorMap defaultColorMap = new DefaultIndexedColorMap();
        byte[] rgbBlue = {(byte) 172, (byte) 212, (byte) 242};
        byte[] rgbYellow = {(byte) 200, (byte) 220, (byte) 137};
        this.myYellow = new XSSFColor(rgbYellow, defaultColorMap);
        this.myBlue = new XSSFColor(rgbBlue, defaultColorMap);
}

    private void createFonts() {
        bodyFont = wb.createFont();
        bodyFont.setFontName("Calibri");
        bodyFont.setFontHeightInPoints((short) 10);

        titleFont = wb.createFont();
        titleFont.setFontName("Calibri");
        titleFont.setFontHeightInPoints((short) 14);
        titleFont.setItalic(true);
        titleFont.setBold(true);

        headerFont = wb.createFont();
        headerFont.setFontName("Calibri");
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 10);
    }

    private void formatSheet(String date) {

        for (int i = 0; i < maxRow; i++) {
            XSSFRow row = sheet.createRow(i);
            for (int k = 0; k < maxCol; k++) {
                XSSFCell cell = row.createCell(k);
            }
        }

        //row 0 blue background with company name
        String companyName = "EMS Inc";
        sheet.getRow(0).getCell(0).setCellValue(companyName);
        for (Cell cell : sheet.getRow(0)) {
            cell.setCellStyle(titleStyle);
        }

        //row 1 blue background with title
        String title = "Performance Summary " + date.replace("\"", "");
        sheet.getRow(1).getCell(0).setCellValue(title);
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

        XSSFRow headerRow = sheet.getRow(3);
        for (int i = 0; i < maxCol; i++) {
            headerRow.getCell(i).setCellStyle(headerStyle);
            headerRow.getCell(i).setCellValue(header[i]);
        }
    }

    private void setColumnStyles() {
        final int topRow = 4;

        //column A style
        for (int i = topRow; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            row.getCell(0).setCellStyle(agentNameStyle);

        }

        //column B C D E style
        for (int i = topRow; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            for (int k = 1; k < 5; k++) {
                row.getCell(k).setCellStyle(twoDecimalStyle);
            }
        }

        //column F style
        for (int i = topRow; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            row.getCell(5).setCellStyle(percentageStyle);
        }

        //column G H Style
        for (int i = topRow; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            for (int k = 6; k < 8; k++) {
                row.getCell(k).setCellStyle(twoDecimalStyle);
            }
        }

        //set column I style
        for (int i = topRow; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            row.getCell(8).setCellStyle(lightGreenStyle);
        }

        //set column J style
        for (int i = topRow; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            row.getCell(9).setCellStyle(lightGreenStyle);
        }
    }

    private void setBorders() {
        PropertyTemplate pt = new PropertyTemplate();
        pt.drawBorders(new CellRangeAddress(
                4, //first row
                maxRow - 1, //last row
                1, //first col
                maxCol - 1 //last col
        ), BorderStyle.THIN, BorderExtent.HORIZONTAL);
        pt.applyBorders(sheet);

        pt = new PropertyTemplate();
        pt.drawBorders(new CellRangeAddress(
                4, //first row
                maxRow - 1, //last row
                0, //first col
                0 //last col
        ), BorderStyle.THIN, BorderExtent.ALL);
        pt.applyBorders(sheet);

        pt = new PropertyTemplate();
        pt.drawBorders(new CellRangeAddress(
                4, //first row
                maxRow - 1, //last row
                8, //first col
                maxCol - 1 //last col
        ), BorderStyle.THIN, BorderExtent.RIGHT);
        pt.applyBorders(sheet);

        pt = new PropertyTemplate();
        pt.drawBorders(new CellRangeAddress(
                4, //first row
                maxRow - 1, //last row
                8, //first col
                maxCol - 1 //last col
        ), BorderStyle.THIN, BorderExtent.VERTICAL);
        pt.applyBorders(sheet);
    }

    private void setFormulas() {
        final int topRow = 4;

        int column = 5;
        XSSFRow row;
        for (int i = topRow; i < maxRow; i++) {
            row = sheet.getRow(i);
            int cellNum = i + 1;
            row.getCell(column)
                    .setCellFormula("IFERROR(E" + cellNum + "/B" + cellNum + ", \"-\")");
        }

        column = 6;
        for (int i = topRow; i < maxRow; i++) {
            row = sheet.getRow(i);
            int cellNum = i + 1;
            row.getCell(column)
                    .setCellFormula("IFERROR(C"
                            + cellNum + "-D" + cellNum + ", \"-\")");
        }

        column = 7;
        for (int i = topRow; i < maxRow; i++) {
            row = sheet.getRow(i);
            int cellNum = i + 1;
            row.getCell(column)
                    .setCellFormula("D" + cellNum + "+E" + cellNum);
        }

        column = 8;
        for (int i = topRow; i < maxRow; i++) {
            row = sheet.getRow(i);
            int cellNum = i + 1;
            row.getCell(column)
                    .setCellFormula("IFERROR(C"
                            + cellNum + "/B" + cellNum + ", \"-\")");
        }

        column = 9;
        for (int i = topRow; i < maxRow; i++) {
            row = sheet.getRow(i);
            int cellNum = i + 1;
            row.getCell(column)
                    .setCellFormula("IFERROR(H"
                            + cellNum + "/(H" + cellNum
                            + "+G" + cellNum + "), \"-\")");
        }
    }

    void setCellValues(Agent agent, int rowNum) {
        System.out.println(agent.toString());

        int index = 4 + rowNum;
        sheet.getRow(index).getCell(0).setCellType(CellType.STRING);
        for (int i = 1; i < 5; i++) {
            sheet.getRow(index).getCell(i).setCellType(CellType.NUMERIC);
        }

        double dLoginTime = agent.getLoginTime().toMillis() / 1000;
        double dWorkingTime = agent.getWorkingTime().toMillis() / 1000;
        double dTalkTime = agent.getTalkTime().toMillis() / 1000;
        double dAcwTime = agent.getAcwTime().toMillis() / 1000;

        sheet.getRow(index).getCell(0).setCellValue(agent.getLastName());
        sheet.getRow(index).getCell(1).setCellValue(dLoginTime / 60);
        sheet.getRow(index).getCell(2).setCellValue(dWorkingTime / 60);
        sheet.getRow(index).getCell(3).setCellValue(dTalkTime / 60);
        sheet.getRow(index).getCell(4).setCellValue(dAcwTime / 60);
    }

    private void createStyles() {
        titleStyle = wb.createCellStyle();
        titleStyle.setFillForegroundColor(myBlue);
        titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleStyle.setFont(titleFont);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);

        headerStyle = wb.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setFont(headerFont);

        agentNameStyle = wb.createCellStyle();
        agentNameStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        agentNameStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        twoDecimalStyle = wb.createCellStyle();
        XSSFDataFormat twoDecimalFormat = wb.createDataFormat();
        twoDecimalStyle.setDataFormat(twoDecimalFormat.getFormat("0.00"));
        twoDecimalStyle.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
        twoDecimalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        twoDecimalStyle.setAlignment(HorizontalAlignment.CENTER);
        twoDecimalStyle.setFont(bodyFont);

        percentageStyle = wb.createCellStyle();
        XSSFDataFormat percentageFormat = wb.createDataFormat();
        percentageStyle.setDataFormat(percentageFormat.getFormat("0.00%"));
        percentageStyle.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
        percentageStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        percentageStyle.setAlignment(HorizontalAlignment.CENTER);
        percentageStyle.setFont(bodyFont);

        lightGreenStyle = wb.createCellStyle();
        lightGreenStyle.setDataFormat(percentageFormat.getFormat("0.00%"));
        lightGreenStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        lightGreenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        lightGreenStyle.setAlignment(HorizontalAlignment.CENTER);

        separatorStyle = wb.createCellStyle();
        separatorStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        separatorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    private void setConditionalFormatting() {
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
            CellRangeAddress.valueOf("I5:I" + maxRow)
        };
        sheetCF.addConditionalFormatting(regions, cfRules);
    }

    private String composeSheetName(String dateline) {
        String sheetName = parseDate(dateline);
        //If you try to create a sheet with a duplicate name, it will fail
        int numOfSheets = this.getWorkbook().getNumberOfSheets();
        if (numOfSheets > 1) {
            for (Sheet localSheet : this.getWorkbook()) {
                if (sheetName.equalsIgnoreCase(localSheet.getSheetName())) {
                    sheetName += " ID " + Math.random() * 1000;
                    break;
                }
            }
        }
        return sheetName;
    }

    private String parseDate(String dateline) {
        //String dateline looks like "From 12/9/2018 12:00:00 AM To 12/15/2018 11:59:59 PM"
        //element 1 is the date in format MM/dd/yyyy
        String[] v = dateline.split(" ");

        //take the date and split again, which should give day, month, and year
        String dateArr[] = v[1].split("/");

        //date formatter fails if the day or month are not two digits, 
        //so this pre-pends a zero if needed.
        StringJoiner join = new StringJoiner("/");
        for (int i = 0; i < 3; i++) {
            if (dateArr[i].length() < 2) {
                dateArr[i] = "0" + dateArr[i];
            }
            join.add(dateArr[i]);
        }

        String wellFormedDate = join.toString();

        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/dd/yyyy");
        LocalDate date = LocalDate.parse(wellFormedDate, formatter);
        System.out.println(date.getMonth() + " " + String.valueOf(date.getDayOfMonth()));
        String month = date.getMonth().toString().substring(0, 3);
        return month + " " + String.valueOf(date.getDayOfMonth());
    }

    XSSFWorkbook getWorkbook() {
        return this.wb;
    }
}
