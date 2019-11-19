package efficiencyreporterauto;

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
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;
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
    

    //Change ReportFormat to read a map of key/value pairs for the keyword and 
    //output filename.
    public final static String[] WORKGROUP_NAMES = {"MSRB", "Orbit", "Shared", "YKHC"};

    public final static String FILEPATH = "S:\\Reports\\Efficiency Reports\\";

    static String selectTargetFile(String name) {
        String targetFile = ReportFormat.FILEPATH;
        switch (name) {
            case "MSRB":
                targetFile += "MSRB Efficiency 2019.xlsx";
                break;
            case "Orbit":
                targetFile += "Orbit Efficiency 2019.xlsx";
                break;
            case "Shared":
                targetFile += "Shared Support Efficiency 2019.xlsx";
                break;
            case "YKHC":
                targetFile += "YKHC Efficiency 2019.xlsx";
                break;
            default:
                targetFile += "new_efficiency_report.xlsx";
        }
        return targetFile;
    }
    
    private WorkingRateThresholds thresholds;

    private final XSSFWorkbook wb;
    private XSSFSheet sheet;
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
    private XSSFCellStyle lightGreenDecimalStyle;
    private XSSFColor myBlue;
    private XSSFColor myYellow;
    private int maxRow;
    private int maxCol;

    private final int HEADER_SIZE = 4;
    private final int COLUMN_SIZE = 14;

    public ReportFormat(int max, String date) {
        System.out.println("default thresholds used");
        this.wb = new XSSFWorkbook();
        initializeSheet(wb, max, date);
        buildSheet(date);
    }

    public ReportFormat(Workbook workbook, int maxRow, String date) {
        setThresholds();
        this.wb = (XSSFWorkbook) workbook;
        initializeSheet(wb, maxRow, date);
        buildSheet(date);
    }

    private void setThresholds() {
        thresholds = new WorkingRateThresholds();
    }
        
    private void initializeSheet(Workbook wb, int maxRow, String date) {
        this.sheet = (XSSFSheet) wb.createSheet(composeSheetName(date));
        this.maxRow = maxRow + HEADER_SIZE;
        this.maxCol = COLUMN_SIZE;
    }

    private void buildSheet(String date) {
        createColors();
        createFonts();
        createStyles();
        formatSheet(date);
        setColumnStyles();
        setFormulas();
        setConditionalFormatting();
        setBorders();
    }

    private void createColors() {
        EfficiencyReportColorMap colorMap = new EfficiencyReportColorMap();
        byte[] rgbBlue = {(byte) 172, (byte) 212, (byte) 242};
        byte[] rgbYellow = {(byte) 255, (byte) 220, (byte) 137};
        this.myBlue = new XSSFColor(rgbBlue, colorMap);
        this.myYellow = new XSSFColor(rgbYellow, colorMap);
    }

    private void createFonts() {
        bodyFont = wb.createFont();
        bodyFont.setFontName("Calibri");
        bodyFont.setFontHeightInPoints((short) 9);

        titleFont = wb.createFont();
        titleFont.setFontName("Calibri");
        titleFont.setFontHeightInPoints((short) 14);
        titleFont.setItalic(true);
        titleFont.setBold(true);

        headerFont = wb.createFont();
        headerFont.setFontName("Calibri");
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 9);
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
        int width = 14 * 256;
        for (int i = 1; i < maxCol; i++) {
            sheet.setColumnWidth(i, width);
        }

        final String[] header = {"Agent name", "Login Time", "Working Time", "Talk Time",
            "ACW Time", "% ACW Time", "Available Time", "Inbound Calls", "Oubound Calls", 
            "Handle Time", "Adjusted Login Time", "Working Rate", "Occupancy", "TCPH"};

        XSSFRow headerRow = sheet.getRow(3);
        for (int i = 0; i < maxCol; i++) {
            headerRow.getCell(i).setCellStyle(headerStyle);
            headerRow.getCell(i).setCellValue(header[i]);
        }
    }

    private void setColumnStyles() {
        final int topRow = 4;

        //column A style agent name
        for (int i = topRow; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            row.getCell(0).setCellStyle(agentNameStyle);

        }

        //column B C D E style login time, working time, talk time, acw time
        for (int i = topRow; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            for (int k = 1; k < 5; k++) {
                row.getCell(k).setCellStyle(twoDecimalStyle);
            }
        }

        //column F style %acwTime
        for (int i = topRow; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            row.getCell(5).setCellStyle(percentageStyle);
        }

        //column G H I J K Style available time, inbound calls, outbound calls,
        //handle time, adjusted login time
        for (int i = topRow; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            for (int k = 6; k < 11; k++) {
                row.getCell(k).setCellStyle(twoDecimalStyle);
            }
        }
    
        //Column L style working rate
        for (int i = topRow; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            row.getCell(11).setCellStyle(lightGreenStyle);
        }

        //Column M style occupancy
        for (int i = topRow; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            row.getCell(12).setCellStyle(lightGreenStyle);
        }
        
        //Column N style TCPH
        for (int i = topRow; i < maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            row.getCell(13).setCellStyle(lightGreenDecimalStyle);
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
                7, //first col
                maxCol - 1 //last col
        ), BorderStyle.THIN, BorderExtent.RIGHT);
        pt.applyBorders(sheet);

        pt = new PropertyTemplate();
        pt.drawBorders(new CellRangeAddress(
                4, //first row
                maxRow - 1, //last row
                7, //first col
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

        column = 9;
        for (int i = topRow; i < maxRow; i++) {
            row = sheet.getRow(i);
            int cellNum = i + 1;
            row.getCell(column)
                    .setCellFormula("D" + cellNum + "+E" + cellNum);
        }

        column = 11;
        for (int i = topRow; i < maxRow; i++) {
            row = sheet.getRow(i);
            int cellNum = i + 1;
            row.getCell(column)
                    .setCellFormula("IFERROR(C"
                            + cellNum + "/B" + cellNum + ", \"-\")");
        }

        column = 12;
        for (int i = topRow; i < maxRow; i++) {
            row = sheet.getRow(i);
            int cellNum = i + 1;
            row.getCell(column)
                    .setCellFormula("IFERROR(J"
                            + cellNum + "/(J" + cellNum
                            + "+G" + cellNum + "), \"-\")");
        }
        
        column = 13;
        for (int i = topRow; i < maxRow; i++) {
            row = sheet.getRow(i);
            int cellNum = i + 1;
            row.getCell(column)
                    .setCellFormula("(H" + cellNum
                            + "+" + "I" + cellNum + ")/(K" 
                            + cellNum + " / 60)");
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
        double dNegativeTime = agent.getBreaksAndOtherTime().toMillis() / 1000;
        double dAdjustedLoginTime = dLoginTime - dNegativeTime;

        sheet.getRow(index).getCell(0).setCellValue(agent.getFullname());
        sheet.getRow(index).getCell(1).setCellValue(dLoginTime / 60);
        sheet.getRow(index).getCell(2).setCellValue(dWorkingTime / 60);
        sheet.getRow(index).getCell(3).setCellValue(dTalkTime / 60);
        sheet.getRow(index).getCell(4).setCellValue(dAcwTime / 60);
        sheet.getRow(index).getCell(10).setCellValue(dAdjustedLoginTime / 60);
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
        agentNameStyle.setFont(bodyFont);

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
        lightGreenStyle.setFont(bodyFont);
        
        lightGreenDecimalStyle = wb.createCellStyle();
        lightGreenDecimalStyle.setDataFormat(twoDecimalFormat.getFormat("0.00"));
        lightGreenDecimalStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        lightGreenDecimalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        lightGreenDecimalStyle.setAlignment(HorizontalAlignment.CENTER);
        lightGreenDecimalStyle.setFont(bodyFont);

        separatorStyle = wb.createCellStyle();
        separatorStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        separatorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    private void setConditionalFormatting() {
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();
        ConditionalFormattingRule greaterThanGoodScore = sheetCF.createConditionalFormattingRule(ComparisonOperator.GE, thresholds.getGoodScore());
        ConditionalFormattingRule greaterThanMidlingScore = sheetCF.createConditionalFormattingRule(ComparisonOperator.GE, thresholds.getMidlingScore());
        ConditionalFormattingRule lessThanPoorScore = sheetCF.createConditionalFormattingRule(ComparisonOperator.LT, thresholds.getPoorScore());

        PatternFormatting poorScoreFormat = lessThanPoorScore.createPatternFormatting();
        poorScoreFormat.setFillBackgroundColor(IndexedColors.RED.getIndex());

        PatternFormatting midlingScoreFormat = greaterThanMidlingScore.createPatternFormatting();
        midlingScoreFormat.setFillBackgroundColor(myYellow);

        PatternFormatting goodScoreFormat = greaterThanGoodScore.createPatternFormatting();
        goodScoreFormat.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.getIndex());

        ConditionalFormattingRule[] cfRules = {greaterThanGoodScore, greaterThanMidlingScore, lessThanPoorScore};

        CellRangeAddress[] regions = {
            CellRangeAddress.valueOf("L5:L" + maxRow)
        };
        sheetCF.addConditionalFormatting(regions, cfRules);
    }

    private String composeSheetName(String dateline) {
        String sheetName = getMonthAndDay(dateline);
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

    private String getMonthAndDay(String line) {
        final String[] months = {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul",
            "Aug", "Sep", "Oct", "Nov", "Dec"};

        //The arg, line, looks like: "From 12/9/2018 12:00:00 AM To 12/15/2018 11:59:59 PM"
        String[] dateline = line.split(" ");
        String[] tokenizedDate = dateline[1].split("/");

        int month = Integer.parseInt(tokenizedDate[0]) - 1;
        String day = tokenizedDate[1];
        return months[month] + " " + day;
    }

    XSSFWorkbook getWorkbook() {
        return this.wb;
    }
}
