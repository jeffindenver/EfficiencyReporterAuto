package efficiencyreporterauto;

import excelops.ExcelOps;
import fileops.FileOps;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author JShepherd
 */
public class Model {

    private ReportFormat reportFormat;
    private final List<Agent> agents;
    private List<String> source;
    private String dateline;
    private String outputFilename;

    public Model() {
        agents = new ArrayList<>();
        source = new ArrayList<>();
    }

    void setSource(List<String> list) {
        this.source = new ArrayList<>(list);
    }

    void initializeAgents() {
        agents.clear();

        int USER_ID = 0;
        int correctLength = 6;
        
        List<String> target = new ArrayList<>();
        for (String line : source) {
            String splitLine[] = line.split(",");
            if (splitLine.length == correctLength) {
                if (!target.contains(splitLine[USER_ID]) //avoid duplicate names
                        && !splitLine[USER_ID].equalsIgnoreCase("User ID")) { //avoid header
                    target.add(splitLine[USER_ID]);
                }
            }
        }
        for (String userID : target) {
            agents.add(new Agent(userID));
        }
    }

    void addFullName() {
        int fName = 1;
        int lName = 2;
        
        for (Agent agent : agents) {
            for (String line : source) {
                if (line.contains(agent.getUserID()) && agent.getFname().isEmpty()) {
                    String[] splitLine = line.split(",");
                    agent.setFname(splitLine[fName]);
                    agent.setLname(splitLine[lName]);
                }
            }
        }
    }

    void extractDate() {
        int DATELINE = 3;
        dateline = source.get(DATELINE);
    }

    void initalizeReport(String outputTarget) throws InvalidFormatException, IOException {
        this.setOutputFilename(outputTarget);
        ExcelOps excelOps = new ExcelOps();
        int maxRow = agents.size();
        Workbook wb = excelOps.openWorkbook(this.getOutputFilename());
        this.reportFormat = new ReportFormat(wb, maxRow, dateline);
    }

    void cleanList(String target, String replacement) {
        for (String line : source) {
            if (line.contains(target)) {
                String tempLine = line.replace(target, replacement);
                source.set(source.indexOf(line), tempLine);
            }
        }
    }

    void calculateStats() {
        for (Agent agent : agents) {
            for (String line : source) {
                if (line.contains(agent.getUserID())) {
                    addStat(agent, line);
                }
            }
        }
    }

    private void addStat(Agent agent, String line) {

        String stats[] = line.split(",");

        if (stats.length == 6) {

            String statusKey = stats[3];
            String statusGroup = stats[4];
            String duration = stats[5];

            if (!statusKey.equalsIgnoreCase("gone home")) {
                agent.addLoginTime(duration);
            }

            if (statusGroup.equalsIgnoreCase("followup")
                    || statusGroup.equalsIgnoreCase("available")) {
                agent.addWorkingTime(duration);
            }

            if (statusKey.equalsIgnoreCase("on call")
                    || statusKey.equalsIgnoreCase("on email")
                    || statusKey.equalsIgnoreCase("on chat")
                    || statusKey.equalsIgnoreCase("on vm")) {
                agent.addTalkTime(duration);
            }

            if (statusGroup.equalsIgnoreCase("followup")) {
                agent.addAcwTime(duration);
            }
        }
        else {
            System.out.println(line);
            System.out.println("Cannot add stat: Invalid line length");
        }
    }

    void alphaSort() {
        java.util.Collections.sort(this.agents, Agent.lnameComparator);
    }

    void setCellValues() {
        int index = 0;
        for (Agent agent : agents) {
            reportFormat.setCellValues(agent, index);
            index++;
        }
    }

    List<String> readExcelFileToList(String filename) throws IOException, InvalidFormatException {
        ExcelOps excelOps = new ExcelOps();
        Workbook wb = excelOps.openWorkbook(filename);
        int sheetIndex = 0;
        return excelOps.sheetToList(wb, sheetIndex);
    }

    List<String> readFileToList(String filename) throws IOException {
        File file = new File(filename);
        FileOps fo = new FileOps(file, true);
        List<String> tempList = Collections.emptyList();

        if (file.exists()) {
            tempList = fo.readToList();
        }
        return tempList;
    }

    boolean writeToFile(String filename) throws InvalidFormatException, IOException {
        ExcelOps excelOps = new ExcelOps();
        XSSFWorkbook wb = reportFormat.getWorkbook();
        deselectSheets(wb);
        wb.setActiveSheet(wb.getNumberOfSheets() - 1);
        excelOps.writeWorkbook(wb, filename);
        return true;
    }

    boolean writeListToCsvFile(File file) throws IOException {
        FileOps fo = new FileOps(file, true);
        StringBuilder sb = new StringBuilder();

        for (Agent agent : agents) {
            sb.append(agent.toString()).append("\n");
        }

        try {
            fo.writeToFile(sb.toString());
        } catch (IOException e) {
            throw e;
        }
        return true;
    }

    String composeFilepath(String filename, String ext) {
        StringBuilder sb = new StringBuilder();
        String tempName = filename.replace(ext, "");

        sb.append(tempName);
        sb.append("_Processed");
        sb.append(ext);

        return sb.toString();
    }

    private void deselectSheets(Workbook wb) {
        for (Sheet sheet : wb) {
            sheet.setSelected(false);
        }
    }

    public String getOutputFilename() {
        return outputFilename;
    }

    public void setOutputFilename(String outputFilename) {
        this.outputFilename = outputFilename;
    }

    public ReportFormat getReportFormat() {
        return reportFormat;
    }

}
