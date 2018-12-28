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

    private final List<Agent> agents;
    private ReportFormat reportFormat;

    private List<String> source;
    private String dateline;

    public Model() {
        agents = new ArrayList<>();
        source = new ArrayList<>();
    }

    void setSource(List<String> list) {
        this.source = new ArrayList<>(list);
    }

    void initializeAgents() {
        agents.clear();

        List<String> target = new ArrayList<>();
        for (String line : source) {
            String splitLine[] = line.split(",");
            if (splitLine.length == 6) {
                if (!target.contains(splitLine[0]) //avoid duplicate names
                        && !splitLine[0].equalsIgnoreCase("User ID")) { //avoid header
                    target.add(splitLine[0]);
                }
            }
        }
        for (String userID : target) {
            agents.add(new Agent(userID));
        }
    }

    void cleanList(String target, String replacement) {
        for (String line : source) {
            if (line.contains(target)) {
                String tempLine = line.replace(target, replacement);
                source.set(source.indexOf(line), tempLine);
            }
        }
    }

    void addFullName() {
        for (Agent agent : agents) {
            for (String line : source) {
                if (line.contains(agent.getUserID()) && agent.getFname().isEmpty()) {
                    String[] splitLine = line.split(",");
                    agent.setFname(splitLine[1]);
                    agent.setLname(splitLine[2]);
                }
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

    void alphaSort() {
        java.util.Collections.sort(this.agents, Agent.lnameComparator);
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

    //With the following two methods, the cell values should be set elsewhere
    boolean writeToExistingFile(String filename) throws InvalidFormatException, IOException {
        ExcelOps excelOps = new ExcelOps();
        int index = 0;
        for (Agent agent : agents) {
            reportFormat.setCellValues(agent, index);
            index++;
        }
        XSSFWorkbook wb = reportFormat.getWorkbook();
        deselectSheets(wb);
        wb.setActiveSheet(wb.getNumberOfSheets() - 1);
        excelOps.writeWorkbook(wb, filename);
        return true;
    }

    boolean writeListToXlsxFile(String filename) {
        ExcelOps excelOps = new ExcelOps();
        int maxRow = agents.size();

        //@TODO Test this with a report that does not have an existing file
        reportFormat = new ReportFormat(maxRow, dateline);
        int index = 0;
        for (Agent agent : agents) {
            reportFormat.setCellValues(agent, index);
            index++;
        }
        excelOps.writeWorkbook(reportFormat.getWorkbook(), filename);
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

    void extractDate() {
        int DATELINE = 3;
        dateline = source.get(DATELINE);

    }

    private void deselectSheets(Workbook wb) {
        for (Sheet sheet : wb) {
            sheet.setSelected(false);
        }
    }

    void initalizeReport(String filename) throws InvalidFormatException, IOException {
        ExcelOps excelOps = new ExcelOps();
        int maxRow = agents.size();
        XSSFWorkbook wb = (XSSFWorkbook) excelOps.openWorkbook(filename);
        this.reportFormat = new ReportFormat(wb, maxRow, dateline);
    }

    void setCellValues() {
        int index = 0;
        for (Agent agent : agents) {
            reportFormat.setCellValues(agent, index);
            index++;
        } 
    }
    
    public ReportFormat getReportFormat() {
        return reportFormat;
    }

}
