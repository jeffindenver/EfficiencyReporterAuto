package efficiencyreporterauto;

import fileops.FileOps;
import java.io.IOException;

/**
 *
 * @author JShepherd
 */
public class WorkingRateThresholds {
    
    private String goodScore;
    private String midlingScore;
    private String poorScore;
    
    public WorkingRateThresholds() {
        String[] values = {"0", "0", "0"};
        try {
            values = readThresholdFile();
        } catch (IOException ex) {
           System.err.println(ex.getMessage());
           System.out.println("Threshold values set to default.");
           values[0] = "0.82";
           values[1] = "0.719";
           values[2] = "0.72";
        }
        this.goodScore = values[0];
        this.midlingScore = values[1];
        this.poorScore = values[2];
    }

    private String[] readThresholdFile() throws IOException {
        FileOps fo = new FileOps("AgentThresholds.txt", true);
        String[] thresholds = fo.readToArray();

        for (String str : thresholds) {
            System.out.println(str);
        }
        return thresholds;
    }

    private void writeThresholdFile(String[] values) throws IOException {
        FileOps fo = new FileOps("AgentThresholds.txt", true);
        for(String item : values) {
            fo.writeToFile(item);
        }
    }
    
    public String getGoodScore() {
        return goodScore;
    }

    public void setGoodScore(String goodScore) {
        this.goodScore = goodScore;
    }

    public String getMidlingScore() {
        return midlingScore;
    }

    public void setMidlingScore(String midlingScore) {
        this.midlingScore = midlingScore;
    }

    public String getPoorScore() {
        return poorScore;
    }

    public void setPoorScore(String poorScore) {
        this.poorScore = poorScore;
    }
    
        
}
