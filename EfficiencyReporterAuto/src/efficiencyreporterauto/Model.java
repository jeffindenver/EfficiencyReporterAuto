package efficiencyreporterauto;



import fileops.FileOps;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 *
 * @author JShepherd
 */
public class Model {

    private final List<Agent> agents;
    private List<String> source;

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
            if (splitLine.length == 4) {
                if (!target.contains(splitLine[0]) //avoid duplicate names
                        && !splitLine[0].equalsIgnoreCase("Last name")) { //avoid header
                    target.add(splitLine[0]);
                }
            }
        }
        target = alphaSort(target);
        for (String name : target) {
            agents.add(new Agent(name));
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

    void calculateStats() {
        for (Agent agent : agents) {
            for (String line : source) {
                if (line.contains(agent.getLastName())) {
                    addStat(agent, line);
                }
            }
        }
    }

    private void addStat(Agent agent, String line) {
        String stats[] = line.split(",");
        String statusKey = stats[1];
        String statusGroup = stats[2];
        String duration = stats[3];

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

    private List<String> alphaSort(List<String> target) {
        java.util.Collections.sort(target);
        return target;
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

    boolean writeListToFile(String filename) {

        FileOps fo = new FileOps(filename, true);

        for (Agent agent : agents) {
            try {
                fo.writeToFile(agent.toString());
                fo.writeToFile("\n");
            } catch (IOException e) {
                System.out.println(e.getMessage());
            }
        }
        return true;
    }

    String composeFilepath(String filename) {
        StringBuilder sb = new StringBuilder();

        String tempName = filename.replace(".csv", "");

        sb.append(tempName);
        sb.append("_Processed");
        sb.append(".csv");

        return sb.toString();
    }
}
