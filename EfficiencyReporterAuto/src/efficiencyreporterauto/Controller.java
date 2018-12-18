package efficiencyreporterauto;

import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetDropEvent;
import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 *
 * @author JShepherd
 */
public class Controller {

    private final GUIview view;
    private final Model model;

    Controller(GUIview view, Model model) {
        this.view = view;
        this.model = model;
    }

    void start() {
        setupListeners();
    }

    private void setupListeners() {
        view.getTextArea().setDropTarget(new DropTarget() {
            private static final long serialVersionUID = 1L;

            @Override
            public synchronized void drop(DropTargetDropEvent evt) {
                try {
                    evt.acceptDrop(DnDConstants.ACTION_COPY);

                    @SuppressWarnings("unchecked")
                    List<File> droppedFiles
                            = (List<File>) evt.getTransferable()
                                    .getTransferData(DataFlavor.javaFileListFlavor);

                    for (File file : droppedFiles) {
                        view.printMessage(file.toString());
                        processFile(file.toString());
                    }
                } catch (UnsupportedFlavorException | IOException e) {
                    view.printError(e.getMessage());
                }
            }

            private void processFile(String filename) {
                getSourceList(filename);

                extractDate();

                initializeAgents();

                cleanData();

                calculateStats();

                writeToExistingFile(determineOutputTarget(filename));

                writeCSVFile(filename);

                writeExcelFile(filename);
            }

        });
    }

    private void getSourceList(String filename) {
        try {
            model.setSource(model.readExcelFileToList(filename));
        } catch (IOException | InvalidFormatException e) {
            view.printError(e.getMessage());
        }
    }
        
    private void writeToExistingFile(String filename) {
        try {
            model.writeToExistingFile(filename);
        } catch (InvalidFormatException | IOException e) {
            System.out.println(e.getMessage());
        }

    }

    private String determineOutputTarget(String filename) {
        String targetFile = filename;
        List<String> workgroupNames = Arrays.asList("CellOne", "Drobo",
                "Homesnap", "Newmark", "Orbit", "Shared", "Xplore", "YKHC");
        for (String name : workgroupNames) {
            if (filename.contains(name)) {
                targetFile = getExistingFilename(name);
            }
        }
        return targetFile;
    }

    private String getExistingFilename(String name) {
        String translation = "S:\\Reports\\Efficiency Reports\\";
        switch (name) {
            case "CellOne":
                translation += "CellOne Efficiency 2018.xlsx";
                break;
            case "Drobo":
                translation += "Drobo Efficiency 2018.xlsx";
                break;
            case "Homesnap":
                translation += "Homesnap Efficiency 2018.xlsx";
                break;
            case "Newmark":
                translation += "Newmark Efficiency 2018.xlsx";
                break;
            case "Orbit":
                translation += "Orbit Efficiency 2018.xlsx";
                break;
            case "Shared":
                translation += "Shared Support Efficiency 2018.xlsx";
                break;
            case "Xplore":
                translation += "Xplore Efficiency 2018.xlsx";
                break;
            case "YKHC":
                translation += "YKHC Efficiency 2018.xlsx";
                break;
            default:
                translation += "new efficiency report.xlsx";
        }
        return translation;
    }



    private void extractDate() {
        model.extractDate();
    }

    private void initializeAgents() {
        model.initializeAgents();
    }

    private void calculateStats() {
        model.calculateStats();
    }

    private void writeCSVFile(String aName) {
        File file = new File(model.composeFilepath(aName, ".csv"));
        try {
            model.writeListToFile(file);            
        } catch (IOException e) {
            view.printError(e.getMessage());
        }
    }

    private void writeExcelFile(String aName) {
        model.writeListToXlsxFile(model.composeFilepath(aName, ".xlsx"));
    }

    private void cleanData() {
        model.cleanList("available, no acd", "no acd");
    }
}
