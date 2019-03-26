package efficiencyreporterauto;

import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetDropEvent;
import java.io.File;
import java.io.IOException;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 *
 * @author JShepherd
 */
@SuppressWarnings("serial")
class Controller {

    private final GraphicalView view;
    private final Model model;
    private String sourceFilename;

    Controller(GraphicalView view, Model model) {
        this.view = view;
        this.model = model;
    }

    void start() {
        setupListeners();
    }

    private void setupListeners() {

        view.getTextArea().setDropTarget(new DropTarget() {

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
                setSourceFilename(filename);

                getSourceList(filename);

                initializeAgents();

                initializeReport(determineOutputTarget(filename));

                cleanData();

                calculateStats();

                alphaSort();

                model.setCellValues();

                writeToFile(model.getOutputFilename());
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

    private void initializeAgents() {
        model.initializeAgents();
        model.addFullName();
    }

    private void initializeReport(String outputTarget) {
        model.extractDate();
        try {
            model.initalizeReport(outputTarget);
        } catch (InvalidFormatException | IOException e) {
            view.printError(e.getMessage());
        }
    }

    private void cleanData() {
        model.cleanList("available, no acd", "no acd");
    }

    private void calculateStats() {
        model.calculateStats();
    }

    private void alphaSort() {
        model.alphaSort();
    }

    private void writeToFile(String filename) {
        try {
            model.writeToFile(filename);
        } catch (InvalidFormatException | IOException e) {
            view.printError(e.getMessage());
        }
    }

    private String determineOutputTarget(String filename) {
        String defaultFilename = getSourceFilename() + "_new_efficiency_report.xlsx";

        String targetFilename = "";

        for (String workgroupName : ReportFormat.WORKGROUP_NAMES) {
            if (filename.contains(workgroupName)) {
                targetFilename = getExistingFilename(workgroupName);
                break;
            } else {
                targetFilename = defaultFilename;
            }
        }
        return targetFilename;
    }

    private String getExistingFilename(String workgroupName) {
        return ReportFormat.selectTargetFile(workgroupName);
    }

    private String getSourceFilename() {
        return this.sourceFilename;
    }
    
    private void setSourceFilename(String filename) {
        this.sourceFilename = filename;
    }
}
