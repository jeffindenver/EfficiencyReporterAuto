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

                initializeReport(filename);

                cleanData();

                calculateStats();

                alphaSort();

                model.setCellValues();

                writeToExistingFile(determineOutputTarget(filename));
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

    private void extractDate() {
        model.extractDate();
    }

    private void initializeAgents() {
        model.initializeAgents();
        model.addFullName();
    }

    private void initializeReport(String filename) {
        try {
            model.initalizeReport(filename);
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

    private void writeToExistingFile(String filename) {
        try {
            model.writeToExistingFile(filename);
        } catch (InvalidFormatException | IOException e) {
            view.printError(e.getMessage());
        }
    }

    private String determineOutputTarget(String filename) {
        String targetFile = filename;

        for (String name : ReportFormat.WORKGROUP_NAMES) {
            if (filename.contains(name)) {
                targetFile = getExistingFilename(name);
            }
        }
        return targetFile;
    }

    private String getExistingFilename(String name) {
        return model.getReportFormat().selectTargetFile(name);
    }
}
