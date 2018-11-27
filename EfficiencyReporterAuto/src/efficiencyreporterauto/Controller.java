package efficiencyreporterauto;

import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetDropEvent;
import java.io.File;
import java.io.IOException;
import java.util.List;

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
                            = (List<File>) evt.getTransferable().getTransferData(DataFlavor.javaFileListFlavor);

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

                initializeAgents();

                cleanData();

                calculateStats();

                writeResults(filename);
            }
        });
    }

    private void getSourceList(String filename) {
        try {
            model.setSource(model.readFileToList(filename));
        } catch (IOException e) {
            view.printError(e.getMessage());
        }
    }

    private void initializeAgents() {
        model.initializeAgents();
    }

    private void calculateStats() {
        model.calculateStats();
    }

    private void writeResults(String aName) {
        String filename = model.composeFilepath(aName);
        model.writeListToFile(filename);
    }

    private void cleanData() {
        model.cleanList("available, no acd", "no acd");
    }
}
