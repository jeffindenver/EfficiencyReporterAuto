package efficiencyreporterauto;

/**
 * This is a third version of "EfficiencyReport" application. This version will 
 * read .xlsx files directly and then modifying or creating the same. 
 * Features of the third version include customizable thresholds for agent 
 * efficiency, which was hard-coded previously, and a customizable map of
 * keywords and filenames (k, v) so that new workgroups can be added.
 *
 * @author JShepherd
 */
public class EfficiencyReport {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {

        Model model = new Model();
        GraphicalView view = new GraphicalView();
        Controller controller = new Controller(view, model);
        controller.start();
    }

}
