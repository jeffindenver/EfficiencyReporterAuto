package efficiencyreporterauto;

/**
 * This is a second version of "EfficiencyReport" application. This version will 
 * read .xlsx files directly and then modifying or creating the same. 
 * Previous version has .csv input and output.
 *
 * @author JShepherd
 */
public class EfficiencyReport {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {

        Model model = new Model();
        GUIview view = new GUIview();
        Controller controller = new Controller(view, model);
        controller.start();
    }

}
