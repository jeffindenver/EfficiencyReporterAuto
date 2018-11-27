package efficiencyreporterauto;



/**
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
