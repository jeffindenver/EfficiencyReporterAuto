package efficiencyreporterauto;

import java.util.Scanner;

/**
 *
 * @author JShepherd
 */
public class View {
    //class to communicate with user, accept input

    void printError(String errorMessage) {
        System.err.println(errorMessage);
    }

    void printMessage(String message) {
        System.out.println(message);
    }

    public String getUserInput(String prompt) {
        Scanner sc = new Scanner(System.in);
        String input = "";
        System.out.println(prompt);
        if (sc.hasNextLine()) {
            input = sc.nextLine();
        }
        return input;
    }

    @Override
    public String toString() {
        return "Command line interface";
    }
}
