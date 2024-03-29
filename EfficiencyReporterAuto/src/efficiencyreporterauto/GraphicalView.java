package efficiencyreporterauto;

import java.awt.Color;
import java.awt.Font;
import java.awt.Insets;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.WindowConstants;

/**
 *
 * @author JShepherd
 */
class GraphicalView {

    private final JFrame frame;
    private final JTextArea textArea;

    GraphicalView() {
        frame = new JFrame();
        frame.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
        frame.setBounds(40, 40, 600, 300);

        textArea = new JTextArea("Efficiency Reporter.");
        JScrollPane scrollPane = new JScrollPane(textArea);
        scrollPane.setOpaque(false);
        Insets inset = new Insets(20, 20, 20, 20);
        textArea.setFont(new Font("Tahoma", Font.PLAIN, 14));
        textArea.setMargin(inset);
        textArea.setEditable(false);
        textArea.setBackground(new Color(0, 166, 255));
        textArea.setForeground(Color.black);

        frame.add(scrollPane);
        frame.setVisible(true);
    }

    JTextArea getTextArea() {
        return textArea;
    }

    void printError(String msg) {
        JOptionPane.showMessageDialog(frame, msg);
    }

    void printMessage(String msg) {
        textArea.append("\n");
        textArea.append(msg);
    }
}
