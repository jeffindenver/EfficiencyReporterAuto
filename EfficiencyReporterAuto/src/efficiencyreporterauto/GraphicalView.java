package efficiencyreporterauto;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Insets;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTabbedPane;
import javax.swing.JTextArea;
import javax.swing.WindowConstants;

/**
 *
 * @author JShepherd
 */
class GraphicalView {

    private final JFrame frame;
    private final JTabbedPane tabbedPane;
    private final JTextArea textArea;

    GraphicalView() {
        frame = new JFrame();
        frame.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
        frame.setBounds(40, 40, 600, 300);

        tabbedPane = new JTabbedPane();
        
        JPanel textPanel = new JPanel();
        textPanel.setPreferredSize(new Dimension(600,300));
        textArea = createTextArea();

        JScrollPane scrollPane = new JScrollPane(textArea);
        scrollPane.setOpaque(false);
        
        textPanel.add(scrollPane);
        
        JPanel optionsPanel = createOptionsArea();

        tabbedPane.addTab("Main", textPanel);
        tabbedPane.addTab("Options", optionsPanel);
        
        frame.add(tabbedPane);
        frame.pack();
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

    private JTextArea createTextArea() {
        JTextArea aTextArea = new JTextArea("Efficiency Reporter.");
        aTextArea.setPreferredSize(new Dimension(600, 300));
        Insets inset = new Insets(20, 20, 20, 20);
        aTextArea.setFont(new Font("Tahoma", Font.PLAIN, 14));
        aTextArea.setMargin(inset);
        aTextArea.setEditable(false);
        aTextArea.setBackground(new Color(0, 166, 255));
        aTextArea.setForeground(Color.black);
        return aTextArea;
    }
    
    private JPanel createOptionsArea() {
        JPanel aJPanel = new JPanel();
        return aJPanel;
    }
        
}
