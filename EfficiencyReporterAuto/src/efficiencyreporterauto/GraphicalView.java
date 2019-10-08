package efficiencyreporterauto;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.GridBagConstraints;
import static java.awt.GridBagConstraints.FIRST_LINE_START;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.Insets;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTabbedPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
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
        
        JPanel optionsPanel = createOptionsPanel();

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
    
    private JPanel createOptionsPanel() {
        JPanel optionsArea = new JPanel(new GridLayout(3,1));
        optionsArea.setBackground(new Color(0, 166, 255));
        
        JPanel labelPanel = new JPanel();
        labelPanel.setOpaque(false);
        JLabel label = new JLabel("Enter the new thresholds below.");
        labelPanel.add(label);
        optionsArea.add(labelPanel );
        optionsArea.add(createThresholdsUserInputArea());
        
        JButton saveButton = new JButton("Save");
        saveButton.setPreferredSize(new Dimension(80,30));
        JPanel buttonPanel = new JPanel();
        buttonPanel.setOpaque(false);
        
        buttonPanel.add(saveButton);
        
        optionsArea.add(buttonPanel);
        return optionsArea;
    }
       
    private JPanel createThresholdsUserInputArea() {
        JPanel optionsPanel = new JPanel(new GridBagLayout());
        optionsPanel.setOpaque(false);
        GridBagConstraints c = new GridBagConstraints();
        c.anchor = FIRST_LINE_START;
        Insets insets = new Insets(10,10,10,10);
        c.insets = insets;

        JLabel goodScore = new JLabel("Good score is greater than :");
        JLabel midScore = new JLabel("Midling score is greater than :");
        JLabel poorScore = new JLabel("Poor score is anything less than :");
        
        c.gridy = 0;
        optionsPanel.add(goodScore, c);
        c.gridy = 1;
        optionsPanel.add(midScore, c);
        c.gridy = 2;
        optionsPanel.add(poorScore, c);
        
        int columns = 2;
        JTextField fieldOne = new JTextField(columns);
        JTextField fieldTwo = new JTextField(columns);
        JTextField fieldThree = new JTextField(columns);
        
        c.gridx = 1;
        c.gridy = 0;
        optionsPanel.add(fieldOne, c);
        
        c.gridy = 1;
        optionsPanel.add(fieldTwo, c);
        
        c.gridy = 2;
        optionsPanel.add(fieldThree, c);

        return optionsPanel;
    }
}
