import javax.swing.*;
import javax.swing.filechooser.FileSystemView;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.Arrays;

class DataEntry {
    public static final JTextArea txtFileInput = new JTextArea();
    public static final JTextArea txtFileTemplate = new JTextArea();
    public static final  JTextField vlOutputLocation = new JTextField(20);
    public static final  JCheckBox ckbIFC = new JCheckBox("IFC");
    public static final  JCheckBox ckbIFA = new JCheckBox("IFA");
    public static final  JTextField vlProject = new JTextField(20);
    public static final  JTextField vlAddress = new JTextField(20);
    public static final  JPasswordField vlDate = new JPasswordField(20);


    public static void main(String[] args) {

        JFrame frame = new JFrame("Combine file IFA & IFC");
        frame.setSize(550, 500);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        JPanel panel = new JPanel();
        frame.add(panel);
        placeComponents(panel);
        frame.setVisible(true);
    }

    public static File[] getInputFile(){
        File[] files;
        JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
        jfc.setDialogTitle("Multiple file and directory selection:");
        jfc.setMultiSelectionEnabled(true);
        jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);

        int returnValue = jfc.showOpenDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            files = jfc.getSelectedFiles();
            return files;

        }else{
            File[] files1 = new File[0];
            return files1;
        }
        
    }

     public static void placeComponents(JPanel panel) {

        panel.setLayout(null);
        JLabel lblProject = new JLabel("Project Number");
        lblProject.setBounds(10,20,200,25);
        panel.add(lblProject);
        JLabel lblAddress = new JLabel("Address");
        lblAddress.setBounds(10,50,200,25);
        panel.add(lblAddress);
        JLabel lblDate = new JLabel("Date");
        lblDate.setBounds(10,80,200,25);
        panel.add(lblDate);
        JLabel lblDescription = new JLabel("Description");
        lblDescription.setBounds(10,110,200,25);
        panel.add(lblDescription);

        ckbIFA.setBounds(200,110,50,25);
        panel.add(ckbIFA);

        ckbIFC.setBounds(260,110,50,25);
        panel.add(ckbIFC);

        vlProject.setBounds(200,20,320,25);
        panel.add(vlProject);

        vlAddress.setBounds(200,50,320,25);
        panel.add(vlAddress);

        vlDate.setBounds(200,80,320,25);
        panel.add(vlDate);

        txtFileInput.setBounds(200, 140, 320, 100);
        panel.add(txtFileInput);

        txtFileTemplate.setBounds(200, 245, 320, 100);
        panel.add(txtFileTemplate);

        vlOutputLocation.setBounds(200,350,320,25);
        panel.add(vlOutputLocation);

        JButton btnLocation = new JButton("Output location");
        btnLocation.setBounds(10, 350, 150, 25);
        panel.add(btnLocation);

        JButton btnInput = new JButton("Select file Input");
        btnInput.setBounds(10, 140, 150, 25);
        panel.add(btnInput);

        JButton btnTemplate = new JButton("Select file Template");
        btnTemplate.setBounds(10, 245, 150, 25);
        panel.add(btnTemplate);

        JButton btnProcess = new JButton("Process");
        btnProcess.setBounds(130, 400, 300, 25);
        panel.add(btnProcess);

        btnInput.addActionListener(e -> {
            File[] file_input = getInputFile();
            txtFileInput.setText("");
            for (File file : file_input) {
                txtFileInput.append(file.getPath() + "\n");
            }
        });

         btnTemplate.addActionListener(e -> {
             File[] file_input = getInputFile();
             txtFileTemplate.setText("");
             for (File file : file_input) {
                 txtFileTemplate.append(file.getPath() + "\n");
             }
         });

         btnLocation.addActionListener(e -> {
             File[] file_input = getInputFile();
             vlOutputLocation.setText("");
             for (File file : file_input) {
                 vlOutputLocation.setText(file.getPath());
             }
         });


    }

}