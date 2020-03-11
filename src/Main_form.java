import javax.swing.*;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.util.ArrayList;
import java.util.List;

class Main_form {
    public static final  JTextField vlSettingIFA = new JTextField(20);
    public static final  JTextField vlSettingIFC = new JTextField(20);
    public static final JTextArea txtFileInput = new JTextArea();
    public static final JTextArea txtFileTemplate = new JTextArea();
    public static final  JTextField vlOutputLocation = new JTextField(20);
    public static final  JRadioButton ckbIFC = new JRadioButton("IFC");
    public static final  JRadioButton ckbIFA = new JRadioButton("IFA");
    public static final  JTextField vlProject = new JTextField(20);
    public static final  JTextField vlAddress = new JTextField(20);
    public static final  JTextField vlDate = new JTextField(20);
    public static final  JLabel lblResult = new JLabel("");
    public static String[] list_input_IFA = {
            "01_Material_List (steel).csv",
            "2_Genis_Transmittal_(layout).csv"
    };

    public static String[] list_input_IFC = {
            "01_Material_List (steel).csv",
            "05_Assembly_Part_List.csv",
            "04_Bolt_Summary.csv",
            "07_MCS_(assy)_List.csv",
            "13_Genis_Transmittal_(single).csv",
            "11_Genis_Transmittal_(assy).csv",
            "12_Genis_Transmittal_(layout).csv"
    };
    public static String[] list_template_IFA = {
            "J-XXX Advance Material list.xls",
            "J-XXX Project Address Transmittal001.xls"
    };

    public static String[] list_template_IFC = {
            "J-XXX_Delivery List_Rev 0.xls",
            "J-XXX Project Address Transmittal002.xls",
            "J-XXX_Material_List_Rev 0.xls",
            "J-XXX_Assembly_Bolt_List_Rev 0.xls",
            "J-XXX_Bolt_Summary_Rev 0.xls"
    };



    public static void main(String[] args) {

        JFrame frame = new JFrame("Combine file IFA & IFC");
        frame.setSize(850, 800);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        JPanel panel = new JPanel();
        frame.add(panel);
        placeComponents(panel);
        frame.setVisible(true);
    }

    public static String getDirectory(){
        JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
        jfc.setDialogTitle("Directory selection:");
        jfc.setMultiSelectionEnabled(false);
        jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        int returnValue = jfc.showOpenDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            File file = jfc.getSelectedFile();
            return file.getPath();

        }else{

            return "";
        }
    }

    public static File[] getMultiFiles(JFileChooser jfc){
        jfc.setDialogTitle("Multiple file and directory selection:");
        jfc.setMultiSelectionEnabled(true);
        jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);

        int returnValue = jfc.showOpenDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            File[] files = jfc.getSelectedFiles();
            return files;

        }else{
            File[] files1 = new File[0];
            return files1;
        }
    }

    public static List<String> getFileInList(String path, String mode){
        try{
            File folder = new File(path);
            File[] files = folder.listFiles();
            int fileCount = folder.list().length;
            List<String> listfile = new ArrayList<String>();

            if (fileCount == 0){
                return listfile;
            }else{
                for (File f : files)
                {
                    if (f.isFile()) {
                        if (mode.equals("inputIFA")){
                            for (String str: list_input_IFA){
                                if (f.getName().equals(str)){
                                    listfile.add(f.getName());
                                }
                            }
                        }


                        if (mode.equals("inputIFC")){
                            for (String str: list_input_IFC){
                                if (f.getName().equals(str)){
                                    listfile.add(f.getName());
                                }
                            }
                        }


                        if (mode.equals("templateIFA")){
                            for (String str: list_template_IFA){
                                if (f.getName().equals(str)){
                                    listfile.add(f.getName());
                                }
                            }
                        }

                        if (mode.equals("templateIFC")){
                            for (String str: list_template_IFC){
                                if (f.getName().equals(str)){
                                    listfile.add(f.getName());
                                }
                            }
                        }
                    }
                }
                return listfile;
            }
        }catch(Exception ex){
            List<String> listfile1 = new ArrayList<String>();
            return listfile1;
        }
    }


    public static File[] getOutputFile(){
        File[] files;
        JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
        files = getMultiFiles(jfc);
        return files;

    }


    public static String action(){
        try{
            Process_Excel px = new Process_Excel();
            String[] inputs = txtFileInput.getText().split("\n");
            String[] templates = txtFileTemplate.getText().split("\n");
            String final_ = "";
            if (ckbIFA.isSelected()){
                String result = px.IFA_process(vlSettingIFA.getText(), vlSettingIFC.getText(), inputs, templates,vlProject.getText(), vlAddress.getText(),vlDate.getText(),vlOutputLocation.getText() );
                final_ =  result;
            }
            if (ckbIFC.isSelected()){
                String result = px.IFC_process(vlSettingIFA.getText(), vlSettingIFC.getText(), inputs, templates,vlProject.getText(), vlAddress.getText(),vlDate.getText(),vlOutputLocation.getText() );
                final_ =  result;
            }
            return final_;

        }catch(Exception ex){
            return "Fail";
        }
    }

    public static void warn() {
        String mode = "";
        txtFileTemplate.setText("");
        if (ckbIFA.isSelected()){
            mode = "templateIFA";
        };

        if (ckbIFC.isSelected()){
            mode = "templateIFC";
        };
//        System.out.println(mode);
        List<String> file_input =  getFileInList(vlSettingIFC.getText(), mode);
        for (String file : file_input) {
            txtFileTemplate.append(file + "\n");
        }
        txtFileTemplate.setEnabled(false);
    }

    public static void warn1() {
        String mode = "";
        txtFileInput.setText("");
        if (ckbIFA.isSelected()){
            mode = "inputIFA";
        };

        if (ckbIFC.isSelected()){
            mode = "inputIFC";
        };
//        System.out.println(mode);
        List<String> file_input =  getFileInList(vlSettingIFA.getText(), mode);
        for (String file : file_input) {
            txtFileInput.append(file + "\n");
        }
        txtFileInput.setEnabled(false);
    }

     public static void placeComponents(JPanel panel) {

        panel.setLayout(null);
         JLabel lblSettingIfa = new JLabel("Paste link Input ==>");
         lblSettingIfa.setBounds(10, 20, 150, 25);
         panel.add(lblSettingIfa);

         JLabel lblSettingIfc = new JLabel("Paste link Template ==>");
         lblSettingIfc.setBounds(10, 50, 180, 25);
         panel.add(lblSettingIfc);

         vlSettingIFA.setBounds(200,20,620,25);
         vlSettingIFA.setEnabled(true);
         panel.add(vlSettingIFA);


         vlSettingIFC.setBounds(200,50,620,25);
         vlSettingIFC.setEnabled(true);
         panel.add(vlSettingIFC);

        JLabel lblProject = new JLabel("Project Number");
        lblProject.setBounds(10,80,200,25);
        panel.add(lblProject);
        JLabel lblAddress = new JLabel("Address");
        lblAddress.setBounds(10,110,200,25);
        panel.add(lblAddress);
        JLabel lblDate = new JLabel("Date");
        lblDate.setBounds(10,140,200,25);
        panel.add(lblDate);
        JLabel lblDescription = new JLabel("Description");
        lblDescription.setBounds(10,170,200,25);
        panel.add(lblDescription);

        ckbIFA.setBounds(200,170,100,25);
        ckbIFA.setSelected(true);
        ckbIFC.setBounds(360,170,100,25);
         ButtonGroup buttonGroup = new ButtonGroup();
         buttonGroup.add(ckbIFA);
         buttonGroup.add(ckbIFC);

        panel.add(ckbIFC);
        panel.add(ckbIFA);

        vlProject.setBounds(200,80,620,25);
        panel.add(vlProject);

        vlAddress.setBounds(200,110,620,25);
        panel.add(vlAddress);

        vlDate.setBounds(200,140,620,25);
        panel.add(vlDate);

        txtFileInput.setBounds(200, 200, 620, 200);
        txtFileInput.setEnabled(false);
        panel.add(txtFileInput);

        txtFileTemplate.setBounds(200, 405, 620, 200);
        txtFileTemplate.setEnabled(false);
        panel.add(txtFileTemplate);

        vlOutputLocation.setBounds(200,610,620,25);
        vlOutputLocation.setEnabled(true);
        panel.add(vlOutputLocation);

        JButton btnLocation = new JButton("Output location");
        btnLocation.setBounds(10, 610, 150, 25);
        panel.add(btnLocation);

        JLabel lblInput = new JLabel("Input files:");
        lblInput.setBounds(10, 200, 150, 25);
        panel.add(lblInput);

        JLabel lblTemplate = new JLabel("Template files:");
        lblTemplate.setBounds(10, 405, 150, 25);
        panel.add(lblTemplate);

        JButton btnProcess = new JButton("Process");
        btnProcess.setBounds(330, 660, 300, 50);
        panel.add(btnProcess);


         lblResult.setBounds(10, 710, 620, 25);
         panel.add(lblResult);

         btnLocation.addActionListener(e -> {
             File[] file_input = getOutputFile();
             vlOutputLocation.setText("");
             for (File file : file_input) {
                 vlOutputLocation.setText(file.getPath());
                 vlOutputLocation.setEnabled(true);
             }
         });

         btnProcess.addActionListener(e -> {
             String result = action();
             lblResult.setText(result);
             lblResult.setEnabled(false);

         });

         ckbIFA.addActionListener(e ->{
             warn1();
             warn();
         });

         ckbIFC.addActionListener(e ->{
             warn1();
             warn();
         });



         vlSettingIFA.getDocument().addDocumentListener(new DocumentListener() {
             public void changedUpdate(DocumentEvent e) {
                 warn1();
             }
             public void removeUpdate(DocumentEvent e) {
                 warn1();
             }
             public void insertUpdate(DocumentEvent e) {
                 warn1();
             }

         });


         vlSettingIFC.getDocument().addDocumentListener(new DocumentListener() {
             public void changedUpdate(DocumentEvent e) {
                 warn();
             }
             public void removeUpdate(DocumentEvent e) {
                 warn();
             }
             public void insertUpdate(DocumentEvent e) {
                 warn();
             }
         });


    }

}