import javax.swing.*;
import javax.swing.filechooser.FileSystemView;
import java.io.File;

class Main_form {
    public static final  JTextField vlSettingIFA = new JTextField(20);
    public static final  JTextField vlSettingIFC = new JTextField(20);
    public static final JTextArea txtFileInput = new JTextArea();
    public static final JTextArea txtFileTemplate = new JTextArea();
    public static final  JTextField vlOutputLocation = new JTextField(20);
    public static final  JCheckBox ckbIFC = new JCheckBox("IFC");
    public static final  JCheckBox ckbIFA = new JCheckBox("IFA");
    public static final  JTextField vlProject = new JTextField(20);
    public static final  JTextField vlAddress = new JTextField(20);
    public static final  JTextField vlDate = new JTextField(20);
    public static final  JLabel lblResult = new JLabel("");
    public static  String cachePathIFA = "";
    public static  String cachePathIFC = "";



    public static void main(String[] args) {

        JFrame frame = new JFrame("Combine file IFA & IFC");
        frame.setSize(850, 600);
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


    public static File[] getOutputFile(){
        File[] files;

        JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());

        if (ckbIFA.isSelected()){
            jfc.setCurrentDirectory(new File( cachePathIFA+"/Output"));
        }
        if (ckbIFC.isSelected()){
            jfc.setCurrentDirectory(new File( cachePathIFC+"/Output"));
        }
        files = getMultiFiles(jfc);
        return files;

    }

    public static File[] getInputFile(){
        File[] files;

        JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
        if (ckbIFA.isSelected()){
            jfc.setCurrentDirectory(new File( cachePathIFA+"/Input"));
        }
        if (ckbIFC.isSelected()){
            jfc.setCurrentDirectory(new File( cachePathIFC+"/Input"));
        }

        files = getMultiFiles(jfc);
        return files;
        
    }

    public static File[] getTemplateFile(){
        File[] files;

        JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
        if (ckbIFA.isSelected()){
            jfc.setCurrentDirectory(new File( cachePathIFA+"/Template"));
        }
        if (ckbIFC.isSelected()){
            jfc.setCurrentDirectory(new File( cachePathIFC+"/Template"));
        }

        files = getMultiFiles(jfc);
        return files;

    }

    public static String action(){
        try{
            Process_Excel px = new Process_Excel();
            String[] inputs = txtFileInput.getText().split("\n");
            String[] templates = txtFileTemplate.getText().split("\n");
            for (int j = 0; j <= inputs.length -1; j++){
                if (inputs[j].contains("01_Material_List (steel).csv")){
                    for (int i = 0; i <= templates.length -1; i++){
                        if (templates[i].contains("J-XXX Advance Material list.xls")){
                            px.IFA_type1(inputs[j], templates[i], vlProject.getText(), vlAddress.getText(), vlDate.getText(),vlOutputLocation.getText());
                        }
                    }
                }

                if (inputs[j].contains("12_Genis_Transmittal_(layout).csv")){
                    for (int i = 0; i <= templates.length -1; i++){
                        if (templates[i].contains("J-XXX Project Address Transmittal001.xls")){
                            px.IFA_type2(inputs[j], templates[i], vlProject.getText(), vlAddress.getText(), vlDate.getText(),vlOutputLocation.getText());
                        }
                    }
                }
            }

            return "Sucess";

        }catch(Exception ex){
            return "Fail";
        }
    }

     public static void placeComponents(JPanel panel) {

        panel.setLayout(null);
         JButton btnSettingIfa = new JButton("Setting IFA");
         btnSettingIfa.setBounds(10, 20, 150, 25);
         panel.add(btnSettingIfa);

         JButton btnSettingIfc = new JButton("Setting IFC");
         btnSettingIfc.setBounds(10, 50, 150, 25);
         panel.add(btnSettingIfc);

         vlSettingIFA.setBounds(200,20,620,25);
         vlSettingIFA.setEnabled(false);
         panel.add(vlSettingIFA);


         vlSettingIFC.setBounds(200,50,620,25);
         vlSettingIFC.setEnabled(false);
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
        panel.add(ckbIFA);

        ckbIFC.setBounds(360,170,100,25);
        panel.add(ckbIFC);

        vlProject.setBounds(200,80,620,25);
        panel.add(vlProject);

        vlAddress.setBounds(200,110,620,25);
        panel.add(vlAddress);

        vlDate.setBounds(200,140,620,25);
        panel.add(vlDate);

        txtFileInput.setBounds(200, 200, 620, 100);
        txtFileInput.setEnabled(false);
        panel.add(txtFileInput);

        txtFileTemplate.setBounds(200, 305, 620, 100);
        txtFileTemplate.setEnabled(false);
        panel.add(txtFileTemplate);

        vlOutputLocation.setBounds(200,410,620,25);
        vlOutputLocation.setEnabled(false);
        panel.add(vlOutputLocation);

        JButton btnLocation = new JButton("Output location");
        btnLocation.setBounds(10, 410, 150, 25);
        panel.add(btnLocation);

        JButton btnInput = new JButton("Select file Input");
        btnInput.setBounds(10, 200, 150, 25);
        panel.add(btnInput);

        JButton btnTemplate = new JButton("Select file Template");
        btnTemplate.setBounds(10, 305, 150, 25);
        panel.add(btnTemplate);

        JButton btnProcess = new JButton("Process");
        btnProcess.setBounds(330, 460, 300, 50);
        panel.add(btnProcess);


         lblResult.setBounds(10, 510, 620, 25);
         panel.add(lblResult);

         btnSettingIfa.addActionListener(e -> {
//             File[] file_input = getInputFile();
             vlSettingIFA.setText(getDirectory());
             cachePathIFA = vlSettingIFA.getText();
         });

         btnSettingIfc.addActionListener(e -> {
//             File[] file_input = getInputFile();
             vlSettingIFC.setText(getDirectory());
             cachePathIFC = vlSettingIFC.getText();
         });

        btnInput.addActionListener(e -> {
            File[] file_input = getInputFile();
            txtFileInput.setText("");
            for (File file : file_input) {
                txtFileInput.append(file.getPath() + "\n");
                txtFileInput.setEnabled(false);
            }
        });

         btnTemplate.addActionListener(e -> {
             File[] file_input = getTemplateFile();
             txtFileTemplate.setText("");
             for (File file : file_input) {
                 txtFileTemplate.append(file.getPath() + "\n");
                 txtFileTemplate.setEnabled(false);
             }
         });

         btnLocation.addActionListener(e -> {
             File[] file_input = getOutputFile();
             vlOutputLocation.setText("");
             for (File file : file_input) {
                 vlOutputLocation.setText(file.getPath());
                 vlOutputLocation.setEnabled(false);
             }
         });

         btnProcess.addActionListener(e -> {
             String result = action();
             lblResult.setText(result);
             lblResult.setEnabled(false);

         });


    }

}