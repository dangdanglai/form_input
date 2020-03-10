import java.io.*;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import static java.lang.Math.abs;

public class Process_Excel {

    public  Boolean Convert_Number(String cell){
        try{
            double f = Double.parseDouble(cell);
            return true;
        }catch(Exception ex){
            return false;
        }
    }

    public  StringBuffer readCsv(String path) throws IOException {
        File file=new File(path);
        FileReader fr=new FileReader(file);
            BufferedReader br=new BufferedReader(fr);
            StringBuffer sb=new StringBuffer();
            String line;
            while((line=br.readLine())!=null)
            {
                sb.append(line);
                sb.append("\n");
            }
            fr.close();
            return sb;
    }

    private void replaceSingleCellValue(Cell cell, String value, String target, String replace){
        cell.setCellValue(value.replace(target, replace));
    }

    private void setInfoOutput(HSSFSheet sheet, String find, String replace, int startRow, int endRow){
        for (int r = startRow; r < endRow; r++){
            Row row = sheet.getRow(r);
            for (int i = 0; i < row.getPhysicalNumberOfCells(); i++){
                Cell cell = row.getCell(i);
                if (cell.toString().contains(find)){
                    replaceSingleCellValue(cell, cell.getStringCellValue(), find, replace);
                }
            }
        }
    }

    private String getOutputFileName(String path_template,String path_output,  String[] find, String[] replace){
        String OutputFileName = (new File(path_template)).getName();
        for (int i=0; i<find.length; i++){
            OutputFileName = OutputFileName.replace(find[i],replace[i]);
        }

        String OutputDic = path_output + "/" + OutputFileName;
        return OutputDic;
    }

    private HSSFWorkbook getWorkbookFromIndex(String path) throws IOException {
        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(path));
        HSSFWorkbook wb = new HSSFWorkbook(fs);
        return  wb;
    }

    private void copyCsvSimple(HSSFSheet sheet,String[] a,  int startRow, int startCsv, int endCsv) {
        Row row_del = sheet.getRow(startRow);
        sheet.removeRow(row_del);
        for (int r = startCsv; r < endCsv; r++){
            HSSFRow row_add = sheet.createRow((short)(r + (startRow-startCsv)));
            String b[];
            b = a[r].split(",");
            for (int j = 0; j <= b.length -1; j++){
                HSSFCell cell_add = row_add.createCell(j);

                if (b[j].trim()!=null ){
                    if (Convert_Number(b[j].trim())){
                        cell_add.setCellValue(b[j].trim());

                    }else{
                        cell_add.setCellValue(b[j].trim());
                    }

                }else{
                    cell_add.setCellType(Cell.CELL_TYPE_BLANK);
                }
            }
        }
    }

    private void exportOutput(HSSFWorkbook wb, String path) throws IOException {
        FileOutputStream fileOut = new FileOutputStream(path);
        wb.write(fileOut);
        fileOut.close();
    }

    public  void clone_row( HSSFSheet sheet, HSSFRow sourceRow, HSSFRow row_add){
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            HSSFCell oldCell = sourceRow.getCell(i);
            HSSFCell newCell = row_add.createCell(i);

            if (oldCell == null) {
//                newCell = null;
                continue;
            }

            newCell.setCellStyle(oldCell.getCellStyle());

            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }
            newCell.setCellType(oldCell.getCellType());

        }
        int numberMergeCell = sheet.getNumMergedRegions();
        for (int i = 0; i < numberMergeCell; i++) {
            CellRangeAddress cellRangeAddress = sheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(row_add.getRowNum(),
                        (row_add.getRowNum() +
                                (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow()
                                )),
                        cellRangeAddress.getFirstColumn(),
                        cellRangeAddress.getLastColumn());
                sheet.addMergedRegion(newCellRangeAddress);
            }
        }

    }

    private void handleFooter(HSSFSheet sheet, int contentRowLimit, int footerStartRow, int footerEndRow, int csvStart, int csvEnd){
        int shift_range = (csvEnd - csvStart + 1) - contentRowLimit;

        if (shift_range < 0){

            for (int j= abs(shift_range); j > 0; j-- ){
                Row row_del = sheet.getRow(footerStartRow - j);
                sheet.removeRow(row_del);
            }
        }

        sheet.shiftRows(footerStartRow, footerEndRow, shift_range );

        for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
            HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
            if (anchor.getRow1() >= footerStartRow && anchor.getRow1() <= footerEndRow){
                anchor.setRow1(anchor.getRow1() + shift_range);
                anchor.setRow2(anchor.getRow2() + shift_range);
            }
        }

    }

    private void copyScaleDown( HSSFSheet sheet,int[] cellArray, String[] stringValueArray , int rowIndex){
            for (int i = 0; i< cellArray.length; i++){
                Cell cell = sheet.getRow(rowIndex).getCell(cellArray[i]);
//                System.out.println(cellArray[i]);
                cell.setCellValue(stringValueArray[i]);
            }
    }

    private void copyScaleUp( HSSFSheet sheet,int[] cellArray, String[] stringValueArray , int rowIndex){
        HSSFRow sourceRow = sheet.getRow((rowIndex-1));
        HSSFRow row_add = sheet.createRow((short)(rowIndex));
        clone_row(sheet, sourceRow, row_add);
        for (int i = 0; i< cellArray.length; i++){
            Cell cell = sheet.getRow(rowIndex).getCell(cellArray[i]);
            if (stringValueArray[i].contains("##**--")){
                cell.setCellType(Cell.CELL_TYPE_FORMULA);
                cell.setCellFormula(stringValueArray[i].substring(6));
            }else{
                cell.setCellValue(stringValueArray[i]);
            }

        }
    }

    private  void IFA_type1(String path_input, String path_template, String PjNumber, String Address, String Date, String path_output) throws IOException {

        String OutputDic = getOutputFileName(path_template, path_output, new String[]{"J-XXX"}, new String[]{PjNumber});
        HSSFWorkbook wb = getWorkbookFromIndex(path_template);
        HSSFSheet sheet = wb.getSheetAt(0);

        setInfoOutput(sheet, "J-XXX", PjNumber, 0,1);
        setInfoOutput(sheet, "Project Address", Address, 1,2);
        setInfoOutput(sheet, "XX.XX.XXX", Date, 2,3);

        StringBuffer sb=readCsv(path_input);
        String a[] = sb.toString().split("\n");
        copyCsvSimple(sheet, a, 6, 4, a.length );
        exportOutput(wb, OutputDic);
    }

    private  void IFA_type2(String path_input, String path_template, String PjNumber, String Address, String Date, String path_output) throws IOException {

        String OutputDic = getOutputFileName(path_template, path_output, new String[]{"J-XXX","Project Address"}, new String[]{PjNumber,Address});

        HSSFWorkbook wb = getWorkbookFromIndex(path_template);
        HSSFSheet sheet = wb.getSheetAt(0);

        Cell cellPj = sheet.getRow(9).getCell(7);
        cellPj.setCellValue(PjNumber);

        Cell cellDate = sheet.getRow(6).getCell(7);
        cellDate.setCellValue(Date);

        Cell cellAddress = sheet.getRow(9).getCell(2);
        cellAddress.setCellValue(Address);

        StringBuffer sb=readCsv(path_input);
        String a[] = sb.toString().split("\n");

        if (a.length > 15){
            handleFooter(sheet, 15, 36, 39, 1, (a.length-1));
        }

        for (int r = 1; r < a.length; r++){

            if ((r+20)>= 35){
                String b[];
                b = a[r].split(",");
                int[] cellArray = {1, 2, 3, 5, 11};
                String[] cellValueArray = {b[0], b[1], "1", "##**--" + "L" + (r+21) ,b[2]};
                copyScaleUp(sheet, cellArray, cellValueArray, r+20);

            }else {
                String b[];
                b = a[r].split(",");
                int[] cellArray = {1, 2, 3, 11};
                String[] cellValueArray = {b[0], b[1], "1", b[2]};
                copyScaleDown(sheet, cellArray, cellValueArray, r+20);
            }
        }

        if (a.length < 15){
            handleFooter(sheet, 15, 36, 39, 1, (a.length -1));
        }
        HSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
        exportOutput(wb, OutputDic);
    }

    public String IFA_process(String pathFolderInput, String pathFolderTemplate, String[] inputs, String[] templates,String PjNumber, String Address, String Date, String path_output ){
        try{

            for (int j = 0; j <= inputs.length -1; j++){
                if (inputs[j].contains("01_Material_List (steel).csv")){
                    for (int i = 0; i <= templates.length -1; i++){
                        if (templates[i].contains("J-XXX Advance Material list.xls")){
                            IFA_type1(pathFolderInput +"/"+ inputs[j], pathFolderTemplate+"/"+templates[i], PjNumber, Address, Date,path_output);
                        }
                    }
                }

                if (inputs[j].contains("12_Genis_Transmittal_(layout).csv")){
                    for (int i = 0; i <= templates.length -1; i++){
                        if (templates[i].contains("J-XXX Project Address Transmittal001.xls")){
                            IFA_type2(pathFolderInput +"/"+ inputs[j], pathFolderTemplate+"/"+templates[i], PjNumber, Address, Date,path_output);
                        }
                    }
                }
            }

            return "Sucess";

        }catch(Exception ex){
            return "Fail";
        }
    }

    private void IFC__rp_Material_List( String path_input, String path_template, String PjNumber, String Address, String Date, String path_output) throws IOException{
        String OutputDic = getOutputFileName(path_template, path_output, new String[]{"J-XXX"}, new String[]{PjNumber});
        HSSFWorkbook wb = getWorkbookFromIndex(path_template);
        HSSFSheet sheet = wb.getSheetAt(0);

        setInfoOutput(sheet, "J-XXX", PjNumber, 0,1);
        setInfoOutput(sheet, "Address", Address, 1,2);
        setInfoOutput(sheet, "DD.MM.YYYY", Date, 2,3);

        StringBuffer sb=readCsv(path_input);
        String a[] = sb.toString().split("\n");
        copyCsvSimple(sheet, a, 6, 4, a.length );
        exportOutput(wb, OutputDic);
    }

    private void IFC_rp_Assembly_Bolt_List(String path_input, String path_template, String PjNumber, String Address, String Date, String path_output) throws IOException {
        String OutputDic = getOutputFileName(path_template, path_output, new String[]{"J-XXX"}, new String[]{PjNumber});
        HSSFWorkbook wb = getWorkbookFromIndex(path_template);
        HSSFSheet sheet = wb.getSheetAt(0);

        setInfoOutput(sheet, "J-XXX", PjNumber, 2,3);
        setInfoOutput(sheet, "Address", Address, 3,4);
        setInfoOutput(sheet, "DD.MM.YYYY", Date, 3,4);

        StringBuffer sb=readCsv(path_input);
        String a[] = sb.toString().split("\n");
        copyCsvSimple(sheet, a, 8, 7, a.length );
        exportOutput(wb, OutputDic);
    }


    private void IFC_rp_Bolt_Summary(String path_input, String path_template, String PjNumber, String Address, String Date, String path_output) throws IOException {
        String OutputDic = getOutputFileName(path_template, path_output, new String[]{"J-XXX"}, new String[]{PjNumber});
        HSSFWorkbook wb = getWorkbookFromIndex(path_template);
        HSSFSheet sheet = wb.getSheetAt(0);

        setInfoOutput(sheet, "J-XXX", PjNumber, 3,4);
        setInfoOutput(sheet, "Address", Address, 4,5);
        setInfoOutput(sheet, "DD.MM.YYYY", Date, 3,4);

        StringBuffer sb=readCsv(path_input);
        String a[] = sb.toString().split("\n");
        int count = 0;
        int end_row = 0;
        for (String a_child: a){
            end_row += 1;
            if (a_child.equals(" -------------------------------------------------------------------------")){
                count += 1;
            }
            if (count ==4){
                break;
            }

        }

        copyCsvSimple(sheet, a, 9, 8, end_row );
        exportOutput(wb, OutputDic);
    }


    private  void IFC_Delivery_List_Rev(String path_input, String path_template, String PjNumber, String Address, String Date, String path_output) throws IOException {
        String OutputDic = getOutputFileName(path_template, path_output, new String[]{"J-XXX"}, new String[]{PjNumber});

        HSSFWorkbook wb = getWorkbookFromIndex(path_template);
        HSSFSheet sheet = wb.getSheetAt(0);

        setInfoOutput(sheet, "J-XXX", PjNumber, 0,1);
        setInfoOutput(sheet, "Address", Address, 1,2);
        setInfoOutput(sheet, "DD.MM.YYYY", Date, 2,3);

        StringBuffer sb=readCsv(path_input);
        String a[] = sb.toString().split("\n");

        if (a.length > 67){
            handleFooter(sheet, 64, 71, 73, 3, (a.length-1));

            Row row = sheet.getRow(73 + (a.length-3 - 64));
            Cell cellTotal1 = row.getCell(4);
            cellTotal1.setCellType(Cell.CELL_TYPE_FORMULA);
            cellTotal1.setCellFormula("SUM(E8:E"+ (a.length + 4)+")");


            Cell cellTotal2 = row.getCell(7);
            cellTotal2.setCellType(Cell.CELL_TYPE_FORMULA);
            cellTotal2.setCellFormula("SUM(H8:H"+ (a.length + 4)+")");


        }

        for (int r = 3; r < a.length; r++){

            if (r >= 67){
                String b[];
                b = a[r].split(",");
                int[] cellArray = {0,1, 2, 3,4, 5,6,7,8};
                String[] cellValueArray = {b[0], b[1], b[2], b[3], b[4], b[5], b[6], "##**--" + "E" + (r+4) + "*G"+(r+4) ,b[7]};
                copyScaleUp(sheet, cellArray, cellValueArray, r+4);

                Row row4 = sheet.getRow(r+4);
                Cell cell74 = row4.getCell(4);
                cell74.setCellValue(Integer.parseInt(b[4].trim()));
                cell74.setCellType(Cell.CELL_TYPE_NUMERIC);


            }else {
                String b[];
                b = a[r].split(",");
                int[] cellArray = {0,1, 2, 3,4, 5,6,8};
                String[] cellValueArray = {b[0], b[1], b[2], b[3], b[4], b[5], b[6] ,b[7]};
                copyScaleDown(sheet, cellArray, cellValueArray, r+4);
                Row row4 = sheet.getRow(r+4);
                Cell cell74 = row4.getCell(4);
                cell74.setCellValue(Integer.parseInt(b[4].trim()));
                cell74.setCellType(Cell.CELL_TYPE_NUMERIC);
            }
        }

        if (a.length < 67){
            handleFooter(sheet, 64, 71, 73, 3, (a.length -1));
            Row row = sheet.getRow(73 + (a.length-3 - 64));
            Cell cellTotal1 = row.getCell(4);
//            cellTotal1.setCellType(Cell.CELL_TYPE_FORMULA);
            cellTotal1.setCellFormula("SUM(E8:E"+ (a.length + 4)+")");


            Cell cellTotal2 = row.getCell(7);
//            System.out.println(cellTotal2);
//            cellTotal2.setCellType(Cell.CELL_TYPE_FORMULA);
            cellTotal2.setCellFormula("SUM(H8:H"+ (a.length + 4)+")");

        }
        HSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
        exportOutput(wb, OutputDic);
    }


    private  void IFC_Transmittal002(String[] path_input, String path_template, String PjNumber, String Address, String Date, String path_output) throws IOException {

        String OutputDic = getOutputFileName(path_template, path_output, new String[]{"J-XXX","Project Address"}, new String[]{PjNumber,Address});

        HSSFWorkbook wb = getWorkbookFromIndex(path_template);
        HSSFSheet sheet = wb.getSheetAt(0);

        Cell cellPj = sheet.getRow(9).getCell(7);
        cellPj.setCellValue(PjNumber);

        Cell cellDate = sheet.getRow(6).getCell(7);
        cellDate.setCellValue(Date);

        Cell cellAddress = sheet.getRow(9).getCell(2);
        cellAddress.setCellValue(Address);

        StringBuffer sb=readCsv(path_input[0]);
        StringBuffer sb1=readCsv(path_input[1]);
        StringBuffer sb2=readCsv(path_input[2]);


        String a[] = sb.toString().split("\n");
        String a1[] = sb1.toString().split("\n");
        String a2[] = sb2.toString().split("\n");
        ArrayList<String> a_full = new ArrayList<String>();
//        String[] a_full = new String[]{};
        for (int a_c= 1; a_c < a.length; a_c++){
            a_full.add(a[a_c]);
        }
        a_full.add(" , , ");
        for (int a_c= 2; a_c < a1.length; a_c++){
            a_full.add(a1[a_c]);
        }
        a_full.add(" , , ");

        for (int a_c= 2; a_c < a2.length; a_c++){
            a_full.add(a2[a_c]);
        }


        if (a_full.size() > 176){
            handleFooter(sheet, 176, 197, 200, 0, a_full.size()-1);
        }

        for (int r = 0; r < a_full.size(); r++){

            if (r>= 176){
                String b[];
                b = a_full.get(r).split(",");
                int[] cellArray = {1, 2, 3, 5, 11};
                String[] cellValueArray = {b[0], b[1], "1", "##**--" + "L" + (r+21) ,b[2]};
                copyScaleUp(sheet, cellArray, cellValueArray, r+20);

            }else {
                String b[];
                b = a_full.get(r).split(",");
                int[] cellArray = {1, 2, 3, 5, 11};
                String[] cellValueArray = {b[0], b[1], "1","##**--" + "L" + (r+21), b[2]};
                if ((r+21) == 35 || (r+21) == 89 ){
                    copyScaleUp(sheet, cellArray, cellValueArray, r+21);
                }else{
//                    System.out.println(r);
//                    System.out.println(a_full.get(r));
                    copyScaleDown(sheet, cellArray, cellValueArray, r+21);
                }
            }
        }

        if (a.length < 176){
            handleFooter(sheet, 176, 197, 200, 0, a_full.size()-1);
        }
        HSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
        exportOutput(wb, OutputDic);
    }
    
    public static void main(String args[]) throws IOException
    {
//        String path_input ="/Users/nguyenbaolam/Downloads/Safari Download/Copy File-2/IFC/Input/07_MCS_(assy)_List.csv";
        String path_template ="/Users/nguyenbaolam/Downloads/Safari Download/Copy File-2/IFC/Template/J-XXX Project Address Transmittal002.xls";
        String[] path_input ={
                "/Users/nguyenbaolam/Downloads/Safari Download/Copy File-2/IFC/Input/12_Genis_Transmittal_(layout).csv",
                "/Users/nguyenbaolam/Downloads/Safari Download/Copy File-2/IFC/Input/11_Genis_Transmittal_(assy).csv",
                "/Users/nguyenbaolam/Downloads/Safari Download/Copy File-2/IFC/Input/13_Genis_Transmittal_(single).csv"
        };
        String PjNumber = "SHIN";
        String Address = "ABCDEF";
        String path_output= "/Users/nguyenbaolam/Desktop";
        String Date = "10.02.2020";
        Process_Excel px = new Process_Excel();
//        px.IFC_Delivery_List_Rev(path_input,  path_template,  PjNumber,  Address,  Date,  path_output);
        px.IFC_Transmittal002(path_input,  path_template,  PjNumber,  Address,  Date,  path_output);

    }
}
