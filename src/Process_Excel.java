import java.io.*;
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

    public  void IFA_type1(String path_input, String path_template, String PjNumber, String Address, String Date, String path_output) throws IOException {

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

    public  void IFA_type2(String path_input, String path_template, String PjNumber, String Address, String Date, String path_output) throws IOException {

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



//    public static void main(String args[]) throws IOException
//    {
//        String out = "/Users/nguyenbaolam/Downloads/Safari Download/Copy File/Output";
////        String path = "/Users/nguyenbaolam/Downloads/Safari Download/Copy File/Input/01_Material_List (steel).csv";
//        String path = "/Users/nguyenbaolam/Downloads/Safari Download/Copy File/Input/12_Genis_Transmittal_(layout).csv";
////        String path1 = "/Users/nguyenbaolam/Downloads/Safari Download/Copy File/Template/J-XXX Advance Material list.xls";
//        String path1 = "/Users/nguyenbaolam/Downloads/Safari Download/Copy File/Template/J-XXX Project Address Transmittal001.xls";
////        IFA_type1(path, path1, "ABC", "LAM", "29.02.2020",out);
//        IFA_type2(path, path1, "ABC", "LAM", "29.02.2020",out);
//
//    }
}
