import java.io.*;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.math.BigDecimal;
import java.text.DecimalFormat;

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
//        File file=new File("/Users/nguyenbaolam/Downloads/Safari Download/Copy File/Input/01_Material_List (steel).csv");
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

    public  void IFA_type1(String path_input, String path_template, String PjNumber, String Address, String Date, String path_output) throws IOException {
        String OutputFileName = (new File(path_template)).getName();
        OutputFileName = OutputFileName.replace("J-XXX",PjNumber);
        String OutputDic = path_output + "/" + OutputFileName;

        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(path_template));
        HSSFWorkbook wb = new HSSFWorkbook(fs);
        HSSFSheet sheet = wb.getSheetAt(0);
//        HSSFRow row;
//        HSSFCell cell;

//        int rows = sheet.getPhysicalNumberOfRows();
//
//        int cols = 0; // No of columns
//        int tmp = 0;

//        for(int i = 0; i < 10 || i < rows; i++) {
//            row = sheet.getRow(i);
//            if(row != null) {
//                tmp = sheet.getRow(i).getPhysicalNumberOfCells();
//                if(tmp > cols) cols = tmp;
//            }
//        }
        Cell cellPj = sheet.getRow(0).getCell(0);
        cellPj.setCellValue(cellPj.toString().replace("J-XXX", PjNumber));

        Cell cellAddress = sheet.getRow(1).getCell(0);
        cellAddress.setCellValue(cellAddress.toString().replace("Project Address", Address));

        Cell cellDate = sheet.getRow(2).getCell(0);
        cellDate.setCellValue(cellDate.toString().replace("XX.XX.XXXX", Date));



        StringBuffer sb=readCsv(path_input);
        String a[] = sb.toString().split("\n");
        Row row_del = sheet.getRow(6);
        sheet.removeRow(row_del);
        for (int r = 4; r < a.length- 1; r++){
            HSSFRow row_add = sheet.createRow((short)(r + 2));

            String b[];
            b = a[r].split(",");
            for (int j = 0; j <= b.length -1; j++){
                HSSFCell cell_add = row_add.createCell(j);

                if (b[j].trim()!=null ){
//                    System.out.println(b[j]+"-----"+Convert_Number(b[j]));
                    if (Convert_Number(b[j].trim())){
//                        HSSFCellStyle cellStyle = wb.createCellStyle();
//                        cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#.0"));
//                        cellStyle.setDataFormat(wb.createDataFormat().getFormat("0.0"));
//                        cell_add.setCellType(Cell.CELL_TYPE_NUMERIC);
//                        double f = Double.parseDouble(b[j].trim());
//                        String rs = String.format("%.2f", new BigDecimal(f));
//                        cell_add.setCellValue((rs));
//                        cell_add.setCellValue(Float.parseFloat(b[j].trim()));
//                        cell_add.setCellStyle(cellStyle);
                        cell_add.setCellValue(b[j].trim());

                    }else{
                        cell_add.setCellValue(b[j].trim());
                    }

                }else{
                    cell_add.setCellType(Cell.CELL_TYPE_BLANK);
                }
            }
        }
        FileOutputStream fileOut = new FileOutputStream(OutputDic);
        wb.write(fileOut);
        fileOut.close();
    }

    public  void clone_row(HSSFWorkbook wb, HSSFSheet sheet, HSSFRow sourceRow, HSSFRow row_add){
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            HSSFCell oldCell = sourceRow.getCell(i);
            HSSFCell newCell = row_add.createCell(i);

            // If the old cell is null jump to next cell
            if (oldCell == null) {
//                newCell = null;
                continue;
            }

            // Copy style from old cell and apply to new cell
            HSSFCellStyle newCellStyle = wb.createCellStyle();
//            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
//            newCell.setCellStyle(newCellStyle);
            newCell.setCellStyle(oldCell.getCellStyle());

            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }



            // Set the cell data type
            newCell.setCellType(oldCell.getCellType());
//            HSSFBorderFormatting border = oldCell.getCellStyle().getBorderBottom();

//            newCellStyle.setBorderBottom(oldCell.getCellStyle().getBorderBottom());
//            newCellStyle.setBorderTop(oldCell.getCellStyle().getBorderTop());
//            newCellStyle.setBorderRight(oldCell.getCellStyle().getBorderRight());
//            newCellStyle.setBorderLeft(oldCell.getCellStyle().getBorderLeft());


//             Set the cell data value
//            switch (oldCell.getCellType()) {
//                case Cell.CELL_TYPE_BLANK:
//                    newCell.setCellValue(oldCell.getStringCellValue());
//                    break;
//                case Cell.CELL_TYPE_BOOLEAN:
//                    newCell.setCellValue(oldCell.getBooleanCellValue());
//                    break;
//                case Cell.CELL_TYPE_ERROR:
//                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
//                    break;
//                case Cell.CELL_TYPE_FORMULA:
//                    newCell.setCellFormula(oldCell.getCellFormula());
//                    break;
//                case Cell.CELL_TYPE_NUMERIC:
//                    newCell.setCellValue(oldCell.getNumericCellValue());
//                    break;
//                case Cell.CELL_TYPE_STRING:
//                    newCell.setCellValue(oldCell.getRichStringCellValue());
//                    break;
//            }
        }
//        System.out.println(sheet.getNumMergedRegions());
        int numberMergeCell = sheet.getNumMergedRegions();
        for (int i = 0; i < numberMergeCell; i++) {
//            System.out.println(i);
            CellRangeAddress cellRangeAddress = sheet.getMergedRegion(i);
//            System.out.println(cellRangeAddress);
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

    public  void IFA_type2(String path_input, String path_template, String PjNumber, String Address, String Date, String path_output) throws IOException {
        String OutputFileName = (new File(path_template)).getName();
        OutputFileName = OutputFileName.replace("J-XXX",PjNumber).replace("Project Address",Address);
        String OutputDic = path_output + "/" + OutputFileName;

        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(path_template));
        HSSFWorkbook wb = new HSSFWorkbook(fs);
        HSSFSheet sheet = wb.getSheetAt(0);


        Cell cellPj = sheet.getRow(9).getCell(7);
        cellPj.setCellValue(PjNumber);

        Cell cellDate = sheet.getRow(6).getCell(7);
        cellDate.setCellValue(Date);

        Cell cellAddress = sheet.getRow(9).getCell(2);
        cellAddress.setCellValue(Address);



        StringBuffer sb=readCsv(path_input);
        String a[] = sb.toString().split("\n");

        if (a.length > 14){
            sheet.shiftRows(36, 39, a.length  - (35 - 21 ) - 2 );
            for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
                HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
//                System.out.println(anchor.getRow1());
//                System.out.println(anchor.getRow2());
                if (anchor.getRow1() == 38 && anchor.getRow1() == 38){
                    anchor.setRow1(anchor.getRow1() + a.length  - (35 - 21 )- 2);
                    anchor.setRow2(anchor.getRow2() + a.length  - (35 - 21 ) -2);
                }


                // if needed you could change column too, using one of these:
                // anchor.setCol1(newColumnInt)
                // anchor.setCol1(anchor.getCol1() + moveColsBy)

            }
        }

//        HSSFRow sourceRow = sheet.getRow(21);

        for (int r = 1; r < a.length; r++){

            if ((r+20)>= 35){

                HSSFRow sourceRow = sheet.getRow((r+20-1));
                HSSFRow row_add = sheet.createRow((short)(r + 20));
//                HSSFRow row_add = sheet.createRow((short)(34));

                clone_row(wb, sheet, sourceRow, row_add);
//
                String b[];
                b = a[r].split(",");
                Cell cell0 = sheet.getRow(r+20).getCell(1);
                Cell cell1 = sheet.getRow(r+20).getCell(2);
                Cell cell2 = sheet.getRow(r+20).getCell(3);
                Cell cell3 = sheet.getRow(r+20).getCell(5);
                Cell cell5 = sheet.getRow(r+20).getCell(11);
                cell0.setCellValue(b[0]);
                cell1.setCellValue(b[1]);
                cell2.setCellValue("1");
                cell5.setCellValue(b[2]);
                String formula= "L" + (r+21);
                cell3.setCellType(Cell.CELL_TYPE_FORMULA);
                cell3.setCellFormula(formula);

            }else {
                String b[];
                b = a[r].split(",");
//                for (int i = 0; i < sheet.getRow((r+20)).getLastCellNum(); i++) {
//                    System.out.println(i);
//                    System.out.println(sheet.getRow((r+20)).getCell(i));
//                }
//
                Cell cell0 = sheet.getRow(r+20).getCell(1);
                Cell cell1 = sheet.getRow(r+20).getCell(2);
                Cell cell2 = sheet.getRow(r+20).getCell(3);
                Cell cell5 = sheet.getRow((r+20)).getCell(11);
                cell0.setCellValue(b[0]);
                cell1.setCellValue(b[1]);
                cell2.setCellValue("1");
                cell5.setCellValue(b[2]);


            }
        }
        int del_row = 21 + a.length;

        if (del_row < 35){
            for (int x= 0; x < (37 - del_row);x++ ){
                Row row_del = sheet.getRow(del_row-1 +x);
                sheet.removeRow(row_del);
            }
            sheet.shiftRows(36, 39, -1 *(35 - a.length - 20 + 1) );

            for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
                HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
//                System.out.println(anchor.getRow1());
//                System.out.println(anchor.getRow2());
                if (anchor.getRow1() == 38 && anchor.getRow1() == 38){
                    anchor.setRow1(anchor.getRow1() + -1 *(35 - a.length - 20 + 1));
                    anchor.setRow2(anchor.getRow2() + -1 *(35 - a.length - 20 + 1));
                }


                // if needed you could change column too, using one of these:
                // anchor.setCol1(newColumnInt)
                // anchor.setCol1(anchor.getCol1() + moveColsBy)

            }
        }



//        HSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
        FileOutputStream fileOut = new FileOutputStream(OutputDic);
        wb.write(fileOut);
        fileOut.close();
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
