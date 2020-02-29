import java.io.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import java.text.DecimalFormat;

public class Process_Excel {
    public static final String path = new String();

    public static Boolean Convert_Number(String cell){
        try{
            double f = Double.parseDouble(cell);
            return true;
        }catch(Exception ex){
            return false;
        }
    }

    public static StringBuffer template1() throws IOException {
            File file=new File("/Users/nguyenbaolam/Downloads/Safari Download/Copy File/Input/01_Material_List (steel).csv");
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

    public static void abc() throws IOException {
        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream("/Users/nguyenbaolam/Downloads/Safari Download/Copy File/Template/J-XXX Advance Material list.xls"));
        HSSFWorkbook wb = new HSSFWorkbook(fs);
        HSSFSheet sheet = wb.getSheetAt(0);
        HSSFRow row;
        HSSFCell cell;

        int rows = sheet.getPhysicalNumberOfRows();

        int cols = 0; // No of columns
        int tmp = 0;

        for(int i = 0; i < 10 || i < rows; i++) {
            row = sheet.getRow(i);
            if(row != null) {
                tmp = sheet.getRow(i).getPhysicalNumberOfCells();
                if(tmp > cols) cols = tmp;
            }
        }

        StringBuffer sb=template1();
        String a[] = sb.toString().split("\n");
        Row row_del = sheet.getRow(6);
        sheet.removeRow(row_del);
        for (int r = 4; r < a.length- 1; r++){
            Row row_add = sheet.createRow((short)(r + 2));

            String b[];
            b = a[r].split(",");
            System.out.println(a[r]);
            for (int j = 0; j <= b.length -1; j++){
                Cell cell_add = row_add.createCell(j);

                System.out.println(b[j]);
                if (b[j].trim()!=null ){
                    if (Convert_Number(b[j])){
                        CellStyle cellStyle = wb.createCellStyle();
                        cellStyle.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("#.#"));
                        cell_add.setCellValue(Double.parseDouble(b[j]));
                        cell_add.setCellStyle(cellStyle);
                        cell_add.setCellType(Cell.CELL_TYPE_NUMERIC);
                    }
                    cell_add.setCellValue(b[j]);
                }else{
                    cell_add.setCellType(Cell.CELL_TYPE_BLANK);
                }
            }
        }
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    public static void main(String args[]) throws IOException
    {
//        StringBuffer sb=template1();
//        String a[] = sb.toString().split("\n");
//        for (int i = 4; i <= a.length -1; i++){
////            System.out.println(a[i]);
//            String b[];
//            b = a[i].split(",");
//            for (int j = 0; j <= b.length -1; j++){
//                System.out.println(b[j].trim());
//            }
//        }
        abc();

    }
}
