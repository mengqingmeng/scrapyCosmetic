import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

/**
 * MQM
 * 2018/8/31 11:37
 */
public class ProcessExcel {

    private static final String path = "C:\\Users\\MQM\\WORKSPACE\\JTP_DATA\\陈";
    private static final String Inpath = path+"\\process.xlsx";

    private static final String outPath = path+"\\out.xlsx";
    public static void main(String[] args) {
        try {

            Workbook outWb = new XSSFWorkbook();
            Sheet outSheet = outWb.createSheet();

            int sheetIndex =0;
            int rowIndex = 0;
            InputStream inp = new FileInputStream(Inpath);
            Workbook wb = WorkbookFactory.create(inp);
            Iterator<Sheet> iterator = wb.sheetIterator();
            while (iterator.hasNext()){
                Sheet sheet = iterator.next();
                Row row = sheet.getRow(4);
                if(row!=null){
                    Cell cell = row.getCell(0);
                    System.out.println(rowIndex);
                    Row outRow = outSheet.createRow(rowIndex);
                    Cell outCell = outRow.createCell(0);
                    outCell.setCellValue(cell.toString());
                    rowIndex ++;

                }else{
                    System.out.println("sheetIndex error:"+sheetIndex);
                }

                sheetIndex++;

            }

            OutputStream fileOut = new FileOutputStream(outPath);
            outWb.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
