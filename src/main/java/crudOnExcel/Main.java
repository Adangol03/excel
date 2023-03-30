package crudOnExcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

// how to read and write on excel file.
public class Main {
    public static void main(String[] args) {
        try {
            File file = new File("E:\\MySQL\\SQL_EXAM_26_02_23\\SQL_26_02_23_xlsx\\Physician.xlsx"); //Reading xls file
            FileInputStream fileInputStream = new FileInputStream(file);

            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream); //Create workbook instance to capture the content of the excel sheet
            XSSFSheet xssfSheet = workbook.getSheetAt(0);//get the first Sheet from excel

            Iterator<Row> rowIterator = xssfSheet.iterator(); //Iterate over each row one by one
            while (rowIterator.hasNext()) {//Fetch row till the condition is true
                Row row = rowIterator.next();//each row

                Iterator<Cell> cellIterator = row.cellIterator();//fetch the cell/column corresponding to each row
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue()+"\t");
                            break;
                    }
                }
                System.out.println("\n");
            }


        } catch (Exception e) {
            e.printStackTrace();
        }

        // write on that file.

    }
}