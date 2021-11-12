
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class ReadExcel {

    public static void main(String[] args) throws FileNotFoundException, IOException  {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream("excel.xls"));
        HSSFSheet sheet = workbook.getSheetAt(0);
        int maxrow = sheet.getLastRowNum();
        //System.out.println(maxrow);

        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()){
            Row nextRow = rowIterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            while (cellIterator.hasNext()){
                Cell cell = cellIterator.next();
                System.out.print(cell.getCellType().toString()+" ");

                switch (cell.getCellType()) {
                    case STRING:
                        System.out.println(cell.getStringCellValue()+" ");
                        break;
                    case NUMERIC:
                        System.out.println(cell.getDateCellValue()+" ");
                        break;
                    case BOOLEAN:
                        System.out.println(cell.getBooleanCellValue()+" ");
                        break;

                }
            }
            System.out.println();
        }

    }
}
