import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;


public class ReadSumtelExcel {

    public static void main(String[] args) throws FileNotFoundException, IOException {
        //Создать объект Workbook sumtel для файла sumtel.xlsx
        XSSFWorkbook sumtel = new XSSFWorkbook(new FileInputStream("sumtel.xlsx"));
        //Создать объет лист для нулевого листа книги
        XSSFSheet sheet = sumtel.getSheetAt(0);
        //Опредилить максимально колисчество строк на листе
        int maxRow = sheet.getLastRowNum();
        int iRow = 1;
        //В цыкле считываем строки пока не доудем до максимальной
        while (iRow <= maxRow){
            //Читаем сроку с номером iRow
            XSSFRow row = sheet.getRow(iRow);
            // Выводим строку на экран
            // Номер строки
            System.out.print(iRow+" ");
            // ФИО абонента
            System.out.print(row.getCell(1).getStringCellValue()+" ");
            //Название улицы
            System.out.print(row.getCell(2).getStringCellValue().toUpperCase()+" ");
            //Номер дома
            if (row.getCell(3).getCellType() == CellType.NUMERIC) {
                System.out.print(row.getCell(3).getNumericCellValue());
            } else if (row.getCell(3).getCellType() == CellType.STRING) {
                System.out.print(row.getCell(3).getStringCellValue());
            }
            System.out.print(" ");
            // Номер квартиры
            if (row.getCell(4).getCellType() == CellType.NUMERIC) {
                System.out.println(row.getCell(4).getNumericCellValue());
            } else if (row.getCell(4).getCellType() == CellType.STRING){
                System.out.println(row.getCell(4).getStringCellValue());
            }
            iRow++;

        }

    }
}
