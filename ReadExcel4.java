import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadExcel4 {
    public static void main(String[] args) throws IOException {
        String excelFilePath = ".\\datafiles\\sheet.xlsx";
        FileInputStream fileInputStream = new FileInputStream(excelFilePath);

        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        Iterator<Sheet> sheet = workbook.sheetIterator();

        while (sheet.hasNext()) {
            Sheet sh = sheet.next();
            System.out.println("Sheet name is:" + sh.getSheetName());
            Iterator<Row> iterator = sh.iterator();

            while (iterator.hasNext()) {
                XSSFRow row = (XSSFRow) iterator.next();
                Iterator celliterator = row.cellIterator();
                while (celliterator.hasNext()) {
                    XSSFCell cell = (XSSFCell) celliterator.next();
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue());
                            break;
                    }
                    System.out.print(" | ");
                }
                System.out.println();
            }
        }
        workbook.close();
    }
}