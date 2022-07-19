import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadExcel3 {

    public static void main(String[] args) throws IOException {
        String excelFilePath=".\\datafiles\\employee sheet1.xlsx";
        FileInputStream fileInputStream=new FileInputStream(excelFilePath);

        XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet= workbook.getSheetAt(0);

        Iterator<Row> iterator=sheet.iterator();

        while(iterator.hasNext())
        {
            XSSFRow row= (XSSFRow) iterator.next();
            Iterator celliterator=row.cellIterator();
            while (celliterator.hasNext())
            {
                XSSFCell cell= (XSSFCell) celliterator.next();
                switch (cell.getCellType())
                {
                    case STRING : System.out.print(cell.getStringCellValue()); break;
                    case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
                    case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
                }
                System.out.print(" | ");
            }
            System.out.println();

        }


    }
}
