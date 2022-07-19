import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.security.Key;
import java.util.HashMap;
import java.util.Map;

public class ExcelToHashMap {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook=new XSSFWorkbook(".\\datafiles\\student data.xlsx");
        XSSFSheet sheet= workbook.getSheet("student data");

        int rows= sheet.getLastRowNum();

        HashMap<String,String>  data=new HashMap<String,String>();

        //Reading data from excel to HashMap
        for (int r = 0; r<= rows; r++)
        {
            String key=sheet.getRow(r).getCell(0).getStringCellValue();
            String value=sheet.getRow(r).getCell(1).getStringCellValue();
            data.put(key,value);
        }
        //Reading data from HashMap
        for (Map.Entry entry:data.entrySet())
        {
            System.out.println(entry.getKey()+"     "+entry.getValue());
        }




    }
}
