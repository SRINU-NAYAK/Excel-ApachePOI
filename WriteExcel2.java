import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class WriteExcel2 {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet= workbook.createSheet("EmpInfo");


        ArrayList<Object[]> empdata=new ArrayList<Object[]>();
        empdata.add(new Object[]{"EmpId", "Name", "Job", "Age"});
        empdata.add(new Object[]{"101", "David", "Engineer","20"});
        empdata.add(new Object[]{"102", "Smith", "Doctor","21"});
        empdata.add(new Object[]{"103", "Dany", "Teacher","22"});
        empdata.add(new Object[]{"104", "Bravo", "Analyst","23"});
        empdata.add(new Object[]{"105", "Sunny", "Developer","24"});
        empdata.add(new Object[]{"106", "Srikar", "Dancer","25"});
        empdata.add(new Object[]{"107", "Sriman", "Pilot","26"});
        empdata.add(new Object[]{"108", "Siddu", "Trainer","27"});
        empdata.add(new Object[]{"109", "Sharat", "Boxer","28"});
        empdata.add(new Object[]{"110", "Jony", "Dancer","29"});
        empdata.add(new Object[]{"111", "Priya", "Tester","30"});
        empdata.add(new Object[]{"112", "Raja", "Youtuber","31"});
        empdata.add(new Object[]{"113", "Pranav", "Writer","32"});
        empdata.add(new Object[]{"114", "Praveen", "Singer","33"});
        empdata.add(new Object[]{"115", "Raviteja", "Programmer","34"});



        int rownum=0;
        for (Object[] emp:empdata)
        {
            XSSFRow row=sheet.createRow(rownum++);

            int cellnum=0;
            for(Object value:emp)
            {
                XSSFCell cell=row.createCell(cellnum++);

                if(value instanceof String)
                    cell.setCellValue((String)value);
                if(value instanceof Integer)
                    cell.setCellValue((Integer)value);
                if(value instanceof Boolean)
                    cell.setCellValue((Boolean)value);
            }

        }
        String filepath=".\\datafiles\\employee sheet2.xlsx";
        FileOutputStream outstream=new FileOutputStream(filepath);
        workbook.write(outstream);

        outstream.close();

        System.out.println("Employee.xls file is written successfully....");



    }
}
