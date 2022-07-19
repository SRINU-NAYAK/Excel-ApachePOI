import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class WriteExcel1 {

    public static void main(String[] args) throws IOException {

        //workbook->sheet->rows->cells
        XSSFWorkbook workbook= new XSSFWorkbook();
        XSSFSheet sheet1=workbook.createSheet("EmpInfo1");
        XSSFSheet sheet2=workbook.createSheet("EmpInfo2");
        XSSFSheet sheet3=workbook.createSheet("EmpInfo3");
        XSSFSheet sheet4=workbook.createSheet("EmpInfo4");
        XSSFSheet sheet5=workbook.createSheet("EmpInfo5");



        Object empdata[][]= {	{"EmpID", "Name", "Job","Age"},
                {"101", "David", "Engineer","36"},
                {"102", "Smith", "Doctor","39"},
                {"103", "Scott", "Analyst","40"},
                {"104", "Bravo", "Analyst","23"},
                {"105", "Sunny", "Developer","24"},
                {"106", "Srikar", "Dancer","25"},
                {"107", "Sriman", "Pilot","26"},
                {"108", "Siddu", "Trainer","27"},
                {"109", "Sharat", "Boxer","28"},
                {"110", "Jony", "Dancer","29"},
                {"111", "Priya", "Tester","30"},
                {"112", "Raja", "Youtuber","31"},
                {"113", "Pranav", "Writer","32"},
                {"114", "Praveen", "Singer","33"},
                {"115", "Raviteja", "Programmer","34"},
        };

        int rows=empdata.length;
        int cols=empdata[0].length;

        System.out.println(rows);
        System.out.println(cols);

        for(int r=0;r<rows;r++) //0
        {
            XSSFRow row=sheet1.createRow(r);

            for(int c=0;c<cols;c++)  //0
            {
                XSSFCell cell=row.createCell(c);
                Object value=empdata[r][c];

                if(value instanceof String)
                    cell.setCellValue((String)value);
                if(value instanceof Integer)
                    cell.setCellValue((Integer)value);
                if(value instanceof Boolean)
                    cell.setCellValue((Boolean)value);
            }
        }

        String filepath=".\\datafiles\\employee sheet1.xlsx";
        FileOutputStream outstream=new FileOutputStream(filepath);
        workbook.write(outstream);

        outstream.close();

        System.out.println("Employee.xls file is written successfully....");

    }

}
