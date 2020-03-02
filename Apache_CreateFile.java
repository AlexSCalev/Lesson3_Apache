package Source;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class Apache_CreateFile {
    public void bublicCreateApach(String nameFile) {
//        include necessary object
        XSSFWorkbook workbook = new XSSFWorkbook();
//        Create name sheet
        XSSFSheet sheet = workbook.createSheet("Dev depID=1");
//        Create array for date input
        Object[][] datatypes = {
                {"EmpID", "LastName", "FirstName", "BirthDate", "Position", "Skills", "MangerID"},
                {"001", "LastName", "FirstName", "01.01.2000", "Departamnet Manager", "Communication \n java", "0"},
                {" ", " ", " ", " ", " ", " ", " "},
                {"EmpID", "LastName", "FirstName", "BirthDate", "Position", "Skills", "MangerID"},
                {"002", "LastName2", "FirstName2", "01.01.2001",
                        "Developer", "Sleeps only 2 hours per day" +
                        " \n Overtimes without concerns " +
                        "\n Works for food", "001"},
                {"String", "Non-Primitive", "No fixed size"}
        };

        int rowNum = 0;
        System.out.println("Creating excel");
// In cicle read all rows and input date for each in all column
        for (Object[] datatype : datatypes) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
//            Input date in column
            for (Object field : datatype) {
                Cell cell = row.createCell(colNum++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
        }
// Create file and in inside file and write date
        try {
            FileOutputStream outputStream = new FileOutputStream(nameFile);
            workbook.write(outputStream);
            workbook.close();
//            Check error for File not created
        } catch (FileNotFoundException e) {
            e.printStackTrace();
//            Check if pointer exist
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Done");
    }
}
