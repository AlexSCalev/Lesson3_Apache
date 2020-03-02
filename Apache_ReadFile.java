package Source;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;


public class Apache_ReadFile {

    public void bublicReadFile(String nameFile) {

        try {
            FileInputStream excelFile = new FileInputStream(new File(nameFile));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
// if Iterator hasNext return true we read date in file
            while (iterator.hasNext()) {
// Row pointer == iterator.next() for rows
                Row currentRow = iterator.next();
//                Create pointer column
                Iterator<Cell> cellIterator = currentRow.iterator();
//                Check iterator column
                while (cellIterator.hasNext()) {
// ponter column == interator.next for column
                    Cell currentCell = cellIterator.next();
                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        System.out.print(currentCell.getStringCellValue() + "--");
                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        System.out.print(currentCell.getNumericCellValue() + "--");
                    }

                }
                System.out.println();

            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
