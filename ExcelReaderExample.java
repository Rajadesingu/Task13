package task13;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReaderExample {
    public static void main(String[] args) {
        try {
            //  Excel file path
            String filePath = "C:\\Users\\PERSONAL\\eclipse-workspace\\ExcelFileOperation\\Firstfile.xlsx";

            FileInputStream inputStream = new FileInputStream(new File(filePath));

            Workbook workbook = WorkbookFactory.create(inputStream);

            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                for (Cell cell : row) {
                    System.out.print(cell.toString() + "\t");
                }
                System.out.println(); // Move to the next line after each row
            }

            inputStream.close();
            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
