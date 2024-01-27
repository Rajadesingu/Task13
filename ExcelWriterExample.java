package task13;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriterExample {

    public static void main(String[] args) throws IOException {

        XSSFWorkbook book = new XSSFWorkbook();
        XSSFSheet sheet = book.createSheet();

        Object[][] data = {

                { "Name", "Age", "City" }, 
                { "Ram", 20, "Chennai" }, 
                { "Ragu", 21, "Chennai" }, 
                { "Radha", 22, "Chennai" } 
        };

        int rowCount = 0;
        for (Object[] row1 : data) {

            XSSFRow row = sheet.createRow(rowCount++);

            int columnCount = 0;
            for (Object col : row1) {

                XSSFCell cell = row.createCell(columnCount++);

                if (col instanceof String) {
                    cell.setCellValue((String) col);

                } else if (col instanceof Integer) {

                    cell.setCellValue((Integer) col);

                }

            }
        }

        try (FileOutputStream output = new FileOutputStream("Firstfile.xlsx");) {
            book.write(output);
        } finally {
            book.close();
        }

    }
}
