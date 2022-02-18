package hssf;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * Illustrates how to create cell values.
 */
public class CreateCells {
    public static void main(String[] args) throws IOException {
        try (HSSFWorkbook wb = new HSSFWorkbook()) {
            HSSFSheet sheet = wb.createSheet("new sheet");

            // Create a row and put some cells in it. Rows are 0 based.
            HSSFRow row = sheet.createRow(0);
            // Create a cell and put a value in it.
            HSSFCell cell = row.createCell(0);
            cell.setCellValue(1);

            // Or do it on one line.
            row.createCell(1).setCellValue(1.2);
            row.createCell(2).setCellValue("This is a string");
            row.createCell(3).setCellValue(true);

            // Write the output to a file
            try (FileOutputStream fileOut = new FileOutputStream(".\\ExcelFile\\CreateCells.xls")) {
                wb.write(fileOut);
            }
        }
    }
}
