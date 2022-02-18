package xssf;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * An example of how to merge regions of cells.
 */
public class MergingCells {
    public static void main(String[] args) throws IOException {
        try (Workbook wb = new XSSFWorkbook()) { //or new HSSFWorkbook();
            Sheet sheet = wb.createSheet("new sheet");

            Row row = sheet.createRow((short) 1);
            Cell cell = row.createCell((short) 1);
            cell.setCellValue(new XSSFRichTextString("This is a test of merging"));

            sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 2));

            // Write the output to a file
            try (FileOutputStream fileOut = new FileOutputStream(".\\ExcelFile\\merging_cells.xlsx")) {
                wb.write(fileOut);
            }
        }
    }
}
