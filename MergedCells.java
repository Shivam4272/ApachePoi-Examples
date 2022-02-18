package hssf;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * An example of how to merge regions of cells.
 */
public class MergedCells {
   public static void main(String[] args) throws IOException {
        try (HSSFWorkbook wb = new HSSFWorkbook()) {
             HSSFSheet sheet = wb.createSheet("new sheet");

             HSSFRow row = sheet.createRow(1);
             HSSFCell cell = row.createCell(1);
             cell.setCellValue("This is a test of merging");

             sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 2));

             // Write the output to a file
             try (FileOutputStream fileOut = new FileOutputStream(".\\ExcelFile\\MergedCells.xls")) {
                  wb.write(fileOut);
             }
        }
    }
}

