package hssf;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
//IT GIVES OUTPUT IN HSSF FORMAT 
/**
 * Shows how various alignment options work.
 */
public class Alignment {
    public static void main(String[] args) throws IOException {
        try (HSSFWorkbook wb = new HSSFWorkbook()) {
            HSSFSheet sheet = wb.createSheet("new sheet");
            HSSFRow row = sheet.createRow(2);
            createCell(wb, row, 0, HorizontalAlignment.CENTER);
            createCell(wb, row, 1, HorizontalAlignment.CENTER_SELECTION);
            createCell(wb, row, 2, HorizontalAlignment.FILL);
            createCell(wb, row, 3, HorizontalAlignment.GENERAL);
            createCell(wb, row, 4, HorizontalAlignment.JUSTIFY);
            createCell(wb, row, 5, HorizontalAlignment.LEFT);
            createCell(wb, row, 6, HorizontalAlignment.RIGHT);

            // Write the output to a file
            try (FileOutputStream fileOut = new FileOutputStream(".\\ExcelFile\\Alignment.xls")) {
                wb.write(fileOut);
            }
        }
    }

    /**
     * Creates a cell and aligns it a certain way.
     *
     * @param wb        the workbook
     * @param row       the row to create the cell in
     * @param column    the column number to create the cell in
     * @param align     the alignment for the cell.
     */
    private static void createCell(HSSFWorkbook wb, HSSFRow row, int column, HorizontalAlignment align) {
        HSSFCell cell = row.createCell(column);
        cell.setCellValue("Align It");
        HSSFCellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(align);
        cell.setCellStyle(cellStyle);
    }
}

