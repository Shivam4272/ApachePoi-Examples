package xssf;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * How to use newlines in cells
 */
public class NewLinesInCells {

    public static void main(String[]args) throws IOException {
        try (Workbook wb = new XSSFWorkbook()) {   //or new HSSFWorkbook();
            Sheet sheet = wb.createSheet();

            Row row = sheet.createRow(2);
            Cell cell = row.createCell(2);
            cell.setCellValue("Use \n with word wrap on to create a new line");

            //to enable newlines you need set a cell styles with wrap=true
            CellStyle cs = wb.createCellStyle();
            cs.setWrapText(true);
            cell.setCellStyle(cs);

            //increase row height to accommodate two lines of text
            row.setHeightInPoints(2 * sheet.getDefaultRowHeightInPoints());

            //adjust column width to fit the content
            sheet.autoSizeColumn(2);

            try (FileOutputStream fileOut = new FileOutputStream(".\\ExcelFile\\newlinesincells.xlsx")) {
                wb.write(fileOut);
            }
        }
    }
}