package xssf;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public abstract class SelectedSheet {

    public static void main(String[]args) throws IOException {
        try (Workbook wb = new XSSFWorkbook()) { //or new HSSFWorkbook();

            wb.createSheet("row sheet");
            wb.createSheet("another sheet");
            Sheet sheet3 = wb.createSheet(" sheet 3 ");
            sheet3.setSelected(true);
            wb.setActiveSheet(2);

            // Create various cells and rows for spreadsheet.

            try (FileOutputStream fileOut = new FileOutputStream(".\\ExcelFile\\selectedSheet.xlsx")) {
                wb.write(fileOut);
            }
        }
    }

}
