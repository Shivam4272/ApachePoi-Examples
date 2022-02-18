package hssf;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.FormulaError;
//IT GIVES OUTPUT IN HSSF FORMAT 
public class CellTypes {
    public static void main(String[] args) throws IOException {
        try (HSSFWorkbook wb = new HSSFWorkbook()) {
            HSSFSheet sheet = wb.createSheet("new sheet");
            HSSFRow row = sheet.createRow(2);
            row.createCell(0).setCellValue(1.1);
            row.createCell(1).setCellValue("a string");
            row.createCell(2).setCellValue(true);
            row.createCell(3).setCellErrorValue(FormulaError.NUM);

            // Write the output to a file
            try (FileOutputStream fileOut = new FileOutputStream(".\\ExcelFile\\CellTypes.xls")) {
                wb.write(fileOut);
            }
        }
    }
}
