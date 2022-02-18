package hssf;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.WorkbookUtil;

/**
 * Creates a new workbook with a sheet that's been explicitly defined.
 */
public abstract class NewSheet {
    public static void main(String[] args) throws IOException {
        try (HSSFWorkbook wb = new HSSFWorkbook()) {
            wb.createSheet("new sheet");
            // create with default name
            wb.createSheet();
            final String name = "second sheet";
            // setting sheet name later
            wb.setSheetName(1, WorkbookUtil.createSafeSheetName(name));

            try (FileOutputStream fileOut = new FileOutputStream(".\\ExcelFile\\NewSheet.xls")) {
                wb.write(fileOut);
            }
        }
    }
}

