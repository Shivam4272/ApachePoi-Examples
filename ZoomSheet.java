package hssf;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * Sets the zoom magnication for a sheet.
 */
public class ZoomSheet
{
    public static void main(String[] args) throws IOException {
        try (HSSFWorkbook wb = new HSSFWorkbook()) {
            HSSFSheet sheet1 = wb.createSheet("new sheet");
            sheet1.setZoom(75);   // 75 percent magnification

            try (FileOutputStream fileOut = new FileOutputStream(".\\ExcelFile\\ZoomSheet.xls")) {
                wb.write(fileOut);
            }
        }
    }
}
