package xssf;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Demonstrates how to insert pictures in a SpreadsheetML document
 */
@SuppressWarnings({"java:S106","java:S4823","java:S1192"})
public final class WorkingWithPictures {
    private WorkingWithPictures() {}

    public static void main(String[] args) throws IOException {

        //create a new workbook
        try (Workbook wb = new XSSFWorkbook()) {
            CreationHelper helper = wb.getCreationHelper();

            //add a picture in this workbook.
            InputStream is = new FileInputStream(".\\images\\me.jpg");
            byte[] bytes = IOUtils.toByteArray(is);
            is.close();
            int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);

            //create sheet
            Sheet sheet = wb.createSheet();

            //create drawing
            Drawing<?> drawing = sheet.createDrawingPatriarch();

            //add a picture shape
            ClientAnchor anchor = helper.createClientAnchor();
            anchor.setCol1(1);
            anchor.setRow1(1);
            Picture pict = drawing.createPicture(anchor, pictureIdx);

            //auto-size picture
            pict.resize(2);

            //save workbook
            String file = ".\\ExcelFile\\WorkingWithpicture.xlsx";
            try (OutputStream fileOut = new FileOutputStream(file)) {
                wb.write(fileOut);
            }
        }
    }
}
