package hssf;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFCreationHelper;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.Font;

/**
 * Demonstrates how to create hyperlinks.
 */
public class Hyperlinks {

    public static void main(String[] args) throws IOException  {
        try (HSSFWorkbook wb = new HSSFWorkbook()) {
            HSSFCreationHelper helper = wb.getCreationHelper();

            //cell style for hyperlinks
            //by default hyperlinks are blue and underlined
            HSSFCellStyle hlinkStyle = wb.createCellStyle();
            HSSFFont hlinkFont = wb.createFont();
            hlinkFont.setUnderline(Font.U_SINGLE);
            hlinkFont.setColor(HSSFColorPredefined.BLUE.getIndex());
            hlinkStyle.setFont(hlinkFont);

            HSSFCell cell;
            HSSFSheet sheet = wb.createSheet("Hyperlinks");

            //URL
            cell = sheet.createRow(0).createCell(0);
            cell.setCellValue("URL Link");
            HSSFHyperlink link = helper.createHyperlink(HyperlinkType.URL);
            link.setAddress("https://poi.apache.org/");
            cell.setHyperlink(link);
            cell.setCellStyle(hlinkStyle);

            //link to a file in the current directory
            cell = sheet.createRow(1).createCell(0);
            cell.setCellValue("File Link");
            link = helper.createHyperlink(HyperlinkType.FILE);
            link.setAddress("link1.xls");
            cell.setHyperlink(link);
            cell.setCellStyle(hlinkStyle);

            //e-mail link
            cell = sheet.createRow(2).createCell(0);
            cell.setCellValue("Email Link");
            link = helper.createHyperlink(HyperlinkType.EMAIL);
            //note, if subject contains white spaces, make sure they are url-encoded
            link.setAddress("mailto:poi@apache.org?subject=Hyperlinks");
            cell.setHyperlink(link);
            cell.setCellStyle(hlinkStyle);

            //link to a place in this workbook

            //create a target sheet and cell
            HSSFSheet sheet2 = wb.createSheet("Target Sheet");
            sheet2.createRow(0).createCell(0).setCellValue("Target Cell");

            cell = sheet.createRow(3).createCell(0);
            cell.setCellValue("Worksheet Link");
            link = helper.createHyperlink(HyperlinkType.DOCUMENT);
            link.setAddress("'Target Sheet'!A1");
            cell.setHyperlink(link);
            cell.setCellStyle(hlinkStyle);

            try (FileOutputStream out = new FileOutputStream(".\\ExcelFile\\hssf-Hyperlinks.xls")) {
                wb.write(out);
            }
        }
    }
}
