package cbox;

import java.io.*;
import java.util.Iterator;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.util.Units;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import javax.xml.namespace.QName;

class checkbox1 
{

	public checkbox1() throws Exception 
	{
		XSSFWorkbook wb  = (XSSFWorkbook)WorkbookFactory.create(new FileInputStream(".\\ExcelFile\\checkbox.xlsx"));
		Sheet sheet = wb.getSheetAt(0);
	
	//manually we have to enter no of rows and columns in which checkbox exists
		for (int i=0;i<=9;i++) 
		{
			Row row=sheet.getRow(i);
			if(row==null)
			{
				row=sheet.createRow(i);
				for(int k=0;k<2;k++)
				{
				Cell cell=row.createCell(k);
				getControlAt((XSSFCell)cell);
				System.out.print("\t");
				}
				System.out.println();
				
			}
			else
			{
			for (int c = 0; c <2; c++) 
			{
				Cell cell = row.getCell(c);
				if(cell!=null)
				{
					System.out.print(cell);	
							
				}			
				else
				{
					
					cell = row.createCell(c);
					getControlAt((XSSFCell)cell);
				}
						
				System.out.print("\t");
			}
			System.out.println();
			}
		}
		
		
		wb.close();
	}

 public void getControlAt(XSSFCell cell) throws Exception 
 {
	 XSSFSheet sheet = cell.getSheet();
	 Row row =  cell.getRow();
	 int r = row.getRowNum();
	 int c = cell.getColumnIndex();
	 

	 int drheight = (int)Math.round(sheet.getDefaultRowHeightInPoints() * Units.PIXEL_DPI / Units.POINT_DPI);
	 int rheight = (int)Math.round(row.getHeightInPoints() * Units.PIXEL_DPI / Units.POINT_DPI);
	 row = null;
	 if(r > 0) row = sheet.getRow(r-1);
	 int rheightbefore = (row!=null)?(int)Math.round(row.getHeightInPoints() * Units.PIXEL_DPI / Units.POINT_DPI):drheight;
	 row = sheet.getRow(r+1);
	 int rheightafter = (row!=null)?(int)Math.round(row.getHeightInPoints() * Units.PIXEL_DPI / Units.POINT_DPI):drheight;

	 String name = null;
	 String objectType = null;
	 String checked = null;
	 
	 XmlCursor xmlcursor = null;
	 if (sheet.getCTWorksheet().getLegacyDrawing() != null)
	 {
		 String legacyDrawingId = sheet.getCTWorksheet().getLegacyDrawing().getId();
		 POIXMLDocumentPart part = sheet.getRelationById(legacyDrawingId);
		 XmlObject xmlDrawing = XmlObject.Factory.parse(part.getPackagePart().getInputStream());
		 xmlcursor = xmlDrawing.newCursor();
		 QName qnameClientData = new QName("urn:schemas-microsoft-com:office:excel", "ClientData", "x");
		 QName qnameAnchor = new QName("urn:schemas-microsoft-com:office:excel", "Anchor", "x");
		 boolean controlFound = false;
		 while (xmlcursor.hasNextToken()) 
		 {
			 XmlCursor.TokenType tokentype = xmlcursor.toNextToken();
			 if (tokentype.isStart())
			 {
				 if (qnameClientData.equals(xmlcursor.getName())) 
				 {
					 controlFound = true;
					 XmlObject clientdata = xmlcursor.getObject();
					 XmlObject[] xmlchecked = clientdata.selectPath("declare namespace x='urn:schemas-microsoft-com:office:excel' x:Checked");
					 if (xmlchecked.length > 0) 
					 {
						 checked = "Checked";
					 } 
					 else
					 {
						 checked = "Not checked";
					 }
					 while (xmlcursor.hasNextToken()) 
					 {
						 tokentype = xmlcursor.toNextToken(); 
						 if (tokentype.isAttr()) 
						 {
							 if (new QName("ObjectType").equals(xmlcursor.getName())) 
							 {
								 objectType = xmlcursor.getTextValue();
								 name = objectType + " in row " + (r+1);
							 } 	
						 }	 
						 else 
						 {
							 break;
						 }	
					 }	
				 } 
				 else if (qnameAnchor.equals(xmlcursor.getName()) && controlFound) 
				 {
					 controlFound = false;
					 String anchorContent = xmlcursor.getTextValue().trim();
					 String[] anchorparts = anchorContent.split(",");
					 int fromCol = Integer.parseInt(anchorparts[0].trim());
					 int fromColDx = Integer.parseInt(anchorparts[1].trim());
					 int fromRow = Integer.parseInt(anchorparts[2].trim());
					 int fromRowDy = Integer.parseInt(anchorparts[3].trim());
					 int toCol = Integer.parseInt(anchorparts[4].trim());
					 int toColDx = Integer.parseInt(anchorparts[5].trim());
					 int toRow = Integer.parseInt(anchorparts[6].trim());
					 int toRowDy = Integer.parseInt(anchorparts[7].trim());

					 if (fromCol == c /*needs only starting into the column*/
					&& (fromRow == r || (fromRow == r-1 && fromRowDy > rheightbefore/2f)) 
					&& (toRow == r || (toRow == r+1 && toRowDy < rheightafter/2f))) 
					 {
//System.out.print(fromCol + ":" +fromColDx + ":" + fromRow + ":" + fromRowDy + ":" + toCol + ":" + toColDx + ":" + toRow + ":" + toRowDy);
						 break;
					 }
				 } 
			 } 
		 }
	 }

	 if (xmlcursor!=null && xmlcursor.hasNextToken())
	 {
		
		 System.out.print(name + ":r/c:" +r+ "/" + c + ":" + checked);
	 }
	 else
	 {
		 //System.out.print("NULL" + ":r/c:" +0+ "/" + 0 + ":" + "NULL");
	 }
	
 }

 public static void main(String[] args) throws Exception
 {
	 checkbox1 o = new checkbox1();
 
 }

}

