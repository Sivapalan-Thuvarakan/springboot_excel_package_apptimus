package com.example.demo;

import java.io.File;
import java.io.FileOutputStream;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ch.qos.logback.core.net.SyslogOutputStream;

public class HashMapDemo1 {

//	public static void print(HashMap<String,HashMap<String,String>> x)
//	{	
//		for (String i : x.keySet()) {
//			  System.out.println("key: " + i + " value: " + x.get(i));
//			  for(String y :x.get(i).keySet())
//			  {
////				  System.out.println("key: " + y + " value: " + x.get(i).get(y));
//				  System.out.println(x.get("1").get("Name"));
//			  }
//		}
//
//	}
	
  public static void writeExcel(String excelName,HashMap<String,HashMap<String,String>> x,String[] data)
  {
	  Workbook workbook = new XSSFWorkbook();

	  Sheet sheet = workbook.createSheet(excelName);
//	  sheet.setColumnWidth(1, 4000);

	  Row header = sheet.createRow(0);//created first row
	  CellStyle headerStyle = workbook.createCellStyle();//create column header
	  headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
	  headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	  XSSFFont font = ((XSSFWorkbook) workbook).createFont();
	  int noOfColumn=x.size();
	  int totalDataSize=data.length;
	  int noOfRows=totalDataSize/noOfColumn;
	  int count=0;
		for (String i : x.keySet()) 
		{
//		  System.out.println("key: " + i + " value: " + x.get(i));
			 
			
				  
				  sheet.setColumnWidth(count,Integer.parseInt(x.get(i).get("Width")));
				  font.setFontName(x.get(i).get("FontName"));
				  font.setFontHeightInPoints((short)Integer.parseInt(x.get(i).get("FontHeight")) );
				  font.setBold(Boolean.parseBoolean(x.get(i).get("Bold")));
				  headerStyle.setFont(font);
				  Cell headerCell = header.createCell(count);
				  headerCell.setCellValue(x.get(i).get("Name"));
				  headerCell.setCellStyle(headerStyle);
//				  System.out.println(count);
				  count++;
//				  System.out.println("key: " + y + " value: " + x.get(i).get(y));
				
		}		
		
		 CellStyle style = workbook.createCellStyle();
		 style.setWrapText(true);
		 int rowCount=0;
		 int dataCount=0;
		while(dataCount<totalDataSize)//8
		{
			while(rowCount<noOfRows)
			{
				Row row = sheet.createRow(rowCount+1);//4

				int columnCount=0;
				while(columnCount<noOfColumn)//2
				{
					
					  Cell cell = row.createCell(columnCount);
//					  cell.setCellValue((double) data[dataCount]);
					  cell.setCellValue(data[dataCount]);
					  cell.setCellStyle(style);
					  System.out.println();
					  dataCount++;//1,2
					  columnCount++;//1,2,1,2
				}
				rowCount++;
			}
		}

	  
	  File currDir = new File(".");
	  String path = currDir.getAbsolutePath();
	  String fileLocation = path.substring(0, path.length() - 1) + "test.xlsx";

	 try {
		 FileOutputStream outputStream = new FileOutputStream(fileLocation);
		  workbook.write(outputStream);
		  workbook.close();	
	} catch (Exception e) {
		// TODO: handle exception
	}
  }
}
