package org.excel.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;import org.apache.http.impl.cookie.DateUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelData {

	public static void main(String[] args) throws IOException {
		File loc = new File("D:\\Agil\\PhotonWorkspace\\Excel\\src\\test\\resources\\Excel\\Book2.xlsx");
		FileInputStream st = new FileInputStream(loc);
		Workbook w = new XSSFWorkbook(st);
		
		Sheet sheet = w.getSheet("Sheet1");
		
		Row row = sheet.getRow(4);
		
		Cell c = row.getCell(0);
		System.out.println(c);
		
	
			
				int type = c.getCellType();
				System.out.println(type);
				
				
				
				
				
				if (type==1) {
					String value = c.getStringCellValue();
					System.out.println(value);
					}
				else {
					if(DateUtil.isCellDateFormatted(c)) {
					Date dateCellValue = c.getDateCellValue();
					SimpleDateFormat ddd= new SimpleDateFormat("dd-MMMM-yyyy");
					String format = ddd.format(dateCellValue);
					System.out.println(format);
					}
					
					else {
					double num = c.getNumericCellValue();
					long ln = (long) num;
					String valueOf = String.valueOf(ln);
					System.out.println(valueOf);
				
				
					}
				
				}
			  
			
			
	}
}

