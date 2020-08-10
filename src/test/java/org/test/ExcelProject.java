package org.test;

import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 

public class ExcelProject {

		public static void main(String[] args) throws Throwable {
			File f = new File("C:\\Users\\NANO SYSTEMS\\eclipse-workspace\\Excel\\target\\book2.xlsx");
			FileInputStream f1 = new FileInputStream(f);
			Workbook w =new XSSFWorkbook(f1);
			Sheet s=w.getSheet("Sheet1");
			for(int i=0;i<s.getPhysicalNumberOfRows();i++) {
				Row row=s.getRow(i);
				for(int j=0;j<row.getPhysicalNumberOfCells();j++) {
					Cell cell =row.getCell(j);
					int c =cell.getCellType();
					if(c==1) {
						
						String value=cell.getStringCellValue();
						System.out.println(value);
					}
					else if(c==0) {
						if(DateUtil.isCellDateFormatted(cell)) {
							Date d=cell.getDateCellValue();
							SimpleDateFormat sd=new SimpleDateFormat("dd/mm/yyyy");
							String Value = sd.format(d);
							System.out.println(Value);
						}
					
					  else
					  {
						 
							double d=cell.getNumericCellValue();
							long L =(long)d;
							String value = String.valueOf(L);
							System.out.println(value);
							
						}
						
						
						}
						}
			    }
		   }
     }
