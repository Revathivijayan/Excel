package org.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;


public class ExcelInfo {
	
public static WebDriver driver;
	public static void main(String[] args) throws Throwable {
		int i=0;
		System.setProperty("webdriver.chrome.driver","C:\\Users\\NANO SYSTEMS\\eclipse-workspace\\Web.com.test\\chromedriv\\chromedriver.exe");
		   driver = new ChromeDriver();
		   driver.get("https://www.flipkart.com/");
		   String source = driver.getCurrentUrl();
		   System.out.println(source);
		      WebElement element = driver.findElement(By.xpath("//button[text()='✕']"));
		      element.click();
		      WebElement ele = driver.findElement(By.xpath("//input[@name='q']"));
		      ele.sendKeys("mi phone");
		      WebElement ele1 = driver.findElement(By.xpath("//button[@type='submit']"));
		      ele1.click();
		      Thread.sleep(3000);
		      List<WebElement>redmi1= driver.findElements(By.xpath("(//div[(@class='_3wU53n')])"));

		      File f = new File("C:\\Users\\NANO SYSTEMS\\eclipse-workspace\\Excel\\target\\book2.xlsx");
				Workbook w =new XSSFWorkbook();
				Sheet s=w.createSheet("write");
				for (WebElement row : redmi1) {
					String text = row.getText();
					Row r2 = s.createRow(i);
					Cell cell =r2.createCell(0);
					cell.setCellValue(text);
					i++; 
    }
	     FileOutputStream f1 = new FileOutputStream(f);
	     w.write(f1);
	     System.out.println("done");
	     driver.findElement(By.xpath("(//div[(@class='_3wU53n')])[4]")).click();
	     String handle = driver.getWindowHandle();
	     Set<String>set = driver.getWindowHandles();
	     for(String s1:set) {
	    	 if(!s1.equals(handle)) {
	    		 driver.switchTo().window(s1);
	    	 }
	     }
	     String text = driver.findElement(By.xpath("//span[@class='_35KyD6']")).getText();
	     System.out.println(text);
	     File fs = new File("C:\\\\Users\\\\NANO SYSTEMS\\\\eclipse-workspace\\\\Excel\\\\target\\\\book2.xlsx");
	     FileInputStream f2 = new FileInputStream(fs);
	     Workbook w1 =new XSSFWorkbook(f2);
			Sheet s2=w1.getSheet("write");
			for (i=0;i<s2.getPhysicalNumberOfRows();i++) {
				Row row = s2.getRow(i);
				for(int j=0;j<row.getPhysicalNumberOfCells();j++) {
					Cell cell = row.getCell(j);
					int type = cell.getCellType();
			if(type==1) {
				
			String value = cell.getStringCellValue();
			System.out.println("two values equals");
			}
			else {
				System.out.println("not equales");
			}
				}
			}
	     
	     
	}  
}
	