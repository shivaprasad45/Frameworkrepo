package com.abc.MagentoScript;

import java.io.FileInputStream;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class MagentoScript 
{
	public static XSSFSheet sheet;
	public static WebDriver driver;
	
	public static String getcellvalue(int r,int c)
	{
		XSSFRow row=sheet.getRow(r);
		XSSFCell cell=row.getCell(c);
		String text=cell.getStringCellValue();
		return text;
	}
	
	public static void main(String[] args) throws Exception
	{
		System.setProperty("webdriver.chrome.driver","C:\\Users\\DEEL\\Desktop\\Selenium Components\\New Drivers\\chromedriver.exe");
		FileInputStream fis=new FileInputStream("E:\\Shiva\\HybridDrivenFrameWork\\src\\com\\abc\\Utilities\\Hybrid.xlsx");
		XSSFWorkbook book=new XSSFWorkbook(fis);
		sheet=book.getSheetAt(0);
		int rows=sheet.getPhysicalNumberOfRows();
		System.out.println("Number of rows:"+rows);
		
		for(int i=1;i<rows;i++)
		{
			String action=getcellvalue(i,2);
			System.out.println(action);
			
			switch (action)
			{
				case "open":
				 driver =new ChromeDriver();
				 driver.manage().window().maximize();
				 driver.manage().timeouts().implicitlyWait(30,TimeUnit.SECONDS);
				 break;
				 
				case "navigate":
					driver.get(getcellvalue(i,4));
					break;
				
				case "click":
					driver.findElement(By.xpath(getcellvalue(i,3))).click();
					break;
					
				case "type":
					driver.findElement(By.xpath(getcellvalue(i,3))).sendKeys(getcellvalue(i,4));
					break;
					
				case "close":
					driver.quit();
					break;
					
				default:System.out.println("Invalid Option");
				 break;
			 }
		}
	}
}


