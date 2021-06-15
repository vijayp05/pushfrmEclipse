package excelread;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class READ_EXCEL_FILE_UPDATED {

	public static void main(String[] args) throws IOException {
		
		System.setProperty("webdriver.gecko.driver","C:\\Selenium_2019\\geckodriver.exe");
		WebDriver driver = new FirefoxDriver();		
	            
        //reading of the file 
		FileInputStream fis = new FileInputStream("C:\\vijay\\test.xls");
		
		//workbook object creation and linking to the FileInputSteam class 
		HSSFWorkbook workbook = new HSSFWorkbook(fis);
		
       	//Sheet object creation and reading the sheet which is linked to work book
        HSSFSheet sheet = workbook.getSheetAt(0);
		                        //I have added test data in the cell A1 as "SoftwareTestingMaterial.com"
		                        //Cell A1 = row 0 and column 0. It reads first row as 0 and Column A as 0.
		 Row row = sheet.getRow(0);
		 Cell cell = row.getCell(0);
		                   
		 String cellval=cell.getStringCellValue();
		 
		 
		 Row row1 = sheet.getRow(0);
		 Cell cell1 = row.getCell(1);
		 
		 String cellval1=cell1.getStringCellValue();
		 //System.out.println(cellval);
		 System.out.println(cellval1);
		 
		 
		// launch Fire fox and direct it to the Base URL
		 String baseUrl = "https://login.yahoo.com/";
	        driver.get(baseUrl); 
		 
		//locator by using name
	        driver.findElement(By.name("username")).sendKeys(cellval);
	        driver .findElement(By.name("signin")).click();
	        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
			driver.findElement(By.name("password")).sendKeys(cellval1);
			driver.findElement(By.id("login-signin")).click();
			

	}

}





