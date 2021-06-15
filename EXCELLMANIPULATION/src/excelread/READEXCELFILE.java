package excelread;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class READEXCELFILE {
	
public void readExcel (String filePath,String fileName,String sheetName) throws IOException {
	
	
	//create an object of a file to open a excel file
	
	File file=new File("C:\\vijay\\test.xlsx");
	//C:\\vijay"+"\\"+"test.xlsx
	
	
	//create an object of FileInputStream class to read excel file
	
	FileInputStream inputStream=new FileInputStream(file);
	Workbook demoWorkbook=null;
	//find the file extension by using substring method
	
	String fileExtensionName=fileName.substring(fileName.indexOf("."));
	if(fileExtensionName.equals(".xlsx")) {
		demoWorkbook= new XSSFWorkbook(inputStream);
	}
		
		else if(fileExtensionName.equals(".xls")) {
			demoWorkbook= new HSSFWorkbook(inputStream);
			
		}
	//read the sheet in side the work book by its name
	
	//Sheet demoSheet=demoWorkbook.getSheet(sheetName);
	
	
	Sheet demoSheet=(Sheet) demoWorkbook.getSheet(sheetName);
	//find number of rows in excel file
	
	int rowCount=((org.apache.poi.ss.usermodel.Sheet) demoSheet).getLastRowNum()-((org.apache.poi.ss.usermodel.Sheet) demoSheet).getFirstRowNum();
	//create a loop over all the rows		
	
	for (int i=0;i<rowCount; i++) {
		
		Row row=((org.apache.poi.ss.usermodel.Sheet) demoSheet).getRow(i);
		
		//Create a loop to print cell values in a row
		for (int j=0;j<row.getLastCellNum(); j++) {
			//print excel data
			
			System.out.println(row.getCell(j).getStringCellValue()+"|| ");
		}
		
		
	}

	
}
	
	public static void main(String[] args) throws IOException {
		//create an object of READEXCELFILE class
		
		READEXCELFILE objExcelFile=new READEXCELFILE();
		
		//String filePath = System.getProperty("C:\\vijay")+"\\src\\excelExportAndFileIO";
		String filePath = System.getProperty("C:\\vijay");
		objExcelFile.readExcel("C:\\vijay", "test.xlsx", "Sheet1");
		
	

	}

}
