package excelread;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.Table.Cell;

public class WRITE_EXCEL {

	public static void main(String[] args) throws IOException {
		
		//create an object of Workbook and pass the FileInputStream object into it to create a pipeline between the sheet and eclipse.
		 FileInputStream fis = new FileInputStream("C:\\vijay\\Test.xls");
		 HSSFWorkbook workbook = new HSSFWorkbook();
		 
		
		HSSFSheet sheet = workbook.getSheetAt(0);
		
		                Row row = sheet.createRow(0);
		             org.apache.poi.ss.usermodel.Cell cell = row.createCell(0);
		 
		 cell.setCellValue("Writing excell sheet test");
		 FileOutputStream fos = new FileOutputStream("C:\\vijay\\Test.xls");
		 workbook.write(fos);
		 fos.close();
		 System.out.println("END OF WRITING DATA IN EXCEL");	
		

	}

}
