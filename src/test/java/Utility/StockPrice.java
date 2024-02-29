package Utility;

import java.io.File;  
import java.io.FileInputStream;
import java.io.IOException;
import java.security.KeyStore.Entry;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.jar.JarException;

import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  

public class StockPrice {

	
		public static void main(String[] args) throws Exception {
				
		{  

			//Reading data from excel file
			try  
			{  
			File file = new File("C:\\Users\\Admin\\Desktop\\LTI MINDTREE\\Assignment.xlsx");     
			FileInputStream fis = new FileInputStream(file);     
			XSSFWorkbook wb = new XSSFWorkbook(fis);   
			XSSFSheet sheet = wb.getSheetAt(0);     
			
			//iterating over excel file
			Iterator<Row> itr = sheet.iterator();      
			while (itr.hasNext())                 
			{  
			Row row = itr.next();  
			Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  
			while (cellIterator.hasNext())   
			{  
			Cell cell = cellIterator.next();  
			switch (cell.getCellType())               
			{  
			case STRING:    //field that represents string cell type  
			System.out.print(cell.getStringCellValue() + "\t\t\t");  
			break;  
			case NUMERIC:    //field that represents number cell type  
			System.out.print(cell.getNumericCellValue() + "\t\t\t");  
			break;  
			default:  
			}  
			}  
			System.out.println("");  
			}  
			}  
			catch(Exception e)  
			{  
			e.printStackTrace();  
			}  
			
			}  

			 
			}
	}




