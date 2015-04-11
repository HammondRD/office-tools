package excel;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class ReadExcel {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try {
			
		  //  FileInputStream file = new FileInputStream(new File("LatiKatalog_29_ 06_14-v2.xlsx"));
		    long startTime1 = System.nanoTime();  

		   // XSSFWorkbook workbook = new XSSFWorkbook(file);
		   
		    File file = new File("laps.xlsx");
		    OPCPackage opcPackage = OPCPackage.open(file.getAbsolutePath());
		    XSSFWorkbook workbook = new XSSFWorkbook(opcPackage);
		     long stopTime2 = System.nanoTime();
		    //Get first sheet from the workbook
		    XSSFSheet sheet = workbook.getSheetAt(0);
		    int lasrRow, lasCol;
		    int firstRow = sheet.getFirstRowNum();
		    lasrRow= sheet.getLastRowNum();
		   
		    System.out.println("Last row num: " + lasrRow);
		    
		   
		    System.out.println("First row num: " + firstRow);
		    
		    System.out.println("First column  " + sheet.getRow(firstRow+1).getFirstCellNum());
		    System.out.println("Last column  " + sheet.getRow(firstRow+1).getLastCellNum());
		    lasCol = sheet.getRow(firstRow).getPhysicalNumberOfCells()+1;
		   // String cellValue = 
		     for(int i=0; i<lasCol; i++){
		    	 System.out.print(sheet.getRow(firstRow).getCell(i)+" | ");
		    }
		     System.out.println("");
		    //Iterate through each rows from first sheet
		   
		    Iterator<Row> rowIterator = sheet.iterator(); 
		  
		   int bCount = 0,numCount = 0,stringCount = 0, rowCount=0, cellCount=0;
		  String[] dbArray;
		  dbArray = new  String[lasrRow];
		 
		   //dbArray = new String[][100000];
		  long startTime = System.nanoTime();
		  
		 for( rowCount = 1;rowCount< lasrRow;rowCount++) {
		        Row row = rowIterator.next();
		        ;
		       // row.
		        //For each row, iterate through each columns
		        Iterator<Cell> cellIterator = row.cellIterator();
		   
		             
		            Cell cell = cellIterator.next();
		            dbArray[rowCount] =  cell.toString();
		     
		            	/*System.out.print("|| From cell - " + cell.toString());
		            	System.out.print("||   ****  From array - " +dbArray[rowCount][cellCount]);*/
		          
		          
		     //   System.out.println("");
		    }
		  long stopTime = System.nanoTime();
	
		 
		  
		   // System.out.print("Boolean: "+bCount +" *  Numeric: "+ numCount + " *  String: "+ stringCount);
		  opcPackage.close();
		    
		  /*  FileOutputStream out = new FileOutputStream(new File("test.xls"));
		    workbook.write(out);
		    out.close();*/
		    long duration = stopTime - startTime;
		    System.out.println("DONE in " + duration/1000000000.0000 + " Sec");
		    duration = stopTime2 - startTime1;
		    System.out.println("Total DONE in " + duration/1000000000.0000 + " Sec");
		} catch (FileNotFoundException e) {
		    e.printStackTrace();
		    System.out.println("ЗАКРОЙТЕ ЕКСЕЛЬ ФАЙЛ С БАЗОЙ");
		} catch (IOException e) {
		    e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
