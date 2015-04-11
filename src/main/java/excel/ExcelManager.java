package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.JFileChooser;
import javax.swing.JPanel;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelManager {

	public ExcelManager(String ext, FileInputStream inp) {
		localExt = ext;
		localInput = inp;
		// localDataInExcel = inDataMatrix;

		if (inp != null) {
			System.out.println("inp   exists");
			// считать лист рабочий книги

			switch (localExt) {
			case "xls":
				try {
					System.out.println("start read xls");
					workbook = new HSSFWorkbook(localInput);
					
					sheet = (HSSFSheet) workbook.getSheetAt(0);

					// readExcelDataToArray();
					System.out.println("xls read ok");
				} catch (IOException e1) {

					e1.printStackTrace();
				}
				break;
			case "xlsx":
				try {
					System.out.println("start read xlsx");
					workbook = new XSSFWorkbook(localInput);

					xSheet = (XSSFSheet) workbook.getSheetAt(0);

					System.out.println("xlsx read ok");
				} catch (IOException e1) {
					System.out.println(" read not ok");
					e1.printStackTrace();
				}

				break;
			default:
				System.out.println("Not Excel formats");
				break;
			}

		}

	}

	private String localExt;
	private Workbook workbook;
	private HSSFSheet sheet;
	private XSSFSheet xSheet;
	private Row row;
	private Cell cell;
	private String[][] localDataInExcel;
	private FileInputStream localInput;

	/**
	 * Find last row num
	 * 
	 * @return lastRowNum последняя строчка
	 */
	public Integer getLastRowNum() {

		switch (localExt) {
		case "xls":

			Globals.lastRowNum = sheet.getLastRowNum();
			break;
		case "xlsx":
			Globals.lastRowNum = xSheet.getLastRowNum();
			break;
		default:
			break;
		}
		return Globals.lastRowNum;

	}

	/**
	 * Find last column name
	 * 
	 * @return Globals.lastColNum последняя ячейка
	 */
	public Integer getLastColumnNum() {
		switch (localExt) {
		case "xls":

			Globals.lastColNum = sheet.getRow(sheet.getFirstRowNum())
					.getLastCellNum();

			break;
		case "xlsx":
			Globals.lastColNum = xSheet.getRow(xSheet.getFirstRowNum())
					.getLastCellNum();

			break;
		default:
			break;
		}

		return Globals.lastColNum;

	}

	public Integer findValueInExcel(String value) {
		switch (localExt) {
		case "xls":
			for (int rows = 0; rows < getLastRowNum(); rows++) {
				for (int cell = 0; cell < getLastColumnNum(); cell++) {
					if (value.equals(getValueAtCell(rows, cell))) {
						Globals.cellRC = rows;
						
					}
				}
			}
			break;
		case "xlsx":

			break;
		default:
			break;
		}

		return Globals.cellRC;

	}

	public String getValueAtCell(int rowNum, int colNum) {
		switch (localExt) {
		case "xls":
			try {
				if (sheet.getRow(rowNum).getCell(colNum) != null) {
					Globals.cellValue = sheet.getRow(rowNum).getCell(colNum)
							.toString();
					// System.out.println(Globals.cellValue+ " = rowNum  = "+
					// rowNum+"  colNum ="+colNum);
				} else {
					Globals.cellValue = "";
				}
			} catch (Exception e) {
				Globals.cellValue = "";
				// e.printStackTrace();
			}

			break;
		case "xlsx":
			try {
				if (xSheet.getRow(rowNum).getCell(colNum) != null) {
					Globals.cellValue = xSheet.getRow(rowNum).getCell(colNum)
							.toString();
				} else {
					Globals.cellValue = "";
				}
			} catch (Exception e) {
				Globals.cellValue = "";
				// e.printStackTrace();
			}
			break;
		default:
			break;
		}
		return Globals.cellValue;

	}
	public void setValueAtCell(String value, int rowNum, int colNum) {
		
		 cell = sheet.getRow(rowNum).getCell(colNum);
		    cell.setCellValue(value);
		    System.out.println(rowNum + " row, "+ colNum + " cell, "+value+" = value" );
		switch (localExt) {
		case "xls":
			try {
		        FileOutputStream out = 
		                new FileOutputStream(Globals.file);
		        workbook.write(out);
		        out.close();
		        System.out.println("Excel written successfully..");
		         
		    } catch (FileNotFoundException e) {
		        e.printStackTrace();
		    } catch (IOException e) {
		        e.printStackTrace();
		    }
			break;
		case "xlsx":
			
			break;
		default:
			break;
		}
	

	}
	public static void main(String[] args) {

	}

}
