package excel;

import java.io.File;

import javax.swing.JFileChooser;

public class Globals {
	protected static Object[][] tableData;
	protected static Integer columnWidths = 100;
	protected static ExcelManager excel;
	protected static int lastRowNum;
	protected static int lastColNum;
	protected static int cellRC;
	protected static String cellValue;
	protected static  String ext;//расширение файла
	 protected static JFileChooser fileChooser;
	 protected static File file;
}
