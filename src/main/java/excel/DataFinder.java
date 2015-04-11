package excel;

import java.awt.*;
import java.awt.event.*;
import java.io.*;

import org.apache.commons.io.*;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.event.TableModelEvent;
import javax.swing.event.TableModelListener;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.TableModel;

public class DataFinder {

	public static void main(String[] args) {
		DataFinderFrame frame = new DataFinderFrame();
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setVisible(true);
	}

	public static File file;

	private void loadSettings() {

	}
}

class DataFinderFrame extends JFrame {
	public DataFinderFrame() {
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch (ClassNotFoundException e) {

			e.printStackTrace();
		} catch (InstantiationException e) {

			e.printStackTrace();
		} catch (IllegalAccessException e) {

			e.printStackTrace();
		} catch (UnsupportedLookAndFeelException e) {

			e.printStackTrace();
		}
		setTitle("Data Finder  - CSV creater");
		setSize(800, 600);
		DataFinderPanel panel = new DataFinderPanel();
		add(panel);
	}
}

class DataFinderPanel extends JPanel {
	DataFinderPanel() {

		setLayout(new BorderLayout());
		JButton readDbButton = new JButton("Connect DB...");
		JButton importDataButton = new JButton("Load Data to check...");
		JButton analyseButton = new JButton("Start Analysing");
		JButton createCSVButton = new JButton("Create Text File...");
		infoField = new JTextArea();
		infoField.setEditable(false);
		infoField.setText("Info board:");
		infoField.setFont(new Font("Serif", Font.BOLD, 16));

		ActionListener readDB = new ReadDatabaseAction();
		ActionListener importData = new ImportDataAction();
		ActionListener analyseData = new AnalyseDataAction();
		ActionListener createCsv = new CreateCsvAction();

		readDbButton.addActionListener(readDB);
		importDataButton.addActionListener(importData);
		analyseButton.addActionListener(analyseData);
		createCSVButton.addActionListener(createCsv);

		JPanel importDataPanel = new JPanel();
		importDataPanel.setLayout(new GridLayout(3, 1));

		add(importDataPanel, BorderLayout.WEST);

		importDataPanel.add(readDbButton);
		importDataPanel.add(importDataButton);
		importDataPanel.add(analyseButton);

		add(createCSVButton, BorderLayout.SOUTH);

		// String[] columnNames = {"First Name",
		// "Last Name",
		// "Sport",
		// "# of Years",
		// "Vegetarian"};
	
		
//		table.setPreferredScrollableViewportSize(new Dimension(600, 70));
//		table.setFillsViewportHeight(true);
		
		
		model.addTableModelListener(new TableModelListener(){

			public void tableChanged(TableModelEvent arg0) {
			
				
			}
			
		});

		
		// add(infoField, BorderLayout.CENTER);
		
	}


	class ReadDatabaseAction implements ActionListener {
		JFileChooser fc;
		
		public void actionPerformed(ActionEvent event) {
			
			openFileDialog();

			DataFinderPanel.dataMatrix = readExcelDataToArray(inp,
					DataFinderPanel.dataMatrix);
			// System.out.println(DataFinderPanel.dataMatrix[0][1]);

		}

	}

	/**
	 * Загружает файл Ексель
	 * 
	 * @return inp
	 */
	private FileInputStream openFileDialog() {

		final JFileChooser fileChooser = new JFileChooser();

		int returnVal = fileChooser.showOpenDialog(DataFinderPanel.this);
		if (returnVal == JFileChooser.APPROVE_OPTION) {
			File file = fileChooser.getSelectedFile();

			ext = FilenameUtils.getExtension(file.getPath());
			System.out.println("Расширение файла: " + ext);
			try {
				inp = new FileInputStream(file);

			} catch (FileNotFoundException ex) {
				ex.printStackTrace();

			} catch (IOException ex) {
				
				ex.printStackTrace();

			}

		}
		return inp;
	}

	/**
	 * Считывает данный из потока и загружает их в массив dataMatrix
	 * 
	 * @param inp
	 *            поток из Экселя
	 * @param inDataMatrix
	 *            Массив куда заносятся данные из файла
	 */
	public String[][] readExcelDataToArray(FileInputStream inp,
			String[][] inDataMatrix) {
		int length; // количество элементов в документе
		try {
			if (ext.equals("xls")) {
				workbook = new HSSFWorkbook(inp);
				sheet = (HSSFSheet) workbook.getSheetAt(0);
				// HSSFRow row;
				// HSSFCell cell;
				// Iterator rows = sheet.rowIterator();
				length = sheet.getLastRowNum();

			} else {
				workbook = new XSSFWorkbook(inp);
				xSheet = (XSSFSheet) workbook.getSheetAt(0);
				length = xSheet.getLastRowNum();
				// XSSFRow row;
				// XSSFCell cell;
			}
			// workbook = new HSSFWorkbook(inp);
			// sheet = (HSSFSheet) workbook.getSheetAt(0);

			inDataMatrix = new String[length + 1][4];

			long start = System.nanoTime();
			/* sheet.getLastRowNum() */
			for (int c = 0; c <= length; c++) {
				// System.out.println(" ");
				try {
					if (ext.equals("xls")) {
						row = sheet.getRow(c);
					} else {
						row = xSheet.getRow(c);
					}
					// System.out.print(c+": ");
					for (int i = 0; i < 4; i++) {
						// .getStringCellValue()
						try {

							cell = row.getCell(i);
							inDataMatrix[i][c] = cell.toString().trim();

						} catch (Exception e) {

							inDataMatrix[i][c] = " ";

						}

						// System.out.print(i+") "+dataMatrix[i][c]);
						// System.out.print(" | ");

					}

				} catch (Exception e) {

					// System.out.println("collumn! i: " + c);
					// e.printStackTrace();
				}

			}
			/* System.out.println(" "); */
			// Р·Р°РЅРµСЃРµРЅРёРµ РґР°РЅРЅС‹С… РІ РјР°С‚СЂРёС†Сѓ
			long stop = System.nanoTime();
			System.out.println("Time: " + (stop - start) / 1000000);

			// find barcode

		} catch (IOException e) {

			e.printStackTrace();
			System.out.println("err1");
		}
		return inDataMatrix;

	}

	class ImportDataAction implements ActionListener {

		public void actionPerformed(ActionEvent e) {
			
			openFileDialog();

			DataFinderPanel.importMatrix = readExcelDataToArray(inp,
					DataFinderPanel.importMatrix);

			// Вывести заголовки из файла
			double matrixLength = importMatrix.length;
			

		}

		String log;
	}

	class AnalyseDataAction implements ActionListener {

		public void actionPerformed(ActionEvent e) {

		}

	}

	class CreateCsvAction implements ActionListener {

		public void actionPerformed(ActionEvent e) {

		}

	}

	public String[] columnNames = { "First Name", "Last Name", "Sport",
			"# of Years", "Vegetarian" };
	public Object[][] data = {
			{ "Kathy", "Smith", "Snowboarding", new Integer(5),
					new Boolean(false) },
			{ "John", "Doe", "Rowing", new Integer(3), new Boolean(true) },
			{ "Sue", "Black", "Knitting", new Integer(2), new Boolean(false) },
			{ "Jane", "White", "Speed reading", new Integer(20),
					new Boolean(true) },
			{ "Joe", "Brown", "Pool", new Integer(10), new Boolean(false) } };

	
	protected static  String ext;
	protected static FileInputStream inp;
	protected static File file;
	protected static Workbook workbook;
	protected static String[][] importMatrix;
	protected static String[][] dataMatrix;
	protected static Cell cell;
	protected static Row row;
	protected static HSSFSheet sheet;
	protected static XSSFSheet xSheet;
	protected static JTextArea infoField;
	protected static TableModel model;
}

/**
 * Открывает каталог в Экселе и загружает данные в массив для дальнейшей работы
 * 
 *
 */

