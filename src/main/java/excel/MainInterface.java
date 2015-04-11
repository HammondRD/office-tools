package excel;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JFileChooser;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.border.EmptyBorder;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.table.DefaultTableModel;
import javax.swing.JScrollPane;

import static java.lang.System.*;

import org.apache.commons.io.FilenameUtils;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;

import javax.swing.JButton;

import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class MainInterface {

	/**
	 * Create the GUI and show it. For thread safety, this method should be
	 * invoked from the event-dispatching thread.
	 */
	public MainInterface() {

	}

	public static void createAndShowGUI() {

		MainInterfaceFrame frame = new MainInterfaceFrame();

	}

	/**
	 * Launch the application.
	 */

	public static void main(String[] args) {
		SwingUtilities.invokeLater(new Runnable() {
			public void run() {
				try {
					
					createAndShowGUI();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

}

class importExcelDataBtnAction implements ActionListener {
	public importExcelDataBtnAction(JTable table) {

		localModel = (DefaultTableModel) table.getModel();
	}

	private JFrame frameLocal;
	public FileInputStream inpLocal;
	private DefaultTableModel localModel;

	public void actionPerformed(ActionEvent e) {

		final JFileChooser fileChooser = new JFileChooser();

		int returnVal = fileChooser.showOpenDialog(frameLocal);
		if (returnVal == JFileChooser.APPROVE_OPTION) {
			Globals.file = fileChooser.getSelectedFile();
out.println(Globals.file.getName()+" has "+(Globals.file.length()/1024/1024)+" MBytes and it was modified at "+ 
			new SimpleDateFormat("dd-MM-yyyy HH-mm-ss").format(new Date(Globals.file.lastModified())));
			Globals.ext = FilenameUtils.getExtension(Globals.file.getPath());

			System.out.println("Расширение файла: " + Globals.ext);
			try {
				inpLocal = new FileInputStream(Globals.file);
				
					Globals.excel = new ExcelManager(Globals.ext, inpLocal);
					System.out.println("set columns -////  "
							+ Globals.excel.getLastColumnNum());
					// добавить нужное количество
					// столбцов-localModel.getColumnCount()

					for (int y = 0; y < (Globals.excel.getLastColumnNum()); y++) {
						localModel
								.addColumn(Globals.excel.getValueAtCell(0, y));

					}
					System.out.println(Globals.excel.getLastRowNum());
					// установить нужное количество строчек
					//
					for (int i = 0; i < Globals.excel.getLastRowNum(); i++) {

						localModel.addRow(new Object[] {});
						localModel.setValueAt(i + 1, i, 0);
						// localModel.setValueAt(excel.getValueAtCell(i+1,
						// 2),i,1);
					}
					// записать данные в таблицу
					for (int rowNum = 0; rowNum < Globals.excel.getLastRowNum(); rowNum++) {
						for (int cellNum = 0; cellNum < Globals.excel
								.getLastColumnNum(); cellNum++) {
							// System.out.println("excel.getLastColumnNum() "+excel.getLastColumnNum());
							localModel.setValueAt(Globals.excel.getValueAtCell(
									rowNum + 1, cellNum), rowNum, cellNum + 1);
							
						}
					}
				
			} catch (IOException ex) {

				ex.printStackTrace();
			

			}

		}

	}

}

class MainInterfaceFrame extends JFrame {
	public MainInterfaceFrame() {
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

		MainInterfacePanel panel = new MainInterfacePanel();
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setTitle(Strings.FRAME_TITLE);
		setVisible(true);
		add(panel);

		// setBounds(100, 100, 820, 523);
		pack();
	}

}

class MainInterfacePanel extends JPanel {
	public MainInterfacePanel() {
		setBorder(new EmptyBorder(5, 5, 5, 5));

		GridBagLayout gridBagLayout = new GridBagLayout();
		gridBagLayout.columnWidths = new int[] { 1, 0 };
		gridBagLayout.rowHeights = new int[] { 1, 0, 0, 0 };
		gridBagLayout.columnWeights = new double[] { 0.0, 1.0, Double.MIN_VALUE };
		gridBagLayout.rowWeights = new double[] { 0.0, 0.0, 0.0,
				Double.MIN_VALUE };

		setLayout(gridBagLayout);
		// ---------------------------------------------
		// setLayout(new BoxLayout(this, BoxLayout.X_AXIS));
		// ------------------

		JButton btnNewButton = new JButton(Strings.BTN_OPEN_FILE);
		GridBagConstraints gbc = new GridBagConstraints();
		// add(btnNewButton, gbc_btnNewButton);
		//ll
		//ooo
		gbc.insets = new Insets(0, 0, 0, 5);
		gbc.gridx = 0;
		gbc.gridy = 0;
		add(btnNewButton, gbc);
		table = new JTable();
		JScrollPane scrollPane = new JScrollPane(table);
		GridBagConstraints gbcscrollPane = new GridBagConstraints();
		gbcscrollPane.insets = new Insets(5, 5, 5, 5);
		gbcscrollPane.gridheight = GridBagConstraints.REMAINDER;
		gbcscrollPane.gridx = 1;
		gbcscrollPane.gridy = 0;
		gbcscrollPane.fill = GridBagConstraints.BOTH;
		gbcscrollPane.ipady = GridBagConstraints.VERTICAL;
		add(scrollPane, gbcscrollPane);

		final JTextField barCodeField = new JTextField(20);
		GridBagConstraints gbcbarCodeField = new GridBagConstraints();
		gbcbarCodeField.insets = new Insets(5, 5, 5, 5);
		gbcbarCodeField.gridx = 0;
		gbcbarCodeField.gridy = 1;
		gbcbarCodeField.gridheight = 1;
		add(barCodeField, gbcbarCodeField);
		
		barCodeField.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				if (Globals.excel != null) {
					System.out.println(Globals.excel.findValueInExcel(
							barCodeField.getText()).toString());

					System.out.println(Globals.excel.getValueAtCell(
							Globals.excel.findValueInExcel(barCodeField
									.getText()), 1));
				} else {
					System.out.println("нет данных, где искать");
				}
			}

		});
		JButton writeToExcelBtn = new JButton(Strings.BTN_WRITE_TO_EXCEL);
		
		// add(btnNewButton, gbc_btnNewButton);
	//gbcbarCodeField.insets = new Insets(0, 0, 0, 5);
		GridBagConstraints writeToExcelBtnGbc = new GridBagConstraints();
		writeToExcelBtnGbc.insets = new Insets(5, 5, 5, 5);
		writeToExcelBtnGbc.gridx = 0;
		writeToExcelBtnGbc.gridy = 2;
		writeToExcelBtnGbc.gridheight = 1;
		add(writeToExcelBtn, writeToExcelBtnGbc);
		writeToExcelBtn.addActionListener(new ActionListener(){

			
			public void actionPerformed(ActionEvent event) {
				if(!barCodeField.getText().isEmpty()){
					System.out.println("запись данных = "+barCodeField.getText());
					Globals.excel.setValueAtCell(barCodeField.getText(), 1, 1);
				}else{
					System.out.println("нет данных для записи");
				}
			}
			
		});
		
		barCodeField.getDocument().addDocumentListener(new DocumentListener() {
			
			public void changedUpdate(DocumentEvent event) {
				System.out.println("findValueInExcel(changedUpdate)");

			}

			public void insertUpdate(DocumentEvent arg0) {
				System.out.println("findValueInExcel(insertUpdate)");

			}

			public void removeUpdate(DocumentEvent arg0) {

			}

		});

		Globals.tableData = new String[][] {};

		btnNewButton.addActionListener(new importExcelDataBtnAction(table) {
		});

		model = (DefaultTableModel) table.getModel();
		model.addColumn("No");

		// ---------------------------------------

	}

	
	protected static Integer length;
	protected static JTable table;

	protected static DefaultTableModel model;

}
