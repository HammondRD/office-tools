package excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JPanel;

public class DataAdjust {

	public DataAdjust() {
		System.out.println("Обработка данных!");
	}

	public void setPictureNames(int rowNum, int cellNum) {

		String price = Globals.excel.getValueAtCell(rowNum, cellNum - 1).trim();

		String pic1 = Globals.excel.getValueAtCell(rowNum, cellNum - 5).trim();
		String pic2 = Globals.excel.getValueAtCell(rowNum, cellNum - 4).trim();
		String pic3 = Globals.excel.getValueAtCell(rowNum, cellNum - 3).trim();
		String artNum = Globals.excel.getValueAtCell(rowNum, cellNum - 12)
				.trim();
		// inpLocal = new FileInputStream(Globals.file);

	
			if (price.length() != 0) {
				if (pic1.length() == 0) {
					pic1 = artNum + ".JPG";
					Globals.localModel
							.setValueAt(pic1, rowNum - 1, cellNum - 4);
					Globals.excel.setValueAtCell(pic1, rowNum - 1, cellNum - 4);
				}
			}
		

	}

	public static void main(String[] args) {

	}

}
