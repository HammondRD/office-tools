package excel;

import java.io.*;
import java.nio.file.FileSystems;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Properties;

public class Config {
    
	public static void main(String[] args) {
		FileInputStream fileProp;
		FileOutputStream fileProOut;
		Properties property = new Properties();
		Iterable<Path> dirs = FileSystems.getDefault().getRootDirectories();
		Path path = Paths.get("..");
		System.out.println(path.toAbsolutePath().toString());
		for (Path name: dirs){
			System.out.println(name);
		}
		try {
			
			boolean check = new File("src/main/resources/config.pro")
					.exists();
			if (check) {
				fileProp = new FileInputStream("src/main/resources/config.pro");
				System.out.println(new File("src/main/resources/config.pro").canRead());
				property.load(fileProp);
				String host = property.getProperty("PROGRAMM_NAME");
				System.out.println(host);
			} else {
				System.out.println("no file");
				fileProOut = new FileOutputStream(
						"src/main/resources/config.pro");
				property.setProperty("PROGRAMM_NAME", "MY_NAME55");

				property.store(new BufferedWriter(new OutputStreamWriter(
						fileProOut, "UTF-8")), null);
			}
		} catch (FileNotFoundException e) {

			e.printStackTrace();
		} catch (IOException e) {

			e.printStackTrace();
		}
	}

}
