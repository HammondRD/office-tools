package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Properties;

public class Settings {
	Properties settings = new Properties();
	OutputStream writeSettingsFile = null;
	InputStream readSettingsFile = null;

	public Settings() {

		if (new File(Strings.DEFAULT_SETTINGS_FILE).exists()) {
			System.out.println("Найден файл с настройками");
			try {
				readSettingsFile = new FileInputStream(
						Strings.DEFAULT_SETTINGS_FILE);
				try {
					settings.load(readSettingsFile);
				} catch (IOException e) {

					e.printStackTrace();
				}
				System.out.println(settings.getProperty("lastSavedDir"));
			} catch (FileNotFoundException e) {

				e.printStackTrace();
			} 

		} else {
			try {
				System.out.println("Создаем файл с настройками");
				writeSettingsFile = new FileOutputStream(
						Strings.DEFAULT_SETTINGS_FILE);
				System.out.println(System.getProperty("user.dir"));
				settings.setProperty(Strings.DEFAULT_DIR,
						System.getProperty("user.dir").toString());
				try {
					settings.store(writeSettingsFile, null);
				} catch (IOException e) {

					e.printStackTrace();
				}
			} catch (FileNotFoundException e) {

				e.printStackTrace();
			} 
		}
	}
	public void settingsPrepare(){
		if (new File(Strings.DEFAULT_SETTINGS_FILE).exists()) {
			System.out.println("Найден файл с настройками");
			try {
				readSettingsFile = new FileInputStream(
						Strings.DEFAULT_SETTINGS_FILE);
				try {
					settings.load(readSettingsFile);
				} catch (IOException e) {

					e.printStackTrace();
				}
				System.out.println(settings.getProperty("lastSavedDir"));
			} catch (FileNotFoundException e) {

				e.printStackTrace();
			} 

		} else {
			try {
				System.out.println("Создаем файл с настройками");
				writeSettingsFile = new FileOutputStream(
						Strings.DEFAULT_SETTINGS_FILE);
				System.out.println(System.getProperty("user.dir"));
				settings.setProperty(Strings.DEFAULT_DIR,
						System.getProperty("user.dir").toString());
				try {
					settings.store(writeSettingsFile, null);
				} catch (IOException e) {

					e.printStackTrace();
				}
			} catch (FileNotFoundException e) {

				e.printStackTrace();
			} 
		}
	}
	public String getProperty(String key) {
		try {
			settings.load(readSettingsFile);
		} catch (IOException e) {
			
			e.printStackTrace();
		}
		System.out.println(settings.getProperty(key));
		String property = settings.getProperty(key);

		return property;

	}
	public void setNewProperty(String key, String value){
		
		if(writeSettingsFile != null){
			settings.setProperty(key,
				value);
			try {settings.store(writeSettingsFile, null);} catch (IOException e) {

			e.printStackTrace();
		}
		}else{
		
			try {
				writeSettingsFile = new FileOutputStream(
						Strings.DEFAULT_SETTINGS_FILE);
				settings.setProperty(key,
						value);
					try {settings.store(writeSettingsFile, null);} catch (IOException e) {

					e.printStackTrace();
				}
			} catch (FileNotFoundException e) {
		
				e.printStackTrace();
			}
				
			
				
			}
			
		
		
	}
	public void closeProperty() {

		if (writeSettingsFile != null) {
			try {
				writeSettingsFile.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		if (readSettingsFile != null) {
			try {
				readSettingsFile.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

}
