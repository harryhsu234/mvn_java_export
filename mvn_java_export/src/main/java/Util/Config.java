package Util;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.Reader;
import java.io.Writer;
import java.util.Hashtable;

import org.dtools.ini.BasicIniFile;
import org.dtools.ini.IniFile;
import org.dtools.ini.IniFileReader;
import org.dtools.ini.IniFileWriter;
import org.dtools.ini.IniItem;
import org.dtools.ini.IniSection;

public class Config {

	String defaultPath = "config.ini";
	String resourcePath = "/config.ini";
	
	/**
	 * Get cofing setion items, if config.ini not exists will create from resources/config.ini
	 * @param String section name
	 * @return Hashtable <item.name,item.value>
	 */
	public Hashtable<String,String> getConfig(String section) throws Exception {
		Hashtable<String,String> result = new Hashtable<String,String>();
		System.out.println(section);
		IniFile ini = getIniFile();// false 不管大小寫
	
		IniSection sec = ini.getSection(section);
		for (IniItem item : sec.getItems()) {
			System.out.println(item.getName() + " = " + item.getValue());
			result.put(item.getName(),item.getValue());
		}

		return result;
	}

	private IniFile getIniFile() throws Exception {
		IniFile ini = new BasicIniFile(false);// false 不管大小寫
		File file = new File(defaultPath);
		if (!file.exists()) {
			file = createConfig();
		}
		IniFileReader reader = new IniFileReader(ini, file);
		reader.read();
		
		return ini;
	}
	
	private File createConfig() throws Exception {
		InputStream in = this.getClass().getResourceAsStream(resourcePath);
		
		System.out.println("Create Config.ini....");
		
		return saveToFile(in);
	}
	
	private File saveToFile(InputStream in) throws Exception {
		File result = new File(defaultPath);
		
		InputStreamReader isr = new InputStreamReader(in, "UTF-8");
		Reader reader = new InputStreamReader(in,"UTF-8");
		BufferedReader fin = new BufferedReader(reader);
		Writer writer = new OutputStreamWriter( new FileOutputStream(result), "UTF-8");
		BufferedWriter fout = new BufferedWriter(writer);
		
		String s;
		while ((s=fin.readLine())!=null) {
		     fout.write(s);
		     fout.newLine();
		}

		fin.close();
	   	fout.close();
	   	
		return result;
	}
	
	public void writeSection(String section,Hashtable ht) {
		
		try {
			IniFile ini = getIniFile(); // false 不管大小寫

			IniSection sec = ini.getSection(section);
			
			for (IniItem item : sec.getItems()) {
				item.setValue(ht.get(item.getName()));
			}
			
			// 寫入ini
			IniFileWriter writer=new IniFileWriter(ini, new File(defaultPath));
			writer.write();
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
}
