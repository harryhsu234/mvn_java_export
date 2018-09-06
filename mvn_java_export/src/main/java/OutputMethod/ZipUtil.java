package OutputMethod;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class ZipUtil {
	
	/**
	 * 壓縮檔案
	 * 
	 * @param file File 物件
	 * @param zipfileName 壓縮的檔名
	 * @throws Exception 
	 */
	public void zip(File file,String zipfileName) throws Exception {
		ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(new File(zipfileName)));
		checkFileType(file, zos, file.getName());
		zos.finish();
		zos.close();
	}
	
	/**
	 * 遞迴檢查 File 的屬性
	 * 
	 * @param file
	 * @param zos
	 * @param fileName
	 * @throws Exception
	 */
	public void checkFileType(File file, ZipOutputStream zos, String fileName) throws Exception {
		if (file.isDirectory()) {
			for (File tmp : file.listFiles()) {
				checkFileType(tmp, zos, fileName + "/" + tmp.getName());
			}
		} else {
			addZipFile(file, zos, fileName);
		}
	}

	/**
	 * 新增 File 至 Zip 檔
	 * 
	 * @param file
	 * @param zos
	 * @param fileName
	 * @throws Exception
	 */
	private static void addZipFile(File file, ZipOutputStream zos, String fileName) throws Exception {
		int l;

		byte[] b = new byte[(int) file.length()];

		FileInputStream fis = new FileInputStream(file);

		ZipEntry entry = new ZipEntry(fileName);

		zos.putNextEntry(entry);

		while ((l = fis.read(b)) != -1) {
			zos.write(b, 0, l);
		}

		entry = null;
		fis.close();
		b = null;
	}

}
