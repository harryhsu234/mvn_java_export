package OutputMethod;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Iterator;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileFilter;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import net.lingala.zip4j.core.ZipFile;
import net.lingala.zip4j.exception.ZipException;

public class Excel_AE_WPI_GROUP extends OutputCommon {

	private class Item {
		String FILENAME, SEQ, BRAND, DESC1, DESC2, DESC3, DESC4, DESC5, UOM, COUNTRY;
		int QTY;
		double PRICE;

		public Item() {
		}

		public Item(String FILENAME, String SEQ, String BRAND, String DESC1, String DESC2, String DESC3, String DESC4,
				String DESC5, int QTY, String UOM, double PRICE, String COUNTRY) {
			this.FILENAME = FILENAME;
			this.SEQ = SEQ;
			this.BRAND = BRAND;
			this.DESC1 = DESC1;
			this.DESC2 = DESC2;
			this.DESC3 = DESC3;
			this.DESC4 = DESC4;
			this.DESC5 = DESC5;
			this.QTY = QTY;
			this.UOM = UOM.toUpperCase();
			this.PRICE = PRICE;
			this.COUNTRY = COUNTRY;
		}

	}

	public final static String programmeTitle = "AE_世平興業";

	XSSFWorkbook wb;
	XSSFSheet ws;

	private ArrayList<File> filesToExtract = new ArrayList<File>();
	private ArrayList<Item> alItem = new ArrayList<Item>();

	public Excel_AE_WPI_GROUP() {
		// TODO Auto-generated constructor stub
	}

	private void doExcel(String selectedFileName) throws Exception {
		String templatePath = "/Excel_AE_WPI_GROUP.xlsx";
		InputStream tmpFile = this.getClass().getResourceAsStream(templatePath);

		wb = new XSSFWorkbook(tmpFile);
		ws = wb.getSheetAt(0);

		int row_pos = 1;
		for (Item obj : this.alItem) {
			int col_pos = 0;

			this.setValue(row_pos, col_pos++, obj.FILENAME);
			this.setValue(row_pos, col_pos++, obj.SEQ);
			this.setValue(row_pos, col_pos++, obj.BRAND);
			this.setValue(row_pos, col_pos++, obj.DESC1);
			this.setValue(row_pos, col_pos++, obj.DESC2);
			this.setValue(row_pos, col_pos++, obj.DESC3);
			this.setValue(row_pos, col_pos++, obj.DESC4);
			this.setValue(row_pos, col_pos++, obj.DESC5);
			this.setValue(row_pos, col_pos++, obj.QTY);
			this.setValue(row_pos, col_pos++, obj.UOM);
			this.setValue(row_pos, col_pos++, obj.PRICE);
			this.setValue(row_pos, col_pos++, obj.COUNTRY);

			row_pos++;
		}

		outputFilePath = "D:\\XML_OUTPUT\\";
		outputFileName = "Excel_WPI_" + selectedFileName + ".xlsx";
		Files.createDirectories(new File(outputFilePath).toPath());
		FileOutputStream stream = new FileOutputStream(outputFilePath + outputFileName);
		wb.write(stream);
		stream.close();

		wb.close();

		System.out.println("JOB_DONE");
		infoBox(outputFileName + " 產生完畢", "JOB_DONE");
	}

	private void extractExcelFile(String sAbsolutePath) throws IOException {

		XSSFWorkbook xwb = new XSSFWorkbook(sAbsolutePath);

		extractInvSheet(sAbsolutePath, xwb.getSheetAt(0));

		xwb.close();
	}

	private void extractExcelFile(File f) throws IOException, InvalidFormatException {

		XSSFWorkbook xwb = new XSSFWorkbook(f);

		extractInvSheet(f.getName(), xwb.getSheetAt(0));

		xwb.close();
	}

	private void extractInvSheet(String sAbsolutePath, XSSFSheet xws) {
		// System.out.println("Start Extract excel:" + sAbsolutePath);
		// alItem.clear();

		Iterator<Row> rowIterator = xws.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();

			// 從第二行開始
			if (row.getRowNum() < 1) {
				continue;
			}

			String flag = getCellString(row, 0);
			if (flag.equals(""))
				break;

			String FILENAME = sAbsolutePath;
			// 第一行
			String SEQ = getCellString(row, 0);
			String BRAND = getCellString(row, 1);
			String DESC3 = getCellString(row, 2);
			String DESC4 = getCellString(row, 3);
			int QTY = (int) getCellDouble(row, 4);
			String UOM = getCellString(row, 5);

			// 第二行
			row = rowIterator.next();
			String DESC1 = getCellString(row, 1);
			String DESC5 = getCellString(row, 3);
			double PRICE = getCellDouble(row, 4);

			// 第三行
			row = rowIterator.next();
			String DESC2 = getCellString(row, 1);
			String COUNTRY = getCellString(row, 2);

			alItem.add(new Item(FILENAME, SEQ, BRAND, DESC1, DESC2, DESC3, DESC4, DESC5, QTY, UOM, PRICE, COUNTRY));

		}
	}

	private class WPIExcelFilter extends FileFilter {
		@Override
		public boolean accept(File pathname) {
			String filename = pathname.getName();
			if (pathname.isDirectory()) {
				return true;

			} else if (filename.endsWith("zip")) {
				return true;
			} else {
				return false;
			}
		}

		@Override
		public String getDescription() {
			return "ZIP Files";
		}
	}

	public void run() throws IOException, ZipException {
		this.filesToExtract.clear();
		this.alItem.clear();

		// get 原始WORD 檔
		JFileChooser fc = new JFileChooser();
		fc.setDialogTitle(programmeTitle);
		fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		fc.setMultiSelectionEnabled(false);

		FileFilter wipFilter = new WPIExcelFilter();
		fc.addChoosableFileFilter(wipFilter);
		fc.setFileFilter(wipFilter);

		int returnVal = fc.showOpenDialog(null);
		File file;
		if (returnVal == JFileChooser.OPEN_DIALOG) {
			file = fc.getSelectedFile();

			String ext = FilenameUtils.getExtension(file.getAbsolutePath()).toLowerCase();

			if (ext.equals("zip")) {
				String destination = outputFilePath + "unzipped-" + file.getName() + "\\";
				String password = "password";
				
					ZipFile zipFile = new ZipFile(file);
					if (zipFile.isEncrypted()) {
						zipFile.setPassword(password);
					}
					zipFile.extractAll(destination);
					
					File folder = new File(destination);
					if(folder.isDirectory()) {
						for(File f : folder.listFiles())
							this.filesToExtract.add(f);
					}
					
				
			} else {
				this.filesToExtract.add(file);
			}

		} else {
			super.infoBox("未選擇檔案", "錯誤訊息");
			return;
		}
		try {

			// 驗證抓到的DOC 檔
			if (this.filesToExtract.size() == 0)
				throw new Exception("未發現可處理的檔案");

			// 處理 word file
			for (File f : this.filesToExtract) {
				extractExcelFile(f);
			}

			check_and_merge_ArrayList();

			// 寫入EXCEL FILE
			doExcel(file.getName());

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			super.infoBox(e.getMessage(), "Error in " + this.getClass());
		}
	}

	private void check_and_merge_ArrayList() throws Exception {

	}

	private void setValue(int row_pos, int col_pos, Object value) throws Exception {
		super.setValue(ws, row_pos, col_pos, value);
	}

}
