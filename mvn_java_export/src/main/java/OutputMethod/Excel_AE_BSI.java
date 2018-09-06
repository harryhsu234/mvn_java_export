package OutputMethod;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.Set;
import java.util.TreeSet;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileFilter;

import org.apache.commons.io.comparator.NameFileComparator;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_AE_BSI extends OutputCommon {

	private class Item {
		double INV_NO;
		Object AMT;
		Object PRICE;
		Object QTY;
		Object ITEM;
		String SEQ, PRODUCT_CODE, DESC_OF_GOODS, UNIT;

		Hashtable<String, String> htItemsPink = new Hashtable<String, String>();

		public Item(double INV_NO, String DESC_OF_GOODS) {
			this.INV_NO = INV_NO;
			this.SEQ = null;
			this.ITEM = null;
			this.PRODUCT_CODE = null;
			this.DESC_OF_GOODS = DESC_OF_GOODS;
			this.QTY = null;
			this.UNIT = null;
			this.PRICE = null;
			this.AMT = null;
		}

		public Item(double INV_NO, String BIG_DESC, String SEQ, int ITEM, String PRODUCT_CODE, String DESC_OF_GOODS,
				double QTY, String UNIT, double PRICE, double AMT, Hashtable<String, String> htItemsPink) {
			this.INV_NO = INV_NO;
			this.SEQ = SEQ;
			this.ITEM = ITEM;
			this.PRODUCT_CODE = PRODUCT_CODE;
			this.DESC_OF_GOODS = DESC_OF_GOODS;
			this.QTY = QTY;
			this.UNIT = UNIT;
			this.PRICE = PRICE;
			this.AMT = AMT;
			this.htItemsPink = htItemsPink;
		}

	}

	private class Declr {
		String DO_NO, QTY, PRICE, AMOUNT;

		public Declr(String DO_NO, String QTY, String PRICE, String AMOUNT) {
			this.DO_NO = DO_NO;
			this.QTY = QTY;
			this.PRICE = PRICE;
			this.AMOUNT = AMOUNT;
		}
	}

	public final static String programmeTitle = "AE_晟宇資訊有限公司";

	XSSFWorkbook wb;
	XSSFSheet ws;

	private ArrayList<File> filesToExtract = new ArrayList<File>();
	private ArrayList<Item> alItem = new ArrayList<Item>();
	private ArrayList<Declr> alDeclr = new ArrayList<Declr>();
	private Hashtable<String, String> htPinkLable = new Hashtable<String, String>();

	public Excel_AE_BSI() {
		// TODO Auto-generated constructor stub
	}

	private void doExcel() throws Exception {
		String templatePath = "/Excel_AE_BSI.xlsx";
		InputStream tmpFile = this.getClass().getResourceAsStream(templatePath);

		wb = new XSSFWorkbook(tmpFile);
		ws = wb.getSheetAt(0);

		int row_pos = 1;
		for (Item obj : this.alItem) {
			int col_pos = 0;

			this.setValue(row_pos, col_pos++, obj.INV_NO);
			this.setValue(row_pos, col_pos++, obj.SEQ);
			this.setValue(row_pos, col_pos++, obj.ITEM);
			this.setValue(row_pos, col_pos++, obj.PRODUCT_CODE);
			this.setValue(row_pos, col_pos++, obj.DESC_OF_GOODS);
			this.setValue(row_pos, col_pos++, obj.QTY);
			this.setValue(row_pos, col_pos++, obj.UNIT);
			this.setValue(row_pos, col_pos++, obj.PRICE);
			this.setValue(row_pos, col_pos++, obj.AMT);

			// 排序粉紅色 原進口報單 的欄位出現順序
			Set<String> keySet = this.htPinkLable.keySet();
			Set<String> sortedSet = new TreeSet<>(); // 排序後的KEY SET
			for (String key : keySet) {
				sortedSet.add(key);
			}

			col_pos = 14;
			for (String key : sortedSet) {
				// 設定標題
				this.setValue(0, col_pos, key);

				String pinkValue = obj.htItemsPink.get(key);
				String cellValue = "" + pinkValue;
				if (cellValue.equals("null"))
					cellValue = "";
				this.setValue(row_pos, col_pos++, cellValue);
			}

			row_pos++;
		}

		ws = wb.getSheetAt(1);
		row_pos = 1;
		for (Declr obj : alDeclr) {
			int col_pos = 0;

			this.setValue(row_pos, col_pos++, obj.DO_NO);
			this.setValue(row_pos, col_pos++, obj.QTY);
			this.setValue(row_pos, col_pos++, obj.PRICE);
			this.setValue(row_pos, col_pos++, obj.AMOUNT);

			row_pos++;
		}

		outputFilePath = "D:\\XML_OUTPUT\\";
		outputFileName = "Excel_晟宇_" + System.currentTimeMillis() + ".xlsx";
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

		extractInvSheet(sAbsolutePath, xwb.getSheet("Shipping INVOICE"));

		xwb.close();
	}

	private void extractExcelFile(File f) throws IOException, InvalidFormatException {

		XSSFWorkbook xwb = new XSSFWorkbook(f);

		extractInvSheet(f.getName(), xwb.getSheet("Shipping INVOICE"));

		xwb.close();
	}

	private void extractInvSheet(String sAbsolutePath, XSSFSheet xws) {
		// 取得粉紅標籤
		getPinkLable(xws);

		// 取得各明細項目
		getItems(xws);

		// 取得最下面的DO NO
		getDoNo(xws);

	}

	private void getPinkLable(XSSFSheet xws) {
		Iterator<Cell> cellIterator = xws.getRow(16).cellIterator();

		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			if (cell.getStringCellValue().contains("原進口報單"))
				this.htPinkLable.put(cell.getStringCellValue(), "");
		}
	}

	private void getDoNo(XSSFSheet xws) {
		Iterator<Row> rowIterator = xws.iterator();
		boolean isDeclr = false;
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();

			if (getCellString(row, 2).trim().equals("Do No.")) {
				isDeclr = true;
			}
			if (isDeclr) {
				if (getCellDouble(row, 4) > 0) {
					String DO_NO = getCellString(row, 3);
					String QTY = (int) getCellDouble(row, 4) + "EAC";
					String PRICE = "* " + getCellDouble(row, 5) + "(代工費)=";
					String AMOUNT = "US$ " + ((double) Math.round(getCellDouble(row, 7) * 1000)) / 1000;

					alDeclr.add(new Declr(DO_NO, QTY, PRICE, AMOUNT));
				} else
					break;
			}
		}
	}

	private void getItems(XSSFSheet xws) {
		Hashtable<String, Integer> htPinkLableIndex = new Hashtable<String, Integer>();
		Iterator<Cell> cellIterator = xws.getRow(16).cellIterator();

		int index = 0;
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			if (cell.getStringCellValue().contains("原進口報單"))
				htPinkLableIndex.put(cell.getStringCellValue(), index);

			index++;
		}

		CellReference cr = new CellReference("I8");
		String d = "" + xws.getRow(cr.getRow()).getCell(cr.getCol()).getNumericCellValue();
		BigDecimal bd = new BigDecimal(d);
		long lonVal = bd.longValue();

		cr = new CellReference("I14");
		String bigDesc = xws.getRow(cr.getRow()).getCell(cr.getCol()).getStringCellValue();

		alItem.add(new Item((double) lonVal, bigDesc));

		Iterator<Row> rowIterator = xws.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();

			// 從第二行開始
			if (row.getRowNum() < 17) {
				continue;
			}

			String flag = getCellString(row, 0);
			if (flag.equals(""))
				break;

			double INV_NO = lonVal;
			String BIG_DESC = bigDesc;
			String SEQ = getCellString(row, 0);
			int ITEM = (int) getCellDouble(row, 1);
			String PRODUCT_CODE = getCellString(row, 2);
			String DESC_OF_GOODS = getCellString(row, 4);
			double QTY = getCellDouble(row, 5);
			String UNIT = "EAC";
			double PRICE = getCellDouble(row, 7);
			double AMT = QTY * PRICE;

			// 抓個明細的粉紅HT
			Hashtable<String, String> htItemsPink = new Hashtable<String, String>();
			for (String key : htPinkLableIndex.keySet()) {
				String value = getCellString(row, htPinkLableIndex.get(key));
				int valueNumber = (int) getCellDouble(row, htPinkLableIndex.get(key));

				if (value.equals(""))
					value = "" + valueNumber;

				htItemsPink.put(key, value);

				// System.out.println("Item get key : " + key);
				// System.out.println("Item get value : " + value);
			}

			alItem.add(new Item(INV_NO, BIG_DESC, SEQ, ITEM, PRODUCT_CODE, DESC_OF_GOODS, QTY, UNIT, PRICE, AMT,
					htItemsPink));

			continue;
		}
	}

	private class BSIExcelFilter extends FileFilter {
		@Override
		public boolean accept(File pathname) {
			String filename = pathname.getName();
			if (pathname.isDirectory()) {
				return true;

			} else if (filename.endsWith("xlsx") || filename.toUpperCase().startsWith("INVOICE")) {
				return true;
			} else {
				return false;
			}
		}

		@Override
		public String getDescription() {
			return "晟宇 Xlsx Files";
		}
	}

	public void run() {
		this.filesToExtract.clear();
		this.alItem.clear();
		this.alDeclr.clear();
		this.htPinkLable.clear();

		// get 原始WORD 檔
		JFileChooser fc = new JFileChooser();
		fc.setDialogTitle(programmeTitle);
		fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		fc.setMultiSelectionEnabled(true);

		FileFilter bsiFilter = new BSIExcelFilter();
		fc.addChoosableFileFilter(bsiFilter);
		fc.setFileFilter(bsiFilter);

		int returnVal = fc.showOpenDialog(null);

		if (returnVal == JFileChooser.OPEN_DIALOG) {
			File[] files = fc.getSelectedFiles();

			if (files.length > 0)
				Arrays.sort(files, NameFileComparator.NAME_INSENSITIVE_COMPARATOR);

			for (File file : files)
				this.filesToExtract.add(file);

		} else {
			super.infoBox("未選擇檔案", "錯誤訊息");
			return;
		}

		try {
			if (this.filesToExtract.size() == 0)
				throw new Exception("未發現可處理的檔案");

			// 處理 word file
			for (File f : this.filesToExtract) {
				String sAbsolutePath = super.xls2xlsx(f.getAbsolutePath(),
						"D:\\XML_OUTPUT\\自動產生檔-" + f.getName() + ".xlsx");
				extractExcelFile(sAbsolutePath);
			}

			check_and_merge_ArrayList();

			// 寫入EXCEL FILE
			doExcel();

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
