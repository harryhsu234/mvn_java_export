package OutputMethod;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Hashtable;
import java.util.Iterator;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileFilter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_AE_DANLI extends OutputCommon {

	private class Declr {
		String desc, curr;
		double amount;

		Declr(String desc, String curr, double amount) {
			this.desc = desc;
			this.curr = curr;
			this.amount = amount;
		}
	}

	private class Item {

		String fileName, inv_no, itemNo, product, materialCode, doNo, pn, description;
		double qty, unitPrice, amount;

		public Item(String fileName, String inv_no, String itemNo, String product, String materialCode, String doNo, String pn,
				String description, double qty, double unitPrice, double amount) {
			super();
			this.fileName = fileName;
			this.inv_no = inv_no;
			this.itemNo = itemNo;
			this.product = product;
			this.materialCode = materialCode;
			this.doNo = doNo;
			this.pn = pn;
			this.description = description;
			this.qty = qty;
			this.unitPrice = unitPrice;
			this.amount = amount;
		}

	}

	public final static String programmeTitle = "AE_丹利_20180110";

	XSSFWorkbook wb;
	XSSFSheet ws;

	private ArrayList<File> filesToExtract = new ArrayList<File>();
	private ArrayList<Item> alItem = new ArrayList<Item>();
	private ArrayList<Declr> alDeclr = new ArrayList<Declr>();
	
	ArrayList<String> alKey = new ArrayList<String>();
	Hashtable<String, Double> htDeclr = new Hashtable<String, Double>();

	public Excel_AE_DANLI() {
		// TODO Auto-generated constructor stub
	}

	private void doExcel() throws Exception {
		String templatePath = "/Excel_AE_DANLI.xlsx";
		InputStream tmpFile = this.getClass().getResourceAsStream(templatePath);

		wb = new XSSFWorkbook(tmpFile);
		ws = wb.getSheet("品名");

		int row_pos = 1;
		for (Item obj : this.alItem) {
			int col_pos = 0;
			this.setValue(row_pos, col_pos++, obj.fileName);
			this.setValue(row_pos, col_pos++, obj.inv_no);
			this.setValue(row_pos, col_pos++, obj.itemNo);
			this.setValue(row_pos, col_pos++, obj.product.replace("Product：", "").trim());
			// this.setValue(row_pos, col_pos++, obj.materialCode.replace(".0", "").trim());
			this.setValue(row_pos, col_pos++, "DO NO:" + obj.doNo);
			this.setValue(row_pos, col_pos++, "P/N:" + obj.pn);
			this.setValue(row_pos, col_pos++, obj.description);
			this.setValue(row_pos, col_pos++, obj.qty);
			this.setValue(row_pos, col_pos++, obj.unitPrice);
			this.setValue(row_pos, col_pos++, obj.qty*obj.unitPrice);
			

			row_pos++;
		}

		ws = wb.getSheet("申報事項");
		row_pos = 1;
		for (Declr obj : this.alDeclr) {
			int col_pos = 0;
			this.setValue(row_pos, col_pos++, obj.desc + obj.amount);
//			this.setValue(row_pos, col_pos++, obj.curr);
//			this.setValue(row_pos, col_pos++, obj.amount);

			row_pos++;
		}

		outputFilePath = "D:\\XML_OUTPUT\\";
		outputFileName = "Excel_DANLI_" + System.currentTimeMillis() + ".xlsx";
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
		int xwb_sheet_count = xwb.getNumberOfSheets();

		for (int i = 0; i < xwb_sheet_count; i++) {
			XSSFSheet xws = xwb.getSheetAt(i);

			extractExcelSheet(sAbsolutePath, xws);
		}

		xwb.close();

	}

	private void checkAlItem() {
		System.out.println("Item count = " + alItem.size());

		for (Item item : alItem) {
			System.out.println(item.fileName);
			System.out.println(item.inv_no);
			System.out.println(item.itemNo);
			System.out.println(item.product);
			System.out.println(item.materialCode);
			System.out.println(item.doNo);
			System.out.println(item.pn);
			System.out.println(item.description);
			System.out.println(item.qty);
			System.out.println(item.unitPrice);
			System.out.println(item.amount);
		}
	}

	
	private void extractExcelSheet(String fileName, XSSFSheet xws) {
		if (!xws.getSheetName().startsWith("CommerciaInvoice")) // 只對 CommerciaInvoice 的頁簽動作
			return;
		
		CellAddress caInvNo = null;
		CellAddress caProduct = null;
		CellAddress caRemark = null; // REMARK:
		CellAddress caDeliveryDate = null; // REMARK:
		
		// begin - scan over cells 掃瞄所有格子
		Iterator<Row> rowIterator = xws.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				CellType cType = cell.getCellTypeEnum();
				// int ctype = cell.getCellType();
				String cell_value = "";

				switch (cType) {
				case BLANK: 
					cell_value = "*blank";
					break;
				case NUMERIC:
					DecimalFormat df = new DecimalFormat("###.#");
					cell_value += df.format(cell.getNumericCellValue());
					break;
				case STRING:
					cell_value = cell.getStringCellValue();
					// DLI Invoice #
					if(cell_value.toUpperCase().startsWith("DLI Invoice #".toUpperCase()))
						caInvNo = new CellAddress(cell.getAddress().getRow(),cell.getAddress().getColumn()+1);
					else if(cell_value.toUpperCase().startsWith("Product：".toUpperCase()))
						caProduct = cell.getAddress();
					else if(cell_value.toUpperCase().startsWith("REMARK:".toUpperCase()))
						caRemark = cell.getAddress();
					else if(cell_value.toUpperCase().startsWith("1. Date of delivery :".toUpperCase()))
						caDeliveryDate = new CellAddress(cell.getAddress().getRow(),cell.getAddress().getColumn()+2);
					break;
				default:
					break;
				}

				if (cType != CellType.BLANK) {
					// System.out.println(cell.getAddress() + " is " + cell_value);
				}
			}
		}
		// end - scan over cells 掃瞄所有格子

		
		DecimalFormat df = new DecimalFormat("###.#");
		// String sInvNo = "" + getCellNumericValue(xws, "H3"); // get Invoice_no
		String sInvNo = df.format(getCellNumericValue(xws, caInvNo)); // get Invoice_no
		String sProduct = getCellStringValue(xws, caProduct); // get Product
		
		
		getItems(fileName, xws, sInvNo, sProduct);
		
		if(caRemark != null)
			getDelrs(xws, caRemark);
	}

	private void getDelrs(XSSFSheet xws, CellAddress caRemark) {
		if(caRemark == null) return;
		
		int row = caRemark.getRow();
		int col = caRemark.getColumn();
		
		
		
		do {
			String sDelr = xws.getRow(row).getCell(col).getStringCellValue();
			sDelr = sDelr.replace("REMARK:", "").trim();
			
			String desc, curr;
			double amount;
			
			String _dollar_sign = "XX_DOLLAR_SIGN_XX";
			sDelr = sDelr.replace("$", _dollar_sign);
			String[] spliters = new String[] { "US"+_dollar_sign, "TW"+_dollar_sign };
			
			String spliter = "";
			for(String _spliter : spliters) {
				if(sDelr.contains(_spliter))
					spliter = _spliter;
				
				if(spliter.length() > 0)
					break;
			}
			
			String[] saDelr = sDelr.split(spliter);
			
			if(spliter.length() > 0 && saDelr.length > 1) {
				desc = saDelr[0].trim();
				amount = Double.parseDouble(saDelr[1].replace(",", "").trim());
			}
			else {
				desc = sDelr.trim();
				amount = 0;
			}
			
			String key = (desc+spliter).replace(_dollar_sign, "$");
			
			if(alKey.contains(key)) {
				double old_amt = htDeclr.get(key);
				htDeclr.put(key, old_amt+amount);
			}
			else {
				alKey.add(key);
				htDeclr.put(key, amount);
			}
			
			row++;
		} while(xws.getRow(row).getCell(col).getCellTypeEnum() == CellType.STRING);
		
		
		
		return;
	}

	private void getItems(String fileName, XSSFSheet xws, String sInvNo, String sProduct) {
		Iterator<Row> rowIterator;
		boolean startItem = false;
		int itemCount = 1;
		Item item = null;
		
		rowIterator = xws.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();

			Cell cell = row.getCell(0);

			if (cell.getCellTypeEnum() == CellType.STRING
					&& cell.getStringCellValue().trim().equalsIgnoreCase("Item #")) {
				startItem = true;
				continue;
			}

			if (startItem) {
				if (itemCount == row.getCell(0).getNumericCellValue()) {
					if (item != null) {
						// 遇到下一個ITEM 編號
						alItem.add(item);
					}
//					String fileName = fileName;
					String inv_no = sInvNo;
					String itemNo = "" + itemCount;
					String product = sProduct;
					
					String materialCode = "";
					if(row.getCell(2).getCellTypeEnum() == CellType.STRING)
						materialCode += row.getCell(2).getStringCellValue();
					else if(row.getCell(2).getCellTypeEnum() == CellType.NUMERIC)
						materialCode += row.getCell(2).getNumericCellValue();
					
					String doNo = row.getCell(3).getStringCellValue();
					String pn = row.getCell(4).getStringCellValue();
					String description = row.getCell(5).getStringCellValue();
					double qty = row.getCell(1).getNumericCellValue();
					double unitPrice = row.getCell(6).getNumericCellValue();
					double amount = row.getCell(7).getNumericCellValue();

					item = new Item(fileName, inv_no, itemNo, product, materialCode, doNo, pn, description, qty, unitPrice, amount);

					itemCount++;
				}

				if (row.getCell(5).getCellTypeEnum() == CellType.BLANK) {
					// ITEM 項目終結
					alItem.add(item);
					startItem = false;
					
				} else if (row.getCell(0).getCellTypeEnum() == CellType.BLANK){  
					String chr10 = "\n";
					
					String description = row.getCell(5).getStringCellValue();
					double unitPrice = row.getCell(6).getNumericCellValue();
					double amount = row.getCell(7).getNumericCellValue();
					
					// 遇到加工費
					// 只加UNIT PRICE 不加DESCRIPTION
					if(!description.startsWith("IC 加工費"))
						item.description += chr10 + description;
					
					item.unitPrice += unitPrice;
					item.amount += amount;
				}
			} // end of startItem
		}
	}

	private String getCellStringValue(XSSFSheet xws, String address) {
		return getCellStringValue(xws,new CellAddress(address));
	}
	private String getCellStringValue(XSSFSheet xws, CellAddress ca) {
		String val = "";
		try {
			Cell cell = xws.getRow(ca.getRow()).getCell(ca.getColumn());
			val = "" + cell.getStringCellValue();
		} catch (Exception ex) {
			System.err.println("Error in converting cell - " + ca.formatAsString());
			System.err.println(ex.getMessage());
		}
		// end - get sheet invoice_no
		return val;
	}
	private double getCellNumericValue(XSSFSheet xws, String address) {
		return getCellNumericValue(xws,new CellAddress(address));
	}
	private double getCellNumericValue(XSSFSheet xws, CellAddress ca) {
		double val = 0;
		try {
			Cell cell = xws.getRow(ca.getRow()).getCell(ca.getColumn());
			val = cell.getNumericCellValue();
		} catch (Exception ex) {
			System.err.println("Error in converting cell - " + ca.formatAsString());
			System.err.println(ex.getMessage());
		}
		// end - get sheet invoice_no
		return val;
	}

	
	private void getFileInDir(File[] files) {
		for (File f : files) {
			if (f.isDirectory())
				getFileInDir(f.listFiles());
			else if (f.isFile() && f.getName().startsWith("EC-HK INVOICE"))
				this.filesToExtract.add(f);
		}
	}

	public void run() {
		this.filesToExtract.clear();

		// get 原始WORD 檔
		JFileChooser fc = new JFileChooser();
		fc.setDialogTitle(programmeTitle);
		fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
		fc.setMultiSelectionEnabled(true);
		
		FileFilter filter = new ExcelFilter();
		fc.addChoosableFileFilter(filter);
		fc.setFileFilter(filter);

		int returnVal = fc.showOpenDialog(null);

		if (returnVal == JFileChooser.OPEN_DIALOG) {
			File[] files = fc.getSelectedFiles();

			// get all files
			getFileInDir(files);

		} else {
			super.infoBox("未選擇檔案", "錯誤訊息");
			return;
		}
		try {

			// 驗證抓到的DOC 檔
			if (this.filesToExtract.size() == 0)
				throw new Exception("未發現可處理的檔案");
			// for(File f : this.filesToExtract) {
			// // extractWordFile(f.getAbsolutePath());
			// }

			// 處理 word file
			for (File f : this.filesToExtract) {
				extractExcelFile(f.getAbsolutePath());
			}

			for(String key:alKey) {
				this.alDeclr.add(new Declr(key, "", htDeclr.get(key)));
			}
			
			checkAlItem();

			// 寫入EXCEL FILE
			doExcel();

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			super.infoBox(e.getMessage(), "Error in " + this.getClass());
		}
	}

	private void setValue(int row_pos, int col_pos, Object value) throws Exception {
		super.setValue(ws, row_pos, col_pos, value);
	}

}
