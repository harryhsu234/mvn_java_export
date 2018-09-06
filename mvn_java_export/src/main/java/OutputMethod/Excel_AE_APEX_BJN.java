package OutputMethod;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Hashtable;
import java.util.Iterator;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileFilter;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Excel_AE_APEX_BJN extends OutputCommon {

	public static String programmeTitle = "精銳科技_轉檔(重慶/北京)";

	XSSFWorkbook wb;
	XSSFSheet ws;

	protected ArrayList<File> filesToExtract = new ArrayList<File>();

	ArrayList<String> alKey = new ArrayList<String>();
	Hashtable<String, Double> htDeclr = new Hashtable<String, Double>();

	public Excel_AE_APEX_BJN() {
		programmeTitle = "精銳科技_轉檔(重慶/北京)";
	}

	protected void doExcel() throws Exception {
		String templatePath = "/Excel_AE_APEX.xlsx";
		InputStream tmpFile = this.getClass().getResourceAsStream(templatePath);

		wb = new XSSFWorkbook(tmpFile);
		ws = wb.getSheet("APEX");

		int row_pos = 1;
		boolean toPrintDesc = true; // 是否列印大品名

		int itemCount = 0;
		for(InvItem obj : this.alInvItem) {
			itemCount++;
			if(toPrintDesc) {

				this.setValue(row_pos, 3, obj.sDesc);
				row_pos++;
				
				toPrintDesc = false;
			}
			
			int col_pos = 0;
			this.setValue(row_pos, col_pos++, itemCount);
			
			String item = (obj.isBond.equals("NB") && obj.sDrno_type.startsWith("G")) ? "NIL" : obj.desc1;
			this.setValue(row_pos, col_pos++, item);
			this.setValue(row_pos, col_pos++, obj.po_no);
			this.setValue(row_pos, col_pos++, obj.desc1);
			this.setValue(row_pos, col_pos++, obj.desc2);
			this.setValue(row_pos, col_pos++, obj.iQty);
			this.setValue(row_pos, col_pos++, obj.iUnit_price);
			this.setValue(row_pos, col_pos++, obj.iAmt);
			this.setValue(row_pos, col_pos++, obj.sDrno_type);
			this.setValue(row_pos, col_pos++, obj.sStatistic_mode);
			this.setValue(row_pos, col_pos++, obj.sQtyUnit);
			this.setValue(row_pos, col_pos++, obj.sTrademark);
			this.setValue(row_pos, col_pos++, obj.isBond); 
			this.setValue(row_pos, col_pos++, obj.sCCCcode);
			this.setValue(row_pos, col_pos++, obj.nw);
			
			row_pos++;
		}

		outputFilePath = "D:\\XML_OUTPUT\\";
		outputFileName = "Excel_APEX_" + System.currentTimeMillis() + ".xlsx";
		Files.createDirectories(new File(outputFilePath).toPath());
		FileOutputStream stream = new FileOutputStream(outputFilePath + outputFileName);
		wb.write(stream);
		stream.close();

		wb.close();

		infoBox(outputFileName + " 產生完畢", "JOB_DONE");
	}

	protected void extractExcelFile(String sAbsolutePath) throws Exception {
		String ext = FilenameUtils.getExtension(sAbsolutePath);
		String filename = FilenameUtils.getBaseName(sAbsolutePath);//(sAbsolutePath);
		if(ext.equalsIgnoreCase("xls")) {
			sAbsolutePath = super.xls2xlsx(sAbsolutePath, "D:\\XML_OUTPUT\\自動產生檔-"+filename+".xlsx");
		}
		
		XSSFWorkbook xwb = new XSSFWorkbook(sAbsolutePath);
		int xwb_sheet_count = xwb.getNumberOfSheets();

		for (int i = 0; i < xwb_sheet_count; i++) {
			XSSFSheet xws = xwb.getSheetAt(i);

			extractExcelSheet(sAbsolutePath, xws);
		}

		xwb.close();
	}

	protected void mergeItems() {
			
//			for(InvItem invItem : alInvItem) {
//				int seq = invItem.itemSeq;
//				
//				for(DescItemPKGDetail pkgd : alDescItemPKGDetail) {
//					if(pkgd.itemIndex == seq) {
//						String htPKG_key = pkgd.boxID + "-"+pkgd.desc;
//						if(pkgd.desc.startsWith("P1101011003"))
//							htPKG_key = htPKG_key;
//						try {
//							
//							double nw = htPKGItem.get(htPKG_key);
//							invItem.nw += nw;
//						}
//						// catch(NullException )
//						catch(Exception ex) {
//							String exm = ex.getMessage();
//						}
//					}
//				}
//				
//				try {
//					invItem.isBond = (htDescItem.get(invItem.itemSeq).sBond.trim().equals("保稅"))? "YB":"NB";
//					
//				}
//				catch(Exception ex) {
//					String exm = ex.getMessage();
//				}
//			}
		System.out.println("MERGE");
		for (InvItem invItem : alInvItem) {
			int seq = invItem.itemSeq;
			
			try {
				// invItem.isBond = (htDescItem.get(invItem.itemSeq).sBond.trim().equals("保稅"))?
				// "YB":"NB";
				if (invItem.sDrno_type.startsWith("G5"))
					invItem.isBond = "NB";
				else if (invItem.sDrno_type.startsWith("B9"))
					invItem.isBond = "YB";
				
				invItem.nw = htPKGItem.get(invItem.description) * invItem.iQty;
			} catch (Exception ex) {
				String exm = ex.getMessage();
			}

		}
		}

	protected void extractExcelSheet(String fileName, XSSFSheet xws) throws Exception {
		String xwsSheetName = xws.getSheetName().toUpperCase().trim();
		if (xwsSheetName.startsWith("INV"))
		{
			extractINVtoArrayList(xws);
		} 
		else if (xwsSheetName.startsWith("PKG")) 
		{
			extractPKGtoArrayList(xws);
		} 
		else {
			// extractDESCtoArrayList(xws);
		}
	}

	Hashtable<Integer, String> htCccCode = new Hashtable<Integer, String>();
	protected void extractINVtoArrayList(XSSFSheet xws) throws Exception {
		System.out.println("執行 INVOICE 頁簽 : " + xws.getSheetName());
		String firstRowIdentifier = "P/O no.";
	
		Iterator<Row> rowIterator;
		boolean startItem = false;
		
		String sDesc = "";
		String sQtyUnit = "";
		String sTerm = "";
		String sCurr = "";
		String sTrademark = "";
		String sCCCcode = "";
	
		// 抓共用值
		startItem = false;
		rowIterator = xws.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
	
			if(row.getCell(0) == null) continue;
			
			Cell cell = row.getCell(0);
			// 確認第一行後 continue 抓第二行的共用值。
			if (cell.getCellTypeEnum() == CellType.STRING
					&& cell.getStringCellValue().trim().equalsIgnoreCase(firstRowIdentifier)) {
				startItem = true;
				continue;
			}
			if (startItem) {
				sDesc = row.getCell(1).getStringCellValue();
				sQtyUnit = row.getCell(3).getStringCellValue();
				sTerm = row.getCell(4).getStringCellValue();
				sCurr = row.getCell(5).getStringCellValue();
	
				startItem = false;
				continue; // 抓到共用值後就結束共用值的抓取動作
			}
	
			String prefix = "*商標:";
			if (cell.getCellTypeEnum() == CellType.STRING && cell.getStringCellValue().replace(" ", "").trim().startsWith(prefix)) {
				sTrademark = cell.getStringCellValue().replace(" ", "").trim().replace(prefix, "");
				continue;
			}
			prefix = "*BRAND:";
			if (cell.getCellTypeEnum() == CellType.STRING && cell.getStringCellValue().replace(" ", "").trim().startsWith(prefix)) {
				sTrademark = cell.getStringCellValue().replace(" ", "").trim().replace(prefix, "");
				continue;
			}
			
			prefix = "稅則";
			if (cell.getCellTypeEnum() == CellType.STRING && cell.getStringCellValue().replace(" ", "").trim().contains(prefix)) {
				String sLine = cell.getStringCellValue().replace(" ", "").trim();
				sCCCcode = sLine.substring(sLine.length()-15 , sLine.length()); // .replace(prefix, "");
				
				if(sLine.contains("第") && sLine.contains("項"))
				{
					// 特殊處理
					String invNos = sLine.substring(sLine.indexOf("第")+1, sLine.indexOf("項"));
					invNos = invNos.replace("、", ",");
					for(String invNos2: invNos.split(",")) {
						if(invNos2.contains("-")) {
							int invNo_from = Integer.parseInt(invNos2.split("-")[0].trim());
							int invNo_to = Integer.parseInt(invNos2.split("-")[1].trim());
							
							while (invNo_from <= invNo_to) {
								htCccCode.put(invNo_from, sCCCcode);
								invNo_from++;
							}
						}
						else
						{
							htCccCode.put(Integer.parseInt(invNos2), sCCCcode);
						}
							
					}
				}
				
				continue;
			}
		}
	
		startItem = false;
		rowIterator = xws.iterator();
		int col_number_of_drno_type = 7;
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			if(row.getCell(0) == null) continue;
			
			Cell cell = row.getCell(0);
			if (cell.getCellTypeEnum() == CellType.STRING
					&& cell.getStringCellValue().trim().equalsIgnoreCase(firstRowIdentifier)) {
				// 如果是第一行，多跳一行到第三行準備撈DETAIL
				rowIterator.next();
				startItem = true;
				continue;
			}
	
			if (startItem) {
				if (row.getCell(4).getNumericCellValue() > 0) {
					int invItemKey = (int) row.getCell(6).getNumericCellValue();
					
					String sPO_No = "";
					if(row.getCell(0).getCellTypeEnum() == CellType.STRING) 
						sPO_No = row.getCell(0).getStringCellValue().trim();
					else if(row.getCell(0).getCellTypeEnum() == CellType.NUMERIC) 
						sPO_No += (int) row.getCell(0).getNumericCellValue();
					
					String sBigDesc = row.getCell(1).getStringCellValue().trim();
					int iQty = (int) row.getCell(3).getNumericCellValue();
					int iUnit_price = (int) row.getCell(4).getNumericCellValue();
					int iAmt = iQty * iUnit_price;
					
					
					// 0 - based
					// 預設的報單類型是放在第七欄位，H欄，
					// 如果H欄是數字格式，有可能是APEX 段輸入錯誤，就改抓第八欄位 I 欄 
					if(row.getCell(col_number_of_drno_type).getCellTypeEnum() == CellType.NUMERIC) 
						col_number_of_drno_type = 8;
					String sDrno_type = row.getCell(col_number_of_drno_type).getStringCellValue().trim();
					
					int slashPos = sBigDesc.indexOf("/");
					String desc1 = "";
					String desc2 = "";
					if(slashPos > 0) {
						desc1 = sBigDesc.substring(0, slashPos);
						desc2 = sBigDesc.substring(slashPos+1, sBigDesc.length());
					}
					else {
						desc1 = sBigDesc;
						desc2 = ""; // sBigDesc.substring(slashPos+1, sBigDesc.length());
					}
					String sStatistic_mode = "";
					if(sDrno_type.trim().equals("")) sStatistic_mode = "";
					else {
						try {
							sStatistic_mode = sDrno_type.substring(sDrno_type.length()-2, sDrno_type.length());
						}
						catch(StringIndexOutOfBoundsException oobEx)
						{
							// throw new Exception("未能正確判斷INV 頁簽的報關類型");
							sStatistic_mode = "";
						}
					}
					
					if(htCccCode.size() > 0) {
						if(htCccCode.get(invItemKey) != null)
							sCCCcode = htCccCode.get(invItemKey);
						else 
							sCCCcode = "";
					}
					
					
					alInvItem.add(new InvItem(invItemKey, sPO_No, sBigDesc, desc1, desc2, iQty, iUnit_price, iAmt, sDrno_type, sStatistic_mode, 
														sDesc, sQtyUnit, sTerm, sCurr, sTrademark, sCCCcode)
								 );
					
				} 
				else {
					startItem = false;
					break;
				}
			}
			
		}
		System.out.println("InvItem count = " + alInvItem.size());
	}

	private void extractDESCtoArrayList(XSSFSheet xws) throws Exception {
		System.out.println("執行 包裝說明 頁簽 : " + xws.getSheetName());
		
		Iterator<Row> rowIterator;
		boolean startItem = false;
		
		rowIterator = xws.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			
			if(startItem && row.getCell(3) == null) break;
			else if(row.getCell(3) == null) continue;
			
			if(row.getCell(3).getCellTypeEnum() == CellType.NUMERIC) {
				startItem = true;
				
				int itemIndex = (int) row.getCell(3).getNumericCellValue();
				// String desc = alInvItem.get(itemIndex-1).description;
				String desc = "";
				try {
				desc = getItemFromArrayList(alInvItem, itemIndex).description;
				}
				catch (Exception e) {
					desc = "";
				}
				String sBond = row.getCell(4).getStringCellValue();
				int boxCount = (int) row.getCell(12).getNumericCellValue();
				String boxNo = row.getCell(13).getStringCellValue();
				
				if(boxNo.trim().isEmpty()) continue;
				
				htDescItem.put(itemIndex, new DescItem(itemIndex, sBond, boxCount, boxNo, desc));
			}
		}
	}
	private InvItem getItemFromArrayList(ArrayList<InvItem> alInvItem, int invItemKey) throws Exception {
		for(InvItem item : alInvItem) {
			if(item.itemSeq == invItemKey)
				return item;
		}
		System.out.println("INV 找不到項次" + invItemKey);
		throw new Exception("INV 找不到項次" + invItemKey);
		// return null;
	}
	protected void extractPKGtoArrayList(XSSFSheet xws) {
//		System.out.println("執行 PACKAGE 頁簽 : " + xws.getSheetName());
//		String firstRowIdentifier = "No.";
//		
//		Iterator<Row> rowIterator;
//		boolean startItem = false;
//		rowIterator = xws.iterator();
//		while (rowIterator.hasNext()) {
//			Row row = rowIterator.next();
//
//			if(row.getCell(0) == null) continue;
//			// System.out.println(" " + row.getCell(0).gets.getStringCellValue().trim());
//
//			Cell cell = row.getCell(0);
//			// 確認第一行後 continue 抓第二行的共用值。
//			if (cell.getCellTypeEnum() == CellType.STRING
//					&& cell.getStringCellValue().trim().equalsIgnoreCase(firstRowIdentifier)) {
//				startItem = true;
//				continue;
//			}
//			else if (startItem && cell.getCellTypeEnum() != CellType.NUMERIC) {
//				break;
//			}
//			
//			if(startItem) {
//				int pkgNo = (int)row.getCell(0).getNumericCellValue();
//				String desc = row.getCell(1).getStringCellValue().sp;
//				int qty = (int)row.getCell(3).getNumericCellValue();
//				double nw = row.getCell(4).getNumericCellValue();
//				
//			
//				String key = pkgNo + "-" + desc;
//				htPKGItem.put(key, nw);
//				// System.out.println("put " + key + " = " + nw);
//			}
//		}
		
		System.out.println("執行 PACKAGE 頁簽2 : " + xws.getSheetName());
		String firstRowIdentifier = "No.";
		
		Iterator<Row> rowIterator;
		boolean startItem = false;
		rowIterator = xws.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();

			if(row.getCell(0) == null) continue;
			// System.out.println(" " + row.getCell(0).gets.getStringCellValue().trim());

			Cell cell = row.getCell(0);
			// 確認第一行後 continue 抓第二行的共用值。
			if (cell.getCellTypeEnum() == CellType.STRING
					&& cell.getStringCellValue().trim().equalsIgnoreCase(firstRowIdentifier)) {
				startItem = true;
				continue;
			}
			else if (startItem && cell.getCellTypeEnum() == CellType.STRING 
								&& !StringUtils.isNumeric(cell.getStringCellValue().trim())) {
				break;
			}
			
			if(startItem) {
				
				String itemName = row.getCell(1).getStringCellValue();
				int qty = (int)row.getCell(3).getNumericCellValue();
				double nw = (double)row.getCell(4).getNumericCellValue();
				// if(gw <= 0) gw=0.5;
				
				double NwPerItem = nw / qty;
				
				String key = itemName;
				
				htPKGItem.put(key, NwPerItem);
			}
		}
	}
	ArrayList<InvItem> alInvItem = new ArrayList<InvItem>();
	Hashtable<Integer, DescItem> htDescItem = new Hashtable<Integer, DescItem>();

	ArrayList<DescItemPKGDetail> alDescItemPKGDetail = new ArrayList<DescItemPKGDetail>();

	Hashtable<String, Double> htPKGItem = new Hashtable<String, Double>();
	class InvItem {
		int itemSeq, iQty, iUnit_price, iAmt;
		String po_no, description, desc1, desc2, sDrno_type, sStatistic_mode;
		
		String sDesc, sQtyUnit, sTerm, sCurr, sTrademark, isBond = "", sCCCcode;
		
		double nw = 0;
		
		InvItem(int itemSeq, String po_no, String description, String desc1, String desc2, int iQty, int iUnit_price, int iAmt, String sDrno_type, String sStatistic_mode, 
							String sDesc, String sQtyUnit, String sTerm, String sCurr, String sTrademark, String sCCCcode) {
			this.itemSeq = itemSeq;
			this.po_no = po_no;
			this.description = description;
			this.desc1 = desc1;
			this.desc2 = desc2;
			this.iQty = iQty;
			this.iUnit_price = iUnit_price;
			this.iAmt = iAmt;
			this.sDrno_type = sDrno_type;
			this.sStatistic_mode = sStatistic_mode;
			
			this.sDesc = sDesc; // 大品名
			this.sQtyUnit = sQtyUnit;
			this.sTerm = sTerm;
			this.sCurr = sCurr;
			this.sTrademark = sTrademark;
			this.sCCCcode = sCCCcode;
		}
		
	}

	class DescItem {
		int itemIndex, boxCount;
		String sBond, boxNo, desc;
		
		DescItem(int itemIndex, String sBond, int boxCount, String boxNo, String desc) {
			this.itemIndex = itemIndex;
			this.sBond = sBond;
			this.boxCount = boxCount;
			this.boxNo = boxNo;
			this.desc = desc;
		
			insertBoxInfo(itemIndex, boxNo, boxCount, desc);
		}
		
		void insertBoxInfo(int itemIndex, String boxNo, int boxCount, String desc) {
			String chr10 = "\n";
			boxNo = boxNo.replace(" ", chr10).replace(".", "");
	
			// 先解析兩行的
			for(String line : boxNo.split(chr10)) {
				if(line.length() < 1) continue; // = line.trim();
				
				int pkgCount = 0;
				if(line.contains("*")) {
					pkgCount = Integer.parseInt(line.split("\\*")[1].trim());
					line = line.substring(0, line.indexOf("*"));
				}
				else if(line.contains("單")) 
				{
					pkgCount = 1;
					line = line.replaceAll("單", "");
				}
				else 
					pkgCount = boxCount;
								
				if(line.contains("-")) {
					int boxID_from = Integer.parseInt(line.split("-")[0].trim());
					int boxID_to = Integer.parseInt(line.split("-")[1].trim());
					
					while (boxID_from <= boxID_to) {
						alDescItemPKGDetail.add(new DescItemPKGDetail(itemIndex, boxID_from, pkgCount, desc));
						// System.out.println(itemIndex + " 箱型: " + boxID_from + "; 箱數: " + pkgCount);
						
						boxID_from++;
					}
				}
				else {
					alDescItemPKGDetail.add(new DescItemPKGDetail(itemIndex, Integer.parseInt(line.trim()), pkgCount, desc));
					// System.out.println(itemIndex + " 箱型: " + line.trim() + "; 箱數: " + pkgCount);
				}
				
			}
		}
	}


	class DescItemPKGDetail {
		int itemIndex, boxID, pkgCount;
		String desc;
		DescItemPKGDetail(int itemIndex, int boxID, int pkgCount, String desc) {
			this.itemIndex = itemIndex;
			this.boxID = boxID;
			this.pkgCount = pkgCount;
			this.desc = desc;
		}
	}


	private String getCellStringValue(XSSFSheet xws, String address) {
		return getCellStringValue(xws, new CellAddress(address));
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
		return getCellNumericValue(xws, new CellAddress(address));
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

	public void getFileInDir(File[] files) {
		for (File f : files) {
			if (f.isDirectory())
				getFileInDir(f.listFiles());
			// else if (f.isFile() && f.getName().startsWith("EC-HK INVOICE"))
			else if (f.isFile())
				this.filesToExtract.add(f);
		}
	}

	public void run() {
	
		this.filesToExtract.clear();

		// get 原始WORD 檔
		JFileChooser fc = new JFileChooser();
		fc.setDialogTitle(this.programmeTitle);
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


			// 處理 word file
			for (File f : this.filesToExtract) {
				
				extractExcelFile(f.getAbsolutePath());
			}
			
			mergeItems();
			
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
