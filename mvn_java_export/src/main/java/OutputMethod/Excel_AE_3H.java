package OutputMethod;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Hashtable;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileFilter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_AE_3H extends OutputCommon {

	private class Item {

		String fileName, tabName, PO, PN, desc, PCS, pn, amt;
		double qty, unitPrice, nw;

		public Item(String fileName, String tabName, String PO, String PN, String desc, double qty, String PCS, double unitPrice, String amt, double nw) {
			super();
			this.fileName = fileName;
			this.tabName = tabName;
			this.PO = PO;
			this.PN = PN;
			this.desc = desc;
			this.qty = qty;
			this.PCS = PCS;
			this.unitPrice = unitPrice;
			this.amt = amt;
			this.nw = nw;
		}

	}

	public final static String programmeTitle = "AE_蜜望實_20180111";

	String descBreaker = "";
	XSSFWorkbook wb;
	XSSFSheet ws;

	private ArrayList<File> filesToExtract = new ArrayList<File>();
	private ArrayList<Item> alItem = new ArrayList<Item>();
	// private ArrayList<Declr> alDeclr = new ArrayList<Declr>();
	
	ArrayList<String> alKey = new ArrayList<String>();
	Hashtable<String, Double> htDeclr = new Hashtable<String, Double>();

	public Excel_AE_3H() {
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
//			this.setValue(row_pos, col_pos++, obj.fileName);
//			this.setValue(row_pos, col_pos++, obj.inv_no);
//			this.setValue(row_pos, col_pos++, obj.itemNo);
//			this.setValue(row_pos, col_pos++, obj.product.replace("Product：", "").trim());
//			// this.setValue(row_pos, col_pos++, obj.materialCode.replace(".0", "").trim());
//			this.setValue(row_pos, col_pos++, "DO NO:" + obj.doNo);
//			this.setValue(row_pos, col_pos++, "P/N:" + obj.pn);
//			this.setValue(row_pos, col_pos++, obj.description);
//			this.setValue(row_pos, col_pos++, obj.qty);
//			this.setValue(row_pos, col_pos++, obj.unitPrice);
//			this.setValue(row_pos, col_pos++, obj.qty*obj.unitPrice);
			

			row_pos++;
		}

		

		outputFilePath = "D:\\XML_OUTPUT\\";
		outputFileName = "Excel_3H_" + System.currentTimeMillis() + ".xlsx";
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
		
		ArrayList<Item> list = new ArrayList<Item>();
		String currentBreaker = "";

		for (int i = 0; i < xwb_sheet_count; i++) {
			XSSFSheet xws = xwb.getSheetAt(i);

			if(xws.getSheetName().startsWith("I")) {
				currentBreaker = (!getCellStringValue(xws, "C28").isEmpty()) ? getCellStringValue(xws, "C28") : getCellStringValue(xws, "B28"); 
				
				int row = 29;
				list.clear();
				
				 do
                 {
					 String fileName = sAbsolutePath;
					 String tabName = xws.getSheetName();
					 String PO = xws.getRow(row).getCell(0).getStringCellValue();//selWS.Cells[row, 1].Value.ToString(),
					 String PN = xws.getRow(row).getCell(1).getStringCellValue();//selWS.Cells[row, 2].Value.ToString();
					 String desc = xws.getRow(row).getCell(2).getStringCellValue();//selWS.Cells[row, 3].Value.ToString();
	                 double qty = xws.getRow(row).getCell(4).getNumericCellValue();// double.Parse(selWS.Cells[row, 5].Value.ToString());
	                 String PCS = xws.getRow(row).getCell(5).getStringCellValue();//selWS.Cells[row, 6].Value.ToString();
	                 double unitPrice = xws.getRow(row).getCell(6).getNumericCellValue();// double.Parse(selWS.Cells[row, 7].Value.ToString());
	                 String amt = xws.getRow(row).getCell(7).getStringCellValue();//selWS.Cells[row, 8].Value.ToString();
	                 double nw = -54689;
					 
                     list.add(new Item(fileName, tabName, PO, PN, desc, qty, PCS, unitPrice, amt, nw));

                     row++;
                 } while (!xws.getRow(row).getCell(1).getStringCellValue().isEmpty());
				
			} // end of if sheet name startsWith("I")
			else if(xws.getSheetName().startsWith("P")) {
				 int row = 29;
                 int p_count = 0;
                 do
                 {
                     p_count++;
                     row++;
                 } while (!xws.getRow(row).getCell(2).getStringCellValue().isEmpty());
			}
			
			if (currentBreaker != descBreaker)
            {
                descBreaker = currentBreaker;
                // ws.Cells[row_pos++, 3].Value = descBreaker;
            }
		}

		xwb.close();

	}

	private void checkAlItem() {
		System.out.println("Item count = " + alItem.size());

		for (Item item : alItem) {
			System.out.println(item.fileName);
			System.out.println(item.tabName);
			System.out.println(item.PO);
			System.out.println(item.PN);
			System.out.println(item.desc);
			System.out.println(item.qty);
			System.out.println(item.PCS);
			System.out.println(item.unitPrice);
			System.out.println(item.amt);
			System.out.println(item.nw);
		}
	}

	
	private void extractExcelSheet(String fileName, XSSFSheet xws) {
		if (xws.getSheetName().startsWith("I")) 
			return;
		else if(xws.getSheetName().startsWith("P"))
			return;
		
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
		boolean notImplement = true;
		if(notImplement) {
			System.err.println("Method not complete");
			super.infoBox("功能未完成", "No function");
			return;
			
		}
		
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
			

			// 處理 word file
			for (File f : this.filesToExtract) {
				extractExcelFile(f.getAbsolutePath());
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

	private void realJob() throws Exception {
		int row_pos = 0;
		for (File f : this.filesToExtract) {
			
			ArrayList<Item> list = new ArrayList<Item>();
            String currentBreaker = "";
            
			XSSFWorkbook xwb = new XSSFWorkbook(f.getAbsolutePath());
			int xwb_sheet_count = xwb.getNumberOfSheets();
			
			for (int i = 0; i < xwb_sheet_count; i++) {
				XSSFSheet xws = xwb.getSheetAt(i); 
				
				if (xws.getSheetName().startsWith("I")) {
					currentBreaker = (!getCellStringValue(xws, "C28").isEmpty()) ? getCellStringValue(xws, "C28") : getCellStringValue(xws, "B28"); 
					
					int row = 29;
					list.clear();
					
					do {
						String fileName = f.getAbsolutePath();
						 String tabName = xws.getSheetName();
						 String PO = xws.getRow(row).getCell(0).getStringCellValue();//selWS.Cells[row, 1].Value.ToString(),
						 String PN = xws.getRow(row).getCell(1).getStringCellValue();//selWS.Cells[row, 2].Value.ToString();
						 String desc = xws.getRow(row).getCell(2).getStringCellValue();//selWS.Cells[row, 3].Value.ToString();
		                 double qty = xws.getRow(row).getCell(4).getNumericCellValue();// double.Parse(selWS.Cells[row, 5].Value.ToString());
		                 String PCS = xws.getRow(row).getCell(5).getStringCellValue();//selWS.Cells[row, 6].Value.ToString();
		                 double unitPrice = xws.getRow(row).getCell(6).getNumericCellValue();// double.Parse(selWS.Cells[row, 7].Value.ToString());
		                 String amt = xws.getRow(row).getCell(7).getStringCellValue();//selWS.Cells[row, 8].Value.ToString();
		                 double nw = -54689;
						 
	                     list.add(new Item(fileName, tabName, PO, PN, desc, qty, PCS, unitPrice, amt, nw));

	                     row++;
	                 } while (!xws.getRow(row).getCell(1).getStringCellValue().isEmpty());
					
				}
				else if(xws.getSheetName().startsWith("P")) {
					
					 int row = 29;
                     int p_count = 0;
                     do
                     {
                         p_count++;
                         row++;
                     } while (!xws.getRow(row).getCell(2).getStringCellValue().isEmpty());

					
                     if(list.size() == p_count) {
                    	 row = 29;
                    	 p_count= 0;
                    	 
                    	 do
                         {
                             
                    		 ((Item)list.get(p_count)).nw = Double.parseDouble(xws.getRow(row).getCell(7).getStringCellValue());
                    		 // ((Item)list[p_count]).nw = 0; //double.Parse(selWS.Cells[row, 8].Value.ToString());

                             p_count++;
                             row++;
                         } while (!xws.getRow(row).getCell(2).getStringCellValue().isEmpty());
                    	 
                    	 if (currentBreaker != descBreaker)
                         {
                             descBreaker = currentBreaker;

                             setValue(row_pos++, 2, descBreaker);
                             // ws.Cells[row_pos++, 3].Value = descBreaker;
                         }
                    	 
                    	 for(Item item : list) {
                    		 int col_pos = 0;
                    		 
//                    		 String fileName, tabName, PO, PN, desc, PCS, pn, amt;
//                    			double qty, unitPrice, nw;
                 			this.setValue(row_pos, col_pos++, item.fileName);
                 			this.setValue(row_pos, col_pos++, item.tabName);
                 			this.setValue(row_pos, col_pos++, item.PO);
                 			this.setValue(row_pos, col_pos++, item.PN);
                 			this.setValue(row_pos, col_pos++, item.desc);
                 			col_pos++;
                 			this.setValue(row_pos, col_pos++, item.qty);
                 			this.setValue(row_pos, col_pos++, item.PCS);
                 			this.setValue(row_pos, col_pos++, item.unitPrice);
                 			this.setValue(row_pos, col_pos++, item.amt);
                 			if(item.nw == -54689) {
                 				
                 			}
                 			else {
                 				this.setValue(row_pos, col_pos++, item.nw);
                 			}
                 			row_pos++;
                 			
                 			
                    	 }
                     }
				}
					
			}
			
				
			
		}
		
	}
	
	private void setValue(int row_pos, int col_pos, Object value) throws Exception {
		// create and set cell
		if (ws.getRow(row_pos) == null) {
			ws.createRow(row_pos);
		}
		if (ws.getRow(row_pos).getCell(col_pos) == null) {
			XSSFCell newCell = ws.getRow(row_pos).createCell(col_pos);
			try {
				newCell.setCellStyle(ws.getRow(3).getCell(col_pos).getCellStyle());
			} catch (Exception ex) {
				String exm = ex.getMessage();
				System.out.println(exm);
				System.out.println(ex.getStackTrace());
			}
		}

		// 判斷填入值是否為NULL
		if (value == null) {
			System.err.println("Value is null");
			return;
		}
		String className = value.getClass().getName();
		System.out.println(value + " is " + className);

		if (className == "java.lang.Integer")
			ws.getRow(row_pos).getCell(col_pos).setCellValue((Integer) value);
		else if (className == "java.lang.Double")
			ws.getRow(row_pos).getCell(col_pos).setCellValue((Double) value);
		else if (className == "java.lang.String")
			ws.getRow(row_pos).getCell(col_pos).setCellValue((String) value);
		else
			throw new Exception("Cell format not supported: " + className);

	}

}
