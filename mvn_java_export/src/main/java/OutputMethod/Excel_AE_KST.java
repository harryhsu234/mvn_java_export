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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_AE_KST extends OutputCommon {

	private class Item {
		String PKD_ITM_CODE, EID_CUST_ITEM, PKD_ITM_NAME, GBM_ALLOWDOC;
		String TOTAL_QUANTITY, PKL_TOPMARK, ITL_DESCRIPTION;
		String GIB, COUNTRY;
		double PKD_TQTY, UNIPRICE, AMT, NW;

		public Item() {
		}

		public Item(String PKD_ITM_CODE, String EID_CUST_ITEM, String PKD_ITM_NAME, String GBM_ALLOWDOC, double AMT,
				String TOTAL_QUANTITY, double UNIPRICE, double PKD_TQTY, String PKL_TOPMARK, String ITL_DESCRIPTION,
				double NW, String GIB, String COUNTRY) {
			super();
			this.PKD_ITM_CODE = PKD_ITM_CODE;
			this.EID_CUST_ITEM = EID_CUST_ITEM;
			this.PKD_ITM_NAME = PKD_ITM_NAME;
			this.GBM_ALLOWDOC = GBM_ALLOWDOC;
			this.PKD_TQTY = PKD_TQTY;

			this.TOTAL_QUANTITY = TOTAL_QUANTITY;
			this.UNIPRICE = UNIPRICE;
			this.AMT = AMT;
			this.PKL_TOPMARK = PKL_TOPMARK;
			this.ITL_DESCRIPTION = ITL_DESCRIPTION;

			this.NW = NW;
			this.GIB = GIB;
			this.COUNTRY = COUNTRY;
		}
	}

	public final static String programmeTitle = "AE_世同金屬轉檔";

	XSSFWorkbook wb;
	XSSFSheet ws;

	private ArrayList<File> filesToExtract = new ArrayList<File>();
	private ArrayList<Item> alItem = new ArrayList<Item>();
	private ArrayList<Double> alNW = new ArrayList<Double>();

	public Excel_AE_KST() {
		// TODO Auto-generated constructor stub
	}

	private void doExcel() throws Exception {
		String templatePath = "/Excel_AE_KST.xlsx";
		InputStream tmpFile = this.getClass().getResourceAsStream(templatePath);

		wb = new XSSFWorkbook(tmpFile);
		ws = wb.getSheetAt(0);

		int row_pos = 1;
		for (Item obj : this.alItem) {
			int col_pos = 0;

			this.setValue(row_pos, col_pos++, obj.PKD_ITM_CODE);
			this.setValue(row_pos, col_pos++, obj.EID_CUST_ITEM);
			this.setValue(row_pos, col_pos++, obj.PKD_ITM_NAME);
			this.setValue(row_pos, col_pos++, obj.GBM_ALLOWDOC);
			this.setValue(row_pos, col_pos++, obj.AMT);

			this.setValue(row_pos, col_pos++, obj.TOTAL_QUANTITY);
			this.setValue(row_pos, col_pos++, obj.UNIPRICE);
			this.setValue(row_pos, col_pos++, obj.PKD_TQTY);
			this.setValue(row_pos, col_pos++, obj.PKL_TOPMARK); // 強制空白
			this.setValue(row_pos, col_pos++, obj.ITL_DESCRIPTION);

			this.setValue(row_pos, col_pos++, obj.NW);
			this.setValue(row_pos, col_pos++, obj.GIB);
			this.setValue(row_pos, col_pos++, obj.COUNTRY);
			

			row_pos++;
		}

		outputFilePath = "D:\\XML_OUTPUT\\";
		outputFileName = "Excel_KST_" + System.currentTimeMillis() + ".xlsx";
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

		if (sAbsolutePath.contains("-INV")) {
			extractInvSheet(sAbsolutePath, xwb.getSheetAt(0));
		} else if (sAbsolutePath.contains("-PK")) {
			extractPkSheet(sAbsolutePath, xwb.getSheetAt(0));
		}

		xwb.close();
	}

	private void extractPkSheet(String sAbsolutePath, XSSFSheet xws) {
		// TODO Auto-generated method stub
		System.out.println("Start Extract PK excel");
		alNW.clear();

		Iterator<Row> rowIterator = xws.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();

			// 從第二行開始
			if (row.getRowNum() < 1) {
				continue;
			}

			if(getCellString(row, 0).equals(""))
				break;

			
			try {
				alNW.add(getCellDouble(row, 32));
			}
			catch(Exception ex) {
				
				alNW.add((double) 0);
			}
			
		}
	}

	private void extractInvSheet(String sAbsolutePath, XSSFSheet xws) {
		// TODO Auto-generated method stub
		System.out.println("Start Extract INV excel");
		alItem.clear();

		Iterator<Row> rowIterator = xws.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			
			// 從第二行開始
			if (row.getRowNum() < 1) {
				continue;
			}

			if(getCellString(row, 0).equals(""))
				break;

			String PKD_ITM_CODE = getCellString(row, 7);
			String EID_CUST_ITEM = getCellString(row, 22);
			String PKD_ITM_NAME = getCellString(row, 8);
			String GBM_ALLOWDOC = "CUSTOM BOM APPROVAL NO. " + getCellString(row, 38);
			double AMT = getCellDouble(row, 34);

			String TOTAL_QUANTITY = getCellString(row, 11);
			double UNIPRICE = getCellDouble(row, 17);
			double PKD_TQTY = getCellDouble(row, 32);
			String PKL_TOPMARK = "";
			String ITL_DESCRIPTION = getCellString(row, 45);

			double NW = 0;
			String GIB = getCellString(row, 33);
			String COUNTRY = getCellString(row, 48);

			alItem.add(new Item(PKD_ITM_CODE, EID_CUST_ITEM, PKD_ITM_NAME, GBM_ALLOWDOC, AMT, TOTAL_QUANTITY, UNIPRICE,
					PKD_TQTY, PKL_TOPMARK, ITL_DESCRIPTION, NW, GIB, COUNTRY));

		}
	}

	
	private class KSTExcelFilter extends FileFilter {
		@Override
		public boolean accept(File pathname) {
			String filename = pathname.getName();
			if (pathname.isDirectory()) {
				return true;

			} else if ((filename.endsWith("xls") || filename.endsWith("xlsx"))
					&& (filename.contains("-INV") || filename.contains("-PK"))) {
				return true;
			} else {
				return false;
			}
		}

		@Override
		public String getDescription() {
			return "KST Excel Files";
		}
	}

	public void run() {
		this.filesToExtract.clear();

		// get 原始WORD 檔
		JFileChooser fc = new JFileChooser();
		fc.setDialogTitle(programmeTitle);
		fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		fc.setMultiSelectionEnabled(false);
		
		FileFilter kstFilter = new KSTExcelFilter();
		fc.addChoosableFileFilter(kstFilter);
		fc.setFileFilter(kstFilter);

		int returnVal = fc.showOpenDialog(null);

		if (returnVal == JFileChooser.OPEN_DIALOG) {
			File file = fc.getSelectedFile();

			// check brother file
			File file_brother = null;
			String fileABSName = file.getAbsolutePath();
			String brotherABSName = "";
			if (fileABSName.contains("-INV")) {
				brotherABSName = fileABSName.replace("-INV", "-PK.");
			} else if (fileABSName.contains("-PK.")) {
				brotherABSName = fileABSName.replace("-PK.", "-INV");
			}

			String ext1 = FilenameUtils.getExtension(fileABSName);
			String ext2 = FilenameUtils.getExtension(brotherABSName);

			String filename1 = FilenameUtils.getBaseName(fileABSName);// (sAbsolutePath);
			if (ext1.equalsIgnoreCase("xls")) {
				try {
					fileABSName = super.xls2xlsx(fileABSName, "D:\\XML_OUTPUT\\自動產生檔-" + filename1 + ".xlsx");
				} catch (Exception ex) {
					System.out.println("EXT1 : " + ex.getMessage());
				}
			}

			String filename2 = FilenameUtils.getBaseName(brotherABSName);// (sAbsolutePath);
			if (ext2.equalsIgnoreCase("xls")) {
				try {
					brotherABSName = super.xls2xlsx(brotherABSName, "D:\\XML_OUTPUT\\自動產生檔-" + filename2 + ".xlsx");
				} catch (Exception ex) {
					System.out.println("EXT2 : " + ex.getMessage());
				}
			}
			file = new File(fileABSName);
			file_brother = new File(brotherABSName);
			if (!file_brother.exists()) {
				super.infoBox("找不到" + file.getName() + " 的兄弟" + file_brother.getName(), "兄弟不見了");
				return;
			}

			this.filesToExtract.add(file);
			this.filesToExtract.add(file_brother);

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
		// TODO Auto-generated method stub
		int inv_size = alItem.size();
		int nw_size = alNW.size();

		if (inv_size == 0) {
			super.infoBox("INV Excel 讀取不到資料", "資料讀取錯誤");
			throw new Exception("INV Excel 讀取不到資料");
		}
		if (nw_size == 0) {
			super.infoBox("PK Excel 讀取不到資料", "資料讀取錯誤");
			throw new Exception("PK Excel 讀取不到資料");
		}
		if (inv_size != nw_size) {
			super.infoBox("資料量有差 INV: " + inv_size + ", PK: " + nw_size, "資料讀取錯誤");
			throw new Exception("資料量有差 INV: " + inv_size + ", PK: " + nw_size);
		}

		// merge
		int itemCount = 0;
		for (Item item : alItem) {
			item.NW = alNW.get(itemCount++);

		}
	}

	private void setValue(int row_pos, int col_pos, Object value) throws Exception {
		super.setValue(ws, row_pos, col_pos, value);
	}

}
