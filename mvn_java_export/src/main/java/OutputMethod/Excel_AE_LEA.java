package OutputMethod;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.Locale;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileFilter;

import org.apache.commons.io.comparator.NameFileComparator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_AE_LEA extends OutputCommon {

	private class Item {
		String SONO,SEQ,Item_Code,DESC_OF_GOODS;
		double QTY;
		String UNIT;
		double PRICE,AMT;
		String TRADEMARK,ST_MTD,CCCode;
		double NW;

		public Item(String SONO, String SEQ, String Item_Code, String DESC_OF_GOODS, double QTY, String UNIT,
				double PRICE, double AMT, String TRADEMARK, String ST_MTD, String CCCode, double NW) {
			this.SONO = SONO;
			this.SEQ = SEQ;
			this.Item_Code = Item_Code;
			this.DESC_OF_GOODS = DESC_OF_GOODS;
			this.QTY = QTY;
			this.UNIT = UNIT;
			this.PRICE = PRICE;
			this.AMT = AMT;
			this.TRADEMARK = TRADEMARK;
			this.ST_MTD = ST_MTD;
			this.CCCode = CCCode;
			this.NW = NW;
		}

	}


	public final static String programmeTitle = "AE_利益得文件整理";

	XSSFWorkbook wb;
	XSSFSheet ws;

	private ArrayList<File> filesToExtract = new ArrayList<File>();
	private ArrayList<Item> alItem = new ArrayList<Item>();

	public Excel_AE_LEA() {
		// TODO Auto-generated constructor stub
	}

	private void doExcel() throws Exception {
		String templatePath = "/Excel_AE_LEA.xlsx";
		InputStream tmpFile = this.getClass().getResourceAsStream(templatePath);

		wb = new XSSFWorkbook(tmpFile);
		ws = wb.getSheetAt(0);

		int item_seq = 0;
		int row_pos = 1;
		for (Item obj : this.alItem) {
			int col_pos = 0;
			this.setValue(row_pos, col_pos++, obj.SONO);

			
			if(obj.SONO.equals("")) {
				item_seq = 0;
			}
			else 
				item_seq++;
			if(item_seq == 0)
				col_pos++;
			else
				this.setValue(row_pos, col_pos++, item_seq);
			
			this.setValue(row_pos, col_pos++, obj.Item_Code);
			this.setValue(row_pos, col_pos++, obj.DESC_OF_GOODS);
			if(obj.QTY == 0)
				col_pos++;
			else
				this.setValue(row_pos, col_pos++, obj.QTY);
			this.setValue(row_pos, col_pos++, obj.UNIT);
			if(obj.PRICE == 0)
				col_pos++;
			else
				this.setValue(row_pos, col_pos++, obj.PRICE);
			if(obj.AMT == 0)
				col_pos++;
			else
				this.setValue(row_pos, col_pos++, obj.AMT);
			this.setValue(row_pos, col_pos++, obj.TRADEMARK);
			this.setValue(row_pos, col_pos++, obj.ST_MTD);
			this.setValue(row_pos, col_pos++, obj.CCCode);
			if(obj.NW == 0)
				col_pos++;
			else
				this.setValue(row_pos, col_pos++, obj.NW);

			row_pos++;
		}

		outputFilePath = "D:\\XML_OUTPUT\\";
		
		Date now = new Date();
		String format3 = new SimpleDateFormat("yyyyMMddHHmmssSSS", Locale.ENGLISH).format(now);
		format3 = format3.substring(0, 15);
		outputFileName = "LEA_" + format3 + ".xlsx";
		
		
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

		getItems(xwb.getSheetAt(0));
		
		xwb.close();
	}

	
	private void getItems(XSSFSheet xws) {
		
		Iterator<Row> rowIterator = xws.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();

			// 從第三行開始
			if (row.getRowNum() < 2) {
				continue;
			}

			String flag = getCellString(row, 9);
			if (flag.equals(""))
				break;
			flag = getCellString(row,10);
			if(flag.contains("@"))
				continue;
			

			String SONO = getCellString(row, 0);
			if(SONO.isEmpty()) {
				SONO = "";
			}
			else
			{
				try {
					int iSono = new BigDecimal(SONO).intValue();
					SONO = ""+iSono;
				}
				catch (Exception ex) {
					SONO = "";
					System.err.println("Excel_AE_LEA parse BigDecimal Error");
				}
			}
			String SEQ = "";
			String Item_Code = getCellString(row, 8);
			String DESC_OF_GOODS = getCellString(row, 9);
			double QTY = getCellDouble(row, 10);
			String UNIT = getCellString(row, 11);
			double PRICE = getCellDouble(row, 15);
			double AMT = getCellDouble(row, 16);
			String TRADEMARK = "";
			String ST_MTD = "";
			String CCCode = getCellString(row, 17);
			double NW = getCellDouble(row, 12);

			alItem.add(new Item(SONO, SEQ, Item_Code, DESC_OF_GOODS, QTY, UNIT, PRICE, AMT, TRADEMARK, ST_MTD, CCCode, NW));

			continue;
		}
	}

	private class LEAExcelFilter extends FileFilter {
		@Override
		public boolean accept(File pathname) {
			String filename = pathname.getName();
			if (pathname.isDirectory()) {
				return true;

			} else if (filename.endsWith("xlsx") || filename.endsWith("xls") ||filename.toUpperCase().startsWith("INVOICE")) {
				return true;
			} else {
				return false;
			}
		}

		@Override
		public String getDescription() {
			return "利益得 Xlsx Files";
		}
	}

	public void run() {
		this.filesToExtract.clear();
		this.alItem.clear();

		// get 原始WORD 檔
		JFileChooser fc = new JFileChooser();
		fc.setDialogTitle(programmeTitle);
		fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		fc.setMultiSelectionEnabled(true);

		FileFilter bsiFilter = new LEAExcelFilter();
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
				

				doExcel();
				this.alItem.clear();
			}


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
