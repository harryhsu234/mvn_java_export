package OutputMethod;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.ArrayList;

import javax.swing.JFileChooser;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_AE_Marvell extends OutputCommon {

	private class MarvDeclr {
		String fileName, invoice_no, desc;
		
		MarvDeclr(String fileName, String invoice_no, String desc) {
			this.fileName = fileName;
			this.invoice_no = invoice_no;
			this.desc = desc;
		}
	}
	
	private class MarvItem {
		String fileName, itemCode, desc, stat, tradeMark, qtyUnit;
		double seq, qty, price, amt, nw, gw;
		
		public MarvItem(String fileName, double seq, String itemCode, String desc, double qty, String qtyUnit, double price, double amt, String stat,
				double nw, double gw, String tradeMark) {
			super();
			this.fileName = fileName;
			this.itemCode = itemCode;
			this.desc = desc;
			this.stat = stat;
			this.seq = seq;
			this.qty = qty;
			this.qtyUnit = qtyUnit;
			this.price = price;
			this.amt = amt;
			this.nw = nw;
			this.gw = gw;
			this.tradeMark = tradeMark;
		}
	}
	public final static String programmeTitle = "AE_Marvell_轉檔_20171102";
	
	XSSFWorkbook wb;
	XSSFSheet ws;
			
	private ArrayList<File> filesToExtract = new ArrayList<File>();
	private ArrayList<MarvItem> alMarvItem = new ArrayList<MarvItem>();
	private ArrayList<MarvDeclr> alMarvDeclr = new ArrayList<MarvDeclr>();
	
	public Excel_AE_Marvell() {
		// TODO Auto-generated constructor stub
	}

	private void doExcel() throws Exception {
		String templatePath = "/Excel_AE_Marvell.xlsx";
		InputStream tmpFile= this.getClass().getResourceAsStream(templatePath);
		
		wb = new XSSFWorkbook(tmpFile);
		ws = wb.getSheet("品名");

		int row_pos = 1;
		for(MarvItem obj : this.alMarvItem) {
			int col_pos = 0;
			this.setValue(row_pos, col_pos++, obj.fileName);
			this.setValue(row_pos, col_pos++, obj.seq);
			this.setValue(row_pos, col_pos++, obj.itemCode);
			this.setValue(row_pos, col_pos++, obj.desc);
			
			
			this.setValue(row_pos, col_pos++, obj.qty);
			this.setValue(row_pos, col_pos++, obj.qtyUnit);
			this.setValue(row_pos, col_pos++, obj.price);
			this.setValue(row_pos, col_pos++, obj.amt);
			this.setValue(row_pos, col_pos++, obj.stat);
			if(obj.seq == 1) {
				this.setValue(row_pos, col_pos++, obj.nw);
				this.setValue(row_pos, col_pos++, obj.gw);
			}
			else {
				col_pos++;
				col_pos++;
			}

			this.setValue(row_pos, col_pos++, obj.tradeMark);
			
			row_pos++;
		}
		
		ws = wb.getSheet("申報事項");
		row_pos = 1;
		for(MarvDeclr obj : this.alMarvDeclr) {
			int col_pos = 0;
			this.setValue(row_pos, col_pos++, obj.fileName);
			this.setValue(row_pos, col_pos++, obj.invoice_no);
			this.setValue(row_pos, col_pos++, obj.desc);
			
			row_pos++;
		}
			
		outputFilePath = "D:\\XML_OUTPUT\\";
        outputFileName = "Excel_Marvell_" + System.currentTimeMillis()+".xlsx";
        Files.createDirectories(new File(outputFilePath).toPath());
		FileOutputStream stream = new FileOutputStream(outputFilePath + outputFileName);
		wb.write(stream);
		stream.close();
		
		
		wb.close();
		
		System.out.println("JOB_DONE");	
		infoBox(outputFileName + " 產生完畢", "JOB_DONE");
	}
	
	private void extractWordFile(String sAbsolutePath) throws IOException {
		HWPFDocument docx = null;
		try {
			docx = new HWPFDocument(new FileInputStream(sAbsolutePath));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	      
	      //using XWPFWordExtractor Class
		WordExtractor we = new WordExtractor(docx);
		
	    System.out.println(we.getText());
	    String[] fullText = we.getParagraphText();
	    
	    we.close();
	    double nw = 0, gw = 0;
	    String stat = "", invoice_no = "", exportno = "", finishCharge = "", MaterialCost = "";
	    String declareText = "";
	    for(String s : fullText) { 
	    	s=s.trim();
	    	
	    	if(s.startsWith("*.N.W. ：")) {
	    		s = s.replace("*.N.W. ：", "").replace("*.G.W. ：", "").replace(" ", "");
	    		nw = Double.parseDouble(s.split("KGS")[0]);
	    		gw = Double.parseDouble(s.split("KGS")[1]);
	    	}
	    	
	    	if(s.startsWith("Stat. Model:")) {
	    		s = s.replace("Stat. Model:", "");
	    		stat = s;
	    	}
	    	
	    	if(s.startsWith("INVOICE DATE：") && s.contains("INVOICE NO：")) {
	    		invoice_no = s.split("INVOICE NO：")[1];
	    	} 
	    	
	    	if(s.startsWith("*.SPIL EXPORT NO：")) {
	    		exportno = s;
	    	} 
	    	
	    	// *The finished charge USD:3261.72
	    	if(s.startsWith("*The finished charge")) {
	    		finishCharge = s;
	    	} 
	    	
	    	// Material cost provided by cust. (F.O.C )USD:3782.16
	    	if(s.startsWith("Material cost provided by cust.")) {
	    		MaterialCost = s;
	    	} 
	    }
	    // \012 是 Excel : 同一個cell 裡的換行符號
	    String newLine = "\012";
	    
	    declareText = "INVOICE NO："+invoice_no + newLine +  exportno + newLine + finishCharge + newLine + MaterialCost + newLine + " "+ newLine;
	    
	    this.alMarvDeclr.add(new MarvDeclr(sAbsolutePath, invoice_no, declareText));
	    
	    boolean isItem = false;
	    int itemCount = 0;
	    String sItemCode = ""; // 料號
	    String sTradeMark = ""; // 商標
	    String sQtyUnit = ""; // 數量單位
	    String sItemDesc = "";
	    for(String s : fullText) {
	    	s= s.trim();
	    	
	    	if(s.contains("S.PN:")) {
	    		sItemCode = s.split(":")[1];
	    		isItem = true;
	    		continue;
	    	}
	    	
	    	if(s.startsWith("BRAND：")) {
	    		// sTradeMark
	    		sTradeMark = s.replace("BRAND：", "").split(" ")[0];
	    		// sQtyUnit
	    		sQtyUnit = s.substring(s.indexOf("(")+1, s.indexOf(")"));
	    	}
	    	
	    	if(isItem) {
	    		if( s.startsWith("LOT NO：") || s.startsWith("---------") ) {
	    			// 不列入品名
	    		}
	    		else if( s.startsWith("Custom BOM Approval No:") ) {
	    			s = s.replace("Custom BOM Approval No:", "").replaceAll(" +", " ").replace(",", "");
	    			
	    			sItemDesc += "Custom BOM Approval No:" + s.split(" ")[0] + " " + s.split(" ")[1];
	    			
	    			itemCount++;
	    			
	    			double qty = Double.parseDouble(s.split(" ")[2]);
	    			double price = Double.parseDouble(s.split(" ")[3]);
	    			double amt = Double.parseDouble(s.split(" ")[4]);
	    			
	    			this.alMarvItem.add(new MarvItem(sAbsolutePath, itemCount, sItemCode,sItemDesc, qty, sQtyUnit, price, amt, stat, nw, gw, sTradeMark));
	    			
	    			isItem = false;
	    			sItemDesc = "";
	    		}
	    		else
	    			sItemDesc += s + newLine;
	    	}
	    }
	}

	private void getFileInDir(File[] files) {
		for(File f : files) {
			if(f.isDirectory())
				getFileInDir(f.listFiles());
			else if(f.isFile() && f.getName().startsWith("MARVELL_S_"))
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
			if(this.filesToExtract.size() == 0) throw new Exception("未發現可處理的Word 檔");
//			for(File f : this.filesToExtract) {
//				// extractWordFile(f.getAbsolutePath());
//			}
			
			// 處理 word file
			for(File f : this.filesToExtract) {
				extractWordFile(f.getAbsolutePath());
			}
		    
			// 寫入EXCEL FILE
			doExcel();
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			super.infoBox(e.getMessage(), "Error in " + this.getClass() );
		}
	}
	
	private void setValue(int row_pos, int col_pos, Object value) throws Exception {
		super.setValue(ws, row_pos, col_pos, value);
	}

}
