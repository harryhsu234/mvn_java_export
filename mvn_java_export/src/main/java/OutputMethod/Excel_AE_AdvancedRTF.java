package OutputMethod;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Hashtable;

import javax.swing.JFileChooser;
import javax.swing.text.BadLocationException;
import javax.swing.text.Document;
import javax.swing.text.rtf.RTFEditorKit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_AE_AdvancedRTF extends OutputCommon {

	private class Declr {
		String exportno, desc;

		Declr(String exportno, String desc) {
			this.exportno = exportno;
			this.desc = desc;
		}
	}

	private class InvItem {
		String fileName, export, itemCode, desc, stat, tradeMark, qtyUnit, ccc_code;
		double seq, qty, price, amt, nw, gw;

		public InvItem(String fileName, String export, double seq, String itemCode, String desc, double qty,
				String qtyUnit, double price, double amt, String stat, double nw, double gw, String tradeMark,
				String ccc_code, String curr_type) {
			super();
			this.fileName = fileName;
			this.export = export;
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
			this.ccc_code = ccc_code;
		}
	}

	public final static String programmeTitle = "AE_AdvancedRTF_轉檔_20171130";

	XSSFWorkbook wb;
	XSSFSheet ws;

	private ArrayList<File> filesToExtract = new ArrayList<File>();
	private ArrayList<InvItem> alInvItem = new ArrayList<InvItem>();
	private ArrayList<Declr> alDeclr = new ArrayList<Declr>();
	private Hashtable<String, Integer> htExportSeq = new Hashtable<String, Integer>();
	private Hashtable<String, String> htExportDNNO = new Hashtable<String, String>();
	private Hashtable<String, String> htExportCurrType = new Hashtable<String, String>();
	private Hashtable<String, Double> htExportCNSTotal = new Hashtable<String, Double>();
	private Hashtable<String, Double> htExportTotalCharge = new Hashtable<String, Double>();

	public Excel_AE_AdvancedRTF() {
		// TODO Auto-generated constructor stub
	}

	private void doExcel() throws Exception {
		String templatePath = "/Excel_AdvancedRTF.xlsx";
		InputStream tmpFile = this.getClass().getResourceAsStream(templatePath);

		wb = new XSSFWorkbook(tmpFile);
		ws = wb.getSheet("品名");

		int row_pos = 1;
		for (InvItem obj : this.alInvItem) {
			int col_pos = 0;
			this.setValue(row_pos, col_pos++, obj.fileName);
			this.setValue(row_pos, col_pos++, obj.export);
			this.setValue(row_pos, col_pos++, obj.seq);
			this.setValue(row_pos, col_pos++, obj.itemCode);
			this.setValue(row_pos, col_pos++, obj.desc);

			this.setValue(row_pos, col_pos++, obj.qty);
			this.setValue(row_pos, col_pos++, obj.qtyUnit);
			this.setValue(row_pos, col_pos++, obj.price);
			this.setValue(row_pos, col_pos++, obj.amt);
			this.setValue(row_pos, col_pos++, obj.stat);

			this.setValue(row_pos, col_pos++, obj.tradeMark);
			this.setValue(row_pos, col_pos++, obj.ccc_code);

			row_pos++;
		}

		ws = wb.getSheet("申報事項");
		row_pos = 1;
		for (Declr obj : this.alDeclr) {
			int col_pos = 0;
			this.setValue(row_pos, col_pos++, obj.exportno);
			this.setValue(row_pos, col_pos++, obj.desc);

			row_pos++;
		}

		outputFilePath = "D:\\XML_OUTPUT\\";
		outputFileName = "Excel_AdvancedRTF_" + System.currentTimeMillis() + ".xlsx";
		Files.createDirectories(new File(outputFilePath).toPath());
		FileOutputStream stream = new FileOutputStream(outputFilePath + outputFileName);
		wb.write(stream);
		stream.close();

		wb.close();

		System.out.println("JOB_DONE");
		infoBox(outputFileName + " 產生完畢", "JOB_DONE");
	}

	private void extractWordFile(String sAbsolutePath) throws IOException {
		Path fileLocation = Paths.get(sAbsolutePath);// Paths.get("D:\\RTF\\報關_PI_1008_SLI_802386939_ZF8.rtf");
		byte[] data;
		try {
			data = Files.readAllBytes(fileLocation);

			RTFEditorKit rtfParser = new RTFEditorKit();
			Document document = rtfParser.createDefaultDocument();
			rtfParser.read(new ByteArrayInputStream(data), document, 0);
			String text = document.getText(0, document.getLength());
			text = text.replaceAll("\u2011", "-"); // 特殊處裡 過濾掉 hex code = 2011 的 橫槓
			// System.out.println(text);

			int index = 0;

			String[] textA = text.split("\n");
			String sExportNo = "";
			String stat = "";
			String tradeMark = "";
			String ccc_code = "";
			String curr_type = "";
			for (String c : textA) {
				if (c.trim().equals("EXPORT  NO :")) {
					sExportNo = textA[index - 1];
					if (!this.htExportSeq.containsKey(sExportNo)) // 如果Export 未被加入
						this.htExportSeq.put(sExportNo, 0);

				}
				if (c.trim().equals("REMARK :")) {
					String[] aRemark = textA[index - 1].split(",");

					stat = aRemark[1];
					tradeMark = aRemark[2].replace("LOGO:", "");
				}
				if (c.trim().startsWith("Microperipheral")) {

					ccc_code = textA[index - 1].trim();
				}

				index++;
			}

			index = 0;
			int group = 1;
			for (String c : textA) {

				System.out.println(index + ":" + c);

				if (c.trim().startsWith("MATERIAL")) {
					curr_type = textA[index - 1].trim();
					if (!this.htExportCurrType.containsKey(sExportNo)) // 如果Export 未被加入
						this.htExportCurrType.put(sExportNo, curr_type);
				}

				if (c.trim().equals("D/N NO:")) {
					String sDNNO = textA[index + 8];
					if (!this.htExportDNNO.containsKey(sExportNo)) // 如果Export 未被加入
						this.htExportDNNO.put(sExportNo, sDNNO);
					else if (!htExportDNNO.get(sExportNo).contains(sDNNO))
						this.htExportDNNO.put(sExportNo, htExportDNNO.get(sExportNo) + "," + sDNNO);
				}

				if (c.trim().equals("CONSIGN TOTAL")) {
					double dCONSIGN_TOTAL = Double.parseDouble(textA[index + 16].trim().replace(",", ""));
					if (!this.htExportCNSTotal.containsKey(sExportNo)) // 如果Export 未被加入
						this.htExportCNSTotal.put(sExportNo, dCONSIGN_TOTAL);
					else
						this.htExportCNSTotal.put(sExportNo, htExportCNSTotal.get(sExportNo) + dCONSIGN_TOTAL);
				}
				if (c.trim().equals("TOTAL CHARGE")) {
					double dTOTAL = Double.parseDouble(textA[index + 16].trim().replace(",", ""));
					if (!this.htExportTotalCharge.containsKey(sExportNo)) // 如果Export 未被加入
						this.htExportTotalCharge.put(sExportNo, dTOTAL);
					else
						this.htExportTotalCharge.put(sExportNo, htExportTotalCharge.get(sExportNo) + dTOTAL);
				}

				if (c.trim().equals("Bonded P/N :")) {
					System.out.println("========= " + group + " =========");
					// System.out.println("廖浩：" + textA[index -1]);
					String newLine = "\012";

					String desc = textA[index] + textA[index - 1] + newLine;
					desc += textA[index - 3] + textA[index - 2] + newLine;
					desc += textA[index - 6] + "  " + textA[index - 7] + newLine;
					desc += textA[index - 4] + textA[index - 11] + newLine;
					desc += "PO : " + textA[index - 8];

					double amt = Double.parseDouble(textA[index - 10].trim().replace(",", ""));
					double price = Double.parseDouble(textA[index - 9].trim().replace(",", ""));
					double qty = Double.parseDouble(textA[index - 5].trim().replace(",", ""));

					int seq = this.htExportSeq.get(sExportNo) + 1;
					this.htExportSeq.put(sExportNo, seq);

					this.alInvItem.add(new InvItem(sAbsolutePath, sExportNo, seq, textA[index - 1], desc, qty, "EAC",
							price, amt, stat, 0, 0, tradeMark, ccc_code, curr_type));

					group++;
				}

				index++;
			}

			// 處理 申報事項

		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (BadLocationException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private void getFileInDir(File[] files) {
		for (File f : files) {
			if (f.isDirectory())
				getFileInDir(f.listFiles());
			else if (f.isFile() && f.getName().startsWith("報關_PI_"))
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
			if (this.filesToExtract.size() == 0)
				throw new Exception("未發現可處理的Word 檔");
			
			// 處理 word file
			for (File f : this.filesToExtract) {
				extractWordFile(f.getAbsolutePath());
			}

			for (String key : this.htExportDNNO.keySet()) {
				System.out.println(key);

				String newLine = "\012";

				String desc = "";
				desc += "EXPORT:[" + key + "]" + newLine;
				desc += "D/N NO: " + this.htExportDNNO.get(key) + "/" + newLine;
				desc += "CONSIGN TOTAL: " + this.htExportCurrType.get(key) + " " + this.htExportCNSTotal.get(key)
						+ newLine;
				desc += "TOTAL CHARGE: " + this.htExportCurrType.get(key) + " " + this.htExportTotalCharge.get(key)
						+ newLine;
				this.alDeclr.add(new Declr(key, desc));

				// String fileName, String invoice_no, String desc
			}

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
