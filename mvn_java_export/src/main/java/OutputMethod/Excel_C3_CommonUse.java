package OutputMethod;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_C3_CommonUse extends OutputCommon {

	public final static String programmeTitle = "C3_共用模板轉出程式";
	
	XSSFWorkbook wb;
	XSSFSheet ws;
	
	public Excel_C3_CommonUse() {
		
	}

	
	private ResultSet getLines(String custom_no, String booking_no) throws Exception {
		Connection conn = connSQL();

		String sql = " SELECT A.DOC_OTR_DESC, B.ITEM_NO as ITEM_NO, A.DCL_DOC_NO, A.DCL_DATE, B.SELLER_ITEM_CODE, B.BUYER_ITEM_CODE, "
				+ " B.DESCRIPTION, B.DOC_UNIT_P, B.INV_UM, B.QTY, B.ST_MTD, B.DOC_TOT_P, A.EXCHG_RATE, B.BOND_NOTE, B.GOODS_MODEL, B.GOODS_SPEC,  "
				+ " B.ORG_IMP_DCL_NO, B.ORG_IMP_DCL_NO_ITEM, B.TRADE_MARK, B.CCC_CODE, B.NET_WT, B.ORG_DCL_NO, B.ORG_DCL_NO_ITEM, B.EXP_NO, B.EXP_SEQ_NO, B.CERT_NO, B.CERT_NO_ITEM "
				+ " FROM DOC_HEAD A "
				+ " LEFT OUTER JOIN DOCINVBD B ON B.AUTO_SEQ_HEAD = A.AUTO_SEQ "
				+ " WHERE A.DCL_DOC_NO = ? and A.DOC_HEAD_DOC_NO = ?  " //and B.ITEM_NO != '*'
				+ " ";
		
		PreparedStatement ps = conn.prepareStatement( sql,ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_READ_ONLY );
		ps.setString(1, custom_no);
		ps.setString(2, booking_no);

		return ps.executeQuery();
	}
	
	
	
	/**
	 * @param custom_no
	 * @throws Exception
	 */
	public void getExcel(String custom_no, String booking_no) throws Exception {
		// get data from GIC
		ResultSet rsLines = getLines(custom_no, booking_no);
		
		// get xlsx template
		String templatePath = "/Excel_C3_CommonUse.xlsx";
		InputStream tmpFile= this.getClass().getResourceAsStream(templatePath);

		// write xlsx
		wb = new XSSFWorkbook(tmpFile);
		ws = wb.getSheetAt(0);
		doLines(rsLines);
		
		
		// 寫入其他申報事項
		rsLines.beforeFirst();
		ws = wb.getSheetAt(1);
		while(rsLines.next()) {
			this.setValue(1, 0, rsLines.getObject("DOC_OTR_DESC"));
			break;
		}
		
		String custom_no_fileName = custom_no.replaceAll("/", "").replaceAll(" ", "").replaceAll("-", "");
		
		outputFilePath = "D:\\XML_OUTPUT\\";
        outputFileName = "C3_共用模板_"+ custom_no_fileName+".xlsx";
        Files.createDirectories(new File(outputFilePath).toPath());
		FileOutputStream stream = new FileOutputStream(outputFilePath + outputFileName);
		wb.write(stream);
		stream.close();
		
		wb.close();
		
		System.out.println("JOB_DONE");	
		infoBox(outputFileName + " 產生完畢", "JOB_DONE");
	}
	
	

	/**
	 * �N����ResultSet �g�JExcel 
	 * @param rsLines
	 * @throws SQLException
	 * @throws Exception
	 */
	private void doLines(ResultSet rsLines) throws SQLException, Exception {
		String[] colNames_Lines = new String[] { 
				"DCL_DATE", "DCL_DOC_NO",  "ITEM_NO", "SELLER_ITEM_CODE", "BUYER_ITEM_CODE", 
				"DESCRIPTION", "QTY",  "INV_UM", "DOC_UNIT_P", "DOC_TOT_P", "ST_MTD", "TRADE_MARK", "CCC_CODE", "NET_WT", "BOND_NOTE", 
				"ORG_IMP_DCL_NO", "ORG_IMP_DCL_NO_ITEM", "GOODS_MODEL", "GOODS_SPEC", "EXP_NO", "EXP_SEQ_NO", "CERT_NO", "CERT_NO_ITEM", "EXCHG_RATE"
		};
		
		int _line_row = 1;
		while(rsLines.next()) {
			int col_pos = 0;
			
			if(rsLines.getString("ITEM_NO").equals("*")) {
				this.setValue(_line_row, 2, rsLines.getString("ITEM_NO"));
				this.setValue(_line_row, 5, rsLines.getString("DESCRIPTION"));
				_line_row++;
				continue;
			}
			
			for(String col_name : colNames_Lines) {
				String chr10 = "\n";

				if(col_name.equalsIgnoreCase("SKIP_ONE")) {
					col_pos++;
					continue;
				}
				else if(col_name.equalsIgnoreCase("ITEM_NO")) {
					this.setValue(_line_row, col_pos++, Integer.parseInt(rsLines.getString(col_name)));
					continue;
				}
				else if(col_name.equalsIgnoreCase("DCL_DOC_NO")) {
					String[] aDCL_DOC_NO = rsLines.getString(col_name).trim().split("/");
					
					// 古時候(關港帽之前是) 2/2/2/4/4 
					// 目前報單號碼格式是     2/2/2/3/5 格式，所以針對第四個PART 進行 TRIM 動作
					String sDCL_DOC_NO = aDCL_DOC_NO[0]+aDCL_DOC_NO[1]+aDCL_DOC_NO[2]+aDCL_DOC_NO[3].trim()+aDCL_DOC_NO[4];
					
					this.setValue(_line_row, col_pos++, sDCL_DOC_NO);
					
					continue;
				}
				else if(col_name.equalsIgnoreCase("DCL_DATE")) {
					Date dDCL_DATE = rsLines.getDate("DCL_DATE");
					Calendar cal = Calendar.getInstance();
					cal.setTime(dDCL_DATE);
					int year = cal.get(cal.YEAR);
					cal.set(year, cal.get(cal.MONTH), cal.get(cal.DATE));
					
					SimpleDateFormat sdfSource = new SimpleDateFormat("yyyy/MM/dd");
					String sDCL_DATE = sdfSource.format(cal.getTime());
				
					this.setValue(_line_row, col_pos++, sDCL_DATE);
					continue;
				}
				else if(col_name.equalsIgnoreCase("ORG_IMP_DCL_NO")) { // "ORG_IMP_DCL_NO", "ORG_IMP_DCL_NO"
					String sORG_IMP_DCL_NO = rsLines.getString("ORG_IMP_DCL_NO");
					if(sORG_IMP_DCL_NO == null || sORG_IMP_DCL_NO.trim().equals("")) sORG_IMP_DCL_NO = rsLines.getString("ORG_DCL_NO");
					if(sORG_IMP_DCL_NO == null) sORG_IMP_DCL_NO = "";
					
					this.setValue(_line_row, col_pos++, sORG_IMP_DCL_NO);
					
					continue;
				}
				else if(col_name.equalsIgnoreCase("ORG_IMP_DCL_NO_ITEM")) {
					String sORG_IMP_DCL_NO_ITEM = rsLines.getString("ORG_IMP_DCL_NO_ITEM");
					if(sORG_IMP_DCL_NO_ITEM == null) sORG_IMP_DCL_NO_ITEM = rsLines.getString("ORG_DCL_NO_ITEM");
					if(sORG_IMP_DCL_NO_ITEM == null) {
						col_pos++;
						continue;
					}
					
					this.setValue(_line_row, col_pos++, Integer.parseInt(sORG_IMP_DCL_NO_ITEM));
					
					continue;
				}
				
				this.setValue(_line_row, col_pos++, rsLines.getObject(col_name));
			}
			
			_line_row++;
		}
		
		
	}
	
	
	
	@SuppressWarnings("unused")
	private void setValue(String colName, Object value) throws Exception {
		CellReference cr = new CellReference(colName);
		int row_pos  = cr.getRow();
		int col_pos = cr.getCol();
		
		setValue(row_pos, col_pos, value);
	}
	
	private void setValue(int row_pos, int col_pos, Object value) throws Exception {
		super.setValue(ws, row_pos, col_pos, value);
	}
	
}
