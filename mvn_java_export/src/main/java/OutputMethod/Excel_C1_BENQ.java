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

public class Excel_C1_BENQ extends OutputCommon {

	public final static String programmeTitle = "C1_BENQ_轉出程式 20171226版";
	
	XSSFWorkbook wb;
	XSSFSheet ws;
	
	public Excel_C1_BENQ() {
		
	}

	
	private ResultSet getLines(String custom_no) throws Exception {
		Connection conn = connSQL();

		PreparedStatement ps = conn.prepareStatement(
				" SELECT A.DOC_OTR_DESC, CAST(B.ITEM_NO AS INT) as ITEM_NO, A.DCL_DOC_NO, A.DCL_DATE, B.SELLER_ITEM_CODE, "
						+ " B.DESCRIPTION, B.INV_UM, B.QTY, B.ST_MTD, A.EXCHG_RATE,  "
						+ " B.ORG_IMP_DCL_NO, B.ORG_IMP_DCL_NO_ITEM "
						+ " FROM DOC_HEAD A "
						+ " LEFT OUTER JOIN DOCINVBD B ON B.AUTO_SEQ_HEAD = A.AUTO_SEQ "
						+ " WHERE A.DCL_DOC_NO = ? and B.ITEM_NO != '*' "
						+ " ");
		ps.setString(1, custom_no);

		return ps.executeQuery();
	}
	
	
	
	/**
	 * @param custom_no
	 * @throws Exception
	 */
	public void getExcel(String custom_no) throws Exception {
		// get data from GIC
		ResultSet rsLines = getLines(custom_no);
		
		// get xlsx template
		String templatePath = "/Excel_C1_BENQ.xlsx";
		InputStream tmpFile= this.getClass().getResourceAsStream(templatePath);
		
		wb = new XSSFWorkbook(tmpFile);
		ws = wb.getSheetAt(0);
		
		// write xlsx
		doLines(rsLines);
		
		String custom_no_fileName = custom_no.replaceAll("/", "").replaceAll(" ", "").replaceAll("-", "");
		
		outputFilePath = "D:\\XML_OUTPUT\\";
        outputFileName = "C1_BENQ_"+ custom_no_fileName + "_" + System.currentTimeMillis()+".xlsx";
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
				"DOC_OTR_DESC", "ITEM_NO", "DCL_DOC_NO", "DCL_DATE", "SELLER_ITEM_CODE", 
				"DESCRIPTION", "INV_UM", "QTY", "ST_MTD", "EXCHG_RATE", 
				"ORG_IMP_DCL_NO", "ORG_IMP_DCL_NO_ITEM", "BOM"
		};
		int _line_row = 1;
		while(rsLines.next()) {
			int col_pos = 0;
			for(String col_name : colNames_Lines) {
				String chr10 = "\n";

				System.out.println("SET " + col_name);
				if(col_name.equalsIgnoreCase("SKIP_ONE")) {
					col_pos++;
					continue;
				}
				else if(col_name.equalsIgnoreCase("DOC_OTR_DESC")) {
					String[] aDOC_OTR_DESC = rsLines.getString(col_name).trim().split(chr10);
					
					// G0000
					String sDOC_OTR_DESC = aDOC_OTR_DESC[0];
					sDOC_OTR_DESC = sDOC_OTR_DESC.replace("明基材料出口字第", "").replace("號", "").trim();
					
					this.setValue(_line_row, col_pos++, sDOC_OTR_DESC);
					
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
				else if(col_name.equalsIgnoreCase("DESCRIPTION")) {
					String _DESCRIPTION = rsLines.getString("DESCRIPTION");
					if(_DESCRIPTION.contains("BOM NO.")) {
						_DESCRIPTION = _DESCRIPTION.split("BOM NO.")[0]; // 不顯示BOM NO. 相關資訊
					}
					
					this.setValue(_line_row, col_pos++, _DESCRIPTION);
					
					continue;
				}
				else if(col_name.equalsIgnoreCase("BOM")) {
					String _DESCRIPTION = rsLines.getString("DESCRIPTION");
					if(_DESCRIPTION.contains("BOM NO.")) {
						String _BOM = _DESCRIPTION.split("BOM NO.")[1];

						this.setValue(_line_row, col_pos++, _BOM);
					}
					else 
						this.setValue(_line_row, col_pos++, "");
					
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
