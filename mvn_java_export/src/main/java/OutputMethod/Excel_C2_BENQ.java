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
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_C2_BENQ extends OutputCommon {

	public final static String programmeTitle = "C2_BENQ_轉出程式 20171226版";
	
	XSSFWorkbook wb;
	XSSFSheet ws;
	
	public Excel_C2_BENQ() {
		
	}

	private ResultSet getHead(String custom_no) throws Exception {
		Connection conn = connSQL();

		String sql = "SELECT A.DOC_HEAD_DOC_NO, A.LOT_NO, A.DCL_DOC_NO, A.DCL_DOC_TYPE, A.DCL_DATE, CONVERT(varchar, A.DOC_IMP_DATE, 111) as DOC_IMP_DATE, "
				+ " A.RL_DATE, A.MAWB, case a.AIR_SEA when '4' then a.hawb else '' end AS HAWB, A.TOT_CTN, A.DOC_IMP_CIF_TWD, " // DCL_AMT
				+ " A.DCL_AMT, A.WAREHOUSE, A.TRANS_VIA, "
				+ " A.CURRENCY, A.EXCHG_RATE, A.P_DOC_ITEM_FRN, ISNULL(A.CARRIER_NAME, ' ') as CARRIER_NAME, A.FLY_NO, A.TERMS_SALES, "
				+ " A.FROM_CODE, A.CNEE_COUNTRY_CODE, CONVERT(varchar, A.DOC_EXP_DATE, 111) as DOC_EXP_DATE, A.DCL_GW, A.CNEE_E_NAME "  
				+ " FROM DOC_H_I A "
				+ " WHERE A.DCL_DOC_NO = ? ";
		
		PreparedStatement ps = conn.prepareStatement(sql,ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_READ_ONLY);
		ps.setString(1, custom_no);

		return ps.executeQuery();
	}
	
	private ResultSet getHead(ArrayList<String> custom_noA) throws Exception {
		Connection conn = connSQL();

		String in_sql = "";
		for(String custom_no : custom_noA) {
			if(!in_sql.equals("")) in_sql += ",";
			
			in_sql += "?";
		}
		
		String sql = "SELECT A.DOC_HEAD_DOC_NO, A.LOT_NO, A.DCL_DOC_NO, A.DCL_DOC_TYPE, A.DCL_DATE, CONVERT(varchar, A.DOC_IMP_DATE, 111) as DOC_IMP_DATE, "
				+ " A.RL_DATE, A.MAWB, case a.AIR_SEA when '4' then a.hawb else '' end AS HAWB, A.TOT_CTN, A.DOC_IMP_CIF_TWD, " // DCL_AMT
				+ " A.DCL_AMT, A.WAREHOUSE, A.TRANS_VIA, "
				+ " A.CURRENCY, A.EXCHG_RATE, A.P_DOC_ITEM_FRN, A.CARRIER_NAME, A.FLY_NO, A.TERMS_SALES, "
				+ " A.FROM_CODE, A.CNEE_COUNTRY_CODE, CONVERT(varchar, A.DOC_EXP_DATE, 111) as DOC_EXP_DATE, A.DCL_GW, A.CNEE_E_NAME "  
				+ " FROM DOC_H_I A "
				+ " WHERE A.DCL_DOC_NO in (" + in_sql + ") "
				+ " ORDER BY A.LOT_NO, A.DCL_DOC_NO ";
		
		PreparedStatement ps = conn.prepareStatement(sql,ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_READ_ONLY);
		
		
		int index = 1;
		for(String custom_no : custom_noA) {
			ps.setString(index, custom_no);
			index++;
		}

		return ps.executeQuery();
	}
	
	
	private ResultSet getLines(String custom_no) throws Exception {
		Connection conn = connSQL();

		String sql = "SELECT A.DOC_HEAD_DOC_NO, A.LOT_NO, A.DCL_DOC_NO, A.DCL_DOC_TYPE, A.DCL_DATE, CONVERT(varchar, A.DOC_IMP_DATE, 111) as DOC_IMP_DATE, "
				+ " A.RL_DATE, A.MAWB, case a.AIR_SEA when '4' then a.hawb else '' end AS HAWB, A.TOT_CTN, A.DOC_IMP_CIF_TWD, " 
				+ " A.CURRENCY, A.EXCHG_RATE, A.P_DOC_ITEM_FRN, ISNULL(A.CARRIER_NAME, ' ') as CARRIER_NAME, A.FLY_NO, A.TERMS_SALES, "
				+ " A.FROM_CODE, A.CNEE_COUNTRY_CODE, CONVERT(varchar, A.DOC_EXP_DATE, 111) as DOC_EXP_DATE, A.DCL_GW, A.CNEE_E_NAME, "  
				+ " CAST(B.ITEM_NO AS INT) as ITEM_NO, B.BUYER_ITEM_CODE,B.DESCRIPTION, "
				+ " B.QTY, B.DOC_UM, B.AFTER_TAX_AMT, B.CCC_CODE, B.NET_WT, B.TAX_METHOD "
				+ " FROM DOC_H_I A "
				+ " LEFT OUTER JOIN DI_INVBD B ON A.AUTO_SEQ= B.AUTO_SEQ_HEAD "
				+ " WHERE A.DCL_DOC_NO = ? and B.ITEM_NO != '*' "
				+ " ORDER BY A.LOT_NO, A.DCL_DOC_NO, CAST(B.ITEM_NO AS INT) ";
		
		PreparedStatement ps = conn.prepareStatement(sql,ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_READ_ONLY);
		ps.setString(1, custom_no);

		return ps.executeQuery();
	}
	
	private ResultSet getLines(ArrayList<String> custom_noA) throws Exception {
		Connection conn = connSQL();

		String in_sql = "";
		for(String custom_no : custom_noA) {
			if(!in_sql.equals("")) in_sql += ",";
			
			in_sql += "?";
		}
		String sql = "SELECT A.DOC_HEAD_DOC_NO, A.LOT_NO, A.DCL_DOC_NO, A.DCL_DOC_TYPE, A.DCL_DATE, CONVERT(varchar, A.DOC_IMP_DATE, 111) as DOC_IMP_DATE, "
				+ " A.RL_DATE, A.MAWB, case a.AIR_SEA when '4' then a.hawb else '' end AS HAWB, A.TOT_CTN, A.DOC_IMP_CIF_TWD, " 
				+ " A.CURRENCY, A.EXCHG_RATE, A.P_DOC_ITEM_FRN, ISNULL(A.CARRIER_NAME, ' ') as CARRIER_NAME, A.FLY_NO, A.TERMS_SALES, "
				+ " A.FROM_CODE, A.CNEE_COUNTRY_CODE, CONVERT(varchar, A.DOC_EXP_DATE, 111) as DOC_EXP_DATE, A.DCL_GW, A.CNEE_E_NAME, "  
				+ " CAST(B.ITEM_NO AS INT) as ITEM_NO, B.BUYER_ITEM_CODE,B.DESCRIPTION, "
				+ " B.QTY, B.DOC_UM, B.AFTER_TAX_AMT, B.CCC_CODE, B.NET_WT, B.TAX_METHOD "
				+ " FROM DOC_H_I A "
				+ " LEFT OUTER JOIN DI_INVBD B ON A.AUTO_SEQ= B.AUTO_SEQ_HEAD "
				+ " WHERE A.DCL_DOC_NO in (" + in_sql + ") and B.ITEM_NO != '*' "
				+ " ORDER BY A.LOT_NO, A.DCL_DOC_NO, CAST(B.ITEM_NO AS INT) ";
		
		PreparedStatement ps = conn.prepareStatement(sql,ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_READ_ONLY);
		
		int index = 1;
		for(String custom_no : custom_noA) {
			ps.setString(index, custom_no);
			index++;
		}

		return ps.executeQuery();
	}
	
	
	/**
	 * @param custom_no
	 * @throws Exception
	 */
	public void getExcel(Object custom_no) throws Exception {
		
		// get data from GIC
		ResultSet rsHead = null;
		ResultSet rsLines = null;
		
		if(custom_no.getClass().getName() == "java.util.ArrayList") {
			System.out.println("Select is ArrayList");

			rsHead = getHead((ArrayList<String>)custom_no);
			rsLines = getLines((ArrayList<String>)custom_no);
			
//			String firstCustomNo = ((ArrayList<String>)custom_no).get(0);
//			custom_no_fileName = firstCustomNo.replaceAll("/", "").replaceAll(" ", "").replaceAll("-", "");
		}
		else if(custom_no.getClass().getName() == "java.lang.String") {
			System.out.println("Select is String");
			rsHead = getHead((String)custom_no);
			rsLines = getLines((String)custom_no);
//			custom_no_fileName = ((String)custom_no).replaceAll("/", "").replaceAll(" ", "").replaceAll("-", "");
		}
		
		// get xlsx template
		String templatePath = "/Excel_C2_BENQ.xlsx";
		InputStream tmpFile= this.getClass().getResourceAsStream(templatePath);
		
		wb = new XSSFWorkbook(tmpFile);
		
		
		// write xlsx
		ws = wb.getSheet("HEAD"); 
		doHead(rsHead);
		
		ws = wb.getSheet("BODY"); 
		doBody(rsLines);
		
		String custom_no_fileName = "";
		rsHead.first();
		custom_no_fileName = rsHead.getString("DOC_HEAD_DOC_NO");
		outputFilePath = "D:\\XML_OUTPUT\\";
        outputFileName = "C2_BENQ_"+ custom_no_fileName + ".xlsx"; // "_" + System.currentTimeMillis()+
        Files.createDirectories(new File(outputFilePath).toPath());
		FileOutputStream stream = new FileOutputStream(outputFilePath + outputFileName);
		wb.write(stream);
		stream.close();
		
		wb.close();
		
		System.out.println("JOB_DONE");	
		infoBox(outputFileName + " 產生完畢", "JOB_DONE");
	}
	
	private void doBody(ResultSet rsLines) throws SQLException, Exception {
		String[] colNames_Lines = new String[] { 
				"LOT_NO", "DCL_DOC_NO", "ITEM_NO", "BUYER_ITEM_CODE","DESCRIPTION", 
				"QTY", "DOC_UM", "AFTER_TAX_AMT", "CCC_CODE", "NET_WT", "TAX_METHOD"
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
				else if(col_name.equalsIgnoreCase("DCL_DOC_NO")) {
					String[] aDCL_DOC_NO = rsLines.getString(col_name).trim().split("/");
					
					// 古時候(關港帽之前是) 2/2/2/4/4 
					// 目前報單號碼格式是     2/2/2/3/5 格式，所以針對第四個PART 進行 TRIM 動作
					String sDCL_DOC_NO = aDCL_DOC_NO[0]+aDCL_DOC_NO[1]+aDCL_DOC_NO[2]+aDCL_DOC_NO[3].trim()+aDCL_DOC_NO[4];
					
					this.setValue(_line_row, col_pos++, sDCL_DOC_NO);
					
					continue;
				}
				else if(col_name.equalsIgnoreCase("DCL_DATE")) {
					Date dDCL_DATE = rsLines.getDate(col_name);
					Calendar cal = Calendar.getInstance();
					cal.setTime(dDCL_DATE);
					
					SimpleDateFormat sdfSource = new SimpleDateFormat("yyyy/MM/dd");
					String sDCL_DATE = sdfSource.format(cal.getTime());
				
					this.setValue(_line_row, col_pos++, sDCL_DATE);
					continue;
				}
				// CCC_CODE
				else if(col_name.equalsIgnoreCase("DESCRIPTION")) {
					String[] aDESCRIPTION = rsLines.getString(col_name).trim().split(chr10);
					
					String sItemName = ""; 
					for(String sDESC : aDESCRIPTION) {
						sDESC = sDESC.trim();
						
						if(sDESC.isEmpty())
							break;
						else if(!sItemName.isEmpty())
							sItemName += chr10 + sDESC;
						else 
							sItemName += sDESC;
					}
					
					this.setValue(_line_row, col_pos++, sItemName);
					continue;
				}
				else if(col_name.equalsIgnoreCase("CCC_CODE")) {
					String sCCC_CODE = rsLines.getString(col_name);
					
					sCCC_CODE = sCCC_CODE.replace("-", "").replace(".", "");
					
					this.setValue(_line_row, col_pos++, sCCC_CODE);
					continue;
				}
				 
				this.setValue(_line_row, col_pos++, rsLines.getObject(col_name));
			}
			
			_line_row++;
		}
	}
	
	
	/**
	 * �N����ResultSet �g�JExcel 
	 * @param rsLines
	 * @throws SQLException
	 * @throws Exception
	 */
	private void doHead(ResultSet rsLines) throws SQLException, Exception {
		String[] colNames_Lines = new String[] { 
				"LOT_NO", "DCL_DOC_NO", "DCL_DOC_TYPE", "DCL_DATE", "DOC_IMP_DATE", 
				"RL_DATE", "MAWB", "HAWB", "TOT_CTN", "DOC_IMP_CIF_TWD",
				"CURRENCY", "EXCHG_RATE", "P_DOC_ITEM_FRN", "CARRIER_NAME_FLY_NO", "TERMS_SALES",
				"FROM_CODE", "CNEE_COUNTRY_CODE", "DOC_EXP_DATE", "DCL_GW", "CNEE_E_NAME",
				"SKIP_ONE", "SKIP_ONE", "SKIP_ONE", "SKIP_ONE", "SKIP_ONE", 
				"SKIP_ONE", "DCL_AMT", "WAREHOUSE", "SKIP_ONE", "TRANS_VIA"
		};
		int _line_row = 1;

		super.RS2AL(rsLines);
		while(rsLines.next()) {
			int col_pos = 0;
			for(String col_name : colNames_Lines) {
				String chr10 = "\n";

				System.out.println("SET " + col_name);
				if(col_name.equalsIgnoreCase("SKIP_ONE")) {
					col_pos++;
					continue;
				}
				else if(col_name.equalsIgnoreCase("CARRIER_NAME_FLY_NO")) {
					String sCARRIER_NAME = rsLines.getString("CARRIER_NAME");
					if(rsLines.wasNull())
						sCARRIER_NAME = "";
					String sFLY_NO = rsLines.getString("FLY_NO");
					if(rsLines.wasNull())
						sFLY_NO = "";
					
					String value = sCARRIER_NAME.trim();
					if(value.length() > 0)
						value += " / ";
					value += sFLY_NO.trim();
					
					this.setValue(_line_row, col_pos++, value);
					
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
					Date dDCL_DATE = rsLines.getDate(col_name);
					Calendar cal = Calendar.getInstance();
					cal.setTime(dDCL_DATE);
					
					SimpleDateFormat sdfSource = new SimpleDateFormat("yyyy/MM/dd");
					String sDCL_DATE = sdfSource.format(cal.getTime());
				
					this.setValue(_line_row, col_pos++, sDCL_DATE);
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
