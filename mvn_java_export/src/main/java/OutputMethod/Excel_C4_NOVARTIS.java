package OutputMethod;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList; 
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_C4_NOVARTIS extends OutputCommon {

	public final static String programmeTitle = "C4_諾華_轉出程式 20180301版";
	
	XSSFWorkbook wb;
	XSSFSheet ws;
	
	public Excel_C4_NOVARTIS() {
		
	}

	
	
	private ResultSet getHeads(ArrayList<String> selectedCustom) throws Exception {
		String sWhere = "";
		for(String custom_no : selectedCustom) {
			if(sWhere.length() > 0) {
				sWhere += ", ";
			}
			sWhere += "'" + custom_no + "'";
		}
		sWhere = "and A.DCL_DOC_NO in (" + sWhere + ") ";
		
		Connection conn = connSQL();

//		String sql = " SELECT A.LOT_NO, CAST(B.ITEM_NO AS INT) as ITEM_NO, A.DCL_DOC_NO, A.DCL_DATE,  "
//				+ " B.SELLER_ITEM_CODE, B.DESCRIPTION,B.DOC_UM,B.DOC_UNIT_P,B.QTY,B.TAX_METHOD, "
//				+ " A.EXCHG_RATE, B.ORG_DCL_NO, B.ORG_DCL_NO_ITEM "
//				+ " FROM DOC_H_I A "
//				+ " LEFT OUTER JOIN DI_INVBD B ON A.AUTO_SEQ = B.AUTO_SEQ_HEAD "
//				+ " WHERE B.ITEM_NO != '*' " + sWhere 
//				+ " ORDER BY A.LOT_NO, A.DCL_DOC_NO, CAST(B.ITEM_NO AS INT) ";
		
		
		String sql = "select a.DCL_DOC_NO, a.dcl_date,  a.DOC_IMP_DATE,  a.mawb, a.WAREHOUSE, a.AIR_SEA, \r\n" + 
				"		a.FROM_DESC, a.FROM_CODE, a.DCL_DOC_TYPE, a.DCL_DOC_DESC, a.FLY_NO,\r\n" + 
				"		a.CURRENCY, a.FOB_AMT, a.DOC_IMP_CIF_AMT, a.DOC_IMP_CIF_TWD,\r\n" + 
				"		a.EXCHG_RATE, a.TOT_CTN, \r\n" + 
				"		a.DOC_CTN_UM, a.DCL_GW, a.DOC_TAX_BASE as DPV_AMT,\r\n" + 
				"		a.DOC_OTR_DESC, a.DCL_PASS_METHOD, " +
				"		a.IMPORT_TAX, a.IMPORT_TAX as AC, \r\n" + 
				"		a.PORT_FEE as AD, \r\n" + 
				"		a.EX_TAX_AMT_1 as AE, \r\n" + 
				"		a.YA_MONEY as AF, \r\n" + 
				"		a.COMMODITY_TAX, a.COMMODITY_TAX as AG, \r\n" + 
				"		a.SALEST_TAX as AH, \r\n" + 
				"		a.DELAY_AMT as AI, \r\n" + 
				"		0 as AJ, \r\n" + 
				"		0 as AK, \r\n" + 
				"		0 as AL, \r\n" + 
				"		a.RL_DATE, a.DUTY_NO, \r\n" + 
				"		a.SHPR_E_NAME, a.SHPR_BAN_ID \r\n" + 
				"from doc_h_i a where  1=1 " + sWhere; 
		
		
		PreparedStatement ps = conn.prepareStatement( sql );
		// ps.setString(1, custom_no);

		return ps.executeQuery();
	}
	
	private ResultSet getBodys(ArrayList<String> selectedCustom) throws Exception {
		String sWhere = "";
		for(String custom_no : selectedCustom) {
			if(sWhere.length() > 0) {
				sWhere += ", ";
			}
			sWhere += "'" + custom_no + "'";
		}
		sWhere = "and A.DCL_DOC_NO in (" + sWhere + ") ";
		
		Connection conn = connSQL();

//		String sql = " SELECT A.LOT_NO, CAST(B.ITEM_NO AS INT) as ITEM_NO, A.DCL_DOC_NO, A.DCL_DATE,  "
//				+ " B.SELLER_ITEM_CODE, B.DESCRIPTION,B.DOC_UM,B.DOC_UNIT_P,B.QTY,B.TAX_METHOD, "
//				+ " A.EXCHG_RATE, B.ORG_DCL_NO, B.ORG_DCL_NO_ITEM "
//				+ " FROM DOC_H_I A "
//				+ " LEFT OUTER JOIN DI_INVBD B ON A.AUTO_SEQ = B.AUTO_SEQ_HEAD "
//				+ " WHERE B.ITEM_NO != '*' " + sWhere 
//				+ " ORDER BY A.LOT_NO, A.DCL_DOC_NO, CAST(B.ITEM_NO AS INT) ";
		
		
		String sql = "SELECT A.DCL_DOC_NO, B.ITEM_NO, B.DESCRIPTION, B.EXP_NO, B.EXP_SEQ_NO, B.CCC_CODE, B.GOV_ASGN_NO, \r\n" + 
				"    B.TERMS, B.DOC_UNIT_P, B.NET_WT, B.QTY, B.DOC_UM, B.AFTER_TAX_AMT, B.TAX_RATE_P, B.TAX_METHOD, B.COMM_TAX_RATE, \r\n" + 
				"    B.ORG_COUNTRY, B.ORG_COUNTRY_NAME, B.EXP_NO AS IMPLICENSE, B.EXP_NO AS DRUG, \r\n" + 
				"    ISNULL(A.CNEE_C_NAME, '') AS CNEE_C_NAME, ISNULL(A.CNEE_E_NAME, '') AS CNEE_E_NAME, \r\n" +
				"    ISNULL(A.CNEE_C_ADDR, '') AS CNEE_C_ADDR, ISNULL(A.CNEE_E_ADDR, '') AS CNEE_E_ADDR \r\n" +
				"FROM DOC_H_I A \r\n" + 
				"LEFT OUTER JOIN DI_INVBD B ON A.AUTO_SEQ = B.AUTO_SEQ_HEAD \r\n" + 
				"WHERE B.ITEM_NO != '*' " + sWhere;
		
		
		PreparedStatement ps = conn.prepareStatement( sql );
		// ps.setString(1, custom_no);

		return ps.executeQuery();
	}
	
	
	/**
	 * @param custom_no
	 * @throws Exception
	 */
	public void getExcel(ArrayList<String> selectedCustom) throws Exception {
		
		// get data from GIC
		ResultSet rsHeads = getHeads(selectedCustom);
		ResultSet rsBodys = this.getBodys(selectedCustom);
		
		// get xlsx template
		String templatePath = "/Excel_C4_NOVARTIS.xlsx";
		InputStream tmpFile= this.getClass().getResourceAsStream(templatePath);
		
		wb = new XSSFWorkbook(tmpFile);
		// write xlsx
		ws = wb.getSheet("HEAD"); //.getSheetAt(0);
		doHeads(rsHeads);
		
		ws = wb.getSheet("BODY"); //.getSheetAt(0);
		doBodys(rsBodys);
		
		outputFilePath = "D:\\XML_OUTPUT\\";
        outputFileName = "C4_諾華_" + System.currentTimeMillis()+".xlsx";
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
	private void doHeads(ResultSet rsLines) throws SQLException, Exception {
		String[] colNames_Lines = new String[] { 
				"SKIP_ONE", "MAWB", "SKIP_ONE", "WAREHOUSE", "AIR_SEA",
				"FROM_DESC", "FROM_CODE", "DCL_DOC_TYPE", "DCL_DOC_DESC", "DCL_DOC_NO",
				"SKIP_ONE", "SORT_OF_VALUE", "SKIP_ONE", "SKIP_ONE", "FLY_NO",
				"DOC_IMP_DATE", "DCL_DATE", "CURRENCY", "FOB_AMT", "DOC_IMP_CIF_AMT",
				"DOC_IMP_CIF_TWD", "EXCHG_RATE", "TOT_CTN", "DOC_CTN_UM", "DCL_GW", 
				"DPV_AMT", "DOC_OTR_DESC", "DCL_PASS_METHOD", "AC", "AD", 
				"AE", "AF", "AG", "AH", "AI", 
				"AJ", "AK", "AL", "RL_DATE", "RL_TIME",
				"DUTY_NO", "SKIP_ONE", "SKIP_ONE", "SHPR_E_NAME", "SHPR_BAN_ID"
		};
		int _line_row = 1;
		while(rsLines.next()) {
			int col_pos = 0;
			for(String col_name : colNames_Lines) {
				String chr10 = "\n";

				System.out.println("SET " + col_name);
				if(col_name.equalsIgnoreCase("SKIP_ONE")) {
					this.setValue(_line_row, col_pos++, "");
					continue;
				}
				else if(col_name.equalsIgnoreCase("DCL_DOC_NO")) {
					// 
					String[] aDCL_DOC_NO = rsLines.getString(col_name).trim().split("/");
					
					// 古時候(關港帽之前是) 2/2/2/4/4 
					// 目前報單號碼格式是     2/2/2/3/5 格式，所以針對第四個PART 進行 TRIM 動作
					String sDCL_DOC_NO = aDCL_DOC_NO[0]+aDCL_DOC_NO[1]+aDCL_DOC_NO[2]+aDCL_DOC_NO[3].trim()+aDCL_DOC_NO[4];
					
					this.setValue(_line_row, col_pos++, sDCL_DOC_NO);
					
					continue;
				}
				else if(col_name.equalsIgnoreCase("SORT_OF_VALUE")) {
					
					double dDOC_IMP_CIF_AMT = rsLines.getDouble("DOC_IMP_CIF_AMT");
					
					
					String sSORT_OF_VALUE = "大單";
					try {
						if(dDOC_IMP_CIF_AMT <= 5000) sSORT_OF_VALUE = "小單";
					}
					catch (Exception ex ){
						
						sSORT_OF_VALUE = "大/小單判斷異常";
					}
					
					this.setValue(_line_row, col_pos++, sSORT_OF_VALUE);
					
					continue;
				}
				else if(col_name.equalsIgnoreCase("AK")) {
					 // SUM(AC : AJ)
					
					double dAK = rsLines.getDouble("AC") + rsLines.getDouble("AD") + rsLines.getDouble("AE") +
								 rsLines.getDouble("AF") + rsLines.getDouble("AG") + rsLines.getDouble("AH") +
								 rsLines.getDouble("AI") + rsLines.getDouble("AJ");
				
					this.setValue(_line_row, col_pos++, dAK);
					continue;
				}
				else if(col_name.equalsIgnoreCase("AL")) {
					
					double dIMPORT_TAX = rsLines.getDouble("AC");
					double dCOMMODITY_TAX = rsLines.getDouble("AG");
					double dDOC_IMP_CIF_TWD = rsLines.getDouble("DOC_IMP_CIF_TWD");
					
					double dAL = dIMPORT_TAX + dCOMMODITY_TAX + dDOC_IMP_CIF_TWD;
					
				
					this.setValue(_line_row, col_pos++, dAL);
					continue;
				}
				else if(col_name.equalsIgnoreCase("RL_DATE")) {
					String value = "";
					try {
						String sRL_DATE = rsLines.getString("RL_DATE");
						int date = Integer.parseInt(sRL_DATE.substring(0, 6)); 
						date += 20000000;
					
						
						value += date;
					}
					catch (Exception ex) {
						System.err.println(ex.getMessage());
						System.err.println(ex.getStackTrace());
					}
					
					this.setValue(_line_row, col_pos++, value);
					continue;
				}
				else if(col_name.equalsIgnoreCase("RL_TIME")) {
					String value = "";
					try {
						String sRL_DATE = rsLines.getString("RL_DATE");
						int time = Integer.parseInt(sRL_DATE.substring(6, 10)); 
						
						value += time;
					}
					catch (Exception ex) {
						System.err.println(ex.getMessage());
						System.err.println(ex.getStackTrace());
					}
					
					this.setValue(_line_row, col_pos++, value);
					
					continue;
				}
				
				
				
				this.setValue(_line_row, col_pos++, rsLines.getObject(col_name));
			}
			
			_line_row++;
		}
	}
	
	private void doBodys(ResultSet rsLines) throws SQLException, Exception {
		String[] colNames_Lines = new String[] { 
				"DCL_DOC_NO", "ITEM_NO", "DESCRIPTION", "EXP_NO", "EXP_SEQ_NO",
				"CCC_CODE", "GOV_ASGN_NO", "TERMS", "DOC_UNIT_P", "NET_WT",
				"QTY", "DOC_UM", "AFTER_TAX_AMT", "TAX_RATE_P", "TAX_METHOD",
				"COMM_TAX_RATE", "SKIP_ONE", "SKIP_ONE", "SKIP_ONE", "ORG_COUNTRY",
				"IMPLICENSE", "DRUG", "SELLER"
		};
		int _line_row = 1;
		while(rsLines.next()) {
			int col_pos = 0;
			for(String col_name : colNames_Lines) {
				String chr10 = "\n";

				System.out.println("SET " + col_name);
				if(col_name.equalsIgnoreCase("SKIP_ONE")) {
					this.setValue(_line_row, col_pos++, "");
					continue;
				}
				else if(col_name.equalsIgnoreCase("DCL_DOC_NO")) {
					// 
					String[] aDCL_DOC_NO = rsLines.getString(col_name).trim().split("/");
					String sDCL_DOC_NO = aDCL_DOC_NO[0]+aDCL_DOC_NO[1]+aDCL_DOC_NO[2]+aDCL_DOC_NO[3].trim()+aDCL_DOC_NO[4];
					
					this.setValue(_line_row, col_pos++, sDCL_DOC_NO);
					
					continue;
				}
				else if(col_name.equalsIgnoreCase("ORG_COUNTRY")) { // "ORG_COUNTRY", "IMPLICENSE", "DRUG", "SELLER"
					String sORG_COUNTRY = rsLines.getString("ORG_COUNTRY_NAME") + " - " +  rsLines.getString("ORG_COUNTRY");
					
					this.setValue(_line_row, col_pos++, sORG_COUNTRY);
					continue;
				}
				else if(col_name.equalsIgnoreCase("DRUG")) {
					String sIMPLICENSE = rsLines.getString("IMPLICENSE");
					
					String sDrug = ""; 
					if(sIMPLICENSE != null) {
						for(char c : sIMPLICENSE.toCharArray())
							if( Character.isLetter(c) ) sDrug += c;
							else break;
					}
					
					
					this.setValue(_line_row, col_pos++, sDrug);
					continue;
				}
				else if(col_name.equalsIgnoreCase("SELLER")) { // a.CNEE_C_NAME, a.CNEE_E_NAME, a.CNEE_C_ADDR, a.CNEE_E_ADDR 
					String sSELLER = "";
					
					String sCNEE_C_NAME = rsLines.getString("CNEE_C_NAME");
					String sCNEE_E_NAME = rsLines.getString("CNEE_E_NAME");
					String sCNEE_C_ADDR = rsLines.getString("CNEE_C_ADDR");
					String sCNEE_E_ADDR = rsLines.getString("CNEE_E_ADDR");
					
					sSELLER += (sCNEE_C_NAME.length() > 0) ? sCNEE_C_NAME + chr10 : "";
					sSELLER += (sCNEE_E_NAME.length() > 0) ? sCNEE_E_NAME + chr10 : "";
					sSELLER += (sCNEE_C_ADDR.length() > 0) ? sCNEE_C_ADDR + chr10 : "";
					sSELLER += (sCNEE_E_ADDR.length() > 0) ? sCNEE_E_ADDR + chr10 : "";
					
					this.setValue(_line_row, col_pos++, sSELLER);
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
