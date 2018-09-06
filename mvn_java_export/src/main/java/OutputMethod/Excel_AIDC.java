package OutputMethod;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellReference;

public class Excel_AIDC extends OutputCommon {
	public final static String programmeTitle = "AIDC Excel";
	String type;
	boolean isAir = false;
	String kind = " ";
	HSSFWorkbook  wb;
	HSSFSheet ws;
	
	public Excel_AIDC(String type,boolean isAir) {
		super();
		this.type=type;
		this.isAir=isAir;
	}
	
	private ResultSet getHeads(ArrayList<String> selectedCustom) throws Exception {
		Connection conn = connSQL();
		
		String sWhere = "";
		for(String custom_no : selectedCustom) {
			if(sWhere.length() > 0) {
				sWhere += ", ";
			}
			sWhere += "'" + custom_no + "'";
		}
		sWhere = "and A.DCL_DOC_NO in (" + sWhere + ") ";
		String sql ="";
		if (type.equals("IMP")) {
			sql = "select a.MAWB,a.HAWB,a.WAREHOUSE,a.TRANS_VIA,a.FROM_DESC,a.FROM_CODE,a.DCL_DOC_TYPE,a.DCL_DOC_DESC,a.DCL_DOC_NO,a.CCC_CLASS,a.CARRIER_NAME,a.FLY_NO,a.DOC_IMP_DATE,a.DCL_DATE,a.CURRENCY,a.FOB_AMT,a.DOC_IMP_CIF_AMT,a.DOC_IMP_CIF_TWD,a.EXCHG_RATE,a.TOT_CTN,a.DOC_CTN_UM,a.DCL_GW,a.DCL_AMT,a.DOC_OTR_DESC,a.DCL_PASS_METHOD,a.IMPORT_TAX,a.PORT_FEE,a.EX_TAX_AMT_1,a.YA_MONEY,a.COMMODITY_TAX,a.SALEST_TAX,a.DELAY_AMT,a.EX_TAX_AMT_2,a.EX_TAX_AMT_3,a.DCL_AMT,a.DOC_TAX_BASE,a.RL_DATE,a.DUTY_NO,a.YA_NO,a.HAWB DO_NO  " 
			    + "  from DOC_H_I a where  1=1 "; 
		} else {
			sql = "select a.MAWB,a.PROC_NO HAWB,a.WAREHOUSE,a.TRANS_VIA,a.FROM_DESC,a.FROM_CODE,a.DCL_DOC_TYPE,(select CODE_CHINESE from CMTBASE_ALL where code_kind='002' and code=a.DCL_DOC_TYPE) as DCL_DOC_DESC,a.DCL_DOC_NO,' ' CCC_CLASS,a.CARRIER_NAME,a.FLY_NO,' ' as DOC_IMP_DATE,a.DCL_DATE,a.CURRENCY,a.CAL_IP_TOT_ITEM_AMT FOB_AMT,a.FOB_AMT DOC_IMP_CIF_AMT,a.FOB_AMT_TWD DOC_IMP_CIF_TWD,a.EXCHG_RATE,a.TOT_CTN,a.DOC_CTN_UM,a.DCL_GW,a.DCL_AMT,a.DOC_OTR_DESC,a.DCL_PASS_METHOD,' ' IMPORT_TAX,a.PORT_FEE,a.EX_TAX_AMT_1,' ' YA_MONEY,' ' COMMODITY_TAX,' ' SALEST_TAX,' ' DELAY_AMT,a.EX_TAX_AMT_2,a.EX_TAX_AMT_3,a.DCL_AMT,' ' DOC_TAX_BASE,a.RL_DATE,' ' DUTY_NO,' ' YA_NO,a.SHIPPING_ORDER_NO DO_NO      " 
				+ "  from DOC_HEAD a where  1=1 "; 
		}
		sql += sWhere;
		System.out.println(sql);
		Statement ps = conn.createStatement();

		return ps.executeQuery(sql);
	}
	
	private ResultSet getBodys(ArrayList<String> selectedCustom) throws Exception {
		Connection conn = connSQL();
		
		String sWhere = "";
		for(String custom_no : selectedCustom) {
			if(sWhere.length() > 0) {
				sWhere += ", ";
			}
			sWhere += "'" + custom_no + "'";
		}
		sWhere = "and A.DCL_DOC_NO in (" + sWhere + ") ";
		String sql = "";
		if (type.equals("IMP")) {
			sql = "SELECT A.DCL_DOC_NO, B.ITEM_NO, B.DESCRIPTION, B.EXP_NO, B.EXP_SEQ_NO, B.CCC_CODE, B.GOV_ASGN_NO, \r\n"  
				+ "    B.TERMS, B.DOC_UNIT_P, B.NET_WT, B.QTY, B.DOC_UM, B.AFTER_TAX_AMT, B.TAX_RATE_P, B.TAX_METHOD, B.COMM_TAX_RATE, \r\n" 
				+ "    B.ORG_COUNTRY, B.ORG_COUNTRY_NAME, B.EXP_NO, B.EXP_NO2 "
				+ "FROM DOC_H_I A \r\n"  
				+ "LEFT OUTER JOIN DI_INVBD B ON A.AUTO_SEQ = B.AUTO_SEQ_HEAD \r\n"  
				+ "WHERE B.ITEM_NO != '*' " ;
		} else {
			 sql = "SELECT A.DCL_DOC_NO, B.ITEM_NO, B.DESCRIPTION, B.EXP_NO, B.EXP_SEQ_NO, B.CCC_CODE, B.GOV_ASGN_NO, \r\n"  
			 	 + "     A.TERMS_SALES TERMS, B.DOC_UNIT_P, B.NET_WT, B.QTY, B.DOC_UM, B.FOB_TWD AFTER_TAX_AMT , '0' TAX_RATE_P,' ' TAX_METHOD, '0' COMM_TAX_RATE, \r\n"  
			 	 + "     B.EXP_NO, B.EXP_NO2 FROM DOC_HEAD A \r\n" 
			 	 + "LEFT OUTER JOIN DOCINVBD B ON A.AUTO_SEQ = B.AUTO_SEQ_HEAD "
				 + "WHERE B.ITEM_NO != '*' " ;
		}
		sql += sWhere;
		
		System.out.println(sql);
		
		Statement ps = conn.createStatement();

		return ps.executeQuery(sql);
		
	}
	
	/**
	 * @param custom_no
	 * @throws Exception
	 */
	public void getExcel(ArrayList<String> selectedCustom) throws Exception {
		
		// get data from GIC
		ResultSet rsHeads = getHeads(selectedCustom);
		ResultSet rsBodys = getBodys(selectedCustom);
		
		// get xlsx template
		String templatePath = "/Excel_AIDC.xls";
		InputStream tmpFile= this.getClass().getResourceAsStream(templatePath);
		
		wb = new HSSFWorkbook(tmpFile);
		// write xlsx
		ws = wb.getSheet("HEAD"); //.getSheetAt(0);
		doHeads(rsHeads);
		
		ws = wb.getSheet("BODY"); //.getSheetAt(0);
		doBodys(rsBodys);
		
		outputFilePath = "D:\\XML_OUTPUT\\";
        outputFileName = "AIDC_" + kind + new SimpleDateFormat("yyyyMMddHHmmss").format( new Date())+".xls";
        Files.createDirectories(new File(outputFilePath).toPath());
		FileOutputStream stream = new FileOutputStream(outputFilePath + outputFileName);
		wb.write(stream);
		stream.close();
		
		wb.close();
		
		System.out.println("JOB_DONE");	
		infoBox(outputFileName + " 產生完畢", "JOB_DONE");
	}

	/**
	 * ResultSet  
	 * @param rsLines
	 * @throws SQLException
	 * @throws Exception
	 */
	private void doHeads(ResultSet rsLines) throws SQLException, Exception {
		String[] colNames_Lines = new String[] { 
				"FORWARDER_ID", "MAWB", "HAWB", "WAREHOUSE", "TRANS_VIA",
				"FROM_DESC", "FROM_CODE", "DCL_DOC_TYPE", "DCL_DOC_DESC", "DCL_DOC_NO",
				"CCC_CLASS", "SORT_OF_VALUE", "SKIP_ONE", "CARRIER_NAME", "FLY_NO",
				"DOC_IMP_DATE", "DCL_DATE", "CURRENCY", "FOB_AMT", "DOC_IMP_CIF_AMT",
				"DOC_IMP_CIF_TWD", "EXCHG_RATE", "TOT_CTN", "DOC_CTN_UM", "DCL_GW", 
				"DOC_IMP_CIF_TWD", "DOC_OTR_DESC", "DCL_PASS_METHOD", "IMPORT_TAX", "PORT_FEE", 
				"EX_TAX_AMT_1", "YA_MONEY", "COMMODITY_TAX", "SALEST_TAX", "DELAY_AMT", 
				"EX_TAX_AMT", "DCL_AMT", "DOC_TAX_BASE", "RL_DATE", "RL_TIME",
				"DUTY_NO", "YA_NO", "DO_NO"
		};
		int _line_row = 1;
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
		while(rsLines.next()) {
			int col_pos = 0;
			for(String col_name : colNames_Lines) {
				String chr10 = "\n";
				String value = "";
				
				System.out.print("SET " + col_name);
				switch (col_name) {
				case "FORWARDER_ID":
					//TODO 未來要給固定職
					break;
				case "SKIP_ONE":
					break;
				case "DCL_DOC_TYPE":
					if (kind.equals(" ")) {
						if (rsLines.getObject(col_name) != null) {
							value = String.valueOf( rsLines.getObject(col_name) );
							kind=value;
						}
					}
					
					break;
				case "DCL_DOC_NO":
					String[] aDCL_DOC_NO = rsLines.getString(col_name).trim().split("/");
					// 古時候(關港帽之前是) 2/2/2/4/4 
					// 目前報單號碼格式是     2/2/2/3/5 格式，所以針對第四個PART 進行 TRIM 動作
					value = aDCL_DOC_NO[0]+aDCL_DOC_NO[1]+aDCL_DOC_NO[2]+aDCL_DOC_NO[3].trim()+aDCL_DOC_NO[4];
					
					break;
				case "SORT_OF_VALUE":
					double dDOC_IMP_CIF_AMT = rsLines.getDouble("DOC_IMP_CIF_AMT");
					value = "大單";
					try {
						if (dDOC_IMP_CIF_AMT <= 5000) value = "小單";
					} catch (Exception ex) {
						value = "大/小單判斷異常";
					}
					
					break;
				case "DOC_IMP_DATE":
				case "DCL_DATE":
					if (rsLines.getObject(col_name)!=null && !rsLines.getString(col_name).equals(" ")) {
						value = sdf.format(rsLines.getDate(col_name));
					}
					
					break;
				case "RL_DATE":
					try {
						String sRL_DATE = rsLines.getString("RL_DATE");
						int date = Integer.parseInt(sRL_DATE.substring(0, 6)); 
						date += 20000000;
						value += date;
					} catch (Exception ex) {
						System.err.println(ex.getMessage());
						System.err.println(ex.getStackTrace());
					}
					
					break;
				case "RL_TIME":
					try {
						String sRL_DATE = rsLines.getString("RL_DATE");
						int time = Integer.parseInt(sRL_DATE.substring(6, 10)); 
						
						value += time;
					} catch (Exception ex) {
						System.err.println(ex.getMessage());
						System.err.println(ex.getStackTrace());
					}
					
					break;
				case "EX_TAX_AMT":
					value = String.valueOf(rsLines.getDouble("EX_TAX_AMT_2") + rsLines.getDouble("EX_TAX_AMT_3"));
					
					break;
				case "DO_NO":
					if (isAir && this.type.equals("IMP")) {
						value = "";
					}
					break;
				default :
					if (rsLines.getObject(col_name) != null) {
						value = String.valueOf( rsLines.getObject(col_name) );
					}
				}
				System.out.println("= " + value);
				this.setValue(_line_row, col_pos++, value);
			}
			
			_line_row++;
		}
	}
	
	private void doBodys(ResultSet rsLines) throws SQLException, Exception {
		String[] colNames_Lines = new String[] { 
				"DCL_DOC_NO", "ITEM_NO", "DESCRIPTION", "EXP_NO", "EXP_SEQ_NO",
				"CCC_CODE", "GOV_ASGN_NO", "TERMS", "DOC_UNIT_P", "NET_WT",
				"QTY", "DOC_UM", "AFTER_TAX_AMT", "TAX_RATE_P", "TAX_METHOD",
				"COMM_TAX_RATE", "DESCRIPTION_A", "DESCRIPTION_B", "EXP_NO2"
		};
		int _line_row = 1;
		while(rsLines.next()) {
			int col_pos = 0;
			
			String[] descArray;
			for(String col_name : colNames_Lines) {
				String chr10 = "\n";
				String value = "";
				System.out.print("SET " + col_name);
				switch (col_name) {
				case "DCL_DOC_NO":
					String[] aDCL_DOC_NO = rsLines.getString(col_name).trim().split("/");
					value = aDCL_DOC_NO[0]+aDCL_DOC_NO[1]+aDCL_DOC_NO[2]+aDCL_DOC_NO[3].trim()+aDCL_DOC_NO[4];
					
					break;
				case "DESCRIPTION": // 最後一行是合約號#項次
					descArray = rsLines.getString("DESCRIPTION").trim().split(chr10);
					for (int i=0;i<descArray.length;i++) {
						if (!descArray[i].startsWith("PO:")) {
							value += descArray[i];
						}
					}
					break;
				case "DESCRIPTION_A":
					descArray = rsLines.getString("DESCRIPTION").trim().split(chr10);
					for (int i=0;i<descArray.length;i++) {
						if (descArray[i].startsWith("PO:")) {
							if (descArray[i].contains("#")) {
								value = descArray[i].split("#")[0].replace("PO:", "");
							}
						}
					}
									
					break;
				case "DESCRIPTION_B":
					descArray = rsLines.getString("DESCRIPTION").trim().split(chr10);
					for (int i=0;i<descArray.length;i++) {
						if (descArray[i].startsWith("PO:")) {
							if (descArray[i].contains("#")) {
								value = descArray[i].split("#")[1];
							}
						}
					}
					break;
				case "EXP_SEQ_NO":
					value = String.valueOf((int) rsLines.getInt(col_name));
					break;
				default :
					if (rsLines.getObject(col_name) != null) {
						value = String.valueOf( rsLines.getObject(col_name) );
					}
				}
				System.out.println(col_pos+")= " + value);
				this.setValue(_line_row, col_pos++, value);
				
			}
			
			_line_row++;
		}
	}
	/**
	 * 取的報關基本資料說明
	 * @param kind  海關類型
	 * @param code 代碼
	 * @return CMTBASE_ALL.CODE_CHINESE
	 */
	private String getCMTBASSDesc(String kind,String code) throws Exception {
		String result = "";
		Connection conn = connSQL();
		
		String sql = " select * from CMTBASE_ALL where CODE_KIND=? and CODE=? ";
		PreparedStatement pstat = conn.prepareStatement(sql);
		pstat.setString(1, kind);
		pstat.setString(2, code);
		ResultSet rs = pstat.executeQuery();
		if (rs.next()) {
			result = rs.getString("CODE_CHINESE");
		}
		rs.close();
		pstat.close();
		conn.close();
		return result;
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
