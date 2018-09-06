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
import java.util.Hashtable;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * OP 打單排行榜
 * 統計OP 輸入的報單量作排行榜
 * @author harry206
 *
 */
public class Excel_Common_OP_Key_Report extends OutputCommon {

	public final static String programmeTitle = "OP 打單統計表";
	
	String sWhere = ""; 
	String isExpImp = ""; 
	Hashtable<String, String> conditionPack = new Hashtable<String, String>();
	
	XSSFWorkbook wb;
	XSSFSheet ws;
	
	public Excel_Common_OP_Key_Report() {
	}
	public Excel_Common_OP_Key_Report(String sWhere, String isExpImp) {
		this.sWhere = sWhere;
		this.isExpImp = isExpImp;
	}
	public Excel_Common_OP_Key_Report(String sWhere, String isExpImp, Hashtable<String, String> conditionPack) {
		this.sWhere = sWhere;
		this.isExpImp = isExpImp;
		this.conditionPack = conditionPack;
	}
	

	
	public void run() throws Exception {
		getExcel();
	}
	
	private ResultSet getTotal() throws Exception {
		Connection conn = OutputCommon.connSQL();

		String exp_sql = "select UPPER(isnull(case OP_CODE when '' then '空白OP' else OP_CODE end,'空白OP')) as OP, COUNT(*) as OP_COUNT \r\n" + 
				"from DOC_HEAD \r\n" + 
				"where 1=1 " + this.sWhere +" \r\n" + 
				"group by UPPER(isnull(case OP_CODE when '' then '空白OP' else OP_CODE end,'空白OP')) " +
				"order by OP_COUNT desc ";

		String imp_sql = "select UPPER(isnull(case OP_CODE when '' then '空白OP' else OP_CODE end,'空白OP')) as OP, COUNT(*) as OP_COUNT \r\n" + 
				"from DOC_H_I \r\n" + 
				"where 1=1 " + this.sWhere +" \r\n" + 
				"group by UPPER(isnull(case OP_CODE when '' then '空白OP' else OP_CODE end,'空白OP')) " +
				"order by OP_COUNT desc ";
		String sql;
		if (this.isExpImp.equals("EXP"))
			sql = exp_sql;
		else
			sql = imp_sql;

		System.out.println(sql);

		PreparedStatement ps = conn.prepareStatement(sql,ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_READ_ONLY);
		// ps.setString(1, custom_no);

		return ps.executeQuery();
		
	}
	
	private ResultSet getOPDetail(String op_code) throws Exception {
		Connection conn = OutputCommon.connSQL();

		String exp_sql = "select UPPER(isnull(case OP_CODE when '' then '空白OP' else OP_CODE end,'空白OP')) as OP_CODE, CONVERT(VARCHAR(5), DCL_DATE, 101) as '報關日期', DCL_DOC_NO, \r\n" + 
				"SHPR_CODE, isnull(SHPR_C_NAME, SHPR_E_NAME) as SHPR_NAME, MAWB, HAWB, FROM_CODE, TO_CODE, RL_DATE \r\n" + 
				"from doc_head \r\n" + 
				"where 1=1 and UPPER(isnull(case OP_CODE when '' then '空白OP' else OP_CODE end,'空白OP')) = '" + op_code.toUpperCase() + "' " + this.sWhere;

		String imp_sql = "select UPPER(isnull(case OP_CODE when '' then '空白OP' else OP_CODE end,'空白OP')) as OP_CODE, CONVERT(VARCHAR(5), DCL_DATE, 101) as '報關日期', DCL_DOC_NO, \r\n" + 
				"SHPR_CODE, isnull(SHPR_C_NAME, SHPR_E_NAME) as SHPR_NAME, MAWB, HAWB, FROM_CODE, TO_CODE, \r\n" + 
				"'20'+SUBSTRING(RL_DATE, 1, 2)+'/'+SUBSTRING(RL_DATE, 3, 2)+'/'+SUBSTRING(RL_DATE, 5, 2) as RL_DATE \r\n" + 
				"from DOC_H_I \r\n " + 
				"where 1=1 and UPPER(isnull(case OP_CODE when '' then '空白OP' else OP_CODE end,'空白OP')) = '" + op_code.toUpperCase() + "' " + this.sWhere;
		String sql;
		if (this.isExpImp.equals("EXP"))
			sql = exp_sql;
		else
			sql = imp_sql;

//		System.out.println(sql);

		PreparedStatement ps = conn.prepareStatement(sql,ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_READ_ONLY);
		// ps.setString(1, custom_no);

		return ps.executeQuery();
	}
	
	/**
	 * @param custom_no
	 * @throws Exception
	 */
	public void getExcel() throws Exception {
		// get data from GIC
		ArrayList<ResultSet> alRsOPDetail = new ArrayList<ResultSet>();
		alRsOPDetail.clear();

		ResultSet rsTotal = getTotal();
		while(rsTotal.next()) {
			String op_code = rsTotal.getString("OP").trim();
			
			alRsOPDetail.add(getOPDetail(op_code));
		}
		rsTotal.beforeFirst();
		
//		for(ResultSet rsOPDetail : alRsOPDetail) {
//			while(rsOPDetail.next()) {
//				System.out.print(rsOPDetail.getString("OP_CODE"));
//				System.out.print(" - ");
//				System.out.print(rsOPDetail.getString("報關日期"));
//				System.out.print(" - ");
//				System.out.print(rsOPDetail.getString("SHPR_NAME"));
//				System.out.print(" - ");
//				System.out.println(rsOPDetail.getString("DCL_DOC_NO"));
//			}
//			rsOPDetail.beforeFirst();
//		}

		// get xlsx template
		String templatePath = "/Excel_Common_OP_Key_Report.xlsx";
		InputStream tmpFile= this.getClass().getResourceAsStream(templatePath);
		
		wb = new XSSFWorkbook(tmpFile);
		
		doTotal(rsTotal);
		doOPDetail(alRsOPDetail);
		
		// write xlsx
		outputFilePath = "D:\\XML_OUTPUT\\";
        outputFileName = "OP_Key_Report_" + System.currentTimeMillis()+".xlsx";
        Files.createDirectories(new File(outputFilePath).toPath());
		FileOutputStream stream = new FileOutputStream(outputFilePath + outputFileName);
		wb.write(stream);
		stream.close();
		
		wb.close();
		
		System.out.println("JOB_DONE");	
		infoBox(outputFileName + " 產生完畢", "JOB_DONE");
	}
	
	// 各個 OP之單量
	private void doOPDetail(ArrayList<ResultSet> alRsOPDetail) throws SQLException, Exception {
		String[] colNames_Lines = new String[] { 
				"OP_CODE", "報關日期", "DCL_DOC_NO", "SHPR_CODE", "SHPR_NAME",
				"MAWB", "HAWB", "FROM_CODE", "TO_CODE", "RL_DATE"
		};
		int index = wb.getSheetIndex("TMP");
		for(ResultSet rs : alRsOPDetail) {
			// 抓第一筆的資料.OP_CODE 作為新sheet name
			rs.first();
			String op_code = rs.getString("OP_CODE");
			rs.beforeFirst();
			
			ws = wb.cloneSheet(index, op_code);
			
			setValue("B1", conditionPack.get("DCL_DATE"));
			setValue("B2", conditionPack.get("JOB_TYPE"));
			setValue("B3", conditionPack.get("isRelease"));
			
			int _line_row = 5;
			while(rs.next()) {
				int col_pos = 0;
				for(String col_name : colNames_Lines) {
					String chr10 = "\n";
					
					if(col_name.equalsIgnoreCase("SKIP_ONE")) {
						col_pos++;
						continue;
					}
						
					this.setValue(_line_row, col_pos++, rs.getObject(col_name));
				}
				_line_row++;
			}
		}
		
		wb.removeSheetAt(index);
		
	}

	// 出口 OP 每月/日 單量統計總表
	private void doTotal(ResultSet rsTotal) throws Exception {
		// TODO Auto-generated method stub
		ws = wb.getSheet("總表");
		
		setValue("B1", conditionPack.get("DCL_DATE"));
		setValue("B2", conditionPack.get("JOB_TYPE"));
		setValue("B3", conditionPack.get("isRelease"));
		
		String[] colNames_Lines = new String[] { 
				"OP", "OP_COUNT"
		};
		int _line_row = 5;
		while(rsTotal.next()) { 
			int col_pos = 0;
			for(String col_name : colNames_Lines) {
				String chr10 = "\n";
				
				if(col_name.equalsIgnoreCase("SKIP_ONE")) {
					col_pos++;
					continue;
				}
					
				this.setValue(_line_row, col_pos++, rsTotal.getObject(col_name));
				
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
