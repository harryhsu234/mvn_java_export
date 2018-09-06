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

public class Excel_C4_BENQ2 extends OutputCommon {

	public final static String programmeTitle = "C4_BENQ_第二版格式";

	XSSFWorkbook wb;
	XSSFSheet ws;

	public Excel_C4_BENQ2() {

	}

	private ResultSet getHeader(String custom_no) throws Exception {
		Connection conn = connSQL();
		
		String sql = " select a.DCL_DOC_NO, 'y' as ISHEAD, a.DCL_DATE, CONVERT(varchar, a.DOC_IMP_DATE, 111) as DOC_IMP_DATE, CONVERT(varchar, a.DOC_EXP_DATE, 111) as DOC_EXP_DATE, \r\n" + 
					" a.DCL_DOC_TYPE, a.FROM_CODE, a.TO_CODE, a.WAREHOUSE, a.TRANS_VIA, \r\n" + 
					" a.TERMS_SALES, 0 as report_mounth, a.CURRENCY, a.DOC_CTN_UM, a.DOC_OTR_DESC, \r\n" + 
					" a.P_DOC_ITEM_FRN, a.EXCHG_RATE, a.DCL_GW, a.DCL_NW, a.TOT_CTN, \r\n" + 
					" a.FOB_AMT, a.DOC_IMP_CIF_AMT, a.DOC_IMP_CIF_TWD \r\n" + 
					" from doc_h_i a where a.DCL_DOC_NO = ? ";

		PreparedStatement ps = conn.prepareStatement(sql);
		ps.setString(1, custom_no);

		return ps.executeQuery();
	}

	private ResultSet getLines(String custom_no) throws Exception {
		Connection conn = connSQL();

		String sql = " SELECT 'TW' as SOURCE_CODE, B.ITEM_NO, B.CCC_CODE, B.BUYER_ITEM_CODE, B.SELLER_ITEM_CODE, \r\n" + 
				"        B.TERMS, \r\n" + 
				"        B.TAX_METHOD, a.CURRENCY, B.DOC_UM, B.ST_UM, B.ORG_COUNTRY, " + 
				"		 B.ORG_DCL_NO, B.ORG_DCL_NO_ITEM, B.NET_WT, \r\n" + 
				"        B.DOC_UNIT_P, B.QTY, B.ST_QTY, B.DOC_TOT_P, \r\n" + 
				"        a.EXCHG_RATE, B.TAX_RATE_P, C.DESCRIPTION AS DESC1, B.DESCRIPTION AS DESC2 \r\n" + 
				" FROM DOC_H_I A \r\n" + 
				" LEFT OUTER JOIN DI_INVBD B ON A.AUTO_SEQ= B.AUTO_SEQ_HEAD \r\n" + 
				" left outer join DOCINVBD C on a.AUTO_SEQ = C.AUTO_SEQ_HEAD and C.item_no = '*' \r\n" + 
				" where B.ITEM_NO != '*' and a.DCL_DOC_NO = ? ";
		
		PreparedStatement ps = conn.prepareStatement(sql);
		ps.setString(1, custom_no);

		
		return ps.executeQuery();
	}

	/**
	 * @param custom_no
	 * @throws Exception
	 */
	public void getExcel(String custom_no) throws Exception {
		// get data from GIC
		ResultSet rsHeader = getHeader(custom_no);
		ResultSet rsLines = getLines(custom_no);

		// get xlsx template
		String templatePath = "/Excel_C4_BENQ2.xlsx";
		InputStream tmpFile = this.getClass().getResourceAsStream(templatePath);

		wb = new XSSFWorkbook(tmpFile);
		ws = wb.getSheetAt(0);

		doHeader(rsHeader);
		doLines(rsLines);

		String custom_no_fileName = custom_no.replaceAll("/", "").replaceAll(" ", "").replaceAll("-", "");

		outputFilePath = "D:\\XML_OUTPUT\\";
		outputFileName = "BENQv2_" + custom_no_fileName + "_" + System.currentTimeMillis() + ".xlsx";
		Files.createDirectories(new File(outputFilePath).toPath());
		FileOutputStream stream = new FileOutputStream(outputFilePath + outputFileName);
		wb.write(stream);
		stream.close();

		wb.close();

		System.out.println("JOB_DONE");
		infoBox(outputFileName + " 產生完畢", "JOB_DONE");
	}

	/**
	 * �NResultSet ���G�g�JExcel �����Y
	 * 
	 * @param rsHeader
	 * @throws SQLException
	 * @throws Exception
	 */
	@SuppressWarnings("static-access")
	private void doHeader(ResultSet rsHeader) throws SQLException, Exception {
		String[] colNames_Header = new String[] { 
				"DCL_DOC_NO", "ISHEAD", "DCL_DATE", "DOC_IMP_DATE", "DOC_EXP_DATE",
				"DCL_DOC_TYPE", "FROM_CODE", "TO_CODE", "WAREHOUSE", "TRANS_VIA", 
				"TERMS_SALES", "REPORT_MONTH", "CURRENCY", "DOC_CTN_UM", "DOC_OTR_DESC", 
				"SKIP_ONE", "SKIP_ONE", "SKIP_ONE", "FOB_AMT", "EXCHG_RATE", "DCL_GW",
				"DCL_NW", "TOT_CTN", "FOB_AMT", "DOC_IMP_CIF_AMT", "DOC_IMP_CIF_TWD" };
		int _header_row = 1;
		while (rsHeader.next()) {
			int col_pos = 0;
			for (String col_name : colNames_Header) {
				if (col_name.equalsIgnoreCase("SKIP_ONE")) {
					col_pos++;
					continue;
				} else if (col_name.equalsIgnoreCase("DCL_DATE")) {
					Date dDCL_DATE = rsHeader.getDate("DCL_DATE");
					Calendar cal = Calendar.getInstance();
					cal.setTime(dDCL_DATE);
					int year = cal.get(cal.YEAR) - 1911;
					cal.set(year, cal.get(cal.MONTH), cal.get(cal.DATE));

					SimpleDateFormat sdfSource = new SimpleDateFormat("yyy-MM-dd");
					String sDCL_DATE = sdfSource.format(cal.getTime());

					this.setValue(_header_row, col_pos++, sDCL_DATE);
					continue;
				} else if (col_name.equalsIgnoreCase("REPORT_MONTH")) {
					Date dDCL_DATE = rsHeader.getDate("DCL_DATE");
					Calendar cal = Calendar.getInstance();
					cal.setTime(dDCL_DATE);
					int month = cal.get(Calendar.MONTH);
					int _REPORT_MONTH = (month != 0) ? month : 12;

					this.setValue(_header_row, col_pos++, _REPORT_MONTH);
					continue;
				}

				this.setValue(_header_row, col_pos++, rsHeader.getObject(col_name));
			}
		}

	}

	/**
	 * �N����ResultSet �g�JExcel
	 * 
	 * @param rsLines
	 * @throws SQLException
	 * @throws Exception
	 */
	private void doLines(ResultSet rsLines) throws SQLException, Exception {
		String[] colNames_Lines = new String[] { "SOURCE_CODE", "ITEM_NO", "SKIP_ONE", "SKIP_ONE", "SKIP_ONE",
				"CCC_CODE", "BUYER_ITEM_CODE", "SELLER_ITEM_CODE", "TERMS", "SKIP_ONE", "TAX_METHOD", "CURRENCY",
				"DOC_UM", "ST_UM", "ORG_COUNTRY", "SKIP_ONE", "SKIP_ONE", "ORG_DCL_NO", "ORG_DCL_NO_ITEM", "NET_WT",
				"DOC_UNIT_P", "QTY", "ST_QTY", "DOC_TOT_P", "EXCHG_RATE", "TAX_RATE_P", "DESC1", "DESC2", "DESC3" };
		int _line_row = 3;
		while (rsLines.next()) {
			int col_pos = 0;
			for (String col_name : colNames_Lines) {
				String chr10 = "\n";

				System.out.println("SET " + col_name);
				if (col_name.equalsIgnoreCase("SKIP_ONE")) {
					col_pos++;
					continue;
				} else if (col_name.equalsIgnoreCase("DESC1")) {

					String _DESC1 = rsLines.getString(col_name);
					if (rsLines.getString("ITEM_NO").equals("1") && _DESC1 != null) {
						this.setValue(_line_row, col_pos++, rsLines.getString(col_name));
					} else
						col_pos++;

					continue;
				} else if (col_name.equalsIgnoreCase("DESC2")) {
					String _DESC = rsLines.getString("DESC2");
					if (_DESC.contains(chr10))
						_DESC = _DESC.split(chr10)[0];

					this.setValue(_line_row, col_pos++, _DESC);

					continue;
				} else if (col_name.equalsIgnoreCase("DESC3")) {
					String _DESC = rsLines.getString("DESC2");
					if (_DESC.contains(chr10))
						_DESC = _DESC.split(chr10)[1];
					else
						_DESC = "";

					this.setValue(_line_row, col_pos++, _DESC);

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
		int row_pos = cr.getRow();
		int col_pos = cr.getCol();

		setValue(row_pos, col_pos, value);
	}

	private void setValue(int row_pos, int col_pos, Object value) throws Exception {
		super.setValue(ws, row_pos, col_pos, value);
	}

}
