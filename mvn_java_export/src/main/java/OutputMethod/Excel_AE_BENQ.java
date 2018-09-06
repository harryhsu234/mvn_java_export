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

public class Excel_AE_BENQ extends OutputCommon {

	public final static String programmeTitle = "AE_BENQ_轉出程式 20171017版";

	XSSFWorkbook wb;
	XSSFSheet ws;

	public Excel_AE_BENQ() {

	}

	private ResultSet getHeader(String custom_no) throws Exception {
		Connection conn = connSQL();

		PreparedStatement ps = conn
				.prepareStatement(" select a.AUTO_SEQ, 'y' as ISHEAD, a.DCL_DOC_NO, a.DCL_DATE, a.DCL_DOC_TYPE, "
						+ "'TWTPE' as POL, 'TWZZZ' as POD, WAREHOUSE, TRANS_VIA, TERMS_SALES, 10 as REPORT_MONTH, "
						+ " CURRENCY,DOC_CTN_UM, FOB_AMT,EXCHG_RATE,DCL_GW,DCL_NW, TOT_CTN,FOB_AMT,FOB_AMT_TWD  "
						+ " from doc_head a  where a.DCL_DOC_NO = ? ");
		ps.setString(1, custom_no);

		return ps.executeQuery();
	}

	private ResultSet getLines(String custom_no) throws Exception {
		Connection conn = connSQL();

		PreparedStatement ps = conn.prepareStatement(
				" select 'TW' as source_code, a.item_no, a.INVOICE_NO, a.ccc_code, a.BUYER_ITEM_CODE, a.SELLER_ITEM_CODE, "
						+ " b.TERMS_SALES, a.ST_MTD, 'TWD' as curr, a.inv_um, a.DOC_UM, 'TW' productionCountry, "
						+ " a.TRADE_MARK, a.NET_WT, a.DOC_UNIT_P, a.QTY, a.ST_QTY, a.DOC_TOT_P, c.DESCRIPTION as DESC1, a.DESCRIPTION as DESC2 "
						+ " from DOCINVBD a " + " left outer join doc_head b on a.AUTO_SEQ_HEAD = b.AUTO_SEQ "
						+ " left outer join DOCINVBD c on a.AUTO_SEQ_HEAD = c.AUTO_SEQ_HEAD and c.item_no = '*' "
						+ " where b.DCL_DOC_NO = ? and a.item_no != '*' ");
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
		String templatePath = "/Excel_AE_BENQ.xlsx";
		InputStream tmpFile = this.getClass().getResourceAsStream(templatePath);

		wb = new XSSFWorkbook(tmpFile);
		ws = wb.getSheetAt(0);

		doHeader(rsHeader);
		doLines(rsLines);

		String custom_no_fileName = custom_no.replaceAll("/", "").replaceAll(" ", "").replaceAll("-", "");

		outputFilePath = "D:\\XML_OUTPUT\\";
		outputFileName = "BENQ_" + custom_no_fileName + "_" + System.currentTimeMillis() + ".xlsx";
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
		String[] colNames_Header = new String[] { "DCL_DOC_NO", "ISHEAD", "DCL_DATE", "SKIP_ONE", "SKIP_ONE",
				"DCL_DOC_TYPE", "POL", "POD", "WAREHOUSE", "TRANS_VIA", "TERMS_SALES", "REPORT_MONTH", "CURRENCY",
				"DOC_CTN_UM", "SKIP_ONE", "SKIP_ONE", "SKIP_ONE", "SKIP_ONE", "FOB_AMT", "EXCHG_RATE", "DCL_GW",
				"DCL_NW", "TOT_CTN", "FOB_AMT", "SKIP_ONE", "FOB_AMT_TWD" };
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
		String[] colNames_Lines = new String[] { "SOURCE_CODE", "ITEM_NO", "INVOICE_NO", "SKIP_ONE", "SKIP_ONE",
				"CCC_CODE", "BUYER_ITEM_CODE", "SELLER_ITEM_CODE", "TERMS_SALES", "ST_MTD", "SKIP_ONE", "CURR",
				"INV_UM", "DOC_UM", "PRODUCTIONCOUNTRY", "SKIP_ONE", "TRADE_MARK", "SKIP_ONE", "SKIP_ONE", "NET_WT",
				"DOC_UNIT_P", "QTY", "ST_QTY", "DOC_TOT_P", "SKIP_ONE", "SKIP_ONE", "DESC1", "DESC2", "DESC3" };
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
