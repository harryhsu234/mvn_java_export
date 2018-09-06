package OutputMethod;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Hashtable;

import org.jdom2.Document;
import org.jdom2.Element;
import org.jdom2.output.Format;
import org.jdom2.output.XMLOutputter;

/**
 * Generate AUO Break Import Xml
 * 
 * @author jasonpan
 *
 */
public class Xml_AUO_Import extends AUO_Basic {
	
	public final static String programmeTitle = "AUO Break Import XML";
	String kind = ""; // 報關單類型
	
	public Xml_AUO_Import(String type) {
		super();
		this.type = type;
	}
	
	public static void main(String[] args) {
		String type = "GLS_IMPORT";
		Xml_AUO_Import wpg = new Xml_AUO_Import(type);
		try {
			// wpg.readXML();
			wpg.getXML("CBG2066060E030");
		} catch (Exception e) {
			e.printStackTrace();
			infoBox(e.getMessage(), "ERROR!!");
		}
	}

	/**
	 * 產生兩個XML
	 * 
	 * @param custom_no
	 * @throws Exception
	 */
	public void getXML(String custom_no) throws Exception {
		conn = connSQL();

		Document document;
		String outputFilePath = "D:\\XML_OUTPUT\\";
		String outputFileName = "";
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd'T'HHmmss.SSS'Z'");
		String _custom_no = custom_no.replaceAll("/", "").replaceAll(" ", "").replaceAll("-", "");
		String[] nos = custom_no.split("/");
		kind = nos[1];
		if (kind == null || kind.trim().length()==0)
			kind = "G2";
		System.out.println("Kind--> "+nos[1]);
		// Import
		document = getImpDocument(_custom_no);

		String _filename_date = sdf.format(new Date());
		String midleName = goodNo;
		outputFileName = companyName + "_" + _filename_date + "_" + midleName + "_" + _custom_no + ".xml";
		System.out.println(outputFileName);

		try {
			Files.createDirectories(new File(outputFilePath).toPath());
			File f = new File(outputFilePath + outputFileName);

			XMLOutputter xmlOutputter = new XMLOutputter(Format.getPrettyFormat());
			xmlOutputter.output(document, new FileOutputStream(f));

		} catch (Exception e1) {
			e1.printStackTrace();
		}

		// Export
		document = getExpDocument(_custom_no);

		_filename_date = sdf.format(new Date());
		midleName = "EXP";
		outputFileName = companyName + "_" + _filename_date + "_" + midleName + "_" + _custom_no + ".xml";
		System.out.println(outputFileName);

		try {
			Files.createDirectories(new File(outputFilePath).toPath());
			File f = new File(outputFilePath + outputFileName);

			XMLOutputter xmlOutputter = new XMLOutputter(Format.getPrettyFormat());
			xmlOutputter.output(document, new FileOutputStream(f));

		} catch (Exception e1) {
			e1.printStackTrace();
		}

		if (conn != null) {
			conn.close();
		}

		infoBox(outputFilePath + outputFileName + " created.", "Job Done~");
	}
	
	private Document getImpDocument(String custom_no) throws Exception {
		String linesKey = "";
		String code = "";
		
		Document document = new Document();
		Element root = new Element(type);
		document.setRootElement(root);

		root.addContent(getFromRole(type));
		Hashtable<String,String> rs = getHeader(custom_no);
		System.out.println(custom_no);
		if (rs!=null) {
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
			linesKey = rs.get("AUTO_SEQ");
			String transType = rs.get("F30_GOODSCOME");
			goodNo = rs.get("PROC_NO");
			String releaseDate = rs.get("RL_DATE");
			releaseDate = String.valueOf(2000 + Integer.valueOf(releaseDate.substring(0, 2))) + "/"	+ releaseDate.substring(2, 4) + "/" + releaseDate.substring(4, 6);
			code=rs.get("SHPR_CODE");
			
			Element f30 = new Element("PCZCF30");

			addElement(f30, "GOODS_NO", goodNo);
			addElement(f30, "F30_TYPE", "I");
			addElement(f30, "F30_CUNO", rs.get("DCL_DOC_NO"));
			addElement(f30, "F30_RDATE", rs.get("DCL_DATE"));
			addElement(f30, "RELEASE_DATE", releaseDate);
			addElement(f30, "F30_IDATE", rs.get("DOC_IMP_DATE"));
			addElement(f30, "F30_EDATE", rs.get("DOC_EXP_DATE"));
			addElement(f30, "F30_MPNO", rs.get("MAWB"));
			addElement(f30, "F30_SPNO", rs.get("HAWB"));
			addElement(f30, "F30_AIR", rs.get("F30_GOODSCOME"));
			addElement(f30, "F30_TRWAY", rs.get("SIPA_MESSAGE_TYPE_WH"));
			addElement(f30, "F30_CKIND", rs.get("DCL_DOC_TYPE"));
			addElement(f30, "DECL_REL_WAY", rs.get("DCL_PASS_METHOD"));
			addElement(f30, "F30_SNAME", rs.get("SHPR_E_NAME"));
			addElement(f30, "F30_SNUM", rs.get("SHPR_BAN_ID"));
			addElement(f30, "F30_SADDR", rs.get("SHPR_E_ADDR"));
			addElement(f30, "F30_SSNUM", rs.get("SHPR_BONDED_ID"));
			addElement(f30, "F30_BNAME", rs.get("CNEE_E_NAME"));
			addElement(f30, "F30_BNUM", rs.get("CNEE_BAN_ID"));
			addElement(f30, "F30_BADDR", rs.get("CNEE_E_ADDR"));
			addElement(f30, "F30_BSNUM", rs.get("CNEE_BONDED_ID"));
			addElement(f30, "F30_PNUM", rs.get("IN_BONDED_BAN"));
			addElement(f30, "F30_PSNUM", rs.get("IN_BONDED_CODE"));
			addElement(f30, "BROKER", rs.get("FWD_BAN_ID")); 
			addElement(f30, "FORWARDER", "NIL"); // TODO undefined
			addElement(f30, "F30_SNATN", rs.get("CNEE_COUNTRY_CODE"));
			addElement(f30, "F30_OPORT", rs.get("FROM_CODE"));
			addElement(f30, "F30_SEND", "NIL"); // TODO undefined
			addElement(f30, "ARRIVE_DATE", releaseDate);
			addElement(f30, "ETD", rs.get("DOC_IMP_DATE"));
			addElement(f30, "ETA", rs.get("DOC_EXP_DATE"));
			String boat = "";
			String boatNo = "";
			if (transType.equals("A")) {
				boatNo = rs.get("CALL_SIGN");
				boat = rs.get("FLY_NO");
			} else {
				boatNo = rs.get("CALL_SIGN");
				boat = rs.get("FLY_NO");
			}
			Hashtable details = getLinesSummary(linesKey);
			addElement(f30, "F30_BOATNO", getDefault(boatNo, "N"));
			addElement(f30, "F30_BOAT", getDefault(boat, "N"));
			addElement(f30, "F30_GOOD", String.valueOf(details.get("DESC")));
			addElement(f30, "F30_SFEE", formatAmount(rs.get("DOC_IMP_CIF_AMT"), 2));
			addElement(f30, "F30_TSFEE", formatAmount(rs.get("DOC_IMP_CIF_TWD"), 0));
			addElement(f30, "F30_RATE", formatAmount(rs.get("EXCHG_RATE"), 5));
			addElement(f30, "F30_CUR", rs.get("CURRENCY"));
			addElement(f30, "BOX_QTY", rs.get("TOT_CTN"));
			addElement(f30, "F30_TOTAL", rs.get("DCL_GW"));
			String netw = rs.get("DCL_NW");
			if (netw==null) {
				netw="0";
			}
			addElement(f30, "F30_NETW", String.valueOf(Math.round(Double.valueOf( netw).doubleValue())));
			addElement(f30, "F30_UNIT", rs.get("DOC_CTN_UM"));
			addElement(f30, "F30_TCOND", rs.get("TERMS_SALES"));
			addElement(f30, "F30_STORE", rs.get("WAREHOUSE"));
			addElement(f30, "F30_TFEE", rs.get("FRT_AMT"));
			addElement(f30, "F30_INSU", rs.get("INS_AMT"));
			addElement(f30, "F30_PLUS", rs.get("ADD_AMT"));
			addElement(f30, "F30_MINUS", rs.get("SUBTRACT_AMT"));
			addElement(f30, "QTY", String.valueOf(details.get("QTY")));
			addElement(f30, "F30_BASE", formatAmount(rs.get("DOC_TAX_BASE"), 0));
			addElement(f30, "CUSTOM_FEE", formatAmount(rs.get("IMPORT_TAX"), 0));
			addElement(f30, "CONSTRUCT_FEE", formatAmount(rs.get("PORT_FEE"), 0));
			addElement(f30, "TRADE_FEE", formatAmount(rs.get("EX_TAX_AMT_1"), 0));
			addElement(f30, "DEPOSIT_FEE", formatAmount(rs.get("YA_MONEY"), 0));
			addElement(f30, "GOODS_TAX", formatAmount(rs.get("COMMODITY_TAX"), 0));
			addElement(f30, "OP_TAX", formatAmount(rs.get("SALEST_TAX"), 0));
			addElement(f30, "DELAY_FEE", formatAmount(rs.get("DELAY_AMT"), 0));
			addElement(f30, "F30_TTAX", formatAmount(rs.get("DCL_AMT"), 0));
			addElement(f30, "F30_LFEE", rs.get("FOB_AMT"));
			addElement(f30, "F30_CTAX", rs.get("DCL_DOC_TAX_ACC"));

			root.addContent(f30);
		}

		Element att = getAttachment(custom_no, type, code, kind);
		if (att==null) {
			infoBox(custom_no+" PDF 不存在指定路徑", custom_no+"錯誤");
		}
		root.addContent(att);
		Element glsConInfo = new Element("GlsContactInformation");
		addElement(glsConInfo, "GlsContactName", user);
		addElement(glsConInfo, "GlsEmailAddress", email);
		root.addContent(glsConInfo);

		ResultSet rs1 = getLines(linesKey);
		while (rs1.next()) {
			String ccode = rs1.getString("CCC_CODE").replace("-", "").replace(".", "");
			double rate = 0;
			if (rs1.getString("TAX_RATE_P")!=null) {
				rate = Double.valueOf(rs1.getString("TAX_RATE_P"));
				if (rate != 0 ) {
					rate=rate/100;
				}
			}
			
			Element f31 = new Element("PCZCF31");
			
			addElement(f31, "F31_SERNO", rs1.getString("ITEM_NO"));
			addElement(f31, "F31_NATN", rs1.getString("ORG_COUNTRY"));
			addElement(f31, "F31_BNO", rs1.getString("SELLER_ITEM_CODE"));
			addElement(f31, "PR_NO", " ");
			addElement(f31, "INVOICE_NO", rs1.getString("INN"));
			addElement(f31, "F31_PRMT", "NIL");
			addElement(f31, "F31_HSID", rs1.getString("CCC_CODE"));
			addElement(f31, "F31_UPR", formatAmount(rs1.getString("INV_UNIT_P"),6));
			addElement(f31, "F31_CUR", rs1.getString("CURRENCY"));
			addElement(f31, "F31_NETW", String.valueOf((double) Math.round(rs1.getDouble("NET_WT") * 100) / 100));
			addElement(f31, "F31_AMT", rs1.getString("QTY"));
			addElement(f31, "F31_UNIT", rs1.getString("DOC_UM"));
			addElement(f31, "FTAX", rs1.getString("ISCALC_WT"));
			addElement(f31, "F31_HCOMM", code);
			addElement(f31, "F31_GRATE", formatAmount(String.valueOf(rate), 4));
			addElement(f31, "F31_FLFEE", formatAmount(rs1.getString("DOC_TOT_P"),0));
//			addElement(f31, "F31_GRATE", String.valueOf(rate));
//			addElement(f31, "F31_FLFEE", String.valueOf(Math.round(rs1.getDouble("DOC_TOT_P"))));
			addElement(f31, "F31_COND", rs1.getString("TERMS"));
			addElement(f31, "GOOD_DESC", rs1.getString("DESCRIPTION"));
			addElement(f31, "F31_TWAY", rs1.getString("TAX_METHOD"));

			root.addContent(f31);
		}
		rs1.close();

		Element f32 = new Element("PCZCF32");
		addElement(f32, "F32_GNO", "NIL");
		addElement(f32, "F32_GTYPE", "NIL");
		root.addContent(f32);

		root.addContent(getToRole(type));

		return document;
	}

	private Document getExpDocument(String custom_no) throws Exception {
		String linesKey = "";
		String code = "";
		Document document = new Document();
		Element root = new Element("GLS_EXPORT");
		document.setRootElement(root);

		root.addContent(getFromRole(type));

		Hashtable<String,String> rs = getHeader(custom_no);
		System.out.println(custom_no);
		if (rs !=null) {
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
			linesKey = rs.get("AUTO_SEQ");
			String transType = rs.get("F30_GOODSCOME");
			goodNo = rs.get("PROC_NO");
			String releaseDate = rs.get("RL_DATE");
			releaseDate = String.valueOf(2000 + Integer.valueOf(releaseDate.substring(0, 2))) + "/"
					+ releaseDate.substring(2, 4) + "/" + releaseDate.substring(4, 6);
			code=rs.get("CNEE_CODE");
			Element f30 = new Element("PCZCF30");

			addElement(f30, "GOODS_NO", getDefault(goodNo, "N")); // TODO undefined
			addElement(f30, "F30_TYPE", "E");
			addElement(f30, "F30_CUNO", rs.get("DCL_DOC_NO"));
			addElement(f30, "F30_RDATE", rs.get("DCL_DATE"));
			addElement(f30, "RELEASE_DATE", releaseDate);
			addElement(f30, "F30_MPNO", rs.get("MAWB"));
			addElement(f30, "F30_SPNO", rs.get("HAWB"));
			addElement(f30, "F30_AIR", rs.get("F30_GOODSCOME"));
			addElement(f30, "F30_TRWAY", rs.get("AIR_SEA"));
			addElement(f30, "F30_CKIND", rs.get("DCL_DOC_TYPE"));
			addElement(f30, "DECL_REL_WAY", rs.get("DCL_PASS_METHOD"));
			
			addElement(f30, "F30_SNAME", rs.get("CNEE_E_NAME"));
			addElement(f30, "F30_SNUM", rs.get("CNEE_BAN_ID"));
			addElement(f30, "F30_SADDR", rs.get("CNEE_E_ADDR"));
			addElement(f30, "F30_SSNUM", rs.get("CNEE_BONDED_ID"));
			
			addElement(f30, "F30_BNAME", rs.get("SHPR_E_NAME"));
			addElement(f30, "F30_BNUM", rs.get("SHPR_BAN_ID"));
			addElement(f30, "F30_BADDR", rs.get("SHPR_E_ADDR"));
			addElement(f30, "F30_BSNUM", rs.get("SHPR_BONDED_ID"));
			
			addElement(f30, "F30_ONUM", "NIL"); // TODO undefined
			addElement(f30, "F30_RECV", rs.get("CNEE_E_NAME"));
			addElement(f30, "BROKER",  rs.get("FWD_BAN_ID")); // TODO undefined
			addElement(f30, "F30_OPORT",  rs.get("FROM_CODE")); // TODO undefined
			addElement(f30, "F30_TAX", "N"); // TODO undefined
			addElement(f30, "BNATN", rs.get("CNEE_COUNTRY_CODE"));
			addElement(f30, "F30_DNATN", rs.get("FROM_CODE"));
			addElement(f30, "ETD", rs.get("DOC_IMP_DATE"));
			addElement(f30, "ETA", rs.get("DOC_EXP_DATE"));
			String boat = "";
			String boatNo = "";
			if (transType.equals("A")) {
				boatNo = rs.get("FLY_NO");
				boat = rs.get("FLY_NO");
			} else {
				boatNo = rs.get("CALL_SIGN");
				boat = rs.get("FLY_NO");
			}
			Hashtable details = getLinesSummary(linesKey);
			addElement(f30, "F30_BOATNO", getDefault(boatNo, "N"));
			addElement(f30, "F30_BOAT", getDefault(boat, "N"));
			addElement(f30, "SO_NO", " "); // TODO undefined
			addElement(f30, "SEA_CFS_CFS", ""); // TODO undefined
			addElement(f30, "F30_LFEE", formatAmount("0", 2));
			addElement(f30, "F30_TLFEE", formatAmount("0", 2));
			addElement(f30, "F30_RATE", formatAmount(rs.get("EXCHG_RATE"), 5));
			addElement(f30, "F30_CUR", rs.get("CURRENCY"));
			addElement(f30, "BOX_QTY", rs.get("TOT_CTN"));
			addElement(f30, "F30_TOTAL", rs.get("DCL_GW"));
			addElement(f30, "F30_NETW", String.valueOf(Math.round(Double.valueOf(rs.get("DCL_NW")).doubleValue() * 10) / 10));
			addElement(f30, "F30_UNIT", rs.get("DOC_CTN_UM"));
			addElement(f30, "F30_STORE", rs.get("WAREHOUSE"));
			addElement(f30, "F30_TFEE", formatAmount("0", 2));
			addElement(f30, "F30_INSU", formatAmount("0", 2));
			addElement(f30, "F30_PLUS", formatAmount("0", 2));
			addElement(f30, "F30_MINUS", formatAmount("0", 2));
			addElement(f30, "QTY", String.valueOf(details.get("QTY")));
			addElement(f30, "CONSTRUCT_FEE", formatAmount("0", 2));
			addElement(f30, "TRADE_FEE", formatAmount("0", 2));
			addElement(f30, "F30_TTAX", formatAmount("0", 2));
			addElement(f30, "F30_DOC", ""); // TODO undefined

			root.addContent(f30);
		}

		Element att = getAttachment(custom_no, "GLS_EXPORT",code, kind);
		if (att==null) {
			infoBox("PDF 不存在指定路徑", "錯誤");
		}
		root.addContent(att);
		
		Element glsConInfo = new Element("GlsContactInformation");
		addElement(glsConInfo, "GlsContactName", user);
		addElement(glsConInfo, "GlsEmailAddress", email);
		root.addContent(glsConInfo);

		ResultSet rs1 = getLines(linesKey);
		while (rs1.next()) {
			String ccode = rs1.getString("CCC_CODE").replace("-", "").replace(".", "");
			double rate = 0;
			if (rs1.getString("TAX_RATE_P")!=null) {
				rate = Double.valueOf(rs1.getString("TAX_RATE_P"));
				if (rate != 0 ) {
					rate=rate/100;
				}
			}
			
			Element f31 = new Element("PCZCF31");

			addElement(f31, "F31_SERNO", rs1.getString("ITEM_NO"));
			addElement(f31, "F31_SNO", rs1.getString("SELLER_ITEM_CODE")); 
			addElement(f31, "INVOICE_NO", rs1.getString("INN"));
			addElement(f31, "F31_PRMT", "NIL");
			addElement(f31, "F31_HSID", rs1.getString("CCC_CODE"));
			addElement(f31, "F31_UPR", formatAmount(rs1.getString("INV_UNIT_P"),6));
			addElement(f31, "F31_NETW", String.valueOf((double) Math.round(rs1.getDouble("NET_WT") * 100) / 100));
			addElement(f31, "F31_AMT", rs1.getString("QTY"));
			addElement(f31, "F31_UNIT", rs1.getString("DOC_UM"));
			addElement(f31, "F31_GRATE", String.valueOf(rate));
			addElement(f31, "F31_FLFEE", String.valueOf(Math.round(rs1.getDouble("DOC_TOT_P"))));
			addElement(f31, "GOOD_DESC", rs1.getString("DESCRIPTION"));
			addElement(f31, "F31_SWAY",  rs1.getString("TAX_METHOD")); 

			root.addContent(f31);
		}

		Element f32 = new Element("PCZCF32");
		addElement(f32, "F32_GNO", "NIL");
		addElement(f32, "F32_GTYPE", "NIL");
		root.addContent(f32);

		root.addContent(getToRole(type));

		return document;
	}

	private Hashtable getHeader(String custom_no) throws Exception {
		Hashtable result = new Hashtable();
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
		
		String sql = "SELECT * from DOC_H_I where REPLACE(REPLACE(DCL_DOC_NO,'/',''),' ','') = ? ";
		
		PreparedStatement ps = conn.prepareStatement(sql);
		ps.setString(1, custom_no);

		ResultSet rs = ps.executeQuery();
		ResultSetMetaData rsmd = rs.getMetaData();
		while (rs.next()) {
			for (int i=1;i<=rsmd.getColumnCount();i++) {
				String column = rsmd.getColumnName(i);
				String value = getString(rs.getObject(column));
				
				switch (column) {
				case "DCL_DOC_NO":
					value=value.replaceAll("/", "").replaceAll(" ", "");
					System.out.println(column+" : "+value);
					break;
				case "DCL_DATE":
				case "DOC_IMP_DATE":
				case "DOC_EXP_DATE":
				case "LASTUPD_TIME":
					value = sdf.format(rs.getDate(column));
					break;
				}
				result.put(column, value);
			}
		}
		return result;
	}

	private ResultSet getLines(String custom_no) throws Exception {
		String sql = "SELECT * from DI_INVBD where AUTO_SEQ_HEAD = ? and ITEM_NO != '*'";
		PreparedStatement ps = conn.prepareStatement(sql);
		ps.setString(1, custom_no);

		return ps.executeQuery();
	}

	private Hashtable getResult(PreparedStatement ps) throws Exception {
		Hashtable result = new Hashtable();
		
		ResultSet rs = ps.executeQuery();
		ResultSetMetaData rsmd = rs.getMetaData();
		while (rs.next()) {
			for (int i=1;i<=rsmd.getColumnCount();i++) {
				String column = rsmd.getColumnName(i);
				String value = getString(rs.getObject(column));
				SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
				switch (column) {
				case "AUTO_SEQ":
					value="";
					break;
				case "DCL_DOC_NO":
					value=value.replaceAll("/", "").replaceAll(" ", "");
					System.out.println(column+" : "+value);
					break;
				}
				result.put(column, value);
			}
		}
		return result;
	}
	
	private Hashtable getLinesSummary(String linesKey) {
		Hashtable result = new Hashtable();
		try {
			String sql = "SELECT sum(QTY) QTY from DI_INVBD where AUTO_SEQ_HEAD = '" + linesKey + "' ";
			Statement stat = conn.createStatement();
			ResultSet rs = stat.executeQuery(sql);
			if (rs.next()) {
				result.put("QTY", rs.getString("QTY"));
			}

			sql = "SELECT DESCRIPTION from DI_INVBD where AUTO_SEQ_HEAD = '" + linesKey + "' and ITEM_NO='1' ";
			stat = conn.createStatement();
			rs = stat.executeQuery(sql);
			if (rs.next()) {
				result.put("DESC", rs.getString("DESCRIPTION"));
			}

		} catch (Exception e) {
			e.printStackTrace();
		}

		return result;
	}
}
