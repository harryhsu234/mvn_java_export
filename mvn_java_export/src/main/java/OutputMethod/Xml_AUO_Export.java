package OutputMethod;

import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Hashtable;

import org.jdom2.Document;
import org.jdom2.Element;

/**
 * Generate AUO Break Import Xml
 * @author jasonpan
 *
 */
public class Xml_AUO_Export extends AUO_Basic {
	public final static String programmeTitle = "AUO XML";
	
	public Xml_AUO_Export(String type) {
		super();
		this.type = type;
	}
	
	public static void main(String[] args) {
		String type = "GLS_EXPORT";
		Xml_AUO_Export wpg = new Xml_AUO_Export(type);
		try {
//			Config config = new Config();
//			config.getConfig("GLS_EXPORT_TEC");
			wpg.getXML("CBB2066061B070");
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
		String outputFilePath = "D:\\PDF\\";
		outputFileName = "";
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd'T'HHmmss.SSS'Z'");
		String _custom_no = custom_no.replaceAll("/", "").replaceAll(" ", "").replaceAll("-", "");
		
		// Import
		document = getImpDocument(_custom_no);

		String _filename_date = sdf.format(new Date());
		String midleName = goodNo;
		outputFileName = companyName + "_" + _filename_date + "_" + midleName + "_" + _custom_no + ".xml";
		System.out.println(outputFileName);
		saveDocTofile(document);
		// Export
		document = getExpDocument(_custom_no);

		_filename_date = sdf.format(new Date());
		midleName = "EXP";
		outputFileName = companyName + "_" + _filename_date + "_" + midleName + "_" + _custom_no + ".xml";
		System.out.println(outputFileName);

		saveDocTofile(document);

		if (conn != null) {
			conn.close();
		}

		infoBox(outputFilePath + outputFileName + " created.", "Job Done~");
	}
	
	private Document getImpDocument(String custom_no) throws Exception {
		String linesKey = "";
		String code = "";
		String term = "",transType = "",curr="";
		Document document = new Document();
		// Element root = new Element(type);
		Element root = new Element("GLS_IMPORT");
		document.setRootElement(root);
		
		root.addContent(getFromRole(type));
		
		Hashtable<String,String> rs = getHeader(custom_no);
		
		if (rs!=null) {
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
			linesKey = rs.get("AUTO_SEQ");
			goodNo = rs.get("PROC_NO");
			term = rs.get("TERMS_SALES");
			curr = rs.get("CURRENCY");
			transType = rs.get("AIR_SEA");
			
			if (transType !=null && transType.equals("1")) {
				transType = "S";
			} else {
				transType = "A";
			}
			getToRole(type);
			code=rs.get("CNEE_CODE");
			
			Element f30 = new Element("PCZCF30");
			goodNo = rs.get("FREEFIELD4");
			if (goodNo == null || goodNo.trim().length()==0) {
				goodNo= rs.get("PROC_NO");
			}
			addElement(f30, "GOODS_NO",goodNo);
			addElement(f30, "F30_TYPE", "I");
			addElement(f30, "F30_CUNO", rs.get("DCL_DOC_NO"));
			addElement(f30, "F30_RDATE", rs.get("DCL_DATE"));
			addElement(f30, "RELEASE_DATE", rs.get("RL_DATE"));
			addElement(f30, "F30_IDATE",  rs.get("RL_DATE")); // TODO undefinied
			addElement(f30, "F30_EDATE",  rs.get("RL_DATE")); // TODO undefinied
			addElement(f30, "F30_MPNO", rs.get("MAWB"));
			addElement(f30, "F30_SPNO", rs.get("HAWB"));
			addElement(f30, "F30_AIR", transType);
			addElement(f30, "F30_TRWAY", rs.get("TRANS_VIA").substring(0,1));
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
			
//			addElement(f30, "F30_SNAME", rs.get("SHPR_E_NAME"));
//			addElement(f30, "F30_SNUM", rs.get("SHPR_BAN_ID"));
//			addElement(f30, "F30_SADDR", rs.get("SHPR_E_ADDR"));
//			addElement(f30, "F30_SSNUM", rs.get("SHPR_BONDED_ID"));

//			addElement(f30, "F30_BNAME", rs.get("CNEE_E_NAME"));
//			addElement(f30, "F30_BNUM", rs.get("CNEE_BAN_ID"));
//			addElement(f30, "F30_BADDR", rs.get("CNEE_E_ADDR"));
//			addElement(f30, "F30_BSNUM", rs.get("CNEE_BONDED_ID"));
			
			addElement(f30, "F30_PNUM", rs.get("IN_BONDED_BAN"));
			addElement(f30, "F30_PSNUM", rs.get("IN_BONDED_CODE"));
			
			addElement(f30, "BROKER", broker); 
			addElement(f30, "FORWARDER", rs.get("CARRIER_ID")); // TODO undefined
			addElement(f30, "F30_SNATN", rs.get("CNEE_COUNTRY_CODE"));
			addElement(f30, "F30_OPORT", rs.get("FROM_CODE"));
			addElement(f30, "F30_SEND", "NIL"); // TODO undefined
			addElement(f30, "ARRIVE_DATE", rs.get("RL_DATE"));
			addElement(f30, "ETD", rs.get("RL_DATE"));
			addElement(f30, "ETA", rs.get("RL_DATE"));
			Hashtable details = getLinesSummary(linesKey);
			
			addElement(f30, "F30_BOATNO", getDefault(rs.get("CALL_SIGN"), "N"));
			addElement(f30, "F30_BOAT", getDefault(rs.get("FLY_NO"), "N"));
			addElement(f30, "F30_GOOD", limitString(String.valueOf(details.get("DESC")),150));
			addElement(f30, "F30_SFEE", formatAmount(rs.get("CAL_IP_TOT_ITEM_AMT"), 2));
			addElement(f30, "F30_TSFEE", formatAmount(rs.get("FOB_AMT_TWD"), 0));
			addElement(f30, "F30_RATE", formatAmount(rs.get("EXCHG_RATE"), 5));
			addElement(f30, "F30_CUR", rs.get("CURRENCY"));
			addElement(f30, "BOX_QTY", rs.get("TOT_CTN"));
			addElement(f30, "F30_TOTAL", formatAmount(rs.get("DCL_GW"),1));
			String netw = rs.get("DCL_NW");
			if (netw==null) {
				netw="0";
			}
			addElement(f30, "F30_NETW", formatAmount(String.valueOf(Math.round(Double.valueOf( netw).doubleValue())),1));
			addElement(f30, "F30_UNIT", rs.get("DOC_CTN_UM"));
			addElement(f30, "F30_TCOND", rs.get("TERMS_SALES"));
			addElement(f30, "F30_STORE", rs.get("WAREHOUSE"));
			addElement(f30, "F30_TFEE", formatAmount(rs.get("FRT_AMT"),2));
			addElement(f30, "F30_INSU", formatAmount(rs.get("INS_AMT"),2));
			addElement(f30, "F30_PLUS", formatAmount(rs.get("ADD_AMT"),2));
			addElement(f30, "F30_MINUS", formatAmount(rs.get("SUBTRACT_AMT"),2));
			addElement(f30, "QTY", String.valueOf(details.get("QTY")));
			addElement(f30, "F30_BASE", formatAmount(rs.get("DOC_TAX_BASE"), 0));
			addElement(f30, "CUSTOM_FEE", formatAmount(rs.get("IMPORT_TAX"), 2));
			addElement(f30, "CONSTRUCT_FEE", formatAmount(rs.get("PORT_FEE"), 2));
			addElement(f30, "TRADE_FEE", formatAmount(rs.get("EX_TAX_AMT_1"), 2));
			addElement(f30, "DEPOSIT_FEE", formatAmount(rs.get("YA_MONEY"), 2));
			addElement(f30, "GOODS_TAX", formatAmount(rs.get("COMMODITY_TAX"), 2));
			addElement(f30, "OP_TAX", formatAmount(rs.get("SALEST_TAX"), 2));
			addElement(f30, "PAPER_ISSUE_FEE", "0.00");
			addElement(f30, "PAPER_ISSUE_FEE_TAX", "0.00");
			addElement(f30, "PICK_FEE", "0.00");
			addElement(f30, "TRANSIT_FEE", "0.00");
			addElement(f30, "TRANSIT_FEE_TAX", "0.00");
			addElement(f30, "SERVICE_FEE", "0.00");
			addElement(f30, "SERVICE_FEE_TAX", "0.00");
			addElement(f30, "APORTWH_FEE", "0.00");
			addElement(f30, "APORTWH_FEE_TAX", "0.00");
			addElement(f30, "LIFTER_FEE", "0.00");
			addElement(f30, "LIFTER_FEE_TAX", "0.00");
			addElement(f30, "PARKWH_FEE", "0.00");
			addElement(f30, "RCUS_FEE", "0.00");
			addElement(f30, "RCUS_FEE_TAX", "0.00");
			addElement(f30, "TRANS_FEE", "0.00");
			addElement(f30, "TRANS_FEE_TAX", "0.00");
			addElement(f30, "REG_FEE", "0.00");
			addElement(f30, "ICE_FEE", "0.00");
			addElement(f30, "DELAY_FEE", formatAmount(rs.get("DELAY_AMT"), 2));
			addElement(f30, "KEYIN_FEE", "0.00");
			addElement(f30, "INSPECT_FEE", "0.00");
			addElement(f30, "F30_TTAX", formatAmount(rs.get("DCL_AMT"), 2));
			addElement(f30, "F30_LFEE", formatAmount(rs.get("FOB_AMT"),2));
			addElement(f30, "F30_DOC", " ");
			addElement(f30, "F30_CTAX", rs.get("DCL_DOC_TAX_ACC"));

			root.addContent(f30);
		}

		Element att = getAttachment(custom_no, "GLS_IMPORT",code);
		if (att==null) {
			infoBox(custom_no+" PDF 不存在指定路徑", custom_no+"錯誤");
		}
		root.addContent(att);
		getFromRole(type);
		Element glsConInfo = new Element("GlsContactInformation");
		addElement(glsConInfo, "GlsContactName", user);
		addElement(glsConInfo, "GlsEmailAddress", email);
		root.addContent(glsConInfo);

		ResultSet rs1 = getLines(linesKey);
		while (rs1.next()) {
			String ccode = rs1.getString("CCC_CODE").replace("-", "").replace(".", "");

			Element f31 = new Element("PCZCF31");

			addElement(f31, "F31_SERNO", rs1.getString("ITEM_NO"));
			addElement(f31, "F31_NATN", rs1.getString("ORG_COUNTRY"));
			addElement(f31, "F31_BNO", rs1.getString("BUYER_ITEM_CODE"));
			addElement(f31, "PR_NO", " ");
			addElement(f31, "INVOICE_NO", rs1.getString("INVOICE_NO"));
			addElement(f31, "F31_PRMT", "NIL");
			addElement(f31, "F31_HSID", rs1.getString("CCC_CODE"));
			addElement(f31, "F31_UPR", formatAmount(rs1.getString("INV_UNIT_P"),6));
			addElement(f31, "F31_CUR", curr);
			addElement(f31, "F31_NETW", String.valueOf((double) Math.round(rs1.getDouble("NET_WT") * 10) / 10));
			addElement(f31, "F31_AMT", formatAmount(rs1.getString("QTY"),4));
			addElement(f31, "F31_UNIT", rs1.getString("DOC_UM"));
			addElement(f31, "FTAX", "N"); // TODO undefined
			addElement(f31, "F31_HCOMM", ccode);
			addElement(f31, "F31_GRATE", formatAmount("0", 4));
			addElement(f31, "F31_FLFEE", formatAmount(rs1.getString("DOC_TOT_P"),0));
			addElement(f31, "F31_COND", term);
			addElement(f31, "GOOD_DESC", rs1.getString("DESCRIPTION"));
			addElement(f31, "F31_TWAY", rs1.getString("ST_MTD"));

			root.addContent(f31);
		}
		rs1.close();

		Element f32 = new Element("PCZCF32");
		addElement(f32, "F32_GNO", "NIL");
		addElement(f32, "F32_GTYPE", "NIL");
		root.addContent(f32);

		root.addContent(getToRole(type));

		try {
			if (transType.equals("A")) {
				transType = "AIR";
			} else {
				transType = "SEA";
			}
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
			
			Date date = sdf.parse(rs.get("RL_DATE"));
			sdf = new SimpleDateFormat("yyyyMMdd'T'");
			String strDate  = sdf.format(date) + rs.get("RL_TIME")+"00.000Z";
			Xml_AUO_3B3 b3 = new Xml_AUO_3B3();
			b3.setGoodNo(goodNo);
			b3.setAmount(rs.get("F30_TOTAL"));
			b3.setQty(rs.get("BOX_QTY"));
			b3.setId(rs.get("SHPR_BAN_ID"));
			b3.setName(rs.get("SHPR_E_NAME"));
			b3.setTransType(transType);
			b3.setPort(rs.get("FROM_CODE"));
			b3.setRlDate(strDate);
			b3.getXML();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return document;
	}

	private Document getExpDocument(String custom_no) throws Exception {
		String linesKey = "";
		String type = "GLS_EXPORT";
		String term = "", transType = "", code = "",curr="";
		Document document = new Document();
		Element root = new Element(type);
		document.setRootElement(root);

		root.addContent(getFromRole(type));

		Hashtable<String,String> rs = getHeader(custom_no);
		System.out.println(custom_no);
		if (rs !=null) {
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
			linesKey = rs.get("AUTO_SEQ");
			transType = rs.get("F30_GOODSCOME");
			goodNo = rs.get("PROC_NO");
			curr = rs.get("CURRENCY");
			term = rs.get("TERMS_SALES");
			transType = rs.get("AIR_SEA");
			if (transType !=null && transType.equals("1")) {
				transType = "S";
			} else {
				transType = "A";
			}
			getToRole(type);
			code = rs.get("SHPR_CODE");
			
			Element f30 = new Element("PCZCF30");

			addElement(f30, "GOODS_NO", getDefault(rs.get("PROC_NO"), "N")); // TODO undefined
			addElement(f30, "F30_TYPE", "E");
			addElement(f30, "F30_CUNO", rs.get("DCL_DOC_NO"));
			addElement(f30, "F30_RDATE", rs.get("DCL_DATE"));
			addElement(f30, "RELEASE_DATE",  rs.get("RL_DATE"));
			addElement(f30, "F30_MPNO", rs.get("MAWB"));
			addElement(f30, "F30_SPNO", rs.get("HAWB"));
			addElement(f30, "F30_AIR", transType);
			addElement(f30, "F30_TRWAY", rs.get("TRANS_VIA").substring(0,1));
			addElement(f30, "F30_CKIND", rs.get("DCL_DOC_TYPE"));
			addElement(f30, "DECL_REL_WAY", rs.get("DCL_PASS_METHOD"));
//			addElement(f30, "F30_SNAME", rs.get("SHPR_E_NAME"));
//			addElement(f30, "F30_SNUM", rs.get("SHPR_BAN_ID"));
//			addElement(f30, "F30_SADDR", rs.get("SHPR_E_ADDR"));
//			addElement(f30, "F30_SSNUM", rs.get("SHPR_BONDED_ID"));
			
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
			addElement(f30, "BROKER",  broker); // TODO undefined
			addElement(f30, "F30_OPORT",  rs.get("FROM_CODE")); // TODO undefined
			addElement(f30, "F30_TAX", "N"); // TODO undefined
			addElement(f30, "BNATN", rs.get("CNEE_COUNTRY_CODE"));
			addElement(f30, "F30_DNATN", rs.get("FROM_CODE"));
			addElement(f30, "ETD", rs.get("RL_DATE"));
			addElement(f30, "ETA", rs.get("RL_DATE"));
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
			
			addElement(f30, "F30_LFEE", formatAmount(rs.get("FOB_AMT"),2));
			addElement(f30, "F30_TLFEE", formatAmount(rs.get("FOB_AMT_TWD"), 0));
			addElement(f30, "F30_RATE", formatAmount(rs.get("EXCHG_RATE"), 5));
			addElement(f30, "F30_CUR", rs.get("CURRENCY"));
			addElement(f30, "BOX_QTY", rs.get("TOT_CTN"));
			addElement(f30, "F30_TOTAL", formatAmount(rs.get("DCL_GW"),1));
			addElement(f30, "F30_NETW", String.valueOf(Math.round(Double.valueOf(rs.get("DCL_NW")).doubleValue() * 10) / 10));
			addElement(f30, "F30_UNIT", rs.get("DOC_CTN_UM"));
			addElement(f30, "F30_STORE", rs.get("WAREHOUSE"));
			addElement(f30, "F30_TFEE", formatAmount(rs.get("FRT_AMT"),2));
			addElement(f30, "F30_INSU", formatAmount(rs.get("INS_AMT"),2));
			addElement(f30, "F30_PLUS", formatAmount(rs.get("ADD_AMT"),2));
			addElement(f30, "F30_MINUS", formatAmount(rs.get("SUBTRACT_AMT"),2));
			addElement(f30, "QTY", String.valueOf(details.get("QTY")));
			addElement(f30, "CONSTRUCT_FEE", formatAmount("0", 2));
			addElement(f30, "TRADE_FEE", formatAmount(rs.get("EX_TAX_AMT_1"), 2));
			addElement(f30, "F30_TTAX", formatAmount(rs.get("DCL_AMT"), 2));
			addElement(f30, "F30_DOC", ""); // TODO undefined

			root.addContent(f30);
		}

		Element att = getAttachment(custom_no, type,code);
		if (att==null) {
			infoBox("PDF 不存在指定路徑", "錯誤");
		}
		root.addContent(att);
		getFromRole(type);
		Element glsConInfo = new Element("GlsContactInformation");
		addElement(glsConInfo, "GlsContactName", user);
		addElement(glsConInfo, "GlsEmailAddress", email);
		root.addContent(glsConInfo);

		ResultSet rs1 = getLines(linesKey);
		while (rs1.next()) {
			String ccode = rs1.getString("CCC_CODE").replace("-", "").replace(".", "");

			Element f31 = new Element("PCZCF31");

			addElement(f31, "F31_SERNO", rs1.getString("ITEM_NO"));
			addElement(f31, "F31_SNO", rs1.getString("SELLER_ITEM_CODE")); 
			addElement(f31, "INVOICE_NO", rs1.getString("INVOICE_NO"));
			addElement(f31, "F31_PRMT", "NIL");
			addElement(f31, "F31_HSID", rs1.getString("CCC_CODE"));
			addElement(f31, "F31_UPR", formatAmount(rs1.getString("INV_UNIT_P"),6));
			addElement(f31, "F31_CUR", curr);
			addElement(f31, "F31_NETW", String.valueOf((double) Math.round(rs1.getDouble("NET_WT") * 10) / 10));
			addElement(f31, "F31_AMT", formatAmount(rs1.getString("QTY"),4));
			addElement(f31, "F31_UNIT", rs1.getString("DOC_UM"));
			addElement(f31, "F31_GRATE", formatAmount("0", 4));
			addElement(f31, "F31_FLFEE", formatAmount(String.valueOf(Math.round(rs1.getDouble("DOC_TOT_P"))),0));
			addElement(f31, "GOOD_DESC", rs1.getString("DESCRIPTION"));
			addElement(f31, "F31_SWAY",  rs1.getString("ST_MTD"));
			if (rs1.getString("BOND_NOTE")!=null && rs1.getString("BOND_NOTE").length()>0)
				addElement(f31, "F31_BOND",   rs1.getString("BOND_NOTE")); 
			
			root.addContent(f31);
		}

		Element f32 = new Element("PCZCF32");
		addElement(f32, "F32_GNO", "NIL");
		addElement(f32, "F32_GTYPE", "NIL");
		root.addContent(f32);

		root.addContent(getToRole(type));

		try {
			if (transType.equals("A")) {
				transType = "AIR";
			} else {
				transType = "SEA";
			}
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
			
			Date date = sdf.parse(rs.get("RL_DATE"));
			sdf = new SimpleDateFormat("yyyyMMdd'T'");
			String strDate  = sdf.format(date) + rs.get("RL_TIME")+"00.000Z";
			Xml_AUO_3B3 b3 = new Xml_AUO_3B3();
			b3.setGoodNo(goodNo);
			b3.setAmount(rs.get("F30_TOTAL"));
			b3.setQty(rs.get("BOX_QTY"));
			b3.setId(rs.get("SHPR_BAN_ID"));
			b3.setName(rs.get("SHPR_E_NAME"));
			b3.setTransType(transType);
			b3.setPort(rs.get("FROM_CODE"));
			b3.setRlDate(strDate);
			b3.getXML();
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return document;
	}
	
	private Hashtable getHeader(String custom_no) throws Exception {
		Hashtable result = new Hashtable();
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
		
		String sql = "SELECT * from DOC_HEAD  where REPLACE(REPLACE(DCL_DOC_NO,'/',''),' ','') = ? ";
		
		PreparedStatement ps = conn.prepareStatement(sql);
		ps.setString(1, custom_no);

		ResultSet rs = ps.executeQuery();
		ResultSetMetaData rsmd = rs.getMetaData();
		while (rs.next()) {
			for (int i=1;i<=rsmd.getColumnCount();i++) {
				String column = rsmd.getColumnName(i);
				String value = getString(rs.getObject(column));
				if (rs.getObject(column)!=null) {
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
					case "RL_DATE":
						value = convertChinaYear(rs.getString(column));
						break;
					}
				}
				
				result.put(column, value);
			}
		}
		rs.close();
		ps.close();
		
		return result;
	}

	private ResultSet getLines(String custom_no) throws Exception {
		String sql = "SELECT * from DOCINVBD where AUTO_SEQ_HEAD = ? and ITEM_NO != '*'";
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
			String sql = "SELECT sum(QTY) QTY from DOCINVBD where AUTO_SEQ_HEAD = '" + linesKey + "' ";
			System.out.println(sql);
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
