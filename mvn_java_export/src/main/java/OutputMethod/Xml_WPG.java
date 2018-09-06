package OutputMethod;
import java.awt.Font;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;

import org.jdom2.Document;
import org.jdom2.Element;
import org.jdom2.output.Format;
import org.jdom2.output.XMLOutputter;

/**
 * XML WPG 
 * @author Harry Hsu
 *
 */
public class Xml_WPG extends OutputCommon {
	
	public final static String programmeTitle = "AE/AI_WPG_轉出程式 20171017版";
	
	
	private String filename_date = "";
	private String filename_hawb = "";
	private String filename_mawb = "";

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Xml_WPG wpg = new Xml_WPG();
		try {
			wpg.createGUI();
		} catch (Exception e) {
			e.printStackTrace();
			infoBox(e.getMessage(), "ERROR!!");
		}
	}

	private void createGUI() throws ClassNotFoundException, InstantiationException, IllegalAccessException,
			UnsupportedLookAndFeelException {
		UIManager.setLookAndFeel("javax.swing.plaf.metal.MetalLookAndFeel");

		final JFrame frame = new JFrame();
		frame.setSize(400, 200);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		;
		frame.setLayout(new GridLayout(3, 1));

		JLabel jLabel01 = new JLabel("報單號碼");
		jLabel01.setFont(new Font(jLabel01.getFont().getName(), Font.PLAIN, 30));

		String custom_no = "";
		// custom_no = this.testNO;
		final JTextField jTF01 = new JTextField(custom_no);
		jTF01.setFont(new Font(jLabel01.getFont().getName(), Font.PLAIN, 28));

		JButton jBtn01 = new JButton("RUN XML!!");
		jBtn01.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				String custom_no = jTF01.getText();

				try {
					if (custom_no == null || custom_no.length() == 0)
						throw new Exception("報單號碼不得為空白");

					getXML(custom_no);
				} catch (Exception e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
					infoBox(e1.getMessage(), "ERROR in ActionEvent!!");
				}

			}
		});

		frame.add(jLabel01);
		frame.add(jTF01);
		frame.add(jBtn01);

		frame.setVisible(true);
	}

	

	public void getXML(String custom_no) throws Exception {
		String[] inputA = custom_no.split(",");
		String _payTerm = "";
		if (inputA.length >= 2) {
			custom_no = inputA[0];
			_payTerm = inputA[1];
		}

		Document document;
		// check custom no is 660(for IMP) or 66a(for EXP)
		if (custom_no.contains("/660 /")) {
			// is IMP
			document = getImpDocument(custom_no, _payTerm);
		} else if (custom_no.contains("/66A /")) {
			// is EXP
			document = getExpDocument(custom_no, _payTerm);
		} else {
			// not IMP and not EXP
			infoBox(custom_no + "is not 660 or 66A, 轉出XML 作業取消", "轉出XML 作業取消");
			return;
		}

		String outputFilePath = "D:\\XML_OUTPUT\\";
		String outputFileName = "";
		String _custom_no = custom_no.replaceAll("/", "").replaceAll(" ", "").replaceAll("-", "");
		if (filename_hawb.equals(""))
			filename_hawb = filename_mawb;
		String _filename_hawb = filename_hawb.replaceAll("/", "").replaceAll(" ", "");
		String _filename_date = filename_date.replaceAll("/", "").replaceAll(" ", "").replaceAll("-", "");
		outputFileName = _custom_no + "_" + _filename_hawb + "_" + _filename_date + ".xml";

		try {
			Files.createDirectories(new File(outputFilePath).toPath());
			File f = new File(outputFilePath + outputFileName);

			XMLOutputter xmlOutputter = new XMLOutputter(Format.getPrettyFormat());
			xmlOutputter.output(document, new FileOutputStream(f));

		} catch (Exception e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();

		} finally {

		}

		infoBox(outputFilePath + outputFileName + "  created.", "Job Done~");
	}

	

	private Document getImpDocument(String custom_no, String _payTerm) throws Exception {
		String linesKey = "";

		Document document = new Document();
		Element root = new Element("Header");
		document.setRootElement(root);

		ResultSet rs = getImpHeader(custom_no);
		// if(!rs.next())
		// throw new Exception("No data in Header file");

		// rs.beforeFirst();

		while (rs.next()) {
			linesKey = rs.getString("AUTO_SEQ");

			addElement(root, "Forwarder", "TEC");
			addElement(root, "TaxID", "86865094");
			addElement(root, "DocNo", rs.getString("DCL_DOC_NO"));

			// SimpleDateFormat sdfSource = new SimpleDateFormat( "yyyy/MM/dd"
			// );
			// String sDCL_DATE = rs.getString("DCL_DATE");
			//
			// // 民國年轉換西元年
			// Date d = sdfSource.parse(sDCL_DATE);
			// d.setYear( d.getYear()+1911 );
			//
			//
			// sdfSource = new SimpleDateFormat( "yyyy-MM-dd" );
			// sDCL_DATE = sdfSource.format(d);
			// filename_date = sDCL_DATE;
			// addElement(root, "DocDate", sDCL_DATE);

			Date dDCL_DATE = rs.getDate("DCL_DATE");
			SimpleDateFormat sdfSource = new SimpleDateFormat("yyyy-MM-dd");

			String sDCL_DATE = sdfSource.format(dDCL_DATE);
			filename_date = sDCL_DATE;
			addElement(root, "DocDate", sDCL_DATE);

			filename_hawb = rs.getString("HAWB");
			addElement(root, "HAWB", rs.getString("HAWB"));
			addElement(root, "MAWB", rs.getString("MAWB"));
			addElement(root, "DocType", rs.getString("DCL_DOC_TYPE"));
			addElement(root, "BuyerTaxID", rs.getString("SHPR_BAN_ID"));
			addElement(root, "BuyerName", rs.getString("SHPR_E_NAME"));
			addElement(root, "SellerTaxID", rs.getString("CNEE_BAN_ID"));
			addElement(root, "SellerName", rs.getString("CNEE_E_NAME"));
			addElement(root, "FromCountry", rs.getString("CNEE_COUNTRY_CODE"));
			addElement(root, "FromCity", rs.getString("FROM_CODE"));
			addElement(root, "TotalGrossWeight", rs.getString("DCL_GW"));

			// String sTotalPCS = "0";
			// if(rs.getString("DOC_CTN_UM") == "PLT")
			// sTotalPCS = rs.getString("TOT_CTN");

			addElement(root, "TotalPCS", rs.getString("TOT_CTN"));
			addElement(root, "Transport", "3");
			addElement(root, "DepartingPort", rs.getString("FROM_CODE"));
			addElement(root, "ImportAirline", rs.getString("FLY_NO"));
			addElement(root, "Currency", rs.getString("CURRENCY"));
			addElement(root, "FOBValue", rs.getString("FOB_AMT"));
			addElement(root, "Freight", rs.getString("FRT_AMT"));
			addElement(root, "InsuranceFee", rs.getString("INS_AMT"));

			double dADD_AMT = rs.getDouble("ADD_AMT");
			double dSUBTRACT_AMT = rs.getDouble("SUBTRACT_AMT");
			double dExpensesValue = dADD_AMT + dSUBTRACT_AMT;

			addElement(root, "ExpensesValue", String.valueOf(dExpensesValue)); // 看不懂，要說明
			addElement(root, "CIFValue", rs.getString("DOC_IMP_CIF_AMT"));
			addElement(root, "ExchangeRate", rs.getString("EXCHG_RATE"));
			addElement(root, "DutyMemo", rs.getString("DUTY_NO")); // 無測試值

			addElement(root, "DutyMemoDate", sDCL_DATE);

			addElement(root, "ImportDuty", rs.getString("IMPORT_TAX")); // 無測試值
			addElement(root, "TradePromotionFee", rs.getString("EX_TAX_AMT_1"));
			addElement(root, "ValueAddFee", rs.getString("SALEST_TAX"));
			addElement(root, "TotalAmount", rs.getString("DCL_AMT"));
			addElement(root, "BusinessTax", rs.getString("DOC_TAX_BASE"));
			addElement(root, "DelinquentFee", rs.getString("DELAY_AMT")); // 無測試值
			addElement(root, "ToWareHouse", "LK");
			addElement(root, "ModDutyPay", rs.getString("DCL_DOC_DUTY_VIA"));
			addElement(root, "ChargeableWeight", rs.getString("CHARGE_TON"));

			if (_payTerm.length() > 0)
				;
			else
				_payTerm = rs.getString("TERMS_SALES");
			addElement(root, "PaymentTerm", _payTerm);

			addElement(root, "TotalCarton", rs.getString("TOT_CTN"));

		}

		rs = getImpLines(linesKey);
		while (rs.next()) {
			Element linesElement = new Element("Lines");
			root.addContent(linesElement);

			String sDESCRIPTION = rs.getString("DESCRIPTION");
			String[] sParts = sDESCRIPTION.split("\r\n");
			String sPartCode = sParts[0];

			String sPartName = "";
			if (sParts.length >= 2)
				sPartName = sParts[1];
			else {
				int zeroA = Integer.parseInt("0A", 16);
				char zeroC = (char) zeroA;

				String[] split0A = sDESCRIPTION.split("" + zeroC);
				if (split0A.length >= 2)
					sPartName = split0A[1];
			}

			addElement(linesElement, "ItemNo", rs.getString("ITEM_NO"));
			addElement(linesElement, "PartCode", sPartCode);
			addElement(linesElement, "PartName", sPartName);
			addElement(linesElement, "COO", rs.getString("ORG_COUNTRY"));
			addElement(linesElement, "Price", rs.getString("DOC_UNIT_P"));
			addElement(linesElement, "NetWeight", rs.getString("NET_WT"));
			addElement(linesElement, "QTY", rs.getString("QTY"));
			addElement(linesElement, "DutyPayingValue", rs.getString("AFTER_TAX_AMT"));
			addElement(linesElement, "DutyRate", rs.getString("TAX_RATE_P"));
			addElement(linesElement, "Payment", rs.getString("TERMS"));
			addElement(linesElement, "HSCode", rs.getString("CCC_CODE"));

		}

		return document;
	}

	private ResultSet getImpHeader(String custom_no) throws Exception {
		Connection conn = connSQL();

		PreparedStatement ps = conn.prepareStatement("SELECT A.*, B.CHARGE_TON from DOC_H_I A "
				+ " left outer join DI_PICK B on A.AUTO_SEQ = B.AUTO_SEQ_HEAD" + " where A.DCL_DOC_NO = ? ");
		ps.setString(1, custom_no);

		return ps.executeQuery();
	}

	private ResultSet getImpLines(String custom_no) throws Exception {
		Connection conn = connSQL();

		PreparedStatement ps = conn
				.prepareStatement("SELECT * from DI_INVBD where AUTO_SEQ_HEAD = ? and ITEM_NO != '*' ");
		ps.setString(1, custom_no);

		return ps.executeQuery();
	}

	private Document getExpDocument(String custom_no, String _payTerm) throws Exception {

		String linesKey = "";

		Document document = new Document();
		Element root = new Element("Header");
		document.setRootElement(root);

		ResultSet rs = getExpHeader(custom_no);
		// if(!rs.next())
		// throw new Exception("No data in Header file");

		// rs.beforeFirst();

		while (rs.next()) {
			linesKey = rs.getString("AUTO_SEQ");

			addElement(root, "Forwarder", "TEC");
			addElement(root, "TaxID", "86865094");
			String sDocNo = rs.getString("DCL_DOC_NO").replaceAll("/66A /", "/66A/");

			System.out.println(rs.getString("DCL_DOC_NO"));
			System.out.println(sDocNo);
			System.out.println("sDocNo.leng " + sDocNo.length());
			addElement(root, "DocNo", sDocNo);

			// SimpleDateFormat sdfSource = new SimpleDateFormat( "yyyy/MM/dd"
			// );
			// String sDCL_DATE = rs.getString("DCL_DATE");
			//
			// // 民國年轉換西元年
			// Date d = sdfSource.parse(sDCL_DATE);
			// d.setYear( d.getYear()+1911 );
			//
			//
			// sdfSource = new SimpleDateFormat( "yyyy-MM-dd" );
			// sDCL_DATE = sdfSource.format(d);
			// filename_date = sDCL_DATE;
			// addElement(root, "DocDate", sDCL_DATE);

			Date dDCL_DATE = rs.getDate("DCL_DATE");
			SimpleDateFormat sdfSource = new SimpleDateFormat("yyyy-MM-dd");

			String sDCL_DATE = sdfSource.format(dDCL_DATE);
			filename_date = sDCL_DATE;
			addElement(root, "DocDate", sDCL_DATE);

			filename_hawb = rs.getString("HAWB");
			addElement(root, "HAWB", rs.getString("HAWB"));
			filename_mawb = rs.getString("MAWB");
			addElement(root, "MAWB", rs.getString("MAWB"));
			addElement(root, "DocType", rs.getString("DCL_DOC_TYPE"));

			addElement(root, "Transport", rs.getString("TRANS_VIA")); // ???
			addElement(root, "ExportAirline", rs.getString("FLY_NO")); // ???

			addElement(root, "SellerTaxID", rs.getString("SHPR_BAN_ID")); // ???
																			// SHPR_BAN_ID
			addElement(root, "SellerName", rs.getString("SHPR_E_NAME")); // ???
																			// SHPR_E_NAME

			addElement(root, "FromCountry", rs.getString("FROM_CODE")); // ???
																		// CNEE_COUNTRY_CODE

			addElement(root, "BuyerTaxID", rs.getString("CNEE_BAN_ID")); // ???
																			// CNEE_BAN_ID

			addElement(root, "BuyerName", rs.getString("CNEE_E_NAME")); // ???
																		// CNEE_E_NAME
			String s = rs.getString("CNEE_E_NAME");
			System.out.println(s); // shows

			addElement(root, "ToCountry", rs.getString("TO_CODE")); // ???
																	// CNEE_COUNTRY_CODE
			addElement(root, "Offset", rs.getString("APP_DUTY_REFUND")); // ???
																			// APP_DUTY_REFUND
			addElement(root, "Currency", rs.getString("CURRENCY")); // 報單上右上角的發票總金額的幣別
			addElement(root, "TotalInvoiceAmount", rs.getString("CAL_IP_TOT_ITEM_AMT")); // 報單上右上角的發票總金額
			addElement(root, "Freight", rs.getString("FRT_AMT"));
			addElement(root, "InsuranceFee", rs.getString("INS_AMT"));

			double dADD_AMT = rs.getDouble("ADD_AMT");
			double dSUBTRACT_AMT = rs.getDouble("SUBTRACT_AMT");
			double dExpensesValue = dADD_AMT + dSUBTRACT_AMT;
			addElement(root, "ExpensesValue", String.valueOf(dExpensesValue)); // 看不懂，要說明

			addElement(root, "FOBValue", rs.getString("FOB_AMT"));

			addElement(root, "InCotermCode", ""); // ???
			addElement(root, "ExchangeRate", rs.getString("EXCHG_RATE"));
			addElement(root, "TotalPCS", rs.getString("TOT_CTN"));

			addElement(root, "PackageDescription", ""); // ???
			addElement(root, "TotalGrossWeight", rs.getString("DCL_GW"));
			addElement(root, "ToWareHouse", "LK");
			addElement(root, "Marks", rs.getString("DOC_MARKS_DESC")); // ???
			System.out.println(rs.getString("DOC_MARKS_DESC"));
			addElement(root, "ContainerNo", ""); // ???
			addElement(root, "OtherDeclarations", rs.getString("DOC_OTR_DESC")); // ???
			System.out.println(rs.getString("DOC_OTR_DESC"));
			addElement(root, "ModeStatistics", ""); // ???

		}

		rs = getExpLines(linesKey);
		while (rs.next()) {
			Element linesElement = new Element("Lines");
			root.addContent(linesElement);
			
			String sDESCRIPTION = rs.getString("DESCRIPTION");
			String[] sParts = sDESCRIPTION.split("\r\n");
			String sPartCode = sParts[0];
			String sPartName = "";
			// String sPartSpec = "";

			if (sParts.length >= 2)
				sPartName = sParts[1];
			else {
				int zeroA = Integer.parseInt("0A", 16);
				char zeroC = (char) zeroA;

				String[] split0A = sDESCRIPTION.split("" + zeroC);
				if (split0A.length >= 2)
					sPartName = split0A[1];
			}

			addElement(linesElement, "ItemNo", rs.getString("ITEM_NO"));
			addElement(linesElement, "PartCode", sPartCode);
			addElement(linesElement, "PartName", sPartName);
			addElement(linesElement, "PartSpecification", rs.getString("TRADE_MARK")); // ???
			addElement(linesElement, "HSCode", rs.getString("CCC_CODE"));

			addElement(linesElement, "Price", rs.getString("DOC_UNIT_P"));
			addElement(linesElement, "NetWeight", rs.getString("NET_WT"));
			addElement(linesElement, "QTY", rs.getString("QTY"));
			addElement(linesElement, "FOBValue", rs.getString("DOC_FOB")); // V1
																			// =
																			// FOB_TWD;
																			// V2
																			// =
																			// DOC_FOB
			addElement(linesElement, "ModeStatistics", rs.getString("ST_MTD")); // ???

		}

		return document;
	}

	private ResultSet getExpHeader(String custom_no) throws Exception {
		Connection conn = connSQL();
		// 使用者希望可以不要輸入 報單號碼內的斜線與空白
		custom_no = custom_no.replaceAll("/", "").replaceAll(" ", "").replaceAll("-", "");
		PreparedStatement ps = conn.prepareStatement("SELECT A.*, B.CHARGE_TON from DOC_HEAD A "
				+ " left outer join DI_PICK B on A.AUTO_SEQ = B.AUTO_SEQ_HEAD"
				+ " where REPLACE(REPLACE(A.DCL_DOC_NO, ' ', ''), '/', '') = ? ");
		ps.setString(1, custom_no);

		return ps.executeQuery();
	}

	private ResultSet getExpLines(String custom_no) throws Exception {
		Connection conn = connSQL();

		// 使用者希望可以不要輸入 報單號碼內的斜線與空白
		custom_no = custom_no.replaceAll("/", "").replaceAll(" ", "").replaceAll("-", "");
		PreparedStatement ps = conn.prepareStatement(
				"SELECT * from DOCINVBD where REPLACE(REPLACE(AUTO_SEQ_HEAD, ' ', ''), '/', '') = ? and ITEM_NO != '*' ");
		ps.setString(1, custom_no);

		return ps.executeQuery();
	}

	private void addElement(Element parentELe, String childName, String text) {
		if (text == null)
			text = "";

		// if(childName.equals("DocNo"))
		// parentELe.addElement(childName).addAttribute(QName.get("space",
		// Namespace.XML_NAMESPACE),
		// "preserve").addText(text);
		// else
		parentELe.addContent(new Element(childName).setText(text));
	}

	
}
