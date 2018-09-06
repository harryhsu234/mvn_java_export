package OutputMethod;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Hashtable;

import org.apache.commons.codec.binary.Base64;
import org.jdom2.Element;
import org.jdom2.output.Format;
import org.jdom2.output.XMLOutputter;
import org.jdom2.Document;

import Util.Config;

public class AUO_Basic extends OutputCommon {
	String user = "";
	String email = "";
	String tel = "";
	String taxid = "";
	String method = "";
	String goodNo = "";
	String companyName = "TPEEX_TO_AU";
	String direct = "E"; 
	String broker = ""; 
	String type = "";
	String classCode = "Broker";
	String outputFilePath = "D:\\PDF\\";
	
	Connection conn = null;
	
	protected boolean saveDocTofile(Document doc) {
		boolean result = false;
		try {
			Files.createDirectories(new File(outputFilePath).toPath());
			File f = new File(outputFilePath + outputFileName);

			XMLOutputter xmlOutputter = new XMLOutputter(Format.getPrettyFormat());
			xmlOutputter.output(doc, new FileOutputStream(f));
			result = true;
		} catch (Exception e1) {
			e1.printStackTrace();
		}

		return result;
	}
	
	protected Element getFromRole(String type) {
		readConfigRole(type, "FROM");
		
		Element result = new Element("fromRole");
		Element node1 = new Element("PartnerRoleDescription");
		Element node11 = new Element("ContactInformation");
		Element node111 = new Element("contactName");
		Element node1111 = new Element("FreeFormText").setText(user);
		Element node112 = new Element("EmailAddress").setText(email);
		Element node113 = new Element("telephoneNumber");
		Element node1131 = new Element("CommunicationsNumber").setText(tel);
		Element node12 = new Element("GlobalPartnerRoleClassificationCode").setText("Transpotation Service Provider");
		Element node13 = new Element("PartnerDescription");
		Element node131 = new Element("BusinessDescription");
		Element node1311 = new Element("GlobalBusinessIdentifier").setText(taxid);
		Element node1312 = new Element("GlobalSupplyChainCode").setText("Electronic Components");
		Element node132 = new Element("GlobalPartnerClassificationCode").setText(classCode);

		node111.addContent(node1111);
		node113.addContent(node1131);
		node11.addContent(node111);
		node11.addContent(node112);
		node11.addContent(node113);

		node131.addContent(node1311);
		node131.addContent(node1312);

		node13.addContent(node131);
		node13.addContent(node132);

		node1.addContent(node11);
		node1.addContent(node12);
		node1.addContent(node13);

		result.addContent(node1);

		return result;
	}

	protected Element getToRole(String type) {
		readConfigRole(type, "TO");
		
		Element result = new Element("toRole");
		Element node1 = new Element("PartnerRoleDescription");
		Element node11 = new Element("ContactInformation");
		Element node111 = new Element("contactName");
		Element node1111 = new Element("FreeFormText").setText(user);
		Element node112 = new Element("EmailAddress").setText(email);
		Element node113 = new Element("telephoneNumber");
		Element node1131 = new Element("CommunicationsNumber").setText(tel);
		Element node12 = new Element("GlobalPartnerRoleClassificationCode").setText("In-transit Information User");
		Element node13 = new Element("PartnerDescription");
		Element node131 = new Element("BusinessDescription");
		Element node1311 = new Element("GlobalBusinessIdentifier").setText(taxid);
		Element node1312 = new Element("GlobalSupplyChainCode").setText("Electronic Components");
		Element node132 = new Element("GlobalPartnerClassificationCode").setText("End User");
		
		node111.addContent(node1111);
		node113.addContent(node1131);
		node11.addContent(node111);
		node11.addContent(node112);
		node11.addContent(node113);

		node131.addContent(node1311);
		node131.addContent(node1312);

		node13.addContent(node131);
		node13.addContent(node132);

		node1.addContent(node11);
		node1.addContent(node12);
		node1.addContent(node13);

		result.addContent(node1);

		return result;
	}

	private void readConfigRole(String type, String target) {
		String section = type + "_" + target;
		try {
			Hashtable<String,String> configHT = new Config().getConfig(section);
			
			user = configHT.get("contactName");
			email = configHT.get( "EmailAddress");
			tel = configHT.get("CommunicationsNumber");
			taxid = configHT.get("GlobalBusinessIdentifier");
			if (target.equals("TO")) {
				broker = configHT.get("Broker");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
	
	protected String convertChinaYear(String strDate) {
		String result = "";
		String[] ymd = strDate.split("/");
		System.out.println(strDate);
		int year = Integer.valueOf(ymd[0]).intValue();
		int months = Integer.valueOf(ymd[1]).intValue();
		int days = Integer.valueOf(ymd[2]).intValue();
		
		result = 1911 + year + "/" + months + "/" + days;
		System.out.println(result);
		return result;
	}
	
	protected String getString(Object obj) {
		String result = "";
		if (obj!=null)
			result = String.valueOf(obj);
		
		return result;
	}
	
	/**
	 * 取的預設值
	 * 
	 * @param value
	 * @param type
	 *            S:空白，N:預設NIL，Z:0
	 * @return
	 */
	protected String getDefault(String value, String type) {
		String result = "";

		if (value == null || value.trim().length() == 0) {
			switch (type) {
			case "S":
				result = " ";
				break;
			case "N":
				result = "NIL";
				break;
			case "Z":
				result = "0";
				break;
			}
		} else {
			result = value.trim();
		}

		return result;
	}
	
	/**
	 * 
	 * @param strDate
	 * @param addDays
	 * @return
	 */
	protected String addDate(String strDate, int addDays) {
		String result = strDate;

		try {
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
			Date date = sdf.parse(strDate);
			Calendar c = Calendar.getInstance();
			c.setTime(date);
			c.add(Calendar.DATE, addDays);

			result = sdf.format(c.getTime());
		} catch (Exception e) {
			e.printStackTrace();
		}

		return result;
	}

	protected void addElement(Element parentELe, String childName, String text) {
		if (text == null)
			text = "";

		parentELe.addContent(new Element(childName).setText(text));
	}
	

	protected String formatAmount(String value, int digit) {
		if (value == null || value.trim().length() <= 0) {
			value = "0";
		}
		String formatLayout = "%."+digit+"f";
		
		return value.format(formatLayout, Double.valueOf(value).doubleValue());
	}
	
	/**
	 * Get Broken PDF transfer to xml
	 * 
	 * @param custom_no
	 * @return
	 */
	protected Element getAttachment(String custom_no, String type,String code) {
		return getAttachment(custom_no, type, code, "");
		
//		Element result = new Element("Attachment");
//		String kind = "I";
//		String factory = "";
//		String formType = custom_no.substring(2, 4);
//		System.out.println(formType);
//		try {
//			Base64 base64 = new Base64();
//			if (type.equals("GLS_EXPORT")) {
//				kind = "E";
//			}
//			factory = getFactoryID(code);
//			
//			String inFile = "D:\\PDF\\"  + custom_no + ".PDF";
//			String outFile = "D:\\PDF\\"  + kind + "-" + factory + formType + custom_no + ".PDF";
//			String zipFile = kind + "-" + factory + formType + custom_no + ".ZIP";
//
//			File infile = new File(inFile);
//			File outfile = new File(outFile);
//			if (outfile.exists()) {
//				outfile.delete();
//			}
//			
//			if (copyFile(infile, outfile)) {
//				String zfilePath = "D:\\PDF\\" + zipFile;
//				ZipUtil zipUtil = new ZipUtil();
//				zipUtil.zip(outfile, zfilePath);
//				
//				File zipfile = new File(zfilePath);
//				byte[] data = read(zipfile);
//				addElement(result, "FileName", zipFile);
//				addElement(result, "AttachmentContent", base64.encodeAsString(data));
//			}
//
//		} catch (Exception e) {
//			result = null;
//			e.printStackTrace();
//		}
//
//		return result;
	}
	
	protected Element getAttachment(String custom_no, String type,String code, String outPrefix) {
		Element result = new Element("Attachment");
		String kind = "I";
		String factory = "";
		String formType = custom_no.substring(2, 4);
		System.out.println(formType);
		try {
			Base64 base64 = new Base64();
			if (type.equals("GLS_EXPORT")) {
				kind = "E";
			}
			factory = getFactoryID(code);
			if (outPrefix!=null && outPrefix.trim().length()>0) {
				formType=outPrefix;
			}
			String inFile = "D:\\PDF\\"  + custom_no + ".PDF";
			String outFile = "D:\\PDF\\"  + kind + "-" + factory + formType + custom_no + ".PDF";
			String zipFile =  kind + "-" + factory + formType + custom_no + ".ZIP";

			File infile = new File(inFile);
			File outfile = new File(outFile);
			if (outfile.exists()) {
				outfile.delete();
			}
			
			if (copyFile(infile, outfile)) {
				String zfilePath = "D:\\PDF\\" + zipFile;
				ZipUtil zipUtil = new ZipUtil();
				zipUtil.zip(outfile, zfilePath);
				
				File zipfile = new File(zfilePath);
				byte[] data = read(zipfile);
				addElement(result, "FileName", zipFile);
				addElement(result, "AttachmentContent", base64.encodeAsString(data));
				zipfile.delete();
			}
			outfile.delete();
		} catch (Exception e) {
			result = null;
			e.printStackTrace();
		}

		return result;
	}
	
	protected String getFactoryID(String code) {
		String result = "";
		try {
			Connection conn = OutputCommon.connSQL();
			
			String sql = "select AUO_FACTORY from D_CUST_AUO where COMP_ID='A' and CODE='"+code.trim()+"' ";
			Statement stat = conn.createStatement();
			ResultSet rs = stat.executeQuery(sql);
			if (rs.next()) {
				result = rs.getString(1);
			} else {
				infoBox(code + ": 未設定AUO廠區 ","警告");
			}
			
			conn.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return result;
	}
	
	protected boolean copyFile(File sourceFile, File targetFile) throws IOException {
		boolean result = true;

		try {
			InputStream inStream = new FileInputStream(sourceFile);
			OutputStream outStream = new FileOutputStream(targetFile);

			byte[] buffer = new byte[1024];

			int length;
			// copy the file content in bytes
			while ((length = inStream.read(buffer)) > 0) {
				outStream.write(buffer, 0, length);
			}

			inStream.close();
			outStream.close();

			System.out.println("File is copied successful!");
		} catch (IOException e) {
			e.printStackTrace();
			result = false;
			throw new IOException();
		}
		return result;
	}

	protected byte[] read(File file) {
		ByteArrayOutputStream ous = null;
		InputStream ios = null;

		try {
			byte[] buffer = new byte[4096];
			ous = new ByteArrayOutputStream();
			ios = new FileInputStream(file);
			int read = 0;
			while ((read = ios.read(buffer)) != -1) {
				ous.write(buffer, 0, read);
			}

			if (ous != null)
				ous.close();
			if (ios != null)
				ios.close();

		} catch (Exception e) {
			e.printStackTrace();
		}

		return ous.toByteArray();
	}

	/**
	 * 限制字串長度
	 * @param value 
	 * @param lengths
	 * @return String 不超過長度lengths
	 */
	public String limitString(String value,int lengths) {
		String result = value.replaceAll("\n", "").replaceAll("\r", "");
		
		if (result !=null && result.length()>lengths) {
			result = result.substring(0, lengths);
		}
		
		return result;
	}
}
