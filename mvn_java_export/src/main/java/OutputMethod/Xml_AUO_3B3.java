package OutputMethod;

import java.text.SimpleDateFormat;
import java.util.Date;

import org.jdom2.Document;
import org.jdom2.Element;

public class Xml_AUO_3B3 extends AUO_Basic {
	String root = "Pip3B3ShipmentStatusNotification";
	String goodNo = "";
	String port = "";
	String transType = "";
	String qty = "";
	String amount = "";
	String rlDate = "";
	String id = "";
	String name = "";
	
	public static void main(String[] args) {
		String type = "GLS_IMPORT";
		
		try {
			// wpg.readXML();
			Xml_AUO_3B3 wpg = new Xml_AUO_3B3();
			wpg.getXML();
		} catch (Exception e) {
			e.printStackTrace();
			infoBox(e.getMessage(), "ERROR!!");
		}
	}
	
	public void getXML() throws Exception {
		type = "GLS_EXPORT"; // TODO delete
		classCode = "Freight Forwarder";
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd'T'HHmmss.SSS'Z'");

		String _filename_date = sdf.format(new Date());
		
		// 3B3_TPEEX_THY1714196_20171003T164547.638Z
		Document document = new Document();
		Element root = new Element(this.root);
		document.addContent(root);
		
		root.addContent(getFromRole(type));
		
		addElement(root, "GlobalDocumentFunctionCode", "Request");
		
		Element ss = new Element("ShipmentStatus");
		
		Element customs = new Element("CustomsInformation");
		addElement(customs, "GlobalPortOfDeclarationCode", port);
		addElement(customs, "GlobalPortOfEntryCode", port);
		ss.addContent(customs);
		
		Element op = new Element("OriginatingPartner");
		Element sf = new Element("shipFrom");
		Element pd = new Element("PartnerDescription");
		
		Element bd = new Element("BusinessDescription");
		addElement(bd, "GlobalBusinessIdentifier", id);
		pd.addContent(bd);
		
		Element ci = new Element("ContactInformation");
		Element cn = new Element("contactName");
		addElement(cn, "FreeFormText", name);
		ci.addContent(cn);
		
		Element pa = new Element("PhysicalAddress");
		Element ct = new Element("cityName");
		addElement(ct, "FreeFormText", port);
		pa.addContent(ct);
		ci.addContent(pa);
		pd.addContent(ci);
		sf.addContent(pd);
		op.addContent(sf);
		ss.addContent(op);
		
		Element pdir = new Element("ProofOfDeliveryInformationResource");
		Element nosc = new Element("numberOfShippingContainers");
		addElement(nosc, "CountableAmount", amount);
		pdir.addContent(nosc);
		pdir.addContent(pa.clone());
		ss.addContent(pdir);
		
		Element sm = new Element("Shipment");
		addElement(sm, "GlobalShipmentModeCode", transType);
		sm.addContent( new Element("numberOfShippingContainers").addContent(new Element("CountableAmount").setText(qty)));
		sm.addContent( new Element("shipmentIdentifier").addContent(new Element("ProprietaryReferenceIdentifier").setText("NIL|NIL")));
		sm.addContent( new Element("ShippingContainer").addContent(new Element("shippingContainerIdentifier").addContent(new Element("ProprietarySerialIdentifier").setText(goodNo))));
		ss.addContent(sm);
		
		Element ssd = new Element("ShipmentStatusDetail");
		addElement(ssd, "GlobalShipmentStatusReportingLevelCode", "Current");
		Element sl = new Element("shipmentLocation");
		Element pa1 = new Element("PhysicalAddress");
		pa1.addContent(new Element("cityName").addContent(new Element("FreeFormText").setText("TW")));
		pa1.addContent(new Element("regionName").addContent(new Element("FreeFormText").setText("GMT+8")));
		sl.addContent(pa1);
		ssd.addContent(sl);
		
		ssd.addContent(new Element("shipmentStatusDateTime").addContent(new Element("DateTimeStamp").setText("NIL")));
		ssd.addContent(new Element("shipmentStatusDescription").addContent(new Element("FreeFormText").setText("NIL")));
		
		Element sse = new Element("ShipmentStatusEvent");
		sse.addContent(new Element("DateTimeStamp").setText(rlDate));
		sse.addContent(new Element("GlobalShipmentStatusCode").setText("Customs Release"));
		ssd.addContent(sse);
		ss.addContent(ssd);
		
		root.addContent(ss);
		
		root.addContent(new Element("thisDocumentGenerationDateTime").addContent(new Element("DateTimeStamp").setText(_filename_date)));
		root.addContent(new Element("thisDocumentIdentifier").addContent(new Element("ProprietaryDocumentIdentifier").setText(_filename_date)));
		
		root.addContent(getToRole(type));
		
		String outputFilePath = "D:\\PDF\\";
		outputFileName = "";
	
		String midleName = goodNo;
		outputFileName = "3B3_" + companyName + "_" + _filename_date + "_" + midleName + ".xml";
		System.out.println(outputFileName);
		
		saveDocTofile(document);

	}

	public void setGoodNo(String goodNo) {
		this.goodNo = goodNo;
	}

	public void setPort(String port) {
		this.port = port;
	}

	public void setTransType(String transType) {
		this.transType = transType;
	}

	public void setQty(String qty) {
		this.qty = qty;
	}

	public void setAmount(String amount) {
		this.amount = amount;
	}

	public void setRlDate(String rlDate) {
		this.rlDate = rlDate;
	}

	public void setName(String name) {
		this.name = name;
	}

	public void setId(String id) {
		this.id = id;
	}
	
}
