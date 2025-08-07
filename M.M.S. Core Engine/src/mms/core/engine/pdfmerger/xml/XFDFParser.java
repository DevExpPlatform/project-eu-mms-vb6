package mms.core.engine.pdfmerger.xml;

import java.io.File;
import java.io.StringReader;
import java.util.ArrayList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;

import mms.core.engine.pdfmerger.commons.PDFCutterInfo;
import mms.core.engine.pdfmerger.commons.PDFCutterInfo.fieldTypes;

public class XFDFParser {

	private String 				 	 errMsg	         = "";
	private String 				 	 prjBasePath     = "";
	private String               	 pdfFileName     = "";
	private String               	 pdfStandard	 = "";
	private String 				 	 pdfColorProfile = "";

	private File xfdfFile = null;
	private String xfdf = null;

	public XFDFParser(File xfdxFile) {
		this.xfdfFile = xfdxFile;
	}

	public XFDFParser(String xfdf) {
		this.xfdf = xfdf;
	}

	private ArrayList<PDFCutterInfo> getFields() {
		DocumentBuilderFactory   dbf        = DocumentBuilderFactory.newInstance();
		ArrayList<PDFCutterInfo> rValue = new ArrayList<PDFCutterInfo>();

		try {
			DocumentBuilder db 		= dbf.newDocumentBuilder();
			Document		xmlDoc  = null;

			if (xfdfFile == null) {
				xmlDoc = db.parse(new InputSource(new StringReader(xfdf)));
			} else {
				xmlDoc = db.parse(xfdfFile);
			}

			xmlDoc.getDocumentElement().normalize();

			/*
			 *  Get PDF file reference
			 */
			pdfFileName = xmlDoc.getElementsByTagName("f").item(0).getAttributes().getNamedItem("href").getNodeValue();

			/*
			 * Get Extra Params
			 */
			NodeList extraParamsNode    = xmlDoc.getElementsByTagName("extraparams");
			Element  extraParamsElement = (Element) extraParamsNode.item(0);

			if (extraParamsElement != null) {
				prjBasePath = extraParamsElement.getAttribute("prjbasepath");

				if (extraParamsElement.hasAttribute("pdfstandard")) 
					pdfStandard = extraParamsElement.getAttribute("pdfstandard");

				if (extraParamsElement.hasAttribute("colorprofile")) 
					pdfColorProfile = prjBasePath + "Extras" + System.getProperty("file.separator") + extraParamsElement.getAttribute("colorprofile");
			}

			/*
			 * Get all fields
			 */
			NodeList xfdfFieldNodes = xmlDoc.getElementsByTagName("field");

			for (int s = 0; s < xfdfFieldNodes.getLength(); s++) {
				PDFCutterInfo xfdfField    = new PDFCutterInfo();
				Element   	  xfdfNodeType = (Element) xfdfFieldNodes.item(s);

				xfdfField.elementName = xfdfNodeType.getAttributes().getNamedItem("name").getNodeValue();

				if (xfdfNodeType.hasAttribute("type")) {
					Node fieldType = xfdfNodeType.getAttributes().getNamedItem("type");

					if (fieldType.getNodeValue().equals("img")) {
						xfdfField.elementType  = fieldTypes.IMAGE;
						xfdfField.elementParam = xfdfNodeType.getAttributes().getNamedItem("href").getNodeValue();
					} else if(fieldType.getNodeValue().equals("barcode")) {
						xfdfField.elementType  = fieldTypes.BARCODE;
						xfdfField.elementParam = xfdfNodeType.getAttributes().getNamedItem("symbology").getNodeValue();
					} else if(fieldType.getNodeValue().equals("ml")) {
						xfdfField.elementType  = fieldTypes.XML;
						xfdfField.elementParam = "ml";
						xfdfField.elementNode  = (Node) xfdfNodeType.getElementsByTagName("value").item(0);
					}
				} else {
					xfdfField.elementType  = fieldTypes.TEXT;
				}

				Node xfdfNodeValue = ((Element) xfdfNodeType).getElementsByTagName("value").item(0).getChildNodes().item(0);

				if (xfdfNodeValue != null) 
					xfdfField.elementValue = xfdfNodeValue.getNodeValue();

				rValue.add(xfdfField);
			}
		} catch (Exception e) {
			rValue.clear();

			errMsg      = e.getMessage();
			pdfFileName = "";
			prjBasePath = "";
		}
		
		return rValue;
	}

	public String getColorProfile() {
		return this.pdfColorProfile;
	}

	public String getErrorMessage() {
		return this.errMsg;
	}

	public String getPDFFileName() {
		return this.pdfFileName;
	}

	public String getPDFStardard() {
		return this.pdfStandard;
	}

	public String getPrjBasePath() {
		return this.prjBasePath;
	}

	public ArrayList<PDFCutterInfo> getXFDFFields() {
		return this.getFields();
	}

}
