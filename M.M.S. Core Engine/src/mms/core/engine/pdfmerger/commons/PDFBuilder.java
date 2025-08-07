package mms.core.engine.pdfmerger.commons;

import java.awt.FontFormatException;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.Map;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.Utilities;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfAnnotation;
import com.itextpdf.text.pdf.PdfAppearance;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfCopyFields;
import com.itextpdf.text.pdf.PdfFormField;
import com.itextpdf.text.pdf.PdfName;
import com.itextpdf.text.pdf.PdfNumber;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.itextpdf.text.pdf.PdfStream;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.pdf.TextField;

import mms.core.engine.packer.PDFPacker;
import multivalent.ParseException;
import multivalent.std.adaptor.pdf.PDFReader;

public class PDFBuilder {

	private NodeList 			nodesField			= null;
	private String 				pdfTemplateFileName	= null;
	private String 				pdfTemplateId		= null;
	private Map<String, String> pdfTemplates	    = null;
	private String[] 			pdfTemplatesIndexes = null;
	private String 				pdfVersion		    = "";
	private String 				prjMMSTemplatesDir  = "";

	private void gerMergedFields(PdfStamper pdfStamper) throws DocumentException, IOException {
		pdfStamper.setFullCompression();

		if ((this.nodesField != null) && (this.nodesField.getLength() > 0)) {
			Rectangle 	 fieldCoords 	  = null;
			BaseFont 	 fieldDefFont 	  = getFont("");
			float 	 	 fieldDefFontSize = 8.0f;
			int 	  	 pdfPage 	  	  = 0;
			PdfWriter 	 pdfWriter  	  = pdfStamper.getWriter();

			pdfWriter.setCompressionLevel(PdfStream.BEST_COMPRESSION);
			pdfWriter.getAcroForm().setNeedAppearances(false);

			for (int i = 0; i < nodesField.getLength(); i++) {
				Element  nodeField 		 = (Element) nodesField.item(i);
				int   	 fieldAlignment  = com.itextpdf.text.Element.ALIGN_LEFT;
				int 	 fieldCharLength = 0;
				boolean  fieldCombed     = false;
				BaseFont fieldFont     	 = fieldDefFont;
				float 	 fieldFontSize   = fieldDefFontSize;
				boolean  fieldMultiLine  = false;
				String   fieldName 		 = nodeField.getAttributes().getNamedItem("id").getNodeValue();
				int 	 fieldRotation	 = 0;
				NodeList nodesFieldProps = nodeField.getElementsByTagName("properties");

				if (nodeField.hasAttribute("multiline")) 
					fieldMultiLine = (nodeField.getAttributes().getNamedItem("multiline").getNodeValue().equals("1"));

				if (nodesFieldProps.getLength() == 1) {
					Element nodeFieldProps = (Element) nodesFieldProps.item(0);

					pdfPage     = Integer.valueOf(nodeFieldProps.getAttributes().getNamedItem("pageid").getNodeValue());
					fieldCoords = getBoundingBox(nodeFieldProps.getAttributes().getNamedItem("coords").getNodeValue().split(","), pdfStamper.getReader().getPageSize(pdfPage));

					if (nodeFieldProps.hasAttribute("fontname")) 
						fieldFont = getFont(nodeFieldProps.getAttributes().getNamedItem("fontname").getNodeValue());

					if (nodeFieldProps.hasAttribute("fontsize")) 
						fieldFontSize = Float.valueOf(nodeFieldProps.getAttributes().getNamedItem("fontsize").getNodeValue());

					if (nodeFieldProps.hasAttribute("alignment")) {
						String align = nodeFieldProps.getAttributes().getNamedItem("alignment").getNodeValue();

						if (align.equals("center")) {
							fieldAlignment = com.itextpdf.text.Element.ALIGN_CENTER;
						} else if (align.equals("right")) {
							fieldAlignment = com.itextpdf.text.Element.ALIGN_RIGHT;
						}
					}

					if (nodeFieldProps.hasAttribute("rotation")) 
						fieldRotation = Integer.valueOf(nodeFieldProps.getAttributes().getNamedItem("rotation").getNodeValue());

					if (nodeFieldProps.hasAttribute("maxCharsLength"))
						fieldCharLength = Integer.valueOf(nodeFieldProps.getAttributes().getNamedItem("maxCharsLength").getNodeValue()); 

					if (nodeFieldProps.hasAttribute("comb")) {
						fieldCharLength = Integer.valueOf(nodeFieldProps.getAttributes().getNamedItem("comb").getNodeValue()); 
						fieldCombed     = true;
					}

					TextField myTextField = new TextField(pdfWriter, fieldCoords, fieldName);
					myTextField.setAlignment(fieldAlignment);
					myTextField.setExtensionFont(fieldDefFont);
					myTextField.setFont(fieldFont);
					myTextField.setFontSize(fieldFontSize);
					myTextField.setRotation(fieldRotation);

					if (fieldCharLength > 0) 
						myTextField.setMaxCharacterLength(fieldCharLength);

					if (fieldCombed) 
						myTextField.setOptions(TextField.COMB);
					
					if (fieldMultiLine)
						myTextField.setOptions(TextField.MULTILINE);

					pdfStamper.addAnnotation(myTextField.getTextField(), pdfPage);
				} else {
					PdfFormField   parentField = PdfFormField.createTextField(pdfWriter, false, false, 0);
					PdfContentByte pdfCanvas   = new PdfContentByte(pdfWriter);

					parentField.setFieldName(fieldName);

					if (fieldMultiLine)
						parentField.setFieldFlags(PdfFormField.FF_MULTILINE);

					for (int j = 0; j < nodesFieldProps.getLength(); j++) {
						Element nodeFieldProps = (Element) nodesFieldProps.item(j);

						pdfPage        = Integer.valueOf(nodeFieldProps.getAttributes().getNamedItem("pageid").getNodeValue());
						fieldAlignment = PdfFormField.Q_LEFT;
						fieldCoords    = getBoundingBox(nodeFieldProps.getAttributes().getNamedItem("coords").getNodeValue().split(","), pdfStamper.getReader().getPageSize(pdfPage));
						fieldFont      = fieldDefFont;
						fieldFontSize  = fieldDefFontSize;

						if (nodeFieldProps.hasAttribute("rotation")) 
							fieldRotation = Integer.valueOf(nodeFieldProps.getAttributes().getNamedItem("rotation").getNodeValue());

						if (nodeFieldProps.hasAttribute("fontname")) 
							fieldFont = getFont(nodeFieldProps.getAttributes().getNamedItem("fontname").getNodeValue());

						if (nodeFieldProps.hasAttribute("fontsize")) 
							fieldFontSize = Float.valueOf(nodeFieldProps.getAttributes().getNamedItem("fontsize").getNodeValue());

						if (nodeFieldProps.hasAttribute("alignment")) {
							String align = nodeFieldProps.getAttributes().getNamedItem("alignment").getNodeValue();

							if (align.equals("center")) {
								fieldAlignment = PdfFormField.Q_CENTER;
							} else if (align.equals("right")) {
								fieldAlignment = PdfFormField.Q_RIGHT;
							}
						}

						if (nodeFieldProps.hasAttribute("maxCharsLength"))
							fieldCharLength = Integer.valueOf(nodeFieldProps.getAttributes().getNamedItem("maxCharsLength").getNodeValue()); 

						if (nodeFieldProps.hasAttribute("comb")) {
							fieldCharLength = Integer.valueOf(nodeFieldProps.getAttributes().getNamedItem("comb").getNodeValue()); 
							fieldCombed     = true;
						}

						PdfAppearance pdfAP = pdfCanvas.createAppearance(fieldCoords.getWidth(), fieldCoords.getHeight());
						pdfAP.setFontAndSize(fieldFont, fieldFontSize);
						
						PdfFormField kidField = PdfFormField.createEmpty(pdfWriter);
						kidField.setDefaultAppearanceString(pdfAP);
						kidField.setFlags(PdfAnnotation.FLAGS_PRINT);
						kidField.setPlaceInPage(pdfPage);
						kidField.setMKRotation(fieldRotation);
						kidField.setQuadding(fieldAlignment);
						kidField.setWidget(fieldCoords, null);
						
						if (fieldCharLength > 0) 
							kidField.put(PdfName.MAXLEN, new PdfNumber(fieldCharLength));

						if (fieldCombed) 
							kidField.setFieldFlags(PdfFormField.FF_COMB);
						
						parentField.addKid(kidField);

						pdfStamper.addAnnotation(kidField, pdfPage);
					}

					pdfStamper.addAnnotation(parentField, 1);
				}
			}
		}

		pdfStamper.close();
	}
	
	private Rectangle getBoundingBox(String[] strArray, Rectangle rectangle) {
		float 	  floatArray[] = new float[strArray.length];
		Rectangle rValue 	   = new Rectangle(0, 0, 0, 0);

		for (int i = 0; i < strArray.length; i++) 
			floatArray[i] = Utilities.millimetersToPoints(Float.valueOf(strArray[i].trim()));

		rValue.setLeft(floatArray[0]);
		rValue.setTop((rectangle.getHeight() - floatArray[1]));
		rValue.setRight((floatArray[0] + floatArray[2]));
		rValue.setBottom((rectangle.getHeight() - (floatArray[1] + floatArray[3])));

		return rValue;
	}
	
	public BaseFont getFont(String fontName) throws DocumentException, IOException {
		BaseFont rValue = null;

		if (fontName.equals("")) {
			rValue = BaseFont.createFont("Helvetica", BaseFont.WINANSI, BaseFont.EMBEDDED);
		} else if (fontName.equals("courier")) {
			rValue = BaseFont.createFont("Courier", BaseFont.WINANSI, BaseFont.EMBEDDED);
		} else if (fontName.equals("courier_bold")) {
			rValue = BaseFont.createFont("Courier-Bold", BaseFont.WINANSI, BaseFont.EMBEDDED);
		} else if (fontName.equals("courier_oblique")) {
			rValue = BaseFont.createFont("Courier-Oblique", BaseFont.WINANSI, BaseFont.EMBEDDED);
		} else if (fontName.equals("courier_boldoblique")) {
			rValue = BaseFont.createFont("Courier-BoldOblique", BaseFont.WINANSI, BaseFont.EMBEDDED);
		} else if (fontName.equals("helvetica")) {
			rValue = BaseFont.createFont("Helvetica", BaseFont.WINANSI, BaseFont.EMBEDDED);
		} else if (fontName.equals("helvetica_bold")) {
			rValue = BaseFont.createFont("Helvetica-Bold", BaseFont.WINANSI, BaseFont.EMBEDDED);
		} else if (fontName.equals("helvetica_oblique")) {
			rValue = BaseFont.createFont("Helvetica-Oblique", BaseFont.WINANSI, BaseFont.EMBEDDED);
		} else if (fontName.equals("helvetica_boldoblique")) {
			rValue = BaseFont.createFont("Helvetica-BoldOblique", BaseFont.WINANSI, BaseFont.EMBEDDED);
		} else if (fontName.equals("symbol")) {
			rValue = BaseFont.createFont("Symbol", BaseFont.WINANSI, BaseFont.EMBEDDED);
		} else if (fontName.equals("times_roman")) {
			rValue = BaseFont.createFont("Times-Roman", BaseFont.WINANSI, BaseFont.EMBEDDED);
		} else if (fontName.equals("times_bold")) {
			rValue = BaseFont.createFont("Times-Bold", BaseFont.WINANSI, BaseFont.EMBEDDED);
		} else if (fontName.equals("times_italic")) {
			rValue = BaseFont.createFont("Times-Italic", BaseFont.WINANSI, BaseFont.EMBEDDED);
		} else if (fontName.equals("times_bolditalic")) {
			rValue = BaseFont.createFont("Times-BoldItalic", BaseFont.WINANSI, BaseFont.EMBEDDED);
		} else if (fontName.equals("zapfdingbats")) {
			rValue = BaseFont.createFont("ZapfDingbats", BaseFont.WINANSI, BaseFont.EMBEDDED);
		} else {
			rValue = BaseFont.createFont(prjMMSTemplatesDir + "Fonts/" + fontName, BaseFont.WINANSI, BaseFont.EMBEDDED);
		}

		return rValue;
	}

	private ByteArrayOutputStream getMergedTemplates(String[] pdfFiles, String pdfVersion) throws FileNotFoundException, IOException, DocumentException {
		ByteArrayOutputStream pdfBAOS    = new ByteArrayOutputStream();
		PdfCopyFields 		  pdfMerged  = new PdfCopyFields(pdfBAOS);

		for (int i = 0; i < pdfFiles.length; i++) {
			PdfReader pdfReader = new PdfReader(new FileInputStream(this.prjMMSTemplatesDir + "Builds/" + pdfVersion + "/" + this.pdfTemplates.get(pdfFiles[i])));
			pdfMerged.addDocument(pdfReader);
			pdfReader.close();
		}

		pdfMerged.close();
		
		return pdfBAOS;
	}

	public byte[] getTemplateBytes() throws DocumentException, IOException, ParseException, FontFormatException {
		PdfReader pdfReader = new PdfReader(getMergedTemplates(this.pdfTemplatesIndexes, this.pdfVersion).toByteArray());
		pdfReader.getCatalog().remove(PdfName.METADATA);
		pdfReader.removeUnusedObjects();

		ByteArrayOutputStream pdfBAOS = new ByteArrayOutputStream();

		gerMergedFields(new PdfStamper(pdfReader, pdfBAOS));

		pdfReader.close();

		PDFPacker dstPDF = new PDFPacker(pdfBAOS.toByteArray());
		dstPDF.setQuiet(true);

		return dstPDF.writeBytes();
	}
	
	public String getTemplateFS() throws FileNotFoundException, IOException, DocumentException, ParseException, FontFormatException {
		String pdfOutFileName = this.prjMMSTemplatesDir + this.pdfVersion + "_" + this.pdfTemplateFileName + ".PDF";

		if (new File(pdfOutFileName).exists())
			return pdfOutFileName;

		PdfReader pdfReader = new PdfReader(getMergedTemplates(this.pdfTemplatesIndexes, this.pdfVersion).toByteArray());
		pdfReader.getCatalog().remove(PdfName.METADATA);
		pdfReader.removeUnusedObjects();

		OutputStream pdfOut = new FileOutputStream(pdfOutFileName);

		gerMergedFields(new PdfStamper(pdfReader, pdfOut));
		
		pdfReader.close();

		pdfOut.flush();
		pdfOut.close();
		
		File      srcPDF = new File(pdfOutFileName);
		
		PDFPacker dstPDF = new PDFPacker(new PDFReader(srcPDF));
		dstPDF.setQuiet(true);
		dstPDF.writeFile(srcPDF);
		dstPDF = null;

		return pdfOutFileName;
	}

	public String getTemplateKey() {
		return this.pdfTemplateId;
	}

	public void setTemplates(Map<String, String> pdfTemplates) {
		this.pdfTemplates = pdfTemplates;
	}

	public void setTemplatesPath(String prjMMSTemplatesDir) {
		this.prjMMSTemplatesDir = prjMMSTemplatesDir;
	}

	public void setXMLData(String xmlTemplates) throws ParserConfigurationException, UnsupportedEncodingException, SAXException, IOException {
		DocumentBuilderFactory dbf         = DocumentBuilderFactory.newInstance();
		DocumentBuilder 	   db          = dbf.newDocumentBuilder();
		String 				   elementNode = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><root>" + xmlTemplates + "</root>";
		Document 		       xmlDoc      = db.parse(new InputSource(new ByteArrayInputStream(elementNode.getBytes("UTF-8"))));
		Element 			   xmlElement  = (Element) xmlDoc.getElementsByTagName("template").item(0);
		
		this.pdfVersion 	     = xmlElement.getAttribute("version");
		this.pdfTemplateFileName = xmlElement.getAttribute("filename");
		this.pdfTemplateId       = this.pdfTemplateFileName + "_" + xmlElement.getAttribute("indexes");
		this.pdfTemplatesIndexes = xmlElement.getAttribute("indexes").split(",");
		this.nodesField          = xmlDoc.getElementsByTagName("field");
	}

}