package mms.core.engine.pdfmerger.commons;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.text.DecimalFormat;
import java.util.Hashtable;
import java.util.List;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.w3c.dom.Document;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.AcroFields;
import com.itextpdf.text.pdf.AcroFields.FieldPosition;
import com.itextpdf.text.pdf.AcroFields.Item;
import com.itextpdf.text.pdf.Barcode;
import com.itextpdf.text.pdf.Barcode128;
import com.itextpdf.text.pdf.Barcode39;
import com.itextpdf.text.pdf.BarcodeDatamatrix;
import com.itextpdf.text.pdf.BarcodeInter25;
import com.itextpdf.text.pdf.BarcodeQRCode;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfDictionary;
import com.itextpdf.text.pdf.PdfName;
import com.itextpdf.text.pdf.PdfNumber;
import com.itextpdf.text.pdf.PdfStamper;
import com.itextpdf.text.pdf.TextField;
import com.itextpdf.text.pdf.qrcode.EncodeHintType;
import com.itextpdf.text.pdf.qrcode.ErrorCorrectionLevel;

import mms.core.engine.pdfmerger.barcodes.OMRPBB;
import mms.core.engine.sdf.SDFParser;
import uk.org.okapibarcode.backend.DataMatrix;
import uk.org.okapibarcode.backend.DataMatrix.ForceMode;
import uk.org.okapibarcode.gui.OkapiUI;
import uk.org.okapibarcode.output.Java2DRenderer;

public class PDFManager {

	public enum 			ManagerMode 		{ DBA, XML }

	private PDFCutterInfo 	dataCutter			= null;
	private AcroFields 		pdfAcroFields 		= null;
	private PdfStamper 		pdfFilled 	 		= null;
	private ManagerMode 	pdfManagerMode		= ManagerMode.DBA;
	private String 			prjMMSTemplatesDir	= "";

	public PDFManager(PdfStamper pdfFilled) {
		this.pdfFilled     = pdfFilled;
		this.pdfAcroFields = this.pdfFilled.getAcroFields();
	}

	public void close(boolean setFullCompression) throws DocumentException, IOException {
		pdfFilled.setFreeTextFlattening(true);
		pdfFilled.setFormFlattening(true);

		if (setFullCompression)
			pdfFilled.setFullCompression();

		pdfFilled.close();
		pdfFilled = null;

//    	if (!xfdfPDFStandard.equals("")) {
//			PDFConverter pdfConverter = new PDFConverter();
//			pdfConverter.setPDFStandard(xfdfPDFStandard);
//			pdfConverter.setPDFColorProfile(xfdfColorProfile);
//			pdfConverter.setInPDFFileName(outPDFFName);
//			pdfConverter.setOutPDFFileName(args[1]);
//			pdfConverter.convert();
//		}
	}

	public void execute(PDFCutterInfo xfdfField) throws DocumentException, IOException, ParserConfigurationException, SAXException {
		if (this.pdfManagerMode.equals(ManagerMode.DBA) && (xfdfField.elementValue == null))
			return;

		this.dataCutter = xfdfField;

		switch (xfdfField.elementType) {
		case BARCODE:
			if (this.pdfManagerMode.equals(ManagerMode.DBA)) {
				String tmpElementName  = this.dataCutter.elementName;
				String tmpElementParam = this.dataCutter.elementParam;

				if (this.dataCutter.elementSplitParam.equals("")) {
					if (this.dataCutter.elementName.indexOf("_BC") == -1)
						this.dataCutter.elementName += "_BC";

					String[] elementParams = this.dataCutter.elementParam.split("\\|");

					if (elementParams.length > 1) {
						String elementName = this.dataCutter.elementName;

						for (int i = 0; i < elementParams.length; i++) {
							this.dataCutter.elementName  = elementName + new DecimalFormat("##00").format(i);
							this.dataCutter.elementParam = elementParams[i];

							setBarCode();
						}
					} else {
						setBarCode();
					}

					this.dataCutter.elementName  = tmpElementName;
					this.dataCutter.elementParam = tmpElementParam;
				} else {
					String   tmpElementValue = this.dataCutter.elementValue;
					String[] elementValues 	 = this.dataCutter.elementValue.split("\\" + this.dataCutter.elementSplitParam);

					for (int i = 0; i < elementValues.length; i++) {
						this.dataCutter.elementName = this.dataCutter.elementName.replace("XXX", new DecimalFormat("###000").format(i + 1));

						if (this.dataCutter.elementName.indexOf("_BC") == -1)
							this.dataCutter.elementName += "_BC";

						String[] elementParams = this.dataCutter.elementParam.split("\\|");

						if (elementParams.length > 1) {
							String elementName = this.dataCutter.elementName;

							for (int j = 0; j < elementParams.length; j++) {
								this.dataCutter.elementName  = elementName + new DecimalFormat("##00").format(j);
								this.dataCutter.elementParam = elementParams[j];
								this.dataCutter.elementValue = elementValues[i];

								setBarCode();
							}
						} else {
							if (elementValues.length > 0)
								this.dataCutter.elementValue = elementValues[i];

							setBarCode();
						}

						this.dataCutter.elementName  = tmpElementName;
						this.dataCutter.elementParam = tmpElementParam;
						this.dataCutter.elementValue = tmpElementValue;
					}
				}
			} else {
				setBarCode();
			}

			break;
		case IMAGE:
			setImage();

			break;
		case TEXT:
			setText();

			break;
		case XML:
			setXML();

			break;
		}
	}

	private int getFieldRotation(Item pdfDictionary){
		if (pdfDictionary != null) {
			PdfDictionary widgetDict = pdfDictionary.getWidget(0);

			if (widgetDict != null) {
				PdfDictionary mkDict = widgetDict.getAsDict(PdfName.MK);

				if (mkDict != null) {
					PdfNumber rNum  = mkDict.getAsNumber(PdfName.R);

					if (rNum != null)
						return rNum.intValue();
				}
			}
		}

		return 0;
	}

	private void setBarCode() throws DocumentException, IOException {
	    List<FieldPosition> barcodeArea = pdfAcroFields.getFieldPositions(this.dataCutter.elementName);

	    if (barcodeArea != null) {
	   		int fieldRotation = getFieldRotation(pdfAcroFields.getFieldItem(this.dataCutter.elementName));

	   		for (FieldPosition fieldPosition : barcodeArea) {
	   			Rectangle      fieldPlaceHolder = new Rectangle(fieldPosition.position.getLeft(), fieldPosition.position.getBottom(), fieldPosition.position.getRight(), fieldPosition.position.getTop());
	   			Image 		   imgBarCode       = null;
	   			PdfContentByte pdfContent       = pdfFilled.getOverContent(fieldPosition.page);
	   			String 		   elementParam 	= this.dataCutter.elementParam.toUpperCase();

				switch (elementParam) {
				case "AIMS_500_OMR":
				case "AIMS_500_OMR_CS":
					BarcodeDatamatrix bcAIMSDataMatrix = new BarcodeDatamatrix();
					bcAIMSDataMatrix.setOptions(BarcodeDatamatrix.DM_AUTO);
					bcAIMSDataMatrix.setHeight(20);
					bcAIMSDataMatrix.setWidth(20);
					bcAIMSDataMatrix.generate(this.dataCutter.elementValue);

					imgBarCode = bcAIMSDataMatrix.createImage();

					break;
				case "CODE39":
				case "CODE39_CS":
				case "CODE39EXT":
				case "CODE39EXT_CS":
					Barcode39 bc39 = new Barcode39();

					bc39.setCode(this.dataCutter.elementValue);
					bc39.setBarHeight(fieldPlaceHolder.getHeight());
					bc39.setExtended(elementParam.contains("EXT"));
					bc39.setFont(null);

					imgBarCode = bc39.createImageWithBarcode(pdfContent, null, null);

					break;
				case "CODE128":
				case "CODE128_CS":
				case "EAN128C":
				case "EAN128C_CS":
					Barcode128 bc128 = new Barcode128();

					if (elementParam.equals("EAN128C"))
						bc128.setCodeType(Barcode.CODE128_UCC);

					switch (fieldRotation) {
					case 0:
					case 180:
						bc128.setBarHeight(fieldPlaceHolder.getHeight());

						break;
					case 90:
					case 270:
						bc128.setBarHeight(fieldPlaceHolder.getWidth());

						break;
					}

					bc128.setCode(this.dataCutter.elementValue);
					bc128.setFont(null);

					imgBarCode = bc128.createImageWithBarcode(pdfContent, null, null);

					break;
				case "DATAMATRIX":
				case "DATAMATRIX_CS":
				case "DATAMATRIX_52_52":
				case "DATAMATRIX_52_52_CS":
					BarcodeDatamatrix bcDataMatrix = new BarcodeDatamatrix();
					bcDataMatrix.setOptions(BarcodeDatamatrix.DM_AUTO);

					if (elementParam.startsWith("DATAMATRIX_52_52")) {
						bcDataMatrix.setHeight(52);
						bcDataMatrix.setWidth(52);
						bcDataMatrix.setWs(2);
					} else {
						bcDataMatrix.setHeight(16);
						bcDataMatrix.setWidth(48);
					}

					bcDataMatrix.generate(this.dataCutter.elementValue);

					imgBarCode = bcDataMatrix.createImage();

					break;
				case "DATAMATRIX_PI":
				case "DATAMATRIX_PI_CS":
	                DataMatrix dataMatrix = new DataMatrix();
//                  dataMatrix.setDataType(Symbol.DataType.GS1);
//	                dataMatrix.setDataType(Symbol.DataType.HIBC);
	                dataMatrix.setReaderInit(false);
	                dataMatrix.setPreferredSize(30);
	                dataMatrix.setForceMode(ForceMode.RECTANGULAR);
	                dataMatrix.setModuleWidth(6);
	                dataMatrix.setQuietZoneHorizontal(0);
	                dataMatrix.setQuietZoneVertical(0);
	                dataMatrix.setContent(this.dataCutter.elementValue);

					BufferedImage image = new BufferedImage(dataMatrix.getWidth(), dataMatrix.getHeight(), BufferedImage.TYPE_BYTE_BINARY);

                    Java2DRenderer renderer = new Java2DRenderer(image.createGraphics(), 1, OkapiUI.paperColour, OkapiUI.inkColour);
                    renderer.render(dataMatrix);

					imgBarCode = Image.getInstance(image, null);

					break;
				case "I25":
				case "I25_CS":
					BarcodeInter25 bcI25 = new BarcodeInter25();

					switch (fieldRotation) {
					case 0:
					case 180:
						bcI25.setBarHeight(fieldPlaceHolder.getHeight());

						break;
					case 90:
					case 270:
						bcI25.setBarHeight(fieldPlaceHolder.getWidth());

						break;
					}

					bcI25.setX(1.2f);
					bcI25.setCode(this.dataCutter.elementValue);
					bcI25.setFont(null);

					imgBarCode = bcI25.createImageWithBarcode(pdfContent, null, null);

					break;
				case "OMRPBB":
					OMRPBB bcOMRPBB = new OMRPBB();
					bcOMRPBB.setBoxWidth(fieldPlaceHolder.getWidth());
					bcOMRPBB.setBoxHeight(fieldPlaceHolder.getHeight());
					bcOMRPBB.setCode(Integer.valueOf(this.dataCutter.elementValue));

					imgBarCode = bcOMRPBB.createImageWithBarcode(pdfContent, null);

					break;
				case "QRCODE":
					Hashtable<EncodeHintType, Object> hints = new Hashtable<>();
					hints.put(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.M);
					hints.put(EncodeHintType.CHARACTER_SET, "UTF-8");

					BarcodeQRCode bcQRCode = new BarcodeQRCode(this.dataCutter.elementValue, 1, 1, hints);

					imgBarCode = bcQRCode.getImage();

					break;
				}

	    		imgBarCode.setRotationDegrees(fieldRotation);

			    if (elementParam.endsWith("_CS")) {
				    imgBarCode.scaleAbsoluteWidth(fieldPlaceHolder.getWidth());
				    imgBarCode.scaleAbsoluteHeight(fieldPlaceHolder.getHeight());
			    } else {
				    imgBarCode.scaleToFit(fieldPlaceHolder.getWidth(), fieldPlaceHolder.getHeight());
			    }

			    imgBarCode.setAbsolutePosition((fieldPosition.position.getLeft() + (fieldPlaceHolder.getWidth() - imgBarCode.getScaledWidth()) / 2), (fieldPosition.position.getBottom() + (fieldPlaceHolder.getHeight() - imgBarCode.getScaledHeight()) / 2));

			    pdfContent.addImage(imgBarCode);

			    if (this.pdfManagerMode.equals(ManagerMode.DBA))
			    	this.pdfAcroFields.setField(this.dataCutter.elementName.replace("_BC", "_BCT"), this.dataCutter.elementValue);
	   		}
	   	}
	}

	private void setImage() throws MalformedURLException, IOException, DocumentException {
	    List<FieldPosition> imageArea = this.pdfAcroFields.getFieldPositions(this.dataCutter.elementName);

	    if (imageArea != null) {
		    int fieldRotation = getFieldRotation(this.pdfAcroFields.getFieldItem(this.dataCutter.elementName));

			for (FieldPosition fieldPosition : imageArea) {
				Rectangle      fieldPlaceHolder = new Rectangle(fieldPosition.position.getLeft(), fieldPosition.position.getBottom(), fieldPosition.position.getRight(), fieldPosition.position.getTop());
				String 		   imgFName 		= (this.pdfManagerMode == ManagerMode.DBA ? this.dataCutter.elementValue : this.dataCutter.elementParam);
			    Image          image            = Image.getInstance((imgFName.indexOf("\\") + imgFName.indexOf("/") == -2 ? this.prjMMSTemplatesDir + "Images/" : "") + imgFName);
			    PdfContentByte pdfContent       = pdfFilled.getOverContent(fieldPosition.page);

			    image.setRotationDegrees(fieldRotation);
			    image.scaleToFit(fieldPlaceHolder.getWidth(), fieldPlaceHolder.getHeight());
			    image.setAbsolutePosition(fieldPosition.position.getLeft() + (fieldPlaceHolder.getWidth() - image.getScaledWidth()) / 2, fieldPosition.position.getBottom() + (fieldPlaceHolder.getHeight() - image.getScaledHeight()) / 2);

			    pdfContent.addImage(image);
		    }
	    }
	}

	public void setMMSTemplatesDir(String prjMMSTemplatesDir) {
		this.prjMMSTemplatesDir = prjMMSTemplatesDir;
	}

	public void setMode(ManagerMode manageMode) {
		this.pdfManagerMode = manageMode;
	}

	private void setText() throws IOException, DocumentException {
		int srchAlias = this.dataCutter.elementName.indexOf(" AS");

		if (srchAlias > -1)
			this.dataCutter.elementName = this.dataCutter.elementName.substring(srchAlias + 4);

		if (this.dataCutter.elementValue.contains("¿"))
			this.dataCutter.elementValue = this.dataCutter.elementValue.replace("¿", "€");

		if (!this.dataCutter.elementSplitParam.equals("")){
			String[] elementValues = this.dataCutter.elementValue.split("\\" + this.dataCutter.elementSplitParam);

			for (int i = 0; i < elementValues.length; i++)
				this.pdfAcroFields.setField(this.dataCutter.elementName + new DecimalFormat("###000").format(i + 1), elementValues[i]);
		} else {
			this.pdfAcroFields.setField(this.dataCutter.elementName, this.dataCutter.elementValue);
		}
	}

	private void setXML() throws IOException, DocumentException, ParserConfigurationException, SAXException {
		int 	 elementsCntr  = 1;
		String   elementsName  = this.dataCutter.elementName;
		String[] elementsValue = null;

		if ((this.pdfManagerMode == ManagerMode.DBA) && !this.dataCutter.elementSplitParam.equals("")) {
			elementsValue = this.dataCutter.elementValue.split("\\" + this.dataCutter.elementSplitParam);
			elementsCntr  = elementsValue.length;
		}

		for (int i = 0; i < elementsCntr; i++) {
			if (elementsValue != null) {
				elementsName = this.dataCutter.elementName.replace("XXX", new DecimalFormat("###000").format(i + 1));
				this.dataCutter.elementValue = elementsValue[i];
			}

			List<FieldPosition> textArea = pdfAcroFields.getFieldPositions(elementsName);

		   	if (textArea != null) {
		   		int fieldRotation = getFieldRotation(pdfAcroFields.getFieldItem(elementsName));

		   		for (FieldPosition fieldPosition : textArea) {
			    	PdfContentByte pdfContent       = this.pdfFilled.getOverContent(fieldPosition.page);

//				    float x = fieldPosition.position.getLeft();
//				    float y = fieldPosition.position.getTop();
//				    float w = x + fieldPosition.position.getWidth();
//				    float h = y - fieldPosition.position.getHeight();
//
//					pdfContent.setLineWidth(0.5f);
//					pdfContent.moveTo(x, y);
//					pdfContent.lineTo(w, y);
//					pdfContent.lineTo(w, h);
//					pdfContent.lineTo(x, h);
//					pdfContent.lineTo(x, y);
//					pdfContent.lineTo(w, y);
//					pdfContent.stroke();

				    SDFParser mySDFParser = new SDFParser();
					mySDFParser.setPDFContentByte(pdfContent);
					mySDFParser.setBoundingBox(fieldPosition.position);
					mySDFParser.setFieldRotation(fieldRotation);

					AcroFields.Item acroFieldItem = this.pdfAcroFields.getFieldItem(elementsName);
					TextField 		textField 	  = new TextField(null, null, null);
					PdfDictionary 	merged 		  = acroFieldItem.getMerged(0);

					this.pdfAcroFields.decodeGenericDictionary(merged, textField);

					mySDFParser.setPrjBasePath(this.prjMMSTemplatesDir);
					mySDFParser.setFont(textField.getFont(), textField.getFontSize());

					if (this.pdfManagerMode == ManagerMode.DBA) {
						if (this.dataCutter.elementValue.contains("¿"))
							this.dataCutter.elementValue = this.dataCutter.elementValue.replace("¿", "€");

						DocumentBuilderFactory dbf         = DocumentBuilderFactory.newInstance();
						DocumentBuilder 	   db          = dbf.newDocumentBuilder();
						String 				   elementNode = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><root>" + this.dataCutter.elementValue + "</root>";
						Document 		       xmlDoc      = db.parse(new InputSource(new ByteArrayInputStream(elementNode.getBytes("UTF-8"))));

						mySDFParser.getEntities(xmlDoc.getElementsByTagName("root").item(0));
					} else {
						mySDFParser.getEntities(this.dataCutter.elementNode);
					}
				}
			}
		}
	}

}
