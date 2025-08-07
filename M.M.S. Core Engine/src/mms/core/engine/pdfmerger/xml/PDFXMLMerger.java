package mms.core.engine.pdfmerger.xml;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;

import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;

import mms.core.engine.pdfmerger.commons.PDFCutterInfo;
import mms.core.engine.pdfmerger.commons.PDFManager;
import mms.core.engine.pdfmerger.commons.PDFManager.ManagerMode;

public class PDFXMLMerger {

	private String 			errMsg				= "";
	private String 			pdfIn				= "";
	private InputStream 	pdfInIS				= null;
	private String 			pdfOut				= null;
	private PdfReader 		pdfReader			= null;
	private PdfReader 		pdfTemplate			= null;
	private boolean 		pdfTemplateCache	= false;
	private File   			xfdfFile 			= null;
	private String 			xfdfXML 			= null;

	public boolean execJob() {
		boolean rValue = false;
	
		XFDFParser xfdfParser = null;
	
		if (this.xfdfFile == null) {
			xfdfParser = new XFDFParser(this.xfdfXML);
		} else {
			xfdfParser = new XFDFParser(this.xfdfFile);
		}
	
		ArrayList<PDFCutterInfo> xfdfFields = xfdfParser.getXFDFFields();
	
		if (xfdfFields.size() > 0) {
			try {
				if ((pdfTemplate == null) || (!pdfTemplateCache)) {
					if (pdfInIS == null) {
						pdfTemplate = new PdfReader(((pdfIn == null) ? xfdfParser.getPDFFileName() : pdfIn));
					} else {
						pdfTemplate = new PdfReader(pdfInIS);
					}
				}
	
				PDFManager 			  pdfManage = null;
				ByteArrayOutputStream baos 		= null;
				
				if (this.pdfOut == null) {
					baos      = new ByteArrayOutputStream();
					pdfManage = new PDFManager(new PdfStamper(new PdfReader(pdfTemplate), baos));
				} else {
					pdfManage = new PDFManager(new PdfStamper(new PdfReader(pdfTemplate), new FileOutputStream(this.pdfOut))); 
				}
				
				pdfManage.setMode(ManagerMode.XML);
				pdfManage.setMMSTemplatesDir(xfdfParser.getPrjBasePath());
	
				for (Iterator<PDFCutterInfo> i = xfdfFields.iterator(); i.hasNext();)
					pdfManage.execute((PDFCutterInfo) i.next());
	
				pdfManage.close(true);
	
				if (this.pdfOut == null) {
					this.pdfReader = new PdfReader(baos.toByteArray());
					
					baos.flush();
					baos.close();
				}

				rValue = true;
			} catch (Exception e) {
				errMsg = e.getMessage();
			}
		} else {
			errMsg = xfdfParser.getErrorMessage();
		}
	
		return rValue;
	}

	public String getErrMsg() {
		if (errMsg == null) errMsg = "Unknown Error!";

		return errMsg;
	}

	public PdfReader getPDFReader() {
		return this.pdfReader;
	}

	public void setPDFIn(InputStream pdfInIS) {
		this.pdfInIS = pdfInIS;
	}

	public void setPDFIn(String pdfIn) {
		this.pdfIn = pdfIn;
	}

	public void setPDFOut(String pdfOut) {
		this.pdfOut = pdfOut;
	}

	public void setPDFTemplateCache(boolean pdfTemplateCache) {
		this.pdfTemplateCache = pdfTemplateCache;
	}

	public void setXFDF(File xfdfFile) {
		this.xfdfFile = xfdfFile;
	}

	public void setXFDF(String xfdfXML) {
		this.xfdfXML = xfdfXML;
	}

}