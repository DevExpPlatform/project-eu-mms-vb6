package mms.core.engine.pdfmerger.commons;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.ICC_Profile;
import com.itextpdf.text.pdf.PdfAConformanceLevel;
import com.itextpdf.text.pdf.PdfAWriter;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfWriter;

public class PDFConverter {

	private String 				 colorProfile	= "";
	private String 				 errMsg       	= "";
	private String 				 pdfInFileName	= "";
	private String 				 pdfOutFileName	= "";
	private PdfAConformanceLevel pdfStandard 	= PdfAConformanceLevel.PDF_A_1A;
	
	public boolean convert() {
		PdfContentByte  dstPDFCB     = null;
		Document   	    dstPDFDoc    = new Document();  
		PdfAWriter 	    dstPDFWriter = null;
		boolean 		rValue       = false;
		
		try {
			dstPDFWriter = PdfAWriter.getInstance(dstPDFDoc, new FileOutputStream(pdfOutFileName), pdfStandard);
			dstPDFWriter.setPdfVersion(PdfWriter.PDF_VERSION_1_4);
			dstPDFWriter.createXmpMetadata();
			dstPDFWriter.setTagged();
	
			dstPDFDoc.open();  
			dstPDFCB = dstPDFWriter.getDirectContent();  

			PdfReader srcPDF    = new PdfReader(pdfInFileName);
			int       pageCount = srcPDF.getNumberOfPages();  
	
			for (int i = 0; i < pageCount; i++) {  
				dstPDFDoc.newPage();
				dstPDFCB.addTemplate(dstPDFWriter.getImportedPage(srcPDF, (i + 1)), 0, 0);  
			}  
	
			ICC_Profile icc = ICC_Profile.getInstance(new FileInputStream(colorProfile));
			dstPDFWriter.setOutputIntents("CustomProfile", "", "http://www.color.org", "sRGB IEC61966-2.1", icc);
	
			dstPDFDoc.close();
			srcPDF.close();
			
			rValue = true;
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (DocumentException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	
		return rValue;
	}

	public String getErrMsg() {
		return errMsg;
	}

	public void setInPDFFileName(String pdfInFileName) {
		this.pdfInFileName = pdfInFileName;
	}

	public void setOutPDFFileName(String pdfOutFileName) {
		this.pdfOutFileName = pdfOutFileName;
	}

	public void setPDFColorProfile(String colorProfile) {
		this.colorProfile = colorProfile;
	}

	public void setPDFStandard(String pdfStandard) {
		switch (pdfStandard.hashCode()) {
		case 5418555:
			this.pdfStandard = PdfAConformanceLevel.PDF_A_1A;

			break;
		case 5418556:
			this.pdfStandard = PdfAConformanceLevel.PDF_A_1B;
			
			break;
		}
	}

}
