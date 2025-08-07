package mms.core.engine.pdfmerger.commons;

import org.w3c.dom.Node;

public class PDFCutterInfo {

	public enum  	    fieldTypes 	 		{ BARCODE, IMAGE, TEXT, XML };

	public String   	elementName  		= "";
	public Node     	elementNode  		= null;
	public String   	elementParam 		= ""; 
	public String 		elementSplitParam	= "";
	public fieldTypes 	elementType  		= fieldTypes.TEXT;
	public String   	elementValue 		= "";

}
