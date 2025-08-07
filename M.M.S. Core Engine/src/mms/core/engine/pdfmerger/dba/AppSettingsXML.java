package mms.core.engine.pdfmerger.dba;

import java.io.File;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class AppSettingsXML {

	private String  errorMsg 	  		= ""; 	

	public  String  dbConnectionString 	= null;
	public  String  dbDatabaseName    	= null;
	public  String  dbPassword     		= null;
    public  String  dbServerName   		= null;
    public 	String  dbServerPort   		= null;
	public  String  dbUser		  		= null;
	public  boolean flgCache            = false;
	
	public String getErrorMessage() {
		return this.errorMsg;
	}

	public boolean getSettings() {
		boolean rValue = false;
		
		try {
			File                   xmlFile   = new File("MMSCoreConfig.XML");
			DocumentBuilderFactory dbf       = DocumentBuilderFactory.newInstance();
			DocumentBuilder        db        = dbf.newDocumentBuilder();
			Document               xmlDoc    = db.parse(xmlFile);
			Element                myElement;
			
			xmlDoc.getDocumentElement().normalize();

			/*
			 * DB Connection Settings
			 */
			myElement = (Element) xmlDoc.getElementsByTagName("dbconnection").item(0);
			
			if (myElement.getElementsByTagName("dbConnString").item(0) == null) {
				this.dbServerName   = myElement.getElementsByTagName("srvr").item(0).getChildNodes().item(0).getNodeValue();
				this.dbDatabaseName = myElement.getElementsByTagName("db").item(0).getChildNodes().item(0).getNodeValue();
			} else {
				this.dbConnectionString = myElement.getElementsByTagName("dbConnString").item(0).getChildNodes().item(0).getNodeValue();
			}

			this.dbUser     = myElement.getElementsByTagName("usr").item(0).getChildNodes().item(0).getNodeValue();
			this.dbPassword = myElement.getElementsByTagName("pwd").item(0).getChildNodes().item(0).getNodeValue();

			if (myElement.getElementsByTagName("port").item(0) != null)
				this.dbServerPort = myElement.getElementsByTagName("port").item(0).getChildNodes().item(0).getNodeValue();

			if (myElement.getElementsByTagName("cache").item(0) != null)
				this.flgCache = myElement.getElementsByTagName("cache").item(0).getChildNodes().item(0).getNodeValue().equals("true");

			rValue = true;
		} catch (Exception e) {
			errorMsg = e.getMessage();
		}
		
		return rValue;
	}

}
