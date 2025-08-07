package mms.core.engine;

import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.wb.swt.SWTResourceManager;

import mms.core.engine.dbutils.DBConnector;
import mms.core.engine.dbutils.DBConnector.DBType;
import mms.core.engine.pdfmerger.dba.AppSettingsXML;
import mms.core.engine.pdfmerger.dba.PDFDBManager;
import mms.core.engine.pdfmerger.packages.PDFPackageMerger;
import mms.core.engine.pdfmerger.xml.PDFXMLMerger;

public class MMSCEMain {

	public static void main(String[] args) {
		MMSCEMain appMain = null;
		
		try {
			switch (args[0]) {
			case "0":
				PDFXMLMerger myPDFMXML = new PDFXMLMerger();
				myPDFMXML.setPDFOut(args[2]);
				myPDFMXML.setXFDF(args[1]);
				
				if (!myPDFMXML.execJob()) {
					MessageBox MsgBox = new MessageBox(new Shell(), SWT.OK | SWT.ICON_ERROR);
					MsgBox.setText("Guru Meditation:");
					MsgBox.setMessage(myPDFMXML.getErrMsg());
					MsgBox.open();
				}
				
				break;
			case "1":
				appMain = new MMSCEMain();
				appMain.unattendedMode = args[4].equals("1");
				
				if (appMain.rtnAppLoad()) {
					PDFDBManager pdfDBManager = new PDFDBManager();
					pdfDBManager.setCache(appMain.flgCache);
					pdfDBManager.setDBConn(appMain.appDBConn);
					pdfDBManager.setIdProject(args[1]);
					pdfDBManager.setIdWorkingLoad(args[2]);
					pdfDBManager.setPrjMMSBaseDir(args[3]);
					pdfDBManager.setUnattendedMode(appMain.unattendedMode);
					pdfDBManager.open();

					appMain.rtnAppUnload();
				}
				
				break;
				
			case "2":
				appMain = new MMSCEMain();
				appMain.unattendedMode = args[6].equals("1");

				if (appMain.rtnAppLoad()) {
					PDFPackageMerger pdfPackageMerger = new PDFPackageMerger();
					pdfPackageMerger.setCaption("Start Packaging:");
					pdfPackageMerger.setProgressNum(2);
					pdfPackageMerger.setDBConn(appMain.appDBConn);
					pdfPackageMerger.setIdProject(args[1]);
					pdfPackageMerger.setIdWorkingLoad(args[2]);
					pdfPackageMerger.setPrjMMSBaseDir(args[3]);
					pdfPackageMerger.setPackageStart(args[4]);
					pdfPackageMerger.setPackageEnd(args[5]);
					pdfPackageMerger.setUnattendedMode(appMain.unattendedMode);
					
					if (pdfPackageMerger.open()) {
						System.out.print("OK");
					} else {
						if (appMain.unattendedMode) {
							System.out.println(pdfPackageMerger.getErrMsg());
						} else {
							MessageBox MsgBox = new MessageBox(new Shell(), SWT.OK | SWT.ICON_ERROR);
							MsgBox.setText("Guru Meditation:");
							MsgBox.setMessage(pdfPackageMerger.getErrMsg());
							MsgBox.open();
						}
						
						System.out.println("KO");
					}

					appMain.rtnAppUnload();
				}
				
				break;

			case "--help":
				System.out.println("Vers.: 2.5.4 [18/07/2020]\n" + 
						   		   "> Usage XML  Mode: MMSCoreEngine.exe 0 XFDFFile OutPutFile\n" +
		   		   		   		   "> Usage DBA  Mode: MMSCoreEngine.exe 1 IdProject IdWorkingLoad WorkBaseDir UnattendedMode" +
								   "> Usage PACK Mode: MMSCoreEngine.exe 2 IdProject IdWorkingLoad WorkBaseDir PackageStart PackageEnd UnattendedMode");

				break;
			
			default:
				MessageBox MsgBox = new MessageBox(new Shell(), SWT.OK | SWT.ICON_ERROR);
				MsgBox.setText("Guru Meditation:");
				MsgBox.setMessage("No valid arguments passed.");
				MsgBox.open();

				System.out.println("KO");
				
				break;
			}
		} catch (Exception e) {
			MessageBox MsgBox = new MessageBox(new Shell(), SWT.OK | SWT.ICON_ERROR);
			MsgBox.setText("Guru Meditation:");
			MsgBox.setMessage("Error parsing arguments.");
			MsgBox.open();
		}
	}

	private DBConnector appDBConn		= null;
	private boolean     flgCache        = false;
	private boolean 	unattendedMode 	= false;

	private boolean rtnAppLoad() {
		AppSettingsXML myAppSettingsXML = new AppSettingsXML();
	
		if (!myAppSettingsXML.getSettings()) {
			if (this.unattendedMode) {
				System.out.println(myAppSettingsXML.getErrorMessage());
			} else {
				MessageBox MsgBox = new MessageBox(new Shell(), SWT.OK | SWT.ICON_ERROR);
				MsgBox.setText("XML Load Settings Error:");
				MsgBox.setMessage(myAppSettingsXML.getErrorMessage());
				MsgBox.open();
			}
	
			System.exit(0);
		}
	
		appDBConn = new DBConnector();
		
		if (myAppSettingsXML.dbConnectionString == null) {
			appDBConn.setServerType(DBType.ORACLE);
			appDBConn.setServerName(myAppSettingsXML.dbServerName);
			appDBConn.setDataBase(myAppSettingsXML.dbDatabaseName);
		} else {
			appDBConn.setServerType(DBType.ORACLE_CUSTOM);
			appDBConn.setConnectionString(myAppSettingsXML.dbConnectionString);
		}

		appDBConn.setUserId(myAppSettingsXML.dbUser);
		appDBConn.setUserPwd(myAppSettingsXML.dbPassword);
	
		if (myAppSettingsXML.dbServerPort != null)
			appDBConn.setServerPort(myAppSettingsXML.dbServerPort);
		
		if (!appDBConn.open()) {
			if (this.unattendedMode) {
				System.out.println(appDBConn.getErrorMessage());
			} else {
				MessageBox MsgBox = new MessageBox(new Shell(), SWT.OK | SWT.ICON_ERROR);
				MsgBox.setText("DB Connection Error:");
				MsgBox.setMessage(appDBConn.getErrorMessage());
				MsgBox.open();
			}
	
			System.exit(0);
		}
		
		this.flgCache = myAppSettingsXML.flgCache;
		
		return true;
	}

	private void rtnAppUnload() {
		appDBConn.close();
		appDBConn = null;
	
		SWTResourceManager.dispose();
	}

}    