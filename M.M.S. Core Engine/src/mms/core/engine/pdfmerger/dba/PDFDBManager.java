package mms.core.engine.pdfmerger.dba;

import java.sql.ResultSet;
import java.sql.SQLException;

import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;

import mms.core.engine.dbutils.DBCmndExec;
import mms.core.engine.dbutils.DBConnector;

public class PDFDBManager {

	private DBConnector 	appDBConn 		= null;
	private String 			errMsg			= "";
	private boolean         flgCache        = false;
	private MMSProjectInfo 	prjMMSInfo 		= new MMSProjectInfo();
	private boolean 		unattendedMode 	= false;

	private boolean getPackagesNum() {
		errMsg = "";

		DBCmndExec 	myQuery = new DBCmndExec(appDBConn.dbConn);
		ResultSet 	rS 	    = myQuery.Query("SELECT MAX(ID_PACCO) AS NMR_PACKAGES, COUNT(ID_WORKINGLOAD) AS NMR_PDFS FROM " + this.prjMMSInfo.refTable + " WHERE ID_WORKINGLOAD = " + this.prjMMSInfo.idWorkingLoad);
		boolean 	rValue	= false;

		if (myQuery.getNumRows() > 0) {
			try {
				rS.first();

				this.prjMMSInfo.nmrPDFPacks = rS.getInt("NMR_PACKAGES");
				this.prjMMSInfo.nmrPDFs     = rS.getInt("NMR_PDFS");

				rValue = true;
			} catch (SQLException e) {
				this.errMsg = e.getMessage();
			}
		} else {
			this.errMsg = "No records found...";
		}

		myQuery.QueryClose();
		myQuery = null;

		return rValue;
	}

	private boolean getProjectData() {
		errMsg = "";

		boolean     rValue  = false;
		DBCmndExec 	myQuery = new DBCmndExec(appDBConn.dbConn);
		ResultSet 	rS 		= myQuery.Query("SELECT EDT_SUBPROJECTS.ID_SUBPROJECT, " +
											"EDT_PROJECTS.DESCR_PROJECT, " +
											"EDT_PROJECTS.ID_DATACUTTER, " +
											"EDT_PROJECTS.STR_REFTABLENAME, " +
											"EDT_PROJECTS.STR_PRJWORKDIR, " +
											"EDT_SUBPROJECTS.STR_SUBPRJWORKDIR, " +
											"EDT_SUBPROJECTS.STR_BASEFILENAME, " +
											"EDT_SUBPROJECTS.FLG_SINGLEPDFPACKING, " +
											"EDT_SUBPROJECTS.FLG_EMPTYCACHE, " +
											"EDT_PROJECTS.STR_ORDERFIELDS, " +
											"EDT_PROJECTS.STR_BARCODETABLE, " + 
											"EDT_PROJECTS.ID_SERIALIZER, " + 
											"EDT_PROJECTS.ID_SERIALIZERPARAMS, " + 
											"EDT_PROJECTS.NMR_PRJWEIGHT, " +
											"EDT_PRODUCTS.NMR_PRODOTTOPESO, " +
											"EDT_PRODUCTS.NMR_SCATOLAPESOMAX " +
											"FROM EDT_PROJECTS " +
											"INNER JOIN EDT_SUBPROJECTS " +
											"ON EDT_PROJECTS.ID_PROJECT = EDT_SUBPROJECTS.ID_PROJECT " +
											"INNER JOIN EDT_PRODUCTS " +
											"ON EDT_PRODUCTS.ID_PRODUCT = EDT_PROJECTS.ID_PRODUCT " +
											"WHERE EDT_PROJECTS.ID_PROJECT = " + this.prjMMSInfo.idProject);

		if (myQuery.getNumRows() == 1) {
			try {
				rS.first();

				this.prjMMSInfo.idSubProject 	    = rS.getString("ID_SUBPROJECT");
				this.prjMMSInfo.strDescrProject 	= rS.getString("DESCR_PROJECT");
				this.prjMMSInfo.idDataCutter 	    = rS.getString("ID_DATACUTTER");
				this.prjMMSInfo.refBarCodeTable     = rS.getString("STR_BARCODETABLE");
				this.prjMMSInfo.refTable 	        = rS.getString("STR_REFTABLENAME");
				this.prjMMSInfo.flgSinglePDFPack	= rS.getString("FLG_SINGLEPDFPACKING").equals("1");
				this.prjMMSInfo.emptyPDFCache		= (rS.getInt("FLG_EMPTYCACHE") == 1);
				this.prjMMSInfo.strOrderFields 	    = rS.getString("STR_ORDERFIELDS");
				this.prjMMSInfo.pdfFileName         = rS.getString("STR_BASEFILENAME");
				this.prjMMSInfo.prjMMSTemplatesDir  = this.prjMMSInfo.prjMMSBaseDir + rS.getString("STR_PRJWORKDIR") + "/Templates/"; 
				this.prjMMSInfo.prjMMSBaseDir      += rS.getString("STR_PRJWORKDIR") + "/" + rS.getString("STR_SUBPRJWORKDIR") + "/" + this.prjMMSInfo.idWorkingLoad + "/"; 
				this.prjMMSInfo.nmrBusteMax         = (rS.getInt("NMR_SCATOLAPESOMAX") * 1000) / (rS.getInt("NMR_PRODOTTOPESO") + rS.getInt("NMR_PRJWEIGHT"));
				this.prjMMSInfo.serializerId		= rS.getString("ID_SERIALIZER");
				this.prjMMSInfo.serializerParams	= rS.getString("ID_SERIALIZERPARAMS");
				
				rValue = true;
			} catch (SQLException e) {
				this.errMsg = e.getMessage();
			}
		} else {
			this.errMsg = myQuery.getErrorMessage();
		}

		myQuery.QueryClose();
		myQuery = null;

		return rValue;
	}

	public void open() {
		boolean rValue = false;
		String  txtMsg = "MMS Project Info Error:";

		/*
		 * Step 01
		 */
		if (getProjectData())
			rValue = getPackagesNum();

		/*
		 * Step 02
		 */
		if ((rValue) && (this.prjMMSInfo.nmrPDFPacks == 0)) {
			PDFDBSerializer pdfSerializer = new PDFDBSerializer();
			pdfSerializer.setDBConn(appDBConn);
			pdfSerializer.setMMSPrjInfo(this.prjMMSInfo);
			pdfSerializer.setCaption("Serializing: " + this.prjMMSInfo.strDescrProject);
			pdfSerializer.setProgressNum(1);
			pdfSerializer.setUnattendedMode(this.unattendedMode);

			rValue = pdfSerializer.open();

			txtMsg = "Data Serializer:";
			errMsg = pdfSerializer.getErrMsg();

			if (rValue)
				rValue = getPackagesNum();
		}

		/*
		 * Step 03
		 */
		if (rValue) {
			PDFDBMerger pdfMerger = new PDFDBMerger();
			pdfMerger.setCache(this.flgCache);
			pdfMerger.setDBConn(appDBConn);
			pdfMerger.setMMSPrjInfo(this.prjMMSInfo);
			pdfMerger.setCaption("Generating: " + this.prjMMSInfo.strDescrProject);
			pdfMerger.setProgressNum(2);
			pdfMerger.setUnattendedMode(this.unattendedMode);

			rValue = pdfMerger.open();

			txtMsg = "PDF Merger:";
			errMsg = pdfMerger.getErrMsg();
		}

		if (rValue) {
			System.out.print("OK");
		} else {
			if (this.unattendedMode) {
				System.out.println(txtMsg + " " + errMsg);
			} else {
				MessageBox MsgBox = new MessageBox(new Shell(), SWT.OK | SWT.ICON_ERROR);
				MsgBox.setText(txtMsg);
				MsgBox.setMessage(errMsg);
				MsgBox.open();
				
				System.out.println("KO");
			}
		}
	}

	public void setCache(boolean flgCache) {
		this.flgCache = flgCache;
	}

	public void setDBConn(DBConnector appDBConn) {
		this.appDBConn = appDBConn;
	}

	public void setIdProject(String idProject) {
		this.prjMMSInfo.idProject = idProject;
	}

	public void setIdWorkingLoad(String idWorkingLoad) {
		this.prjMMSInfo.idWorkingLoad = idWorkingLoad;
	}

	public void setPrjMMSBaseDir(String prjMMSBaseDir) {
		this.prjMMSInfo.prjMMSBaseDir = prjMMSBaseDir;
	}

	public void setUnattendedMode(boolean bValue) {
		this.unattendedMode = bValue;
	}

}