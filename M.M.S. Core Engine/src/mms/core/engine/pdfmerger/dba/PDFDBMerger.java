package mms.core.engine.pdfmerger.dba;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.eclipse.swt.widgets.Display;

import com.itextpdf.text.pdf.PdfName;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.itextpdf.text.pdf.PdfStream;

import mms.core.engine.dbutils.DBCmndExec;
import mms.core.engine.dbutils.DBConnector;
import mms.core.engine.gui.APB;
import mms.core.engine.packer.PDFPacker;
import mms.core.engine.pdfmerger.commons.PDFBuilder;
import mms.core.engine.pdfmerger.commons.PDFCutterInfo;
import mms.core.engine.pdfmerger.commons.PDFCutterInfo.fieldTypes;
import mms.core.engine.pdfmerger.commons.PDFManager;
import mms.core.engine.pdfmerger.commons.PDFManager.ManagerMode;
import multivalent.std.adaptor.pdf.PDFReader;

public class PDFDBMerger extends APB {

	class JobExecutor extends Thread {

		public void executeProcess() {
			rValue = false;

			String 	   				fieldName   = "";
			HashMap<String, byte[]> hmPDFCache  = null;
			DBCmndExec  			myQuery     = null;
			byte[]  				pdfSrcBytes = null;
			String 	   				sqlString 	 = "";

			for (Iterator<PDFCutterInfo> i = pdfDataCutter.iterator(); i.hasNext();) 
				sqlString += i.next().elementName + ", ";

			sqlString = "SELECT ID_POSIZIONE, " + sqlString.substring(0, sqlString.length() - 2) + " FROM " + prjMMSInfo.refTable + " WHERE (ID_WORKINGLOAD = " + prjMMSInfo.idWorkingLoad + " AND ID_PACCO = ";

			try {
				int        cntrPDF 		  = 0;
				PDFBuilder pdfBuilder     = null;
				String     tmpTemplateKey = null;

				myQuery = new DBCmndExec(appDBConn.dbConn);

				if (!unattendedMode) {
					myAPBList.get(0).setMaximum(prjMMSInfo.nmrPDFPacks);
					myAPBList.get(0).setETCETAMaxValue(prjMMSInfo.nmrPDFs);
				}

				for (int idPack = 1; idPack <= prjMMSInfo.nmrPDFPacks; idPack++) {
					int 	  	cntrPDFQPos 	= 0;
					String 	  	pdfWorkDir  	= prjMMSInfo.prjMMSBaseDir + "Workings/Package_" + new DecimalFormat("###000").format(idPack) +"/" ;
					ResultSet 	rS 		  		= myQuery.Query(sqlString + idPack + ") ORDER BY ID_POSIZIONE");

					if (!unattendedMode) {
						myAPBList.get(1).setMaximum(myQuery.getNumRows());
						nmrPDFQNum = new DecimalFormat("###000").format(myQuery.getNumRows());

						if (idPack == 1) {
							nmrPacksNum = new DecimalFormat("###000").format(prjMMSInfo.nmrPDFPacks);
							nmrPDFNum   = new DecimalFormat("#####00000").format(prjMMSInfo.nmrPDFs);

							ProgressEvent progressEvent = new ProgressEvent();
							progressEvent.setDaemon(true);
							progressEvent.start();
						}
					}

					fsDeleteDir(new File(pdfWorkDir));
					fsMakeDir(pdfWorkDir);

					while (rS.next()) {
						if (flgCache) {
							if (pdfTemplateQField.equals("")) {
								if (pdfTemplates.size() == 1) {
									if (pdfSrcBytes == null)
										pdfSrcBytes = Files.readAllBytes(new File(prjMMSInfo.prjMMSTemplatesDir + pdfTemplates.get("1")).toPath());
								} else {
									if (pdfBuilder == null) {
										pdfBuilder     = new PDFBuilder();
										pdfBuilder.setTemplates(pdfTemplates);
										pdfBuilder.setTemplatesPath(prjMMSInfo.prjMMSTemplatesDir);

										hmPDFCache     = new HashMap<String, byte[]>();
										tmpTemplateKey = "";
									}

									pdfBuilder.setXMLData(rS.getString("XML_TEMPLATE"));

									String templateKey = pdfBuilder.getTemplateKey();
									
									if (tmpTemplateKey.equals(templateKey)) {
										flgRender = false;
									} else {
										tmpTemplateKey = templateKey;
										pdfSrcBytes    = hmPDFCache.get(templateKey);

										flgRender = (pdfSrcBytes == null);
										
										if (flgRender) {
											pdfSrcBytes = pdfBuilder.getTemplateBytes();

											hmPDFCache.put(templateKey, pdfSrcBytes);
										}
									}
								}
							} else {
								if (hmPDFCache == null)
									hmPDFCache = new HashMap<String, byte[]>();

								pdfSrcBytes = hmPDFCache.get(pdfTemplateQField);

								flgRender = (pdfSrcBytes == null);
								
								if (flgRender) {
									pdfSrcBytes = Files.readAllBytes(new File(prjMMSInfo.prjMMSTemplatesDir + pdfTemplates.get(rS.getString(pdfTemplateQField))).toPath());

									hmPDFCache.put(pdfTemplateQField, pdfSrcBytes);
								}
							}
						} else {
							if (pdfTemplateQField.equals("")) {
								if (pdfTemplates.size() == 1) {
									if (pdfSrcBytes == null)
										pdfSrcBytes = Files.readAllBytes(new File(prjMMSInfo.prjMMSTemplatesDir + pdfTemplates.get("1")).toPath());
								} else {
									if (pdfBuilder == null) {
										pdfBuilder = new PDFBuilder();
										pdfBuilder.setTemplates(pdfTemplates);
										pdfBuilder.setTemplatesPath(prjMMSInfo.prjMMSTemplatesDir);
									}

									pdfBuilder.setXMLData(rS.getString("XML_TEMPLATE"));

									pdfSrcBytes = pdfBuilder.getTemplateBytes();
								}
							} else {
								pdfSrcBytes = Files.readAllBytes(new File(prjMMSInfo.prjMMSTemplatesDir + pdfTemplates.get(rS.getString(pdfTemplateQField))).toPath());
							}
						}

						String dstPDFName = pdfWorkDir + prjMMSInfo.pdfFileName + "_D" + new DecimalFormat("###000").format(rS.getInt("ID_POSIZIONE")) + ".PDF";

						PdfStamper pdfStamper = new PdfStamper(new PdfReader(pdfSrcBytes), new FileOutputStream(dstPDFName));
						pdfStamper.getReader().getCatalog().remove(PdfName.METADATA);
						pdfStamper.getReader().removeUsageRights();
						pdfStamper.getReader().removeUnusedObjects();
						pdfStamper.getWriter().setCompressionLevel(PdfStream.BEST_COMPRESSION);
						pdfStamper.setFullCompression();

						PDFManager pdfManage = new PDFManager(pdfStamper); 
						pdfManage.setMode(ManagerMode.DBA);
						pdfManage.setMMSTemplatesDir(prjMMSInfo.prjMMSTemplatesDir);

						for (int i = 0; i < pdfDataCutter.size(); i++) {
							PDFCutterInfo pdfCutterInfo = pdfDataCutter.get(i);
							pdfCutterInfo.elementValue  = rS.getString(i + 2);

							fieldName = pdfCutterInfo.elementName;

							pdfManage.execute(pdfCutterInfo);
						}

						pdfManage.close(true);
						pdfManage = null;

						if (prjMMSInfo.flgSinglePDFPack) {
							File      srcPDF = new File(dstPDFName);
							PDFPacker dstPDF = new PDFPacker(new PDFReader(srcPDF));

							dstPDF.setQuiet(true);
							dstPDF.setJPEG(true);
							dstPDF.writeFile(srcPDF);
							dstPDF = null;
						}

						nmrPacksPosition = idPack;
						nmrPDFPosition   = ++cntrPDF;
						nmrPDFQPos       = ++cntrPDFQPos;
					}

					myQuery.QueryClose();

					System.gc();
				}

				if (hmPDFCache != null) {
					hmPDFCache.clear();
					hmPDFCache = null;
				}

				myQuery = null;

				rValue = true;
			} catch (Exception e) {
				errMsg = ((e.getMessage() == null) ? "Null Value Error." : e.getMessage()) + "\n\nFieldName: " + fieldName + "\nClass: " + e.getStackTrace()[0].getClassName() + "\nMethod: " + e.getStackTrace()[0].getMethodName() + "\nLine Number: " + e.getStackTrace()[0].getLineNumber();
			}

			if (pdfSrcBytes != null)
				pdfSrcBytes = null;

			if (myQuery != null) {
				myQuery.QueryClose();
				myQuery = null;
			}

			jobRunning = false;

			if (!unattendedMode) {
				Display.getDefault().syncExec(new Runnable() {
					public void run() {
						frmMainShell.close();
					}
				});
			}
		}

		public void run() {
			executeProcess();
		}

	}

	class ProgressEvent extends Thread {

		private int nmrPacksPositionTmp = 0;
		private int nmrPDFPositionTmp	= 0;
		private int nmrPDFQPosTmp		= 0;

		public void run() {
			while (jobRunning) {
				Display.getDefault().syncExec(new Runnable() {
					public void run() {
						if (nmrPDFPositionTmp != nmrPDFPosition) {
							myAPBList.get(0).getETCETA(nmrPDFPosition);

							nmrPDFPositionTmp = nmrPDFPosition;
						}

						if (nmrPacksPositionTmp != nmrPacksPosition) {
							myAPBList.get(0).setSelection(nmrPacksPosition);
							myAPBList.get(0).setLabelRightCaption("Pack: " + new DecimalFormat("###000").format(nmrPacksPosition) + "/" + nmrPacksNum);

							nmrPacksPositionTmp = nmrPacksPosition;
						}

						if (nmrPDFQPosTmp != nmrPDFQPos) {
							myAPBList.get(1).getETCETA(nmrPDFQPos);
							myAPBList.get(1).setSelection(nmrPDFQPos);
							myAPBList.get(1).setLabelRightCaption((flgRender ? "• " : "") + "Doc: " + new DecimalFormat("###000").format(nmrPDFQPos) + "/"  + nmrPDFQNum + " - Item: " + new DecimalFormat("#####00000").format(nmrPDFPosition) + "/" + nmrPDFNum);

							nmrPDFQPosTmp = nmrPDFQPos;
						}

						if ((nmrPDFPosition % 500) == 0)
							Thread.yield();
					}
				});
			}
		}

	}

	private DBConnector 			 appDBConn				= null;
	private boolean                  flgCache               = false;
	private boolean                  flgRender               = false;
	private String 					 nmrPacksNum			= "";
	private int 					 nmrPacksPosition		= 1;
	private String 					 nmrPDFNum				= ""; 
	private int 					 nmrPDFPosition			= 1;
	private String 					 nmrPDFQNum				= "";
	private int 					 nmrPDFQPos				= 1;
	private ArrayList<PDFCutterInfo> pdfDataCutter   		= new ArrayList<PDFCutterInfo>();
	private String 					 pdfTemplateQField 		= "";
	private Map<String, String> 	 pdfTemplates 			= new HashMap<String, String>();
	private MMSProjectInfo 			 prjMMSInfo				= null;

	private boolean fsDeleteDir(File DirName) {
		if (DirName.isDirectory()) {
			String[] Children = DirName.list();

			for (int I = 0; I < Children.length; I++) {
				boolean success = fsDeleteDir(new File(DirName, Children[I]));

				if (!success) 
					return false;
			}
		}

		return DirName.delete();
	}

	private boolean fsMakeDir(String myPath) {
		File NewDir = new File(myPath);

		if (!NewDir.exists()) {
			return NewDir.mkdirs();
		} else {
			return true;
		}
	}

	private boolean getDataCutter() {
		DBCmndExec 	  myQuery 		= new DBCmndExec(appDBConn.dbConn);
		ResultSet 	  rS 			= myQuery.Query("SELECT * FROM EDT_DATACUTTER WHERE ID_DATACUTTER = " + this.prjMMSInfo.idDataCutter + " ORDER BY NMR_FIELDORDER");
		boolean       rValue  		= false;

		if (myQuery.getNumRows() > 0) {
			try {
				while (rS.next()) {
					PDFCutterInfo myDataCutter  = new PDFCutterInfo();
					myDataCutter.elementName = rS.getString("DESCR_FIELDNAME");

					if (rS.getString("FLG_SPLITTER") != null) 
						myDataCutter.elementSplitParam = rS.getString("FLG_SPLITTER");

					if (rS.getString("STR_BARCODETYPE") != null) {
						myDataCutter.elementParam = rS.getString("STR_BARCODETYPE");
						myDataCutter.elementType  = fieldTypes.BARCODE;
					}

					if (rS.getBoolean("FLG_ISIMAGE"))
						myDataCutter.elementType  = fieldTypes.IMAGE;

					if (rS.getBoolean("FLG_ISML")) 
						myDataCutter.elementType  = fieldTypes.XML;

					this.pdfDataCutter.add(myDataCutter);
				}

				rValue = true;
			} catch (SQLException e) {
				errMsg = "Get DataCutter: " + e.getMessage();
			}
		} else {
			errMsg = "Get DataCutter: Unable to find Data Cutter.";
		}

		myQuery.QueryClose();
		myQuery = null;

		return rValue;
	}

	private boolean getDataCutterMerge() {
		DBCmndExec 	myQuery = new DBCmndExec(appDBConn.dbConn);
		ResultSet 	rS 		= myQuery.Query("SELECT * FROM EDT_DATACUTTER WHERE ID_DATACUTTER = " + this.prjMMSInfo.idDataCutter + " AND STR_FIELDMERGER IS NOT NULL ORDER BY STR_FIELDMERGER");
		boolean     rValue  = false;

		if (myQuery.getNumRows() > 0) {
			String fieldMerged  = "";
			String tmpFieldName = "";

			try {
				while (rS.next()) {
					String fieldName = rS.getString("STR_FIELDMERGER");

					fieldName = fieldName.substring(0, fieldName.length() - 3);

					if (!fieldName.equals(tmpFieldName)) {
						if (!tmpFieldName.equals("")) {
							fieldMerged = fieldMerged.substring(0, fieldMerged.length() - 1).replace("|", " || ' ' || ") + " AS " + tmpFieldName; 

							PDFCutterInfo myDataCutter  = new PDFCutterInfo();
							myDataCutter.elementName = fieldMerged;

							this.pdfDataCutter.add(myDataCutter);
						}

						fieldMerged  = "";
						tmpFieldName = fieldName;
					}

					fieldMerged += rS.getString("DESCR_FIELDNAME") + "|"; 
				}

				fieldMerged = fieldMerged.substring(0, fieldMerged.length() - 1).replace("|", " || ' ' || ") + " AS " + tmpFieldName; 

				PDFCutterInfo myDataCutter = new PDFCutterInfo();
				myDataCutter.elementName = fieldMerged;

				this.pdfDataCutter.add(myDataCutter);

				rValue = true;
			} catch (SQLException e) {
				errMsg = "Get Merged DataCutter: " + e.getMessage();
			}
		} else {
			rValue = true;
		}

		myQuery.QueryClose();
		myQuery = null;

		return rValue;
	}

	private boolean getTemplates() {
		boolean     rValue  = false;
		DBCmndExec 	myQuery = new DBCmndExec(appDBConn.dbConn);
		ResultSet 	rS 		= myQuery.Query("SELECT NMR_TEMPLATEORDER, STR_TEMPLATEFILENAME, STR_QFIELD, STR_QVALUE FROM REF_TEMPLATES " +
				"INNER JOIN EDT_TEMPLATES ON REF_TEMPLATES.ID_TEMPLATE = EDT_TEMPLATES.ID_TEMPLATE " +
				"WHERE ID_SUBPROJECT = " + this.prjMMSInfo.idSubProject +
				"ORDER BY ID_SUBPROJECT, STR_QFIELD, STR_QVALUE, NMR_TEMPLATEORDER");
		int 		numRows = myQuery.getNumRows(); 

		if (numRows > 0) {
			int cntr = 0;

			try {
				while (rS.next()) {
					if ((cntr == 0) && (rS.getString("STR_QFIELD") != null)) 
						this.pdfTemplateQField = rS.getString("STR_QFIELD");

					pdfTemplates.put((rS.getString("STR_QVALUE") == null ? rS.getString("NMR_TEMPLATEORDER") : rS.getString("STR_QVALUE")), rS.getString("STR_TEMPLATEFILENAME"));
				}

				rValue = true;
			} catch (SQLException e) {
				errMsg = "Get Templates Info: " + e.getMessage();
			}
		} else {
			errMsg = "Get Templates Info: Unable to read Templates Info.";
		}

		myQuery.QueryClose();
		myQuery = null;

		return rValue;
	}

	protected void jobExecutor() {
		rValue = getDataCutter(); 

		if (rValue)
			rValue = getDataCutterMerge(); 

		if (rValue) 
			rValue = getTemplates();

		if (rValue) {
			JobExecutor jobMerger = new JobExecutor();

			if (unattendedMode) {
				jobMerger.executeProcess();
			} else {
				jobMerger.setPriority(Thread.MAX_PRIORITY);
				jobMerger.start();
			}
		} else {
			if (!unattendedMode) {
				Display.getDefault().syncExec(new Runnable() {
					public void run() {
						frmMainShell.close();
					}
				});
			}
		}
	}

	public void setCache(boolean flgCache) {
		this.flgCache = flgCache;
	}

	public void setDBConn(DBConnector myAppDBConn) {
		this.appDBConn = myAppDBConn;
	}

	public void setMMSPrjInfo(MMSProjectInfo prjMMSInfo) {
		this.prjMMSInfo = prjMMSInfo;
	}

}
