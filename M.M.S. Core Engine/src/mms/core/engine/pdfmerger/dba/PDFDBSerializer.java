package mms.core.engine.pdfmerger.dba;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DecimalFormat;

import org.eclipse.swt.widgets.Display;

import mms.core.engine.dbutils.DBCmndExec;
import mms.core.engine.dbutils.DBConnector;
import mms.core.engine.gui.APB;
import mms.core.engine.pdfmerger.serializers.CustomSerializer;
import mms.core.engine.pdfmerger.serializers.codegen.AIMS_500;
import mms.core.engine.pdfmerger.serializers.codegen.BCR_DS85;
import mms.core.engine.pdfmerger.serializers.codegen.OMR_PBB;

public class PDFDBSerializer extends APB {

	class JobExecutor  extends Thread {

		public void executeProcess() {
			boolean    serialBarCode    	  = (prjMMSInfo.refBarCodeTable != null);
			boolean    customSerializer 	  = (prjMMSInfo.serializerId != null);
			String[]   customSerializerParams = ((prjMMSInfo.serializerParams == null) ? null : prjMMSInfo.serializerParams.split("\\|"));	 
			DBCmndExec myQuery 				  = new DBCmndExec(appDBConn.dbConn);

			errMsg = "";

			myQuery.setTransactional(true);

			if (!unattendedMode)
				myAPBList.get(0).setLabelLeftCaption("Initializing...");

			String sqlString = "UPDATE " + prjMMSInfo.refTable + " SET " + 
							   "ID_PACCO = NULL, " +
							   "ID_POSIZIONE = NULL" + 
							   (serialBarCode ? ", NMR_ITFCODE = NULL" : "") + 
							   (customSerializer ? ", " + customSerializerParams[0] + " = NULL" : "") + 
							   " WHERE ID_WORKINGLOAD = " + prjMMSInfo.idWorkingLoad;  

			if (myQuery.QueryCommand(sqlString)) {
				sqlString = "";

				String prjSQLSortFields = getSortFields(prjMMSInfo.strOrderFields);

				if (customSerializer && (prjSQLSortFields != "")) {
					sqlString = ", " + prjSQLSortFields;

					for (int i = 0; i < (customSerializerParams.length - 1); i++)
						if (!prjSQLSortFields.contains(customSerializerParams[i]))
							sqlString += ", " + customSerializerParams[i];
				}

				sqlString = "SELECT ID_PACCO, ID_POSIZIONE" + 
							(serialBarCode ? ", NMR_ITFCODE" : "") + 
							sqlString + 
							" FROM " + prjMMSInfo.refTable + 
							" WHERE ID_WORKINGLOAD = " + prjMMSInfo.idWorkingLoad + 
							(prjSQLSortFields == "" ? "" : " ORDER BY " + prjSQLSortFields);

				ResultSet rS = myQuery.Query(sqlString, ResultSet.TYPE_SCROLL_SENSITIVE, ResultSet.CONCUR_UPDATABLE);					

				int rowNums = myQuery.getNumRows();

				if (rowNums > 0) {
					if (!unattendedMode) {
						myAPBList.get(0).setMaximum(rowNums);
						recNums = new DecimalFormat("#####00000").format(rowNums);

						Display.getDefault().syncExec(new Runnable() {
							public void run() {
								myAPBList.get(0).setLabelLeftCaption("");
							}
						});

						ProgressEvent progressEvent = new ProgressEvent();
						progressEvent.setDaemon(true);
						progressEvent.start();
					}

					long 			 idBarCode    = 0;
					CustomSerializer mySerializer = null;

					try {
						if (serialBarCode)
							idBarCode = getBarCodeId();

						if (customSerializer) {
							switch (prjMMSInfo.serializerId) {
							case "AIMS_500":
								mySerializer = new AIMS_500();
								
								break;
							case "BCR_DS85":
								mySerializer = new BCR_DS85();
								
								break;
							case "OMR_PBB":
								mySerializer = new OMR_PBB();
								
								break;
							default:
								break;
							}
						}

						while (rS.next()) {
							if (posCntr == prjMMSInfo.nmrBusteMax) {
								packCntr++;
								posCntr = 0;
							}

							rS.updateInt("ID_PACCO", packCntr);
							rS.updateInt("ID_POSIZIONE", ++posCntr);

							if (serialBarCode) {
								String strBarCodeId = String.valueOf(idBarCode++);

								if ((strBarCodeId.length() % 2) != 0) 
									strBarCodeId = "0" + strBarCodeId;

								rS.updateString("NMR_ITFCODE", strBarCodeId);
							}

							if (customSerializer) {
								mySerializer.setPackage(packCntr);
								mySerializer.setPos(posCntr);
								mySerializer.setParams(customSerializerParams);
								mySerializer.setRS(rS);

								String sequence = mySerializer.getSequence();

								if (sequence != null)
									rS.updateString(customSerializerParams[0], sequence);
							}

							rS.updateRow();

							recPos++;
						}

						if (serialBarCode) {
							rValue = setBarCodeId(idBarCode);
						} else { 
							rValue = true;
						}
					} catch (SQLException ex) {
						errMsg = ex.getMessage();
					}
				} else {
					errMsg = "Nessun record trovato per la serializzazione.";
				}
			} else {
				errMsg = "Errore durante lo svuotamento dei campi per la tabella " + prjMMSInfo.refTable;
			}

			jobRunning = false;

			if (rValue) {
				myQuery.TransactionCommit();
			} else {
				myQuery.TransactionRollBack();
			}

			myQuery.setTransactional(false);
			myQuery.QueryClose();

			if (!unattendedMode) {
				Display.getDefault().syncExec(new Runnable() {
					public void run() {
						frmMainShell.close();
					}
				});
			}
		}

		private long getBarCodeId() throws SQLException {
			DBCmndExec myQuery  = new DBCmndExec(appDBConn.dbConn);
			ResultSet  rS 		= myQuery.Query("SELECT NMR_LASTRANGEID FROM REF_" + prjMMSInfo.refBarCodeTable);					
			long 	   rValue 	= 0;

			int rowNums = myQuery.getNumRows();

			if (rowNums > 0) {
				rS.next();

				rValue = rS.getLong("NMR_LASTRANGEID");
			}

			myQuery.QueryClose();
			myQuery = null;

			return rValue;
		}

		private String getSortFields(String prjSortField) {
			if (prjSortField.length() > 0) {
				String[] sqlFields = prjSortField.split("\\|");

				if (sqlFields[0].length() > 0) {
					String sqlSortFields = "";

					for (int i = 0; i < sqlFields.length; i++) {
						String[] sqlFieldValues = sqlFields[i].split("\\;");

						sqlSortFields += sqlFieldValues[0] + (sqlFieldValues.length == 1 ? "" : " DESC") + ", ";
					}

					return sqlSortFields.substring(0, sqlSortFields.length() - 2);
				}
			}

			return "";
		}

		public void run() {
			executeProcess();
		}

		public boolean setBarCodeId(long lastId) {
			DBCmndExec myQuery = new DBCmndExec(appDBConn.dbConn);
			boolean    rValue  = false;

			rValue = myQuery.QueryCommand("UPDATE REF_" + prjMMSInfo.refBarCodeTable + " SET NMR_LASTRANGEID = " + lastId); 

			myQuery.QueryClose();
			myQuery = null;

			return rValue;
		}

	}

	class ProgressEvent extends Thread {

		public void run() {
			while (jobRunning) {
				if (recPosTmp == recPos) {
					if ((recPos % 500) == 0)
						Thread.yield();
				} else {
					Display.getDefault().syncExec(new Runnable() {
						public void run() {
							myAPBList.get(0).getETCETA(recPos);
							myAPBList.get(0).setSelection(recPos);
							myAPBList.get(0).setLabelRightCaption("P: " + new DecimalFormat("###000").format(packCntr) + " - D: " + new DecimalFormat("###000").format(posCntr) + " - R: " + new DecimalFormat("#####00000").format(recPos) + " of " + recNums);
						}
					});

					recPosTmp = recPos;
				}
			}
		}

	}

	private DBConnector 	appDBConn		= null;
	private int 			packCntr   		= 1;
	private int 			posCntr			= 0;
	private MMSProjectInfo 	prjMMSInfo	  	= null;
	private String 			recNums 		= "";
	private int 			recPos 			= 0;
	private int 			recPosTmp 		= 0;

	protected void jobExecutor() {
		JobExecutor jobSerializer = new JobExecutor();
		
		if (unattendedMode) {
			jobSerializer.executeProcess();
		} else {
			myAPBList.get(0).setLabelLeftCaption("Initializing...");

			jobSerializer.start();
		}
	}

	public void setDBConn(DBConnector appDBConn) {
		this.appDBConn = appDBConn;
	}

	public void setMMSPrjInfo(MMSProjectInfo prjMMSInfo) {
		this.prjMMSInfo = prjMMSInfo;
	}

}

