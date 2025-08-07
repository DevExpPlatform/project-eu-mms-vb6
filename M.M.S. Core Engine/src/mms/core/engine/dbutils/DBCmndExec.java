package mms.core.engine.dbutils;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Vector;

@SuppressWarnings("serial")
public class DBCmndExec extends DBConnector {

	private String				ErrorMsg 	= ""; 	
    private ResultSet 			RS			= null;
    private Statement 			Stmnt 		= null;
    private ResultSetMetaData 	RSMD 		= null; 
    private int 				NumRows 	= 0;
    
    public 	DBUtilsRtns		DBUtils		= new DBUtilsRtns();
    
    public DBCmndExec() {}

    public DBCmndExec(Connection extDBConn) {
    	dbConn = extDBConn;
    }
    
    public String getErrorMessage() {
    	return ErrorMsg;
    }
	
	public int getLastIdentity() {
    	int LastIdentity = -1;

    	ErrorMsg = "";

		try {
			Stmnt = dbConn.createStatement();
			RS = Stmnt.executeQuery("SELECT @@IDENTITY");

	        if (RS != null) { 
				while (RS.next()) 
				{
					LastIdentity = RS.getInt(1);
				}
	        }
	        
	        QueryClose();
		} catch (SQLException e) {
        	ErrorMsg = e.getMessage();
		}

		return LastIdentity;
	}

	public int getNumRows() {
		return NumRows;
	}

	public ResultSetMetaData getResultSetMetaData() {
		return RSMD;
	}

	public ResultSet Query(String SQLString) {
    	ErrorMsg  = "";
		NumRows   = 0;
		
		try {
    		Stmnt = dbConn.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
    		RS 	  = Stmnt.executeQuery(SQLString);
    		RSMD  = RS.getMetaData();

    		RS.last();
    		NumRows = RS.getRow();
            RS.beforeFirst();
            
    		return RS;
    	} catch (SQLException e) {
        	ErrorMsg = e.getMessage();

        	return null;
    	}
    }
   
	public ResultSet Query(String SQLString, int resultSetType, int resultSetConcurrency) {
    	ErrorMsg  = "";
		NumRows   = 0;
		
		try {
    		Stmnt = dbConn.createStatement(resultSetType, resultSetConcurrency);
    		RS 	  = Stmnt.executeQuery(SQLString);
    		RSMD  = RS.getMetaData();

    		RS.last();
    		NumRows = RS.getRow();
            RS.beforeFirst();
            
    		return RS;
    	} catch (SQLException e) {
        	ErrorMsg = e.getMessage();

        	return null;
    	}
    }

	public boolean QueryClose() {
		ErrorMsg = "";

		try {
			if (RS != null) {
				RS.close();
			}
    	
			if (Stmnt != null) {
				Stmnt.close();
			}
			
    		return true;
    	} catch (SQLException e) {
        	ErrorMsg = e.getMessage();

        	return false;
    	}
    }
    
	public boolean QueryCommand(String SQLString) {
		QueryClose();
	
		ErrorMsg = "";
	
		try {
			Stmnt = dbConn.createStatement();
			Stmnt.execute(SQLString);
			Stmnt.close();
	
			return true;
		} catch (SQLException e) {
	    	ErrorMsg = e.getMessage();
	
			return false;
		}
	}

	public Vector<String> QueryScalar(String SQLString) {
    	Vector<String> retValues = null;
		
		ErrorMsg 	= "";
		NumRows 	= 0;
		
		try {
    		RS = Query(SQLString);

    		if (NumRows > 0) {
    			retValues = new Vector<String>();
    			
    			while (RS.next()) {
            		String RowString = "";

            		for (int I = 1; I <= RSMD.getColumnCount(); I++) {
    					RowString += RS.getString(I) + ";";
    				}
        		
            		retValues.add(RowString);
        		}
    		}
    		
    		QueryClose();
    	} catch (SQLException e) {
        	ErrorMsg = e.getMessage();
        	
        	return null;
    	}
		
		return retValues;
	}
	
	public boolean QueryUpdate(String upd_Table, String upd_Fields, String upd_Values, String upd_WHERE) {
		boolean	  BValue 	= false;
		String[]  myFields  = null;
		Statement myStmnt	= null;
		String[]  myValues  = null;
		String 	  SQLString = "";
	
		ErrorMsg 			= "";
	
		myFields = upd_Fields.split(",");
		myValues = upd_Values.split(",");
		
		for (int I = 0; I < myFields.length; I++) {
			SQLString += (SQLString.equals("")? "": ", ") + myFields[I].trim() + " = " + myValues[I].trim();
		}
	
		SQLString = "UPDATE " + upd_Table + " SET " + SQLString + (upd_WHERE.equals("")? "": " WHERE " + upd_WHERE);
		
		try {
			myStmnt = dbConn.createStatement();
			myStmnt.execute(SQLString);
			myStmnt.close();
	
			BValue = true;
		} catch (SQLException e) {
	    	ErrorMsg = e.getMessage();
			BValue   = false;
		}
	
		return BValue;
	}

	public boolean TransactionCommit() {
		ErrorMsg = "";

		try {
			dbConn.commit();
			
			return true;
		} catch (SQLException e) {
        	ErrorMsg = e.getMessage();

        	return false;
		}
	}

    public boolean TransactionRollBack(){
    	ErrorMsg = "";

		try {
			dbConn.rollback();

			return true;
		} catch (SQLException e) {
        	ErrorMsg = e.getMessage();

        	return false;
		}
    }

}
