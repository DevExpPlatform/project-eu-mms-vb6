package mms.core.engine.dbutils;

import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;

public class DBUtilsRtns {

	private String errMsg = "";

	public String getErrMsg() {
		return errMsg;
	}

	public HashMap<Integer, HashMap<String, String>> getResultSetHashMap(ResultSet resultSet) {
		errMsg = "";
		
		HashMap<Integer, HashMap<String, String>> rValue = null;

		if (resultSet != null) {
			rValue = new HashMap<Integer, HashMap<String,String>>();
			
			try {
				ResultSetMetaData metaData = resultSet.getMetaData();
				int 			  rowCntr  = 0;
		
				while (resultSet.next()) {
					HashMap<String, String> rowData = new HashMap<String, String>();
		
					for (int i = 1; i <= metaData.getColumnCount(); i++) 
						rowData.put(metaData.getColumnName(i), ((resultSet.getString(i) == null) ? "" : resultSet.getString(i)));
		
					rValue.put(rowCntr, rowData);
		
					rowCntr++;
				}
				
				if (rowCntr == 0) 
					rValue = null;
			} catch (SQLException e) {
				rValue = null;
				errMsg = e.getMessage();
			}
		}
		
		return rValue;
	}

	public int getResultSetRowCount(ResultSet resultSet) {
	    int size = 0;
	    
	    try {
	        resultSet.last();
	        size = resultSet.getRow();
	        resultSet.beforeFirst();
	    }
	    catch(Exception ex) {
	        return 0;
	    }
	    
	    return size;
	}

	public String getStringFormat(int Mode, String InputString, String Filler, int FinalLenght) {
		int FillLenght = FinalLenght - InputString.length();
		String tmpString = "";
	
		for (int I = 0; I < (FillLenght); I++)
			tmpString += Filler;
	
		if (Mode == 0) {
			tmpString = InputString + tmpString;
		} else {
			tmpString += InputString;
		}
	
		return tmpString; 
	}

	public String getSQLDateTime(String InputDate) {
		if (InputDate.trim().equals("")) {
			return "NULL";
		} else {
			String 			 tmpString 		= "";
			SimpleDateFormat myDateFormat 	= new SimpleDateFormat("yyyy-MM-dd" + (InputDate.length() > 10 ? " HH:mm:ss" : ""));

			try {
				Date myDate = new SimpleDateFormat("dd/MM/yyyy").parse(InputDate);

				tmpString = "CONVERT(DATETIME, '" + myDateFormat.format(myDate) + "', 102)";
			} catch (ParseException e) {
			}

			return tmpString;
		}
	}

	public String getSQLDateTimeORACLE(Date InputDate) {
		if (InputDate.equals(null)) {
			return "NULL";
		} else {
			return "TO_DATE('" + new SimpleDateFormat("dd/MM/yyyy HH.mm.ss").format(InputDate) + "', 'DD/MM/YYYY HH24:MI:SS')";
		}
	}

	public String getSQLString(String InputValue, boolean WithQuotes) {
		String TmpString = "";

		if (InputValue.trim() == "") {
			TmpString = "NULL";
		} else {
			TmpString = InputValue.replace("'", "''").trim();

			if (WithQuotes)
				TmpString = "'" + TmpString + "'";
		}

		return TmpString;
	}

	public String getSQLWhereClause(int SrchModeIdx, int SrchTypeIdx, String SrchField, String SrchFieldValue) {
		if (!SrchFieldValue.trim().equals("")) {
			String SQLString = "";

			if (SrchModeIdx == 0) {
				SQLString = " AND";
			} else if (SrchModeIdx == 1) {
				SQLString = " OR";
			}

			SQLString 		+= " (" + SrchField;
			SrchFieldValue	= getSQLString(SrchFieldValue, false);

			switch (SrchTypeIdx) {
			case 0:	// "comincia per"
				SQLString += " LIKE '" + SrchFieldValue + "%'";

				break;
			case 1:	// "contiene"
				SQLString += " LIKE '%" + SrchFieldValue + "%'";

				break;
			case 2:	// "diverso da"
				SQLString += " NOT LIKE '" + SrchFieldValue + "'";

				break;
			case 3:	// "finisce per"
				SQLString += " LIKE '%" + SrchFieldValue + "'";

				break;
			case 4: // "uguale a"
				SQLString += " LIKE '" + SrchFieldValue + "'";

				break;
			}

			SQLString += ")";

			return SQLString;
		} else {
			return "";
		}
	}

}
