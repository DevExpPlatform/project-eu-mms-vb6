package mms.core.engine.dbutils;

import java.io.Serializable;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

import javax.naming.Context;
import javax.naming.InitialContext;
import javax.naming.NamingException;

import oracle.jdbc.pool.OracleDataSource;

public class DBConnector implements Serializable { 

	public enum DBType {
		JNDI, MSACCESS, MSSQL, MYSQL, ODBC, ORACLE, ORACLE_CUSTOM
	}

	private static final long serialVersionUID = -8443096010094793553L;

	private String 		connectionString 	= "";
	private String 		dataBase	  		= "";
	private String 		errorMsg   			= ""; 	
	private DBType 		serverType 			= DBType.MSSQL;
	private String 		serverName 			= "";
	private String 		serverPort 			= "1521";
	private String 		userId	  			= "";
	private String 		userPwd	  			= "";

	public  Connection 	dbConn 				= null;

	public boolean close(){
		try{
			if(this.dbConn != null) {
				this.dbConn.close();
				this.dbConn = null;
			}

			return true;
		}catch(Exception e){
			errorMsg = e.getMessage();

			return false;
		}
	}

	private boolean getConnJDBC(String connString) {
		try {
			this.dbConn = DriverManager.getConnection(connString, userId, userPwd);
	
			return (this.dbConn != null);
		} catch (SQLException e) {
			this.errorMsg = e.getMessage();
	
			return false;
		} 
	}

	private boolean getConnJNDI(String resourceName) {
		try {
			Context 		 initContext = new InitialContext();
			Context 		 envContext  = (Context) initContext.lookup("java:/comp/env");
			OracleDataSource oracleDS    = (OracleDataSource) envContext.lookup(resourceName);
	
			if (oracleDS != null)
				this.dbConn = oracleDS.getConnection();
		} catch (NamingException e) {
			this.errorMsg = e.getMessage();
		} catch (SQLException e) {
			this.errorMsg = e.getMessage();
		}
			
		return (this.dbConn != null);
	}

	public String getErrorMessage() {
		return errorMsg;
	}

	public boolean open() {
		boolean rValue = false;

		this.errorMsg = "";

		if (dbConn == null) {
			switch(this.serverType) {
			case JNDI:
				rValue = getConnJNDI(this.serverName);
				
				break;
			case MSACCESS:
				rValue = getConnJDBC("jdbc:odbc:Driver={Microsoft Access Driver (*.mdb)};DBQ=" + this.serverName + ";DriverID22");

				break;
			case MSSQL:
				rValue = getConnJDBC("jdbc:jtds:sqlserver://" + this.serverName + ":" + this.serverPort + "/" + this.dataBase);
				
				break;
			case MYSQL:
				rValue = getConnJDBC("jdbc:mysql://" + this.serverName + "/" + this.dataBase);
				
				break;
			case ODBC:
				rValue = getConnJDBC("jdbc:odbc:" + this.serverName);

				break;
			case ORACLE:
				rValue = getConnJDBC("jdbc:oracle:thin:@//" + this.serverName + ":" + this.serverPort + "/" + this.dataBase);
				
				break;
			case ORACLE_CUSTOM:
				rValue = getConnJDBC(this.connectionString);
				
				break;
			}
		} else {
			rValue = true;
		}

		return rValue;
	}

	public void setConnectionString(String connectionString) {
		this.connectionString = connectionString;
	}

	public void setDataBase(String DBName) {
		this.dataBase = DBName;
	}

	public void setServerName(String SrvrName) {
		this.serverName = SrvrName; 
	}

	public void setServerPort(String dbServerPort) {
		this.serverPort = dbServerPort;
	}

	public void setServerType(DBType SrvrType) {
		this.serverType = SrvrType;
	}

	public boolean setTransactional(boolean Trnsctnl) {
		boolean BValue = false;

		try {
			this.dbConn.setAutoCommit(!Trnsctnl);

			BValue = true;
		} catch (SQLException e) {
			this.errorMsg = e.getMessage();
		}

		return BValue;
	}

	public void setUserId(String UsrId) {
		this.userId = UsrId; 
	}

	public void setUserPwd(String UsrPwd) {
		this.userPwd = UsrPwd; 
	}

	public boolean test() {
		boolean rValue = false;

		open();
		rValue = (this.dbConn != null);
		close();

		return rValue;
	}

}
