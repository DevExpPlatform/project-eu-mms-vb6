package mms.core.engine.pdfmerger.serializers;

import java.sql.ResultSet;
import java.sql.SQLException;

public abstract class CustomSerializer {

	protected int      packCntr = 0;
	protected String[] params   = null;
	protected int      posCntr  = 0;

	public String getSequence() {
		return null;
	}

	public void setPackage(int packCntr) {
		this.packCntr = packCntr;
	}

	public void setParams(String[] params) {
		this.params = params;
	}

	public void setPos(int posCntr) {
		this.posCntr = posCntr;
	}

	public void setRS(ResultSet rS) throws SQLException {}

}
