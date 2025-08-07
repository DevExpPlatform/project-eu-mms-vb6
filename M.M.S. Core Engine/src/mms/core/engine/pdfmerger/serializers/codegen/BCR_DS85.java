package mms.core.engine.pdfmerger.serializers.codegen;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DecimalFormat;

import mms.core.engine.pdfmerger.serializers.CustomSerializer;

public class BCR_DS85 extends  CustomSerializer {

	private int cntrSequence = 0;
	private int nmrPages 	 = 0;
	
	@Override
	public String getSequence() {
		if (this.nmrPages == -1)
			return null;
		
		String rValue = "";
		
		for (int i = 0; i <= this.nmrPages; i++) {
			rValue += new DecimalFormat("###000").format(this.cntrSequence++) + ((i == this.nmrPages) ? "1" : "2") + this.params[3];

			if (this.cntrSequence == 1000)
				this.cntrSequence = 0;
		}
		
		return rValue.substring(0, (rValue.length() - 4));
	}

	@Override
	public void setRS(ResultSet rS) throws SQLException {
		if (rS.getString(params[2]) == null) {
			this.nmrPages = -1;
		} else {
			this.nmrPages = ((rS.getInt(params[1]) / 2) - 1);
		}
	}

}
