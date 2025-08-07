package mms.core.engine.pdfmerger.serializers.codegen;

import java.sql.ResultSet;
import java.sql.SQLException;

import mms.core.engine.pdfmerger.serializers.CustomSerializer;

public class OMR_PBB extends  CustomSerializer {

	private int nmrPages = 0;

	@Override
	public String getSequence() {
		if (this.nmrPages == 0) {
			return "51";
		} else {
			String rValue = "";
			
			for (int i = 0; i <= this.nmrPages; i++) {
				if (i == 0) {
					rValue = "39" + this.params[3];
				} else if (i == nmrPages) {
					rValue += "43" + this.params[3];
				} else {
					rValue += "63" + this.params[3];
				}
			}

			return rValue.substring(0, (rValue.length() - 4));
		}
	}

	@Override
	public void setRS(ResultSet rS) throws SQLException {
		this.nmrPages = ((rS.getInt(this.params[1]) / 2) - 1);
	}

}
