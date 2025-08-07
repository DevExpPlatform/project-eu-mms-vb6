package mms.core.engine.pdfmerger.serializers.codegen;

import java.security.SecureRandom;
import java.sql.ResultSet;
import java.sql.SQLException;

import mms.core.engine.pdfmerger.serializers.CustomSerializer;

public class AIMS_500 extends CustomSerializer {

	private String   alphabet     = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
	private String   idJob        = null;
	private int 	 nmrPages     = 0;
	private long     seqEnvelope  = 0;
	private int 	 seqSheet     = 0;
	private String[] subParams    = null;

	private String getRandomString(int length) {
		SecureRandom  sr = new SecureRandom();
		StringBuilder sb = new StringBuilder(length);

		for (int i = 0; i < length; i++)
			sb.append(this.alphabet.charAt(sr.nextInt(this.alphabet.length())));

		return new String(sb);
	}

	@Override
	public String getSequence() {
		if (this.nmrPages == -1)
			return null;

		String rValue = "";

		this.seqEnvelope++;

		for (int i = 0; i <= this.nmrPages; i++) {
			rValue += String.format("%03d", this.seqSheet++) + ((i == this.nmrPages) ? "1" : "2") + this.idJob + String.format("%010d", this.seqEnvelope) + this.subParams[1];

			if (this.seqSheet == 1000)
				this.seqSheet = 0;
		}

		if (this.seqEnvelope == 9999999999L)
			this.seqEnvelope = 0;

		return rValue.substring(0, (rValue.length() - 4));
	}

	@Override
	public void setParams(String[] params) {
		super.setParams(params);
	}

	@Override
	public void setRS(ResultSet rS) throws SQLException {
		if (rS.getString(params[2]) == null) {
			this.nmrPages = -1;
		} else {
			this.subParams = this.params[3].split(",");

			if (this.idJob == null)
				this.idJob = this.subParams[0] + "_" + this.getRandomString(7);

			this.nmrPages = ((rS.getInt(this.params[1]) / 2) - 1);
		}
	}

}
