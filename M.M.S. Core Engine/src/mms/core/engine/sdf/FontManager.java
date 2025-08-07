package mms.core.engine.sdf;

import java.io.IOException;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.pdf.BaseFont;

public class FontManager {

	private String 	  errMsg 		= "";
	private String 	  fontBasePath 	= "";
	private BaseColor fontColor 	= new BaseColor(0, 0, 0);  
	private float 	  fontSize		= 12;
	private int 	  fontStyle 	= Font.NORMAL;
	
	public Font getDefaultFont() {
		return FontFactory.getFont(FontFactory.COURIER, 8, Font.NORMAL, new BaseColor(0, 0, 0));
	}
	
	public String getErrMsg() {
		return errMsg;
	}

	public Font getFont(String fontDescr) {
		Font rValue = null;
		
		if (fontDescr.equals("")) {
			rValue = FontFactory.getFont(FontFactory.HELVETICA, fontSize, fontStyle, fontColor);
		} else if (fontDescr.equals("courier")) {
			rValue = FontFactory.getFont(FontFactory.COURIER, fontSize, fontStyle, fontColor);
		} else if (fontDescr.equals("courier_bold")) {
			rValue = FontFactory.getFont(FontFactory.COURIER_BOLD, fontSize, fontStyle, fontColor);
		} else if (fontDescr.equals("courier_boldoblique")) {
			rValue = FontFactory.getFont(FontFactory.COURIER_BOLDOBLIQUE, fontSize, fontStyle, fontColor);
		} else if (fontDescr.equals("courier_oblique")) {
			rValue = FontFactory.getFont(FontFactory.COURIER_OBLIQUE, fontSize, fontStyle, fontColor);
		} else if (fontDescr.equals("helvetica")) {
			rValue = FontFactory.getFont(FontFactory.HELVETICA, fontSize, fontStyle, fontColor);
		} else if (fontDescr.equals("helvetica_bold")) {
			rValue = FontFactory.getFont(FontFactory.HELVETICA_BOLD, fontSize, fontStyle, fontColor);
		} else if (fontDescr.equals("helvetica_boldoblique")) {
			rValue = FontFactory.getFont(FontFactory.HELVETICA_BOLDOBLIQUE, fontSize, fontStyle, fontColor);
		} else if (fontDescr.equals("helvetica_oblique")) {
			rValue = FontFactory.getFont(FontFactory.HELVETICA_OBLIQUE, fontSize, fontStyle, fontColor);
		} else if (fontDescr.equals("symbol")) {
			rValue = FontFactory.getFont(FontFactory.SYMBOL, fontSize, fontStyle, fontColor);
		} else if (fontDescr.equals("times_bold")) {
			rValue = FontFactory.getFont(FontFactory.TIMES_BOLD, fontSize, fontStyle, fontColor);
		} else if (fontDescr.equals("times_bolditalic")) {
			rValue = FontFactory.getFont(FontFactory.TIMES_BOLDITALIC, fontSize, fontStyle, fontColor);
		} else if (fontDescr.equals("times_italic")) {
			rValue = FontFactory.getFont(FontFactory.TIMES_ITALIC, fontSize, fontStyle, fontColor);
		} else if (fontDescr.equals("times_roman")) {
			rValue = FontFactory.getFont(FontFactory.TIMES_ROMAN, fontSize, fontStyle, fontColor);
		} else if (fontDescr.equals("zapfdingbats")) {
			rValue = FontFactory.getFont(FontFactory.ZAPFDINGBATS, fontSize, fontStyle, fontColor);
		} else {
			errMsg = "";
	
			try {
				BaseFont bfFont = BaseFont.createFont(fontBasePath + fontDescr, BaseFont.WINANSI, BaseFont.EMBEDDED);
	
				rValue = new Font(bfFont, fontSize, fontStyle, fontColor);
			} catch (DocumentException e) {
				errMsg = e.getMessage();
				rValue = null;
			} catch (IOException e) {
				errMsg = e.getMessage();
				rValue = null;
			}
		}
	
		return rValue;
	}

	public void setFontBasePath(String fontBasePath) {
		this.fontBasePath = fontBasePath;
	}

	public void setFontColor(int red, int green, int blue) {
		this.fontColor = new BaseColor(red, green, blue);
	}

	public void setFontSize(float fontSize) {
		this.fontSize = fontSize;
	}

	public void setFontStyle(String fontStyle) {
		String strArray[] = fontStyle.split("\\|");

		for (int i = 0; i < strArray.length; i++) {
			int tmpFontStyle = 0;
			
			if (strArray[i].equals("bold")) {
				tmpFontStyle = Font.BOLD;
			} else if (strArray[i].equals("italic")) {
				tmpFontStyle = Font.ITALIC;
			} else if (strArray[i].equals("underline")) {
				tmpFontStyle = Font.UNDERLINE;
			} else if (strArray[i].equals("strikethru")) {
				tmpFontStyle = Font.STRIKETHRU;
			}

			if (i == 0) {
				this.fontStyle = tmpFontStyle;
			} else {
				this.fontStyle |= tmpFontStyle;
			}
		}
	}

}
