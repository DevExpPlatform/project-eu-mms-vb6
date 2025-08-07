package mms.core.engine.sdf.entities.text;

import java.io.IOException;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Chunk;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.BaseFont;

public class SDFText {

	private String 	  errMsg 		= "";
	private BaseFont  font			= null;
	private String 	  fontBasePath 	= "";
	private BaseColor fontColor 	= new BaseColor(0, 0, 0);
	private BaseColor fontColorBack = null;
	private String 	  fontName		= "";
	private float 	  fontSize		= 12;
	private int 	  fontStyle 	= Font.NORMAL;
	private Paragraph paragraph 	= null;
	private boolean   underline		= false;

	public SDFText() {
		paragraph = new Paragraph();
	}

	public void addChunk(String chunkText) {
		boolean addNewLine = false;
		Font    myFont 	   = null;

		if (fontName.equals("")) {
			myFont = new Font(font, fontSize, fontStyle, fontColor);
		} else {
			myFont = getFont(fontName);
		}

		if (myFont == null)
			myFont = FontFactory.getFont(FontFactory.HELVETICA, fontSize, fontStyle, fontColor);

		if (chunkText.endsWith("<br>")) {
			chunkText = chunkText.substring(0, chunkText.length() - 4);
			addNewLine = true;
		}

		boolean bValue = getChunkSymbol(chunkText, "[$SIGMA]", "s", myFont) || getChunkTabbed(chunkText, myFont, 28f);

		if (!bValue) {
			Chunk chunk = new Chunk(chunkText, myFont);

			if (underline) {
				underline = false;

				chunk.setUnderline(0.6f, -2f);
			}

			if (this.fontColorBack != null)
				chunk.setBackground(this.fontColorBack);

			paragraph.add(chunk);
		}

		if (addNewLine)
			paragraph.add(Chunk.NEWLINE);
	}

	public void clear() {
		paragraph.clear();
	}

	private boolean getChunkSymbol(String chunkText, String tag, String symbolChar, Font chunkFont) {
		if (chunkText.contains(tag)) {
			Chunk chunk 	= null;
			Font  fntSymbol = getFont("symbol");
			int   idxEnd 	= 0;
			int   idxNew 	= 0;
			int   idxOld 	= 0;
			int   itmLen 	= tag.length();

			while(idxNew != -1){
				idxNew = chunkText.indexOf(tag, idxNew);

				if (idxNew != -1) {
					if (idxNew != 0) {
						chunk = new Chunk(chunkText.substring(idxOld, idxNew), chunkFont);

						if (underline)
							chunk.setUnderline(0.6f, -2f);

						paragraph.add(chunk);
					}

					/**/
					idxEnd = (idxNew + itmLen);
					chunk  = new Chunk(symbolChar, fntSymbol);

					if (underline)
						chunk.setUnderline(0.6f, -2f);

					paragraph.add(chunk);

					idxNew = idxEnd;
					idxOld = idxEnd;
				} else {
					chunk = new Chunk(chunkText.substring(idxOld), chunkFont);

					if (underline)
						chunk.setUnderline(0.6f, -2f);

					paragraph.add(chunk);
				}
			}

			if (underline)
				underline = false;

			return true;
		}

		return false;
	}

	private boolean getChunkTabbed(String chunkText, Font chunkFont, float f) {
		if (chunkText.contains("[$TAB]")) {
			Chunk chunk 	= null;
			int   idxEnd 	= 0;
			int   idxNew 	= 0;
			int   idxOld 	= 0;
			int   itmLen 	= 6;

			while(idxNew != -1){
				idxNew = chunkText.indexOf("[$TAB]", idxNew);

				if (idxNew != -1) {
					if (idxNew != 0) {
						chunk = new Chunk(chunkText.substring(idxOld, idxNew), chunkFont);

						if (underline)
							chunk.setUnderline(0.6f, -2f);

						paragraph.add(chunk);
					}

					/**/
					idxEnd = (idxNew + itmLen);

					paragraph.add(Chunk.createTabspace(f));

					idxNew = idxEnd;
					idxOld = idxEnd;
				} else {
					chunk = new Chunk(chunkText.substring(idxOld), chunkFont);

					paragraph.add(chunk);
				}
			}

			return true;
		}

		return false;
	}

	public String getErrMsg() {
		return errMsg;
	}

	private Font getFont(String fontDescr) {
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

	public float getFontSize() {
		return paragraph.getFont().getSize();
	}

	public Paragraph getText() {
		return paragraph;
	}

	public void setFont(BaseFont font) {
		this.font = font;
	}

	public void setFont(String fontname) {
		this.fontName = fontname;
	}

	public void setFontBasePath(String fontBasePath) {
		this.fontBasePath = fontBasePath;
	}

	public void setFontColor(int red, int green, int blue) {
		this.fontColor = new BaseColor(red, green, blue);
	}

	public void setFontColorBack(int red, int green, int blue) {
		this.fontColorBack = new BaseColor(red, green, blue);
	}

	public void setFontSize(float fontSize) {
		this.fontSize = fontSize;
	}

	public void setFontStyle(String fontStyle) {
		String strArray[] = fontStyle.split(",");

		for (int i = 0; i < strArray.length; i++) {
			int tmpFontStyle = 0;

			if (strArray[i].trim().equals("bold")) {
				tmpFontStyle = Font.BOLD;
			} else if (strArray[i].trim().equals("italic")) {
				tmpFontStyle = Font.ITALIC;
			} else if (strArray[i].trim().equals("underline")) {
				this.underline = true;
			} else if (strArray[i].trim().equals("strikethru")) {
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
