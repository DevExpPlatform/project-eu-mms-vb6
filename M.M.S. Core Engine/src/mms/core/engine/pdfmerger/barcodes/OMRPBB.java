package mms.core.engine.pdfmerger.barcodes;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.ExceptionConverter;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.Utilities;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfTemplate;

public class OMRPBB {

	private static final byte BARS[][] = { 
			{ 0, 0, 0, 0, 0, 0 },	// No Code
			{ 1, 1, 1, 0, 0, 1 }, 	// 39
			{ 1, 1, 0, 1, 0, 1 }, 	// 43
			{ 1, 1, 0, 0, 1, 1 }, 	// 51
			{ 1, 1, 1, 1, 1, 1 } 	// 63
	};

	private float  barsThickness	= 1.0f;
	private int    code				= 0;
	private float  boxWidth 		= Utilities.millimetersToPoints(12.0f);
	private float  boxHeight		= Utilities.millimetersToPoints(15.0f);

	public Image createImageWithBarcode(PdfContentByte cb, BaseColor barColor) {
		PdfTemplate tp   = cb.createTemplate(0, 0);
		Rectangle   rect = placeBarcode(tp, barColor);

		tp.setBoundingBox(rect);

		try {
			return Image.getInstance(tp);
		} catch (Exception e) {
			throw new ExceptionConverter(e);
		}
	}

	private byte[] getOMRCode() {
		byte bars[] = BARS[0];

		switch (this.code) {
		case 39:
			bars = BARS[1];

			break;
		case 43:
			bars = BARS[2];

			break;
		case 51:
			bars = BARS[3];

			break;
		case 63:
			bars = BARS[4];

			break;
		}

		return bars;
	}

	private Rectangle placeBarcode(PdfContentByte cb, BaseColor barColor) {
		byte  bars[] 	 = getOMRCode(); 
		int   barItems 	 = (bars.length - 1);
		float barStartY  = 0.0f;
		float barsHeight = (this.barsThickness * bars.length);
		float barsRange  = ((this.boxHeight - barsHeight) / barItems);

		if (barColor != null)
			cb.setColorFill(barColor);

		for (int i = barItems; i >= 0; --i) {
			if (bars[i] == 1)
				cb.rectangle(0, barStartY, this.boxWidth, this.barsThickness);

			barStartY += (barsRange + this.barsThickness);
		}

		cb.fill();

		return new Rectangle(this.boxWidth, this.boxHeight);
	}

	public void setBarsThickness(float barsThickness) {
		this.barsThickness = Utilities.millimetersToPoints(barsThickness);
	}

	public void setBarsWidthMM(float barsWidthMM) {
		this.boxWidth = Utilities.millimetersToPoints(barsWidthMM);
	}

	public void setBoxHeight(float boxHeight) {
		this.boxHeight = boxHeight;
	}

	public void setBoxWidth(float boxWidth) {
		this.boxWidth = boxWidth;
	}

	public void setCode(int code) {
		this.code = code;
	}

}
