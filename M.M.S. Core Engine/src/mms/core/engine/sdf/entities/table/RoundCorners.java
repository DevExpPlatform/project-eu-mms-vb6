package mms.core.engine.sdf.entities.table;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPCellEvent;
import com.itextpdf.text.pdf.PdfPTable;

public class RoundCorners implements PdfPCellEvent {

    private boolean 	bottomRight	 = false;
    private boolean 	bottomLeft	 = false;
	private Float 		borderWidth	 = 0.5f;
	private BaseColor 	colorBack 	 = null;
	private BaseColor 	colorBorder  = null;
	private float 		cornerRadius = 4.0f;
    private boolean 	topLeft		 = false;
    private boolean 	topRight	 = false;

    public RoundCorners(BaseColor colorBorder, BaseColor colorBack, float borderWidth, float cornerRadius, boolean topLeft, boolean topRight, boolean bottomRight, boolean bottomLeft) {
        this.bottomRight  = bottomRight;
        this.bottomLeft   = bottomLeft;
        this.borderWidth  = borderWidth;
        this.colorBack    = colorBack;
        this.colorBorder  = colorBorder;
        this.cornerRadius = cornerRadius;
    	this.topLeft 	  = topLeft;
        this.topRight     = topRight;
    }

    @Override
    public void cellLayout(PdfPCell cell, Rectangle rect, PdfContentByte[] canvas) {
        float shift  = 0 ; //(this.borderWidth / 2);
    	float left   = (rect.getLeft() - shift);
        float top    = (rect.getTop() - shift);
        float right  = (rect.getRight() + shift);
        float bottom = (rect.getBottom() + shift);

        if (colorBack != null) {
            PdfContentByte cb = canvas[PdfPTable.BACKGROUNDCANVAS];
        	cb.setColorFill(this.colorBack);

	        if(topLeft) {
	        	cb.moveTo(left, top - this.cornerRadius);
	        	cb.curveTo(left, top, left + this.cornerRadius, top);
	        } else cb.moveTo(left, top);

	        if(topRight) {
	        	cb.lineTo(right - this.cornerRadius, top);
	        	cb.curveTo(right, top, right, top - this.cornerRadius);
	        } else cb.lineTo(right, top);

	        if(bottomRight) {
	        	cb.lineTo(right, bottom + this.cornerRadius);
	        	cb.curveTo(right, bottom, right - this.cornerRadius, bottom);
	        } else cb.lineTo(right, bottom);

	        if(bottomLeft) {
	        	cb.lineTo(left + this.cornerRadius, bottom);
	        	cb.curveTo(left, bottom, left, bottom + this.cornerRadius);
	        } else cb.lineTo(left, bottom);

	        if(topLeft) cb.lineTo(left, top - this.cornerRadius);
	        else cb.lineTo(left, top);

	        cb.closePath();
	        cb.fill();
        }
        
        PdfContentByte cb = canvas[PdfPTable.LINECANVAS];
        cb.setColorStroke(this.colorBorder);
        cb.setLineWidth(this.borderWidth);

        if (this.cornerRadius == 0.0f) {
            if(topLeft) {
                cb.moveTo(left, bottom);
                cb.lineTo(left, top);
            } else cb.moveTo(left, top);
        	
            if(topRight) {
                cb.lineTo(right, top);
            } else cb.moveTo(right, top);

            if(bottomRight) {
                cb.lineTo(right, bottom);
            } else cb.moveTo(right, bottom);

            if(bottomLeft) 
                cb.lineTo(left, bottom);
        } else {
            if(topLeft) {
                cb.moveTo(left, (bottomLeft ? bottom + this.cornerRadius : bottom));
                cb.lineTo(left, top - this.cornerRadius);
                cb.curveTo(left, top, left + this.cornerRadius, top);
                cb.lineTo((topRight ? right - this.cornerRadius : right), top);
            } else cb.moveTo(left, top);

            if(topRight) {
            	cb.lineTo(right - this.cornerRadius, top);
                cb.curveTo(right, top, right, top - this.cornerRadius);
                cb.lineTo(right, (bottomRight ? bottom + this.cornerRadius : bottom));
            } else cb.moveTo(right, top);

            if(bottomRight) {
                cb.lineTo(right, bottom + this.cornerRadius);
                cb.curveTo(right, bottom, right - this.cornerRadius, bottom);
                cb.lineTo((bottomLeft ? left + this.cornerRadius : left), bottom);
            } else cb.moveTo(right, bottom);

            if(bottomLeft) {
                cb.lineTo(left + this.cornerRadius, bottom);
                cb.curveTo(left, bottom, left, bottom + this.cornerRadius);
                cb.lineTo(left, (topLeft ? top - this.cornerRadius : top));
            }
        }
        
        cb.stroke();
    }

}