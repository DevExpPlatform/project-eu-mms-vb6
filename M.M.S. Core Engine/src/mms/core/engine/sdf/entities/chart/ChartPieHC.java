package mms.core.engine.sdf.entities.chart;

import java.awt.Color;
import java.awt.Font;
import java.awt.geom.Rectangle2D;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.block.BlockBorder;
import org.jfree.chart.plot.PiePlot;
import org.jfree.chart.title.LegendTitle;
import org.jfree.chart.title.TextTitle;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.ui.RectangleEdge;
import org.jfree.ui.RectangleInsets;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import com.itextpdf.awt.DefaultFontMapper;
import com.itextpdf.awt.FontMapper;
import com.itextpdf.text.pdf.BaseFont;

import mms.core.engine.pdfmerger.commons.CommonUtils;
import mms.core.engine.sdf.FontManager;

public class ChartPieHC implements ChartManager {

	private BaseFont 		  		fieldFont	  	= null;
	private float    		  		fieldFontSize 	= 8.0f;
	private DefaultFontMapper 		fntMapper 	  	= new DefaultFontMapper();
	private String 			  		fontsPath	  	= "";

	@Override
	public JFreeChart getChart(Element chartElement) {
		String[] 	  	chartData     	= null;
		Font 			chartFont 		= null;
		Color		  	chartFontColor	= new Color(0);
		int 		  	chartFontSize 	= 8;
		int 		  	chartFontStyle 	= java.awt.Font.PLAIN;
		boolean 		chartLegend		= false;
		RectangleEdge 	chartLegendPos	= RectangleEdge.RIGHT;
		String			chartTitle		= null;

		DefaultPieDataset defaultPieDataset = new DefaultPieDataset();
	    NodeList 		  pieDataset        = ((Element) chartElement).getElementsByTagName("dataset");

		if (this.fontsPath.equals("")) {
			chartFont = this.fntMapper.pdfToAwt(fieldFont, (int) fieldFontSize) ;
		} else {
			if (chartElement.hasAttribute("chartfont")) {
		        if (chartElement.hasAttribute("chartfontcolor")) {
					int[] intArray = CommonUtils.getIntegerArray(chartElement.getAttribute("chartfontcolor").split(","));
					chartFontColor = new Color(intArray[0], intArray[1], intArray[2]);
				}

				if (chartElement.hasAttribute("chartfontsize"))
					chartFontSize = Integer.valueOf(chartElement.getAttribute("chartfontsize"));

				if (chartElement.hasAttribute("chartfontstyle")) 
					chartFontStyle = getDecodedFontStyle(chartElement.getAttribute("chartfontstyle"));
				
				FontManager myFontManager = new FontManager();
				myFontManager.setFontBasePath(this.fontsPath);
				
				chartFont = this.fntMapper.pdfToAwt(myFontManager.getFont(chartElement.getAttribute("chartfont")).getBaseFont(), chartFontSize);
				chartFont = chartFont.deriveFont(chartFontStyle);
			} else {
				chartFont = this.fntMapper.pdfToAwt(fieldFont, (int) fieldFontSize) ;
			}
		}

		for (int datasetIdx = 0; datasetIdx < pieDataset.getLength(); datasetIdx++) {
			Element dataSetElement = (Element) pieDataset.item(datasetIdx);
			chartData = dataSetElement.getChildNodes().item(0).getNodeValue().split("\\|");
			defaultPieDataset.setValue(chartData[0], Float.valueOf(chartData[1]));
		}

		if (chartElement.hasAttribute("title")) 
			chartTitle = chartElement.getAttribute("title");

		if (chartElement.hasAttribute("legendpos")) {
			chartLegend = true;
			chartLegendPos = getDecodedLegendPos(chartElement.getAttribute("legendpos"));
		}

		/*
		 * Build Chart 
		 */
		JFreeChart chart = ChartFactory.createPieChart(chartTitle, defaultPieDataset, chartLegend, false, false);
	      
		PiePlot plot = (PiePlot) chart.getPlot();
		plot.setBackgroundPaint(new Color(255, 255, 255));
		plot.setBaseSectionOutlinePaint(chartFontColor);
		plot.setInsets(new RectangleInsets(0, 0, 0, 0));
		plot.setLabelLinksVisible(false);
		plot.setLabelGenerator(null); 
		plot.setLegendItemShape(new Rectangle2D.Double(0.0, 0.0, 6.0, 6.0));
		plot.setOutlineVisible(false);
		plot.setShadowPaint(null);

		/*
		 * Render Chart
		 */
        chart.setBackgroundPaint(new Color(255, 255, 255));
        
        TextTitle textTitle = chart.getTitle();

        if (textTitle != null) {
            textTitle.setFont(chartFont);
            textTitle.setPaint(chartFontColor); 
        }

        if (chartLegend) {
            LegendTitle legend = chart.getLegend();
            legend.setFrame(BlockBorder.NONE);
            legend.setItemFont(chartFont);
            legend.setItemPaint(chartFontColor);
            legend.setPosition(chartLegendPos);
        }

		return chart;
	}

	private int getDecodedFontStyle(String fontStyle) {
		int    rValue     = java.awt.Font.PLAIN;
		String strArray[] = fontStyle.split("\\|");

		for (int i = 0; i < strArray.length; i++) {
			int tmpFontStyle = 0;
			
			if (strArray[i].equals("bold")) {
				tmpFontStyle = java.awt.Font.BOLD;
			} else if (strArray[i].equals("italic")) {
				tmpFontStyle = java.awt.Font.ITALIC;
			}

			if (i == 0) {
				rValue = tmpFontStyle;
			} else {
				rValue |= tmpFontStyle;
			}
		}
		
		return rValue;
	}

	private RectangleEdge getDecodedLegendPos(String legendPos) {
		RectangleEdge rValue = RectangleEdge.TOP; 
		
		switch (legendPos.hashCode()) {
		case 115029:
			rValue = RectangleEdge.TOP; 
			
			break;
		case 3317767:
			rValue = RectangleEdge.LEFT;
			
			break;
		case 108511772:
			rValue = RectangleEdge.RIGHT;
			
			break;
		case -1383228885:
			rValue = RectangleEdge.BOTTOM;
			
			break;
		}
		
		return rValue;
	}

	@Override
	public FontMapper getFontMapper() {
		return fntMapper;
	}

	@Override
	public float getSizeW() {
		return 0;
	}

	@Override
	public void setFont(BaseFont fieldFont, float fieldFontSize) {
		this.fieldFont     = fieldFont;
		this.fieldFontSize = fieldFontSize;
	}

	@Override
	public void setFontsPath(String fontsPath) {
		this.fontsPath = fontsPath;

		if (!fontsPath.equals("")) 
			fntMapper.insertDirectory(fontsPath);
	}

}
