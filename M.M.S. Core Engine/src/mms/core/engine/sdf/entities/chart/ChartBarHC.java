package mms.core.engine.sdf.entities.chart;

import java.awt.Color;
import java.awt.Font;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.CategoryAxis;
import org.jfree.chart.axis.CategoryLabelPositions;
import org.jfree.chart.axis.ValueAxis;
import org.jfree.chart.block.BlockBorder;
import org.jfree.chart.labels.ItemLabelAnchor;
import org.jfree.chart.labels.ItemLabelPosition;
import org.jfree.chart.labels.StandardCategoryItemLabelGenerator;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.renderer.category.StandardBarPainter;
import org.jfree.chart.title.LegendTitle;
import org.jfree.chart.title.TextTitle;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.ui.RectangleEdge;
import org.jfree.ui.RectangleInsets;
import org.jfree.ui.TextAnchor;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import com.itextpdf.awt.DefaultFontMapper;
import com.itextpdf.awt.FontMapper;
import com.itextpdf.text.pdf.BaseFont;

import mms.core.engine.pdfmerger.commons.CommonUtils;
import mms.core.engine.sdf.FontManager;

public class ChartBarHC implements ChartManager {

	private float 					chartSizeW		= 0.0f;
	private BaseFont 		  		fieldFont	  	= null;
	private float    		  		fieldFontSize 	= 8.0f;
	private DefaultFontMapper 		fntMapper 	  	= new DefaultFontMapper();
	private String 			  		fontsPath	  	= "";

	public ChartBarHC(float chartSizeW) {
		this.chartSizeW = chartSizeW;
	}

	@Override
	public JFreeChart getChart(Element chartElement) {
		Color		  			barsColor				= new Color(0);
		String[] 	  			chartData     			= null;
		Font 					chartFont 				= null;
		Color		  			chartFontColor			= new Color(0);
		int 		  			chartFontSize 			= 8;
		int 		  			chartFontStyle 			= java.awt.Font.PLAIN;
		boolean 	  			chartLegend				= false;
		RectangleEdge 			chartLegendPos			= RectangleEdge.RIGHT;
		String		   			chartTitle				= null;
        DefaultCategoryDataset 	defaultCategoryDataset	= new DefaultCategoryDataset();
	    NodeList 			   	categoryDataset 		= ((Element) chartElement).getElementsByTagName("dataset");
	    float 				   	categoryMargin 		  	= 0.1f;
	    float 				   	categoryMarginUp 	  	= 0.02f;
	    float 				   	categoryMarginLow 	  	= 0.02f;

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
		
		if (chartElement.hasAttribute("title")) 
			chartTitle = chartElement.getAttribute("title");

		if (chartElement.hasAttribute("legendpos")) {
			chartLegend = true;
			chartLegendPos = getDecodedLegendPos(chartElement.getAttribute("legendpos"));
		}

		Font tickLabelFont = chartFont.deriveFont(chartFontSize * 0.75f);
		Font itemLabelFont = chartFont.deriveFont(chartFontSize * 0.85f);

		for (int datasetIdx = 0; datasetIdx < categoryDataset.getLength(); datasetIdx++) {
			Element dataSetElement = (Element) categoryDataset.item(datasetIdx);
			
			chartData = dataSetElement.getChildNodes().item(0).getNodeValue().split("\\|");
			
			defaultCategoryDataset.setValue(Float.valueOf(chartData[0]), chartData[1], chartData[2]);
		}

		if (chartElement.hasAttribute("maxbars")) {
			int   maxBars     = Integer.valueOf(chartElement.getAttribute("maxbars"));
			int   numBars     = categoryDataset.getLength();
			float singleBarW  = ((this.chartSizeW - (this.chartSizeW * 0.04f) - (this.chartSizeW * (0.01f * (maxBars - 1)))) / maxBars); 
	        float numBarsW    = (singleBarW * numBars) + (this.chartSizeW * 0.04f) + (this.chartSizeW * (0.01f * (numBars - 1))); 
         	float barsShift   = this.chartSizeW / numBarsW;
         	
         	categoryMargin    = (0.01f * (numBars - 1) *  barsShift);
         	categoryMarginLow = (categoryMarginLow * barsShift);
         	categoryMarginUp  = (categoryMarginUp * barsShift);
         	
         	this.chartSizeW   = numBarsW;
		}

		if (chartElement.hasAttribute("barscolor")) {
			int[] intArray = CommonUtils.getIntegerArray(chartElement.getAttribute("barscolor").split(","));
			barsColor = new Color(intArray[0], intArray[1], intArray[2]);
		}

		/*
		 * Build Chart 
		 */
		JFreeChart chart = ChartFactory.createBarChart(chartTitle, null, null, defaultCategoryDataset, PlotOrientation.VERTICAL, chartLegend, false, false);
        
        CategoryPlot p = chart.getCategoryPlot(); 
        p.setAxisOffset(new RectangleInsets(0, 0, 0, 0));
        p.setRangeGridlinesVisible(false);
        p.setBackgroundPaint(new Color(255, 255, 255));
        p.setInsets(new RectangleInsets(0, 0, 0, 0));
        p.setOutlineVisible(false);
        
        CategoryAxis domainAxis = p.getDomainAxis();
        domainAxis.setAxisLineVisible(false);
        domainAxis.setCategoryLabelPositions(CategoryLabelPositions.createUpRotationLabelPositions(Math.PI/3.5f)); 
        domainAxis.setCategoryLabelPositionOffset(0);
        domainAxis.setCategoryMargin(categoryMargin);
        domainAxis.setLowerMargin(categoryMarginLow);
        domainAxis.setMaximumCategoryLabelLines(2);
        domainAxis.setTickMarksVisible(false);
        domainAxis.setTickLabelFont(tickLabelFont);
        domainAxis.setTickLabelInsets(new RectangleInsets(0, 0, 0, 0));
        domainAxis.setTickLabelPaint(chartFontColor);
        domainAxis.setUpperMargin(categoryMarginUp);

        ValueAxis rangeAxis = p.getRangeAxis();
        rangeAxis.setAxisLineVisible(false);
        rangeAxis.setTickLabelsVisible(false);
        rangeAxis.setTickMarksVisible(false);
        rangeAxis.setUpperMargin(0.2);
        
        BarRenderer renderer = (BarRenderer) p.getRenderer();
        renderer.setBarPainter(new StandardBarPainter());
        renderer.setBaseItemLabelFont(itemLabelFont);
        renderer.setBaseItemLabelPaint(chartFontColor);
        renderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator());
        renderer.setBaseItemLabelsVisible(true);
        renderer.setItemLabelAnchorOffset(0.0);
        renderer.setSeriesPaint(0, barsColor);
        
        ItemLabelPosition p2 = new ItemLabelPosition(ItemLabelAnchor.OUTSIDE6, TextAnchor.TOP_CENTER);
        renderer.setBaseNegativeItemLabelPosition(p2);

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
		return this.chartSizeW;
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
