package mms.core.engine.sdf.entities.chart;

import java.awt.Color;
import java.awt.geom.Rectangle2D;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PiePlot;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.ui.RectangleInsets;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import com.itextpdf.awt.DefaultFontMapper;
import com.itextpdf.awt.FontMapper;
import com.itextpdf.text.pdf.BaseFont;

public class ChartPie implements ChartManager {

//	private BaseFont 		  fieldFont	    = null;
//	private float    		  fieldFontSize = 8.0f;
	private DefaultFontMapper fntMapper 	= new DefaultFontMapper();
//	private String 			  fontsPath		= "";

	@Override
	public JFreeChart getChart(Element chartElement) {
		String[] 	  	  chartData     	= null;
		boolean 	  	  chartLegend		= false;
		String		   	  chartTitle		= null;
		DefaultPieDataset defaultPieDataset = new DefaultPieDataset();
		NodeList 		  pieDataset        = ((Element) chartElement).getElementsByTagName("dataset");

		for (int datasetIdx = 0; datasetIdx < pieDataset.getLength(); datasetIdx++) {
			Element dataSetElement = (Element) pieDataset.item(datasetIdx);
			chartData = dataSetElement.getChildNodes().item(0).getNodeValue().split("\\|");
			defaultPieDataset.setValue(chartData[0], Float.valueOf(chartData[1]));
		}

		JFreeChart pieChart = ChartFactory.createPieChart(chartTitle, defaultPieDataset, chartLegend, false, false);

		PiePlot plot = (PiePlot) pieChart.getPlot();
		plot.setBackgroundPaint(new Color(255, 255, 255));
		//plot.setBaseSectionOutlinePaint(chartFontColor);
		plot.setInsets(new RectangleInsets(0, 0, 0, 0));
		plot.setLabelLinksVisible(false);
		plot.setLabelGenerator(null); 
		plot.setLegendItemShape(new Rectangle2D.Double(0.0, 0.0, 6.0, 6.0));
		plot.setOutlineVisible(false);
		plot.setShadowPaint(null);

		return pieChart;
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
		//this.fieldFont     = fieldFont;
		//this.fieldFontSize = fieldFontSize;
	}

	@Override
	public void setFontsPath(String fontsPath) {
		//this.fontsPath = fontsPath;

		if (!fontsPath.equals("")) 
			fntMapper.insertDirectory(fontsPath);
	}

}
