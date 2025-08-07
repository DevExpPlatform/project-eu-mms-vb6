package mms.core.engine.sdf.entities.chart;

import org.jfree.chart.JFreeChart;
import org.w3c.dom.Element;

import com.itextpdf.awt.FontMapper;
import com.itextpdf.text.pdf.BaseFont;

public interface ChartManager {

	public JFreeChart getChart(Element chartElement);
	
	public FontMapper getFontMapper();
	
	public float getSizeW();
	
	public void setFont(BaseFont fieldFont, float fieldFontSize);

	public void setFontsPath(String fontsPath);
	
}
