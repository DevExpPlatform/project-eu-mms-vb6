package mms.core.engine.sdf.entities.chart;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Font;
import java.io.File;
import java.text.DecimalFormat;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.block.BlockBorder;
import org.jfree.chart.labels.StandardPieSectionLabelGenerator;
import org.jfree.chart.plot.RingPlot;
import org.jfree.chart.title.LegendTitle;
import org.jfree.chart.title.TextTitle;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.data.general.PieDataset;
import org.jfree.ui.HorizontalAlignment;
import org.jfree.ui.RectangleEdge;
import org.jfree.ui.RectangleInsets;
import org.jfree.util.Rotation;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.itextpdf.awt.DefaultFontMapper;
import com.itextpdf.awt.FontMapper;
import com.itextpdf.text.pdf.BaseFont;

import mms.core.engine.pdfmerger.commons.CommonUtils;

public class ChartRing implements ChartManager {

	private Font              fieldFont     = null;
	private float    		  fieldFontSize = 8.0f;
	private DefaultFontMapper fntMapper 	= new DefaultFontMapper();
	private   String 		  fontsPath	    = "";

	@Override
	public JFreeChart getChart(Element chartElement) {
		boolean 	 flgLegendTitle = (chartElement.getElementsByTagName("legendTitle").item(0) != null);
		String 		 lblTextTitle   = this.getNodeLabel(chartElement.getElementsByTagName("textTitle").item(0));
		NamedNodeMap nnmChart		= chartElement.getAttributes();
		JFreeChart   ringChart      = ChartFactory.createRingChart(lblTextTitle, this.getDataSet(chartElement.getElementsByTagName("dataset")), flgLegendTitle, false, false);

		this.setChart(ringChart, nnmChart);
		this.setChartTextTitle(ringChart.getTitle(), chartElement.getElementsByTagName("textTitle").item(0));
		this.setChartPlot((RingPlot) ringChart.getPlot(), chartElement.getElementsByTagName("ringPlot").item(0), chartElement.getElementsByTagName("dataset"));
		this.setChartLegend(ringChart.getLegend(), chartElement.getElementsByTagName("legendTitle").item(0));

		return ringChart;
	}

	private PieDataset getDataSet(NodeList nodeList) {
		DefaultPieDataset dataSet = new DefaultPieDataset();

		for (int i = 0; i < nodeList.getLength(); i++) {
			String[] dataSetValues = nodeList.item(i).getChildNodes().item(0).getNodeValue().split("\\|");

			dataSet.setValue(dataSetValues[0], Float.valueOf(dataSetValues[1]));
		}

		return dataSet;
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

		switch (legendPos) {
		case "bottom":
			rValue = RectangleEdge.BOTTOM;

			break;
		case "left":
			rValue = RectangleEdge.LEFT;

			break;
		case "top":
			rValue = RectangleEdge.TOP;

			break;
		case "right":
			rValue = RectangleEdge.RIGHT;

			break;
		}

		return rValue;
	}

	private Font getFont(String fontName) {
		Font rValue = null;

		try {
			rValue = Font.createFont(Font.TRUETYPE_FONT, new File(this.fontsPath + fontName));
		} catch (Exception e) {
			rValue = null;
		}

		if (rValue == null)
			rValue = this.fieldFont;

		return rValue;
	}

	@Override
	public FontMapper getFontMapper() {
		return fntMapper;
	}

	private String getNodeLabel(Node node) {
		if ((node != null) && (node.getAttributes().getNamedItem("label") != null))
			return node.getAttributes().getNamedItem("label").getNodeValue();

		return null;
	}

	@Override
	public float getSizeW() {
		return 0;
	}

	private void setChart(JFreeChart jfChart, NamedNodeMap nnmChart) {
		if (nnmChart.getNamedItem("backColor") != null) {
			int[] intArray = CommonUtils.getIntegerArray(nnmChart.getNamedItem("backColor").getNodeValue().split(","));
			jfChart.setBackgroundPaint(new Color(intArray[0], intArray[1], intArray[2]));
		} else {
			jfChart.setBackgroundPaint(null);
		}

		if (nnmChart.getNamedItem("chartBorderColor") != null) {
			jfChart.setBorderVisible(true);

			int[] intArray = CommonUtils.getIntegerArray(nnmChart.getNamedItem("chartBorderColor").getNodeValue().split(","));
			jfChart.setBorderPaint(new Color(intArray[0], intArray[1], intArray[2]));

			if (nnmChart.getNamedItem("chartBorderStroke") != null)
				jfChart.setBorderStroke(new BasicStroke(Float.valueOf(nnmChart.getNamedItem("chartBorderStroke").getNodeValue())));
		}

		if (nnmChart.getNamedItem("chartPadding") != null) {
			double[] doubleArray = CommonUtils.getDoubleArray(nnmChart.getNamedItem("chartPadding").getNodeValue().split(","));
			jfChart.setPadding(new RectangleInsets(doubleArray[0], doubleArray[1], doubleArray[2], doubleArray[3]));
		}
	}

	private void setChartLegend(LegendTitle legendTitle, Node nodeLegendTitle) {
		if (legendTitle != null) {
			Font 		 fntLegendTitle = this.fieldFont;
			float 		 fntSize	    = this.fieldFontSize;
			int 		 fntStyle		= Font.PLAIN;
			NamedNodeMap nnmLegendTitle = nodeLegendTitle.getAttributes();

			if (nnmLegendTitle.getNamedItem("backColor") != null) {
				int[] intArray = CommonUtils.getIntegerArray(nnmLegendTitle.getNamedItem("backColor").getNodeValue().split(","));
				legendTitle.setBackgroundPaint(new Color(intArray[0], intArray[1], intArray[2]));
			} else {
				legendTitle.setBackgroundPaint(null);
			}

			if (nnmLegendTitle.getNamedItem("fontName") != null)
				fntLegendTitle = this.getFont(nnmLegendTitle.getNamedItem("fontName").getNodeValue());

			if (nnmLegendTitle.getNamedItem("fontSize") != null)
				fntSize = Float.valueOf(nnmLegendTitle.getNamedItem("fontSize").getNodeValue());

			if (nnmLegendTitle.getNamedItem("fontColor") != null) {
				int[] intArray = CommonUtils.getIntegerArray(nnmLegendTitle.getNamedItem("fontColor").getNodeValue().split(","));
				legendTitle.setItemPaint(new Color(intArray[0], intArray[1], intArray[2]));
			}

			if (nnmLegendTitle.getNamedItem("fontStyle") != null)
				fntStyle = this.getDecodedFontStyle(nnmLegendTitle.getNamedItem("fontStyle").getNodeValue());

			if (nnmLegendTitle.getNamedItem("legendPos") != null)
				legendTitle.setPosition(this.getDecodedLegendPos(nnmLegendTitle.getNamedItem("legendPos").getNodeValue()));

			if (nnmLegendTitle.getNamedItem("borders") != null) {
				double[] doubleArray = CommonUtils.getDoubleArray(nnmLegendTitle.getNamedItem("borders").getNodeValue().split(","));
				legendTitle.setBorder(doubleArray[0], doubleArray[1], doubleArray[2], doubleArray[3]);
			} else {
				legendTitle.setFrame(BlockBorder.NONE);
			}

			if (nnmLegendTitle.getNamedItem("margins") != null) {
				double[] doubleArray = CommonUtils.getDoubleArray(nnmLegendTitle.getNamedItem("margins").getNodeValue().split(","));
				legendTitle.setMargin(new RectangleInsets(doubleArray[0], doubleArray[1], doubleArray[2], doubleArray[3]));
			}

			fntLegendTitle = fntLegendTitle.deriveFont(fntStyle, fntSize);
			legendTitle.setItemFont(fntLegendTitle);
		}
	}

	private void setChartPlot(RingPlot ringPlot, Node nodeRingPlot, NodeList nlDataSet) {
		if (nodeRingPlot != null) {
			Font         fntRingPlot  = this.fieldFont;
			float        fntSize      = this.fieldFontSize;
			int          fntStyle     = Font.PLAIN;
			NamedNodeMap nnmChartPlot = nodeRingPlot.getAttributes();
			Float        sectionDepth = 0.5f;

			if (nnmChartPlot.getNamedItem("backColor") != null) {
				int[] intArray = CommonUtils.getIntegerArray(nnmChartPlot.getNamedItem("backColor").getNodeValue().split(","));
				ringPlot.setBackgroundPaint(new Color(intArray[0], intArray[1], intArray[2]));
			} else {
				ringPlot.setBackgroundPaint(null);
			}

			if (nnmChartPlot.getNamedItem("outLine") != null) {
				ringPlot.setOutlineVisible(true);

				int[] intArray = CommonUtils.getIntegerArray(nnmChartPlot.getNamedItem("chartBorderColor").getNodeValue().split(","));
				ringPlot.setOutlinePaint(new Color(intArray[0], intArray[1], intArray[2]));

				if (nnmChartPlot.getNamedItem("outLineStroke") != null)
					ringPlot.setOutlineStroke(new BasicStroke(Float.valueOf(nnmChartPlot.getNamedItem("chartBorderStroke").getNodeValue())));
			} else {
				ringPlot.setOutlinePaint(null);
				ringPlot.setOutlineVisible(false);
			}

			if (nnmChartPlot.getNamedItem("sectionDepth") != null)
				sectionDepth = Float.valueOf(nnmChartPlot.getNamedItem("sectionDepth").getNodeValue());

			if ((nnmChartPlot.getNamedItem("showLabels") != null) && (nnmChartPlot.getNamedItem("showLabels").getNodeValue().equals("true"))) {
				if (nnmChartPlot.getNamedItem("fontName") != null)
					fntRingPlot = this.getFont(nnmChartPlot.getNamedItem("fontName").getNodeValue());

				if (nnmChartPlot.getNamedItem("fontSize") != null)
					fntSize = Float.valueOf(nnmChartPlot.getNamedItem("fontSize").getNodeValue());

				if (nnmChartPlot.getNamedItem("fontColor") != null) {
					int[] intArray = CommonUtils.getIntegerArray(nnmChartPlot.getNamedItem("fontColor").getNodeValue().split(","));
					ringPlot.setLabelPaint(new Color(intArray[0], intArray[1], intArray[2]));
				}

				if (nnmChartPlot.getNamedItem("fontStyle") != null)
					fntStyle = this.getDecodedFontStyle(nnmChartPlot.getNamedItem("fontStyle").getNodeValue());

				fntRingPlot = fntRingPlot.deriveFont(fntStyle, fntSize);

				ringPlot.setLabelFont(fntRingPlot);
				ringPlot.setSimpleLabels(true);
				ringPlot.setLabelGenerator(new StandardPieSectionLabelGenerator("{1}",new DecimalFormat("#,##0"), new DecimalFormat("0.000%")));
				ringPlot.setLabelBackgroundPaint(null);
				ringPlot.setLabelShadowPaint(null);
				ringPlot.setLabelOutlinePaint(null);
			} else {
				ringPlot.setLabelGenerator(null);
			}

			if (nnmChartPlot.getNamedItem("margins") != null) {
				double[] doubleArray = CommonUtils.getDoubleArray(nnmChartPlot.getNamedItem("margins").getNodeValue().split(","));
				ringPlot.setInsets(new RectangleInsets(doubleArray[0], doubleArray[1], doubleArray[2], doubleArray[3]));
			}

			if (nlDataSet != null) {
				for (int i = 0; i < nlDataSet.getLength(); i++) {
					Node nodeDataSet = nlDataSet.item(i);

					if (nodeDataSet.getAttributes().getNamedItem("color") != null) {
						int[] intArray = CommonUtils.getIntegerArray(nodeDataSet.getAttributes().getNamedItem("color").getNodeValue().split(","));

						ringPlot.setSectionPaint(nodeDataSet.getChildNodes().item(0).getNodeValue().split("\\|")[0], new Color(intArray[0], intArray[1], intArray[2]));
					}
				}
			}

			ringPlot.setDirection(Rotation.CLOCKWISE);
			ringPlot.setSectionDepth(sectionDepth);
			ringPlot.setSectionOutlinesVisible(false);
			ringPlot.setSeparatorsVisible(false);
			ringPlot.setShadowPaint(null);
		}
	}

	private void setChartTextTitle(TextTitle textTitle, Node nodeTextTitle) {
		if (textTitle != null) {
			Font fntTextTitle = this.fieldFont;

			if (nodeTextTitle == null) {
				textTitle.setFont(fntTextTitle);
			} else {
				float 		 fntSize	   = this.fieldFontSize;
				int 		 fntStyle	   = Font.PLAIN;
				NamedNodeMap nnmChartTitle = nodeTextTitle.getAttributes();

				if (nnmChartTitle.getNamedItem("alignment") != null) {
					switch (nnmChartTitle.getNamedItem("alignment").getNodeValue()) {
					case "left":
						textTitle.setHorizontalAlignment(HorizontalAlignment.LEFT);

						break;
					case "right":
						textTitle.setHorizontalAlignment(HorizontalAlignment.RIGHT);

						break;
					}
				}

				if (nnmChartTitle.getNamedItem("fontName") != null)
					fntTextTitle = this.getFont(nnmChartTitle.getNamedItem("fontName").getNodeValue());

				if (nnmChartTitle.getNamedItem("fontSize") != null)
					fntSize = Float.valueOf(nnmChartTitle.getNamedItem("fontSize").getNodeValue());

				if (nnmChartTitle.getNamedItem("fontColor") != null) {
					int[] intArray = CommonUtils.getIntegerArray(nnmChartTitle.getNamedItem("fontColor").getNodeValue().split(","));
					textTitle.setPaint(new Color(intArray[0], intArray[1], intArray[2]));
				}

				if (nnmChartTitle.getNamedItem("fontStyle") != null)
					fntStyle = this.getDecodedFontStyle(nnmChartTitle.getNamedItem("fontStyle").getNodeValue());

				if (nnmChartTitle.getNamedItem("borders") != null) {
					double[] doubleArray = CommonUtils.getDoubleArray(nnmChartTitle.getNamedItem("borders").getNodeValue().split(","));

					textTitle.setBorder(doubleArray[0], doubleArray[1], doubleArray[2], doubleArray[3]);
				}

				if (nnmChartTitle.getNamedItem("margins") != null) {
					double[] doubleArray = CommonUtils.getDoubleArray(nnmChartTitle.getNamedItem("margins").getNodeValue().split(","));

					textTitle.setMargin(new RectangleInsets(doubleArray[0], doubleArray[1], doubleArray[2], doubleArray[3]));
				}

				fntTextTitle = fntTextTitle.deriveFont(fntStyle, fntSize);
				textTitle.setFont(fntTextTitle);
			}
		}
	}

	@Override
	public void setFont(BaseFont fieldFont, float fieldFontSize) {
		this.fieldFont     = this.fntMapper.pdfToAwt(fieldFont, (int) fieldFontSize);
		this.fieldFontSize = fieldFontSize;
	}

	@Override
	public void setFontsPath(String fontsPath) {
		this.fontsPath = fontsPath;

		if (!fontsPath.equals(""))
			fntMapper.insertDirectory(fontsPath);
	}

}
