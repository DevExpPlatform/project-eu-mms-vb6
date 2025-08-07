package mms.core.engine.sdf.entities.chart;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Font;
import java.awt.FontFormatException;
import java.awt.GradientPaint;
import java.awt.Graphics2D;
import java.awt.Paint;
import java.awt.Stroke;
import java.awt.geom.RectangularShape;
import java.io.File;
import java.io.IOException;
import java.util.HashMap;

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
import org.jfree.chart.renderer.category.GroupedStackedBarRenderer;
import org.jfree.chart.renderer.category.StandardBarPainter;
import org.jfree.chart.title.LegendTitle;
import org.jfree.chart.title.TextTitle;
import org.jfree.data.KeyToGroupMap;
import org.jfree.data.category.CategoryDataset;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.text.G2TextMeasurer;
import org.jfree.text.TextBlock;
import org.jfree.text.TextUtilities;
import org.jfree.ui.GradientPaintTransformer;
import org.jfree.ui.HorizontalAlignment;
import org.jfree.ui.RectangleEdge;
import org.jfree.ui.RectangleInsets;
import org.jfree.ui.TextAnchor;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.itextpdf.awt.DefaultFontMapper;
import com.itextpdf.awt.FontMapper;
import com.itextpdf.text.pdf.BaseFont;

import mms.core.engine.pdfmerger.commons.CommonUtils;

@SuppressWarnings("serial")
public class ChartBar implements ChartManager {

	class StandardBarPainterExt extends StandardBarPainter {

		private CategoryDataset dataset = null;

		@Override
		public void paintBar(Graphics2D g2, BarRenderer renderer, int row, int column, RectangularShape bar, RectangleEdge base) {
			if ((dataset != null) && (this.dataset.getValue(row, column).floatValue() > 0.0f)) {
				Paint itemPaint = renderer.getItemPaint(row, column);
				GradientPaintTransformer t = renderer.getGradientPaintTransformer();

				if (t != null && itemPaint instanceof GradientPaint)
					itemPaint = t.transform((GradientPaint) itemPaint, bar);
				
				g2.setPaint(itemPaint);
				g2.fill(bar);

				if (renderer.isDrawBarOutline()) {
					// && state.getBarWidth() > BAR_OUTLINE_WIDTH_THRESHOLD) {
					Stroke stroke = renderer.getItemOutlineStroke(row, column);
					Paint paint = renderer.getItemOutlinePaint(row, column);
				
					if (stroke != null && paint != null) {
						g2.setStroke(stroke);
						g2.setPaint(paint);
						g2.draw(bar);
					}
				}
			}
		}

		public void setDataSet(CategoryDataset dataset) {
			this.dataset = dataset;
		}

    }
	private HashMap<String, String> customLabels  = null;
	private BaseFont 		  		fieldFont	  = null;
	private float    		  		fieldFontSize = 8.0f;
	private DefaultFontMapper 		fntMapper 	  = new DefaultFontMapper();
	private String 			  		fontsPath	  = "";
	private boolean 		  		groupedChart  = false;

	public ChartBar(boolean groupedChart) {
		this.groupedChart = groupedChart;
	}

	@Override
	public JFreeChart getChart(Element chartElement) {
		JFreeChart 			   barChart		   = null;
		DefaultCategoryDataset dataset 		   = new DefaultCategoryDataset();
		boolean 			   flgLegendTitle  = (chartElement.getElementsByTagName("legendTitle").item(0) != null);
		String 		 		   lblCategoryAxis = getNodeLabel(chartElement.getElementsByTagName("categoryAxis").item(0));
		String 		 		   lblValueAxis    = getNodeLabel(chartElement.getElementsByTagName("valueAxis").item(0));
		String 		 		   lblTextTitle    = getNodeLabel(chartElement.getElementsByTagName("textTitle").item(0));
		NamedNodeMap		   nnmChart		   = chartElement.getAttributes();
		NodeList 			   nlDataset       = ((Element) chartElement).getElementsByTagName("dataset");
		PlotOrientation 	   plotOrientation = PlotOrientation.VERTICAL;

		if (nnmChart.getNamedItem("chartOrientation") != null) {
			if (nnmChart.getNamedItem("chartOrientation").getNodeValue().equals("horizontal"))
				plotOrientation = PlotOrientation.HORIZONTAL;
		}

		if (nnmChart.getNamedItem("customLabels") != null) {
			if (nnmChart.getNamedItem("customLabels").getNodeValue().equals("true"))
				this.customLabels = new HashMap<String, String>();
		}

		for (int i = 0; i < nlDataset.getLength(); i++) {
			String[] chartDataSet = (String[]) nlDataset.item(i).getChildNodes().item(0).getNodeValue().split("\\|", -1);
			dataset.setValue(Float.valueOf(chartDataSet[0]), chartDataSet[1], chartDataSet[2]);
			
			if (this.customLabels != null)
				this.customLabels.put(chartDataSet[2] + "_" + chartDataSet[1], chartDataSet[3]);
		}

		if (groupedChart) {
			barChart = ChartFactory.createStackedBarChart(lblTextTitle, lblCategoryAxis, lblValueAxis, dataset, plotOrientation, flgLegendTitle, false, false);
		} else {
			barChart = ChartFactory.createBarChart(lblTextTitle, lblCategoryAxis, lblValueAxis, dataset, plotOrientation , flgLegendTitle, false, false);
		}

		if (nnmChart.getNamedItem("backColor") != null) {
			int[] intArray = CommonUtils.getIntegerArray(nnmChart.getNamedItem("backColor").getNodeValue().split(","));
			barChart.setBackgroundPaint(new Color(intArray[0], intArray[1], intArray[2]));
		}
		
		if (nnmChart.getNamedItem("chartBorderColor") != null) {
			barChart.setBorderVisible(true);

			int[] intArray = CommonUtils.getIntegerArray(nnmChart.getNamedItem("chartBorderColor").getNodeValue().split(","));
			barChart.setBorderPaint(new Color(intArray[0], intArray[1], intArray[2]));

			if (nnmChart.getNamedItem("chartBorderStroke") != null)
				barChart.setBorderStroke(new BasicStroke(Float.valueOf(nnmChart.getNamedItem("chartBorderStroke").getNodeValue())));
		}

		if (nnmChart.getNamedItem("chartPadding") != null) {
			double[] doubleArray = CommonUtils.getDoubleArray(nnmChart.getNamedItem("chartPadding").getNodeValue().split(","));
			barChart.setPadding(new RectangleInsets(doubleArray[0], doubleArray[1], doubleArray[2], doubleArray[3]));
		}

		setChartCategoryPlot(barChart.getCategoryPlot(), chartElement.getElementsByTagName("categoryPlot").item(0));
		
		if (groupedChart) {
	        CategoryPlot plot = (CategoryPlot) barChart.getPlot();
	        plot.setRenderer(setChartBarGroupedStackedBarRenderer(barChart.getCategoryPlot(), chartElement.getElementsByTagName("barGroupRenderer").item(0), dataset));
		} else {
			setChartBarRenderer(barChart.getCategoryPlot(), chartElement.getElementsByTagName("barRenderer").item(0));
		}
			
		setChartTextTitle(barChart.getTitle(), chartElement.getElementsByTagName("textTitle").item(0));
		setChartCategoryAxis(barChart.getCategoryPlot(), chartElement.getElementsByTagName("categoryAxis").item(0));
		setChartValueAxis(barChart.getCategoryPlot(), chartElement.getElementsByTagName("valueAxis").item(0));
		setChartLegend(barChart.getLegend(), chartElement.getElementsByTagName("legendTitle").item(0));

		return barChart;
	}

	private int getDecodedFontStyle(String fontStyle) {
		int    rValue     = Font.PLAIN;
		String strArray[] = fontStyle.split("\\|");

		for (int i = 0; i < strArray.length; i++) {
			int tmpFontStyle = 0;

			if (strArray[i].equals("bold")) {
				tmpFontStyle = Font.BOLD;
			} else if (strArray[i].equals("italic")) {
				tmpFontStyle = Font.ITALIC;
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

		//System.out.println(legendPos);

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

	private Font getFont(String fontName) {
		Font rValue = null;

		try {
			rValue = Font.createFont(Font.TRUETYPE_FONT, new File(fontsPath + fontName));
		} catch (FontFormatException e) {
			rValue = null;
		} catch (IOException e) {
			rValue = null;
		}

		if (rValue == null) 
			rValue = fntMapper.pdfToAwt(fieldFont, (int) fieldFontSize);

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

	private GroupedStackedBarRenderer setChartBarGroupedStackedBarRenderer(CategoryPlot categoryPlot, Node nodeBarRenderer, CategoryDataset dataset) {
	    final GroupedStackedBarRenderer chartBarRenderer = new GroupedStackedBarRenderer();
	
	    if (nodeBarRenderer != null) {
	        int 		 keyGroupMapLen    = 0;
	        int 		 keyGroupMapCount  = 0;
	        NamedNodeMap nnmBarRenderer	   = nodeBarRenderer.getAttributes();
	        
			if (nnmBarRenderer.getNamedItem("groupMap") != null) {
		        String[]      keyGroupMap = nnmBarRenderer.getNamedItem("groupMap").getNodeValue().split("\\|");
				KeyToGroupMap map         = new KeyToGroupMap(keyGroupMap[0]);
	
				keyGroupMapLen    = (keyGroupMap.length - 1);
				keyGroupMapCount = map.getGroupCount();
				
				for (int i = 1; i < keyGroupMap.length; i++) {
					String[] keyGroup = keyGroupMap[i].split(",");
	
					map.mapKeyToGroup(keyGroup[0], keyGroup[1]);
				}
				
		        chartBarRenderer.setSeriesToGroupMap(map); 
			}
	        
			if (nnmBarRenderer.getNamedItem("customBarRenderer") != null) {
				if (nnmBarRenderer.getNamedItem("customBarRenderer").getNodeValue().equals("true")) {
					StandardBarPainterExt render = new StandardBarPainterExt();
					render.setDataSet(dataset);
					
					chartBarRenderer.setBarPainter(render);
//					chartBarRenderer.setBarPainter(new StandardBarPainter());
	
					if (nnmBarRenderer.getNamedItem("itemMargin") != null)
				        chartBarRenderer.setItemMargin(Float.valueOf(nnmBarRenderer.getNamedItem("itemMargin").getNodeValue()));
	
					if (nnmBarRenderer.getNamedItem("barsColor") != null) {
						String[] seriesColorPaint = nnmBarRenderer.getNamedItem("barsColor").getNodeValue().split("\\|");
						
						for (int i = 0; i < seriesColorPaint.length; i++) {
							int[] intArray = CommonUtils.getIntegerArray(seriesColorPaint[i].split(","));
							chartBarRenderer.setSeriesPaint(i, new Color(intArray[0], intArray[1], intArray[2]));
						}
					}
	
					if (nnmBarRenderer.getNamedItem("outLineStroke") != null) {
						chartBarRenderer.setDrawBarOutline(true);
	
						String[] outLineStroke = nnmBarRenderer.getNamedItem("outLineStroke").getNodeValue().split("\\|");
						
						for (int i = 0; i < outLineStroke.length; i++)
							chartBarRenderer.setSeriesOutlineStroke(i, new BasicStroke(Float.valueOf(outLineStroke[i])));
	
						if (nnmBarRenderer.getNamedItem("outLineColor") != null) {
							String[] outLineColor = nnmBarRenderer.getNamedItem("outLineColor").getNodeValue().split("\\|");

							for (int i = 0; i < outLineColor.length; i++) {
								int[] intArray = CommonUtils.getIntegerArray(outLineColor[i].split(","));
								chartBarRenderer.setSeriesOutlinePaint(i, new Color(intArray[0], intArray[1], intArray[2]));
							}
						}
					}
				}
			}
	
			if (nnmBarRenderer.getNamedItem("showBaseItemLabel") != null) {
				if (nnmBarRenderer.getNamedItem("showBaseItemLabel").getNodeValue().equals("true")) {
					float fntSize      = fieldFontSize;
					int   fntStyle	   = Font.PLAIN;
					Font  fntItemLabel = fntMapper.pdfToAwt(fieldFont, (int) fieldFontSize); 
	
					if (nnmBarRenderer.getNamedItem("bilFontName") != null) 
						fntItemLabel = getFont(nnmBarRenderer.getNamedItem("bilFontName").getNodeValue());
	
					if (nnmBarRenderer.getNamedItem("bilFontSize") != null)
						fntSize  = Integer.valueOf(nnmBarRenderer.getNamedItem("bilFontSize").getNodeValue());
	
					if (nnmBarRenderer.getNamedItem("bilFontColor") != null) {
						int[] intArray = CommonUtils.getIntegerArray(nnmBarRenderer.getNamedItem("bilFontColor").getNodeValue().split(","));
						chartBarRenderer.setBaseItemLabelPaint(new Color(intArray[0], intArray[1], intArray[2]));
					}
	
					if (nnmBarRenderer.getNamedItem("bilFontStyle") != null) 
						fntStyle = getDecodedFontStyle(nnmBarRenderer.getNamedItem("bilFontStyle").getNodeValue());
	
					fntItemLabel = fntItemLabel.deriveFont(fntStyle, fntSize);

					chartBarRenderer.setBaseItemLabelFont(fntItemLabel);
					chartBarRenderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.CENTER, TextAnchor.CENTER));
					chartBarRenderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator() {
						@Override
						public String generateLabel(CategoryDataset dataset, int row, int column) {
							return ((customLabels == null) ? generateLabelString(dataset, row, column) : customLabels.get(dataset.getColumnKey(column) + "_" + dataset.getRowKey(row)));
						}
					});

					if (keyGroupMapCount == 1)
						chartBarRenderer.setSeriesPositiveItemLabelPosition((keyGroupMapLen -1), new ItemLabelPosition(ItemLabelAnchor.OUTSIDE12, TextAnchor.BASELINE_CENTER));

					if (nnmBarRenderer.getNamedItem("bilAnchorOffset") != null) 
						chartBarRenderer.setItemLabelAnchorOffset(Float.valueOf(nnmBarRenderer.getNamedItem("bilAnchorOffset").getNodeValue()));
	
					ItemLabelPosition ilp = new ItemLabelPosition(ItemLabelAnchor.OUTSIDE6, TextAnchor.TOP_CENTER);
					chartBarRenderer.setBaseNegativeItemLabelPosition(ilp);

					chartBarRenderer.setBaseItemLabelsVisible(true);
				}
			}
		}

	    return chartBarRenderer;
	}

	private void setChartBarRenderer(CategoryPlot categoryPlot, Node nodeBarRenderer) {
		if (nodeBarRenderer != null) {
			BarRenderer  chartBarRenderer = (BarRenderer) categoryPlot.getRenderer();
			NamedNodeMap nnmBarRenderer	  = nodeBarRenderer.getAttributes();

			if (nnmBarRenderer.getNamedItem("customBarRenderer") != null) {
				if (nnmBarRenderer.getNamedItem("customBarRenderer").getNodeValue().equals("true")) {
					chartBarRenderer.setBarPainter(new StandardBarPainter());

					if (nnmBarRenderer.getNamedItem("barsColor") != null) {
						int[] intArray = CommonUtils.getIntegerArray(nnmBarRenderer.getNamedItem("barsColor").getNodeValue().split(","));
						chartBarRenderer.setSeriesPaint(0, new Color(intArray[0], intArray[1], intArray[2]));
					}

					if (nnmBarRenderer.getNamedItem("outLineStroke") != null) {
						chartBarRenderer.setDrawBarOutline(true);
						chartBarRenderer.setSeriesOutlineStroke(0, new BasicStroke(Float.valueOf(nnmBarRenderer.getNamedItem("outLineStroke").getNodeValue())));;

						if (nnmBarRenderer.getNamedItem("outLineColor") != null) {
							int[] intArray = CommonUtils.getIntegerArray(nnmBarRenderer.getNamedItem("outLineColor").getNodeValue().split(","));
							chartBarRenderer.setSeriesOutlinePaint(0, new Color(intArray[0], intArray[1], intArray[2]));
						}
					}
				}
			}

			if (nnmBarRenderer.getNamedItem("showBaseItemLabel") != null) {
				if (nnmBarRenderer.getNamedItem("showBaseItemLabel").getNodeValue().equals("true")) {
					float fntSize      = fieldFontSize;
					int   fntStyle	   = Font.PLAIN;
					Font  fntItemLabel = fntMapper.pdfToAwt(fieldFont, (int) fieldFontSize); 

					if (nnmBarRenderer.getNamedItem("bilFontName") != null) 
						fntItemLabel = getFont(nnmBarRenderer.getNamedItem("bilFontName").getNodeValue());

					if (nnmBarRenderer.getNamedItem("bilFontSize") != null)
						fntSize  = Integer.valueOf(nnmBarRenderer.getNamedItem("bilFontSize").getNodeValue());

					if (nnmBarRenderer.getNamedItem("bilFontColor") != null) {
						int[] intArray = CommonUtils.getIntegerArray(nnmBarRenderer.getNamedItem("bilFontColor").getNodeValue().split(","));
						chartBarRenderer.setBaseItemLabelPaint(new Color(intArray[0], intArray[1], intArray[2]));
					}

					if (nnmBarRenderer.getNamedItem("bilFontStyle") != null) 
						fntStyle = getDecodedFontStyle(nnmBarRenderer.getNamedItem("bilFontStyle").getNodeValue());

					fntItemLabel = fntItemLabel.deriveFont(fntStyle, fntSize);

					chartBarRenderer.setBaseItemLabelFont(fntItemLabel);
					chartBarRenderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.CENTER, TextAnchor.CENTER));
					chartBarRenderer.setBaseItemLabelGenerator(new StandardCategoryItemLabelGenerator() {
						@Override
						public String generateLabel(CategoryDataset dataset, int row, int column) {
							return ((customLabels == null) ? generateLabelString(dataset, row, column) : customLabels.get(dataset.getColumnKey(column) + "_" + dataset.getRowKey(row)));
						}
					});

					if (nnmBarRenderer.getNamedItem("bilAnchorOffset") != null) 
						chartBarRenderer.setItemLabelAnchorOffset(Float.valueOf(nnmBarRenderer.getNamedItem("bilAnchorOffset").getNodeValue()));
	
					ItemLabelPosition ilp = new ItemLabelPosition(ItemLabelAnchor.OUTSIDE6, TextAnchor.TOP_CENTER);
					chartBarRenderer.setBaseNegativeItemLabelPosition(ilp);

					chartBarRenderer.setBaseItemLabelsVisible(true);
				}
			}
		}
	}

	private void setChartCategoryAxis(CategoryPlot categoryPlot, Node nodeCategoryAxis) {
		Font fntCategoryAxis  = fntMapper.pdfToAwt(fieldFont, (int) fieldFontSize);;
		Font fntTickLabelFont = fntMapper.pdfToAwt(fieldFont, (int) fieldFontSize);;

		if (nodeCategoryAxis == null) {
			CategoryAxis categoryAxis     = categoryPlot.getDomainAxis();
			categoryAxis.setMaximumCategoryLabelLines(4);
			categoryAxis.setLabelFont(fntCategoryAxis);
			categoryAxis.setTickLabelFont(fntTickLabelFont);
		} else {
			NamedNodeMap nnmCategoryAxis  = nodeCategoryAxis.getAttributes();

			if (nnmCategoryAxis.getNamedItem("catlAlignment") != null) {
				if (nnmCategoryAxis.getNamedItem("catlAlignment").getNodeValue().equals("left")) {
					categoryPlot.setDomainAxis(new CategoryAxis() {
						@SuppressWarnings("rawtypes")
						@Override
						protected TextBlock createLabel(Comparable category, float width, RectangleEdge edge, Graphics2D g2) {
							TextBlock label = TextUtilities.createTextBlock(category.toString(), getTickLabelFont(category), getTickLabelPaint(category), width, this.getMaximumCategoryLabelLines(), new G2TextMeasurer(g2));
							label.setLineAlignment(HorizontalAlignment.LEFT);
							
							return label;
						}
					});
				} else if (nnmCategoryAxis.getNamedItem("catlAlignment").getNodeValue().equals("right")) {
					categoryPlot.setDomainAxis(new CategoryAxis() {
						@SuppressWarnings("rawtypes")
						@Override
						protected TextBlock createLabel(Comparable category, float width, RectangleEdge edge, Graphics2D g2) {
							TextBlock label = TextUtilities.createTextBlock(category.toString(), getTickLabelFont(category), getTickLabelPaint(category), width, this.getMaximumCategoryLabelLines(), new G2TextMeasurer(g2));
							label.setLineAlignment(HorizontalAlignment.RIGHT);
							
							return label;
						}
					});
				}
			}

			CategoryAxis categoryAxis     = categoryPlot.getDomainAxis();
			categoryAxis.setMaximumCategoryLabelLines(4);

			/*
			 * axisLineVisible
			 */
			float fntSize  = fieldFontSize;
			int   fntStyle = Font.PLAIN;

			if (nnmCategoryAxis.getNamedItem("axisLineVisible") != null) {
				if (nnmCategoryAxis.getNamedItem("axisLineVisible").getNodeValue().equals("false")) 
					categoryAxis.setAxisLineVisible(false);
			}

			if (categoryAxis.isAxisLineVisible()) {
				if (nnmCategoryAxis.getNamedItem("axisLineStroke") != null) 
					categoryAxis.setAxisLineStroke(new BasicStroke(Float.valueOf(nnmCategoryAxis.getNamedItem("axisLineStroke").getNodeValue())));

				if (nnmCategoryAxis.getNamedItem("axisLineColor") != null) {
					int[] intArray = CommonUtils.getIntegerArray(nnmCategoryAxis.getNamedItem("axisLineColor").getNodeValue().split(","));
					categoryAxis.setAxisLinePaint(new Color(intArray[0], intArray[1], intArray[2]));
				}
			}

			/*
			 * tickMarksVisible
			 */
			if (nnmCategoryAxis.getNamedItem("tickMarksVisible") != null) {
				if (nnmCategoryAxis.getNamedItem("tickMarksVisible").getNodeValue().equals("false")) 
					categoryAxis.setTickMarksVisible(false);
			}

			if (categoryAxis.isTickMarksVisible()) {
				if (nnmCategoryAxis.getNamedItem("tickMarkStroke") != null) 
					categoryAxis.setTickMarkStroke(new BasicStroke(Float.valueOf(nnmCategoryAxis.getNamedItem("tickMarkStroke").getNodeValue())));

				if (nnmCategoryAxis.getNamedItem("tickMarkColor") != null) {
					int[] intArray = CommonUtils.getIntegerArray(nnmCategoryAxis.getNamedItem("tickMarkColor").getNodeValue().split(","));
					categoryAxis.setTickMarkPaint(new Color(intArray[0], intArray[1], intArray[2]));
				}
			}

			/*
			 * Category Axis Tick Label
			 */
			if (nnmCategoryAxis.getNamedItem("catlVisible") != null) {
				if (nnmCategoryAxis.getNamedItem("catlVisible").getNodeValue().equals("false")) 
					categoryAxis.setTickLabelsVisible(false);
			}

			if (categoryAxis.isTickLabelsVisible()) {
				if (nnmCategoryAxis.getNamedItem("catlFontName") != null) 
					fntTickLabelFont = getFont(nnmCategoryAxis.getNamedItem("catlFontName").getNodeValue());

				if (nnmCategoryAxis.getNamedItem("catlFontSize") != null)
					fntSize = Float.valueOf(nnmCategoryAxis.getNamedItem("catlFontSize").getNodeValue());

				if (nnmCategoryAxis.getNamedItem("catlFontColor") != null) {
					int[] intArray = CommonUtils.getIntegerArray(nnmCategoryAxis.getNamedItem("catlFontColor").getNodeValue().split(","));
					categoryAxis.setTickLabelPaint(new Color(intArray[0], intArray[1], intArray[2]));
				}

				if (nnmCategoryAxis.getNamedItem("catlFontStyle") != null) 
					fntStyle = getDecodedFontStyle(nnmCategoryAxis.getNamedItem("catlFontStyle").getNodeValue());

				if (nnmCategoryAxis.getNamedItem("catlMargins") != null) {
					double[] doubleArray = CommonUtils.getDoubleArray(nnmCategoryAxis.getNamedItem("catlMargins").getNodeValue().split(","));
					categoryAxis.setTickLabelInsets(new RectangleInsets(doubleArray[0], doubleArray[1], doubleArray[2], doubleArray[3]));
				}

				fntTickLabelFont = fntTickLabelFont.deriveFont(fntStyle, fntSize);
				categoryAxis.setTickLabelFont(fntTickLabelFont);

				if (nnmCategoryAxis.getNamedItem("catlRotation") != null) {
					double value = (Double.valueOf(nnmCategoryAxis.getNamedItem("catlRotation").getNodeValue()) * (Math.PI / 180));
					categoryAxis.setCategoryLabelPositions(CategoryLabelPositions.createUpRotationLabelPositions(value)); 
				}
			}

			/*
			 * Category Axis Label
			 */
			if (categoryAxis.getLabel() != null) {
				fntSize	 = fieldFontSize;
				fntStyle = Font.PLAIN;

				if (nnmCategoryAxis.getNamedItem("calFontName") != null) 
					fntCategoryAxis = getFont(nnmCategoryAxis.getNamedItem("calFontName").getNodeValue());

				if (nnmCategoryAxis.getNamedItem("calFontSize") != null)
					fntSize = Float.valueOf(nnmCategoryAxis.getNamedItem("calFontSize").getNodeValue());

				if (nnmCategoryAxis.getNamedItem("calFontColor") != null) {
					int[] intArray = CommonUtils.getIntegerArray(nnmCategoryAxis.getNamedItem("calFontColor").getNodeValue().split(","));
					categoryAxis.setLabelPaint(new Color(intArray[0], intArray[1], intArray[2]));
				}

				if (nnmCategoryAxis.getNamedItem("calFontStyle") != null) 
					fntStyle = getDecodedFontStyle(nnmCategoryAxis.getNamedItem("calFontStyle").getNodeValue());

				if (nnmCategoryAxis.getNamedItem("calMargins") != null) {
					double[] doubleArray = CommonUtils.getDoubleArray(nnmCategoryAxis.getNamedItem("calMargins").getNodeValue().split(","));
					categoryAxis.setLabelInsets(new RectangleInsets(doubleArray[0], doubleArray[1], doubleArray[2], doubleArray[3]));
				}

				fntCategoryAxis = fntCategoryAxis.deriveFont(fntStyle, fntSize);
				categoryAxis.setLabelFont(fntCategoryAxis);
			}
		}
	}

	private void setChartCategoryPlot(CategoryPlot categoryPlot, Node item) {
		if (item != null) {
			NamedNodeMap nnmChartPlot = item.getAttributes();

			// backColor
			if (nnmChartPlot.getNamedItem("backColor") != null) {
				int[] intArray = CommonUtils.getIntegerArray(nnmChartPlot.getNamedItem("backColor").getNodeValue().split(","));
				categoryPlot.setBackgroundPaint(new Color(intArray[0], intArray[1], intArray[2]));
			}

			// outlineVisible
			if (nnmChartPlot.getNamedItem("outlineVisible") != null) {
				if (nnmChartPlot.getNamedItem("outlineVisible").getNodeValue().equals("false")) 
					categoryPlot.setOutlineVisible(false);
			}

			if (categoryPlot.isOutlineVisible()) {
				if (nnmChartPlot.getNamedItem("outlineColor") != null) {
					int[] intArray = CommonUtils.getIntegerArray(nnmChartPlot.getNamedItem("outlineColor").getNodeValue().split(","));
					categoryPlot.setOutlinePaint(new Color(intArray[0], intArray[1], intArray[2]));
				}

				if (nnmChartPlot.getNamedItem("outlineStroke") != null)
					categoryPlot.setOutlineStroke(new BasicStroke(Float.valueOf(nnmChartPlot.getNamedItem("outlineStroke").getNodeValue())));
			}

			// domainGridlinesVisible
			if (nnmChartPlot.getNamedItem("domainGridlinesVisible") != null) {
				if (nnmChartPlot.getNamedItem("domainGridlinesVisible").getNodeValue().equals("true")) 
					categoryPlot.setDomainGridlinesVisible(true);
			}

			if (categoryPlot.isDomainGridlinesVisible()) {
				if (nnmChartPlot.getNamedItem("domainGridLineColor") != null) {
					int[] intArray = CommonUtils.getIntegerArray(nnmChartPlot.getNamedItem("domainGridLineColor").getNodeValue().split(","));
					categoryPlot.setDomainGridlinePaint(new Color(intArray[0], intArray[1], intArray[2]));
				}
			}

			// rangeGridlinesVisible
			if (nnmChartPlot.getNamedItem("rangeGridlinesVisible") != null) {
				if (nnmChartPlot.getNamedItem("rangeGridlinesVisible").getNodeValue().equals("false")) 
					categoryPlot.setRangeGridlinesVisible(false);
			}

			if (categoryPlot.isRangeGridlinesVisible()) {
				if (nnmChartPlot.getNamedItem("rangeGridLineColor") != null) {
					int[] intArray = CommonUtils.getIntegerArray(nnmChartPlot.getNamedItem("rangeGridLineColor").getNodeValue().split(","));
					categoryPlot.setRangeGridlinePaint(new Color(intArray[0], intArray[1], intArray[2]));
				}
			}

			// axisOffset
			if (nnmChartPlot.getNamedItem("axisOffset") != null) {
				double[] doubleArray = CommonUtils.getDoubleArray(nnmChartPlot.getNamedItem("axisOffset").getNodeValue().split(","));
				categoryPlot.setAxisOffset(new RectangleInsets(doubleArray[0], doubleArray[1], doubleArray[2], doubleArray[3]));
			}

			// margins
			if (nnmChartPlot.getNamedItem("margins") != null) {
				double[] doubleArray = CommonUtils.getDoubleArray(nnmChartPlot.getNamedItem("margins").getNodeValue().split(","));
				categoryPlot.setInsets(new RectangleInsets(doubleArray[0], doubleArray[1], doubleArray[2], doubleArray[3]));
			}
		}
	}

	private void setChartLegend(LegendTitle legendTitle, Node nodeLegendTitle) {
		if (legendTitle != null) {
			float 		 fntSize	    = fieldFontSize;
			int 		 fntStyle		= Font.PLAIN;
			Font 		 fntLegendTitle = fntMapper.pdfToAwt(fieldFont, (int) fieldFontSize);
			NamedNodeMap nnmLegendTitle = nodeLegendTitle.getAttributes();

			if (nnmLegendTitle.getNamedItem("backColor") != null) {
				int[] intArray = CommonUtils.getIntegerArray(nnmLegendTitle.getNamedItem("backColor").getNodeValue().split(","));
				legendTitle.setBackgroundPaint(new Color(intArray[0], intArray[1], intArray[2]));
			}
			
			if (nnmLegendTitle.getNamedItem("fontName") != null) 
				fntLegendTitle = getFont(nnmLegendTitle.getNamedItem("fontName").getNodeValue());

			if (nnmLegendTitle.getNamedItem("fontSize") != null)
				fntSize = Float.valueOf(nnmLegendTitle.getNamedItem("fontSize").getNodeValue());

			if (nnmLegendTitle.getNamedItem("fontColor") != null) {
				int[] intArray = CommonUtils.getIntegerArray(nnmLegendTitle.getNamedItem("fontColor").getNodeValue().split(","));
				legendTitle.setItemPaint(new Color(intArray[0], intArray[1], intArray[2]));
			}

			if (nnmLegendTitle.getNamedItem("fontStyle") != null) 
				fntStyle = getDecodedFontStyle(nnmLegendTitle.getNamedItem("fontStyle").getNodeValue());

			if (nnmLegendTitle.getNamedItem("legendPos") != null) 
				legendTitle.setPosition(getDecodedLegendPos(nnmLegendTitle.getNamedItem("legendPos").getNodeValue()));

			if (nnmLegendTitle.getNamedItem("borders") != null) {
				double[] doubleArray = CommonUtils.getDoubleArray(nnmLegendTitle.getNamedItem("borders").getNodeValue().split(","));
				legendTitle.setBorder(doubleArray[0], doubleArray[1], doubleArray[2], doubleArray[3]);
			} else {
				legendTitle.setFrame(BlockBorder.NONE);;
			}

			if (nnmLegendTitle.getNamedItem("margins") != null) {
				double[] doubleArray = CommonUtils.getDoubleArray(nnmLegendTitle.getNamedItem("margins").getNodeValue().split(","));
				legendTitle.setMargin(new RectangleInsets(doubleArray[0], doubleArray[1], doubleArray[2], doubleArray[3]));;
			}

			fntLegendTitle = fntLegendTitle.deriveFont(fntStyle, fntSize);
			legendTitle.setItemFont(fntLegendTitle);
		}
	}

	private void setChartTextTitle(TextTitle textTitle, Node nodeTextTitle) {
		if (textTitle != null) {
			Font fntTextTitle = fntMapper.pdfToAwt(fieldFont, (int) fieldFontSize);

			if (nodeTextTitle == null) {
				textTitle.setFont(fntTextTitle);
			} else {
				float 		 fntSize	    	 = fieldFontSize;
				int 		 fntStyle			 = Font.PLAIN;
				NamedNodeMap nnmChartTitle 	 	 = nodeTextTitle.getAttributes();

				if (nnmChartTitle.getNamedItem("alignment") != null) {
					switch (nnmChartTitle.getNamedItem("alignment").getNodeValue().hashCode()) {
					case 3317767:
						textTitle.setHorizontalAlignment(HorizontalAlignment.LEFT);
						
						break;
					case 108511772:
						textTitle.setHorizontalAlignment(HorizontalAlignment.RIGHT);
						
						break;
					}
				}

				if (nnmChartTitle.getNamedItem("fontName") != null)
					fntTextTitle = getFont(nnmChartTitle.getNamedItem("fontName").getNodeValue());

				if (nnmChartTitle.getNamedItem("fontSize") != null)
					fntSize = Float.valueOf(nnmChartTitle.getNamedItem("fontSize").getNodeValue());

				if (nnmChartTitle.getNamedItem("fontColor") != null) {
					int[] intArray = CommonUtils.getIntegerArray(nnmChartTitle.getNamedItem("fontColor").getNodeValue().split(","));
					textTitle.setPaint(new Color(intArray[0], intArray[1], intArray[2]));
				}

				if (nnmChartTitle.getNamedItem("fontStyle") != null) 
					fntStyle = getDecodedFontStyle(nnmChartTitle.getNamedItem("fontStyle").getNodeValue());

				if (nnmChartTitle.getNamedItem("borders") != null) {
					double[] doubleArray = CommonUtils.getDoubleArray(nnmChartTitle.getNamedItem("borders").getNodeValue().split(","));
					textTitle.setBorder(doubleArray[0], doubleArray[1], doubleArray[2], doubleArray[3]);
				}

				if (nnmChartTitle.getNamedItem("margins") != null) {
					double[] doubleArray = CommonUtils.getDoubleArray(nnmChartTitle.getNamedItem("margins").getNodeValue().split(","));
					textTitle.setMargin(new RectangleInsets(doubleArray[0], doubleArray[1], doubleArray[2], doubleArray[3]));;
				}

				fntTextTitle = fntTextTitle.deriveFont(fntStyle, fntSize);
				textTitle.setFont(fntTextTitle);
			}
		}
	}

	private void setChartValueAxis(CategoryPlot categoryPlot, Node nodeValueAxis) {
		Font  		 fntTickLabelFont = fntMapper.pdfToAwt(fieldFont, (int) fieldFontSize);
		Font  		 fntValueAxis  	  = fntMapper.pdfToAwt(fieldFont, (int) fieldFontSize);
		ValueAxis 	 valueAxis 	  	  = categoryPlot.getRangeAxis();

		if (nodeValueAxis == null) {
			valueAxis.setLabelFont(fntValueAxis);
			valueAxis.setTickLabelFont(fntTickLabelFont);
		} else {
			NamedNodeMap nnmValueAxis     = nodeValueAxis.getAttributes();

			if (nnmValueAxis.getNamedItem("axisRange") != null) {
				double[] doubleArray = CommonUtils.getDoubleArray(nnmValueAxis.getNamedItem("axisRange").getNodeValue().split(","));
				valueAxis.setRange(doubleArray[0], doubleArray[1]);
			}

			if (nnmValueAxis.getNamedItem("axisRangeLowerRange") != null)
				valueAxis.setLowerMargin(Float.valueOf(nnmValueAxis.getNamedItem("axisRangeLowerRange").getNodeValue()));

			if (nnmValueAxis.getNamedItem("axisRangeUpperMargin") != null)
				valueAxis.setUpperMargin(Float.valueOf(nnmValueAxis.getNamedItem("axisRangeUpperMargin").getNodeValue()));

			/*
			 * axisLineVisible
			 */
			float fntSize  = fieldFontSize;
			int   fntStyle = Font.PLAIN;

			if (nnmValueAxis.getNamedItem("axisLineVisible") != null) {
				if (nnmValueAxis.getNamedItem("axisLineVisible").getNodeValue().equals("false")) 
					valueAxis.setAxisLineVisible(false);
			}

			if (valueAxis.isAxisLineVisible()) {
				if (nnmValueAxis.getNamedItem("axisLineStroke") != null) 
					valueAxis.setAxisLineStroke(new BasicStroke(Float.valueOf(nnmValueAxis.getNamedItem("axisLineStroke").getNodeValue())));

				if (nnmValueAxis.getNamedItem("axisLineColor") != null) {
					int[] intArray = CommonUtils.getIntegerArray(nnmValueAxis.getNamedItem("axisLineColor").getNodeValue().split(","));
					valueAxis.setAxisLinePaint(new Color(intArray[0], intArray[1], intArray[2]));
				}
			}

			/*
			 * tickMarksVisible
			 */
			if (nnmValueAxis.getNamedItem("tickMarksVisible") != null) {
				if (nnmValueAxis.getNamedItem("tickMarksVisible").getNodeValue().equals("false")) 
					valueAxis.setTickMarksVisible(false);
			}

			if (valueAxis.isTickMarksVisible()) {
				if (nnmValueAxis.getNamedItem("tickMarkStroke") != null) 
					valueAxis.setTickMarkStroke(new BasicStroke(Float.valueOf(nnmValueAxis.getNamedItem("tickMarkStroke").getNodeValue())));

				if (nnmValueAxis.getNamedItem("tickMarkColor") != null) {
					int[] intArray = CommonUtils.getIntegerArray(nnmValueAxis.getNamedItem("tickMarkColor").getNodeValue().split(","));
					valueAxis.setTickMarkPaint(new Color(intArray[0], intArray[1], intArray[2]));
				}
			}

			/*
			 * Value Axis Tick Label
			 */
			if (nnmValueAxis.getNamedItem("catlVisible") != null) {
				if (nnmValueAxis.getNamedItem("catlVisible").getNodeValue().equals("false")) 
					valueAxis.setTickLabelsVisible(false);
			}

			if (valueAxis.isTickLabelsVisible()) {
				if (nnmValueAxis.getNamedItem("vatlFontName") != null) 
					fntTickLabelFont = getFont(nnmValueAxis.getNamedItem("vatlFontName").getNodeValue());

				if (nnmValueAxis.getNamedItem("vatlFontSize") != null)
					fntSize = Float.valueOf(nnmValueAxis.getNamedItem("vatlFontSize").getNodeValue());

				if (nnmValueAxis.getNamedItem("vatlFontColor") != null) {
					int[] intArray = CommonUtils.getIntegerArray(nnmValueAxis.getNamedItem("vatlFontColor").getNodeValue().split(","));
					valueAxis.setTickLabelPaint(new Color(intArray[0], intArray[1], intArray[2]));
				}

				if (nnmValueAxis.getNamedItem("vatlFontStyle") != null) 
					fntStyle = getDecodedFontStyle(nnmValueAxis.getNamedItem("vatlFontStyle").getNodeValue());

				if (nnmValueAxis.getNamedItem("vatlMargins") != null) {
					double[] doubleArray = CommonUtils.getDoubleArray(nnmValueAxis.getNamedItem("vatlMargins").getNodeValue().split(","));
					valueAxis.setTickLabelInsets(new RectangleInsets(doubleArray[0], doubleArray[1], doubleArray[2], doubleArray[3]));
				}

				fntTickLabelFont = fntTickLabelFont.deriveFont(fntStyle, fntSize);
				valueAxis.setTickLabelFont(fntTickLabelFont);
			}

			/*
			 * Value Axis Label
			 */
			if (valueAxis.getLabel() != null) {
				fntSize	 = fieldFontSize;
				fntStyle = Font.PLAIN;

				if (nnmValueAxis.getNamedItem("valFontName") != null) 
					fntValueAxis = getFont(nnmValueAxis.getNamedItem("valFontName").getNodeValue());

				if (nnmValueAxis.getNamedItem("valFontSize") != null)
					fntSize = Float.valueOf(nnmValueAxis.getNamedItem("valFontSize").getNodeValue());

				if (nnmValueAxis.getNamedItem("valFontColor") != null) {
					int[] intArray = CommonUtils.getIntegerArray(nnmValueAxis.getNamedItem("valFontColor").getNodeValue().split(","));
					valueAxis.setLabelPaint(new Color(intArray[0], intArray[1], intArray[2]));
				}

				if (nnmValueAxis.getNamedItem("valFontStyle") != null) 
					fntStyle = getDecodedFontStyle(nnmValueAxis.getNamedItem("valFontStyle").getNodeValue());

				if (nnmValueAxis.getNamedItem("valMargins") != null) {
					double[] doubleArray = CommonUtils.getDoubleArray(nnmValueAxis.getNamedItem("valMargins").getNodeValue().split(","));
					valueAxis.setLabelInsets(new RectangleInsets(doubleArray[0], doubleArray[1], doubleArray[2], doubleArray[3]));
				}

				fntValueAxis = fntValueAxis.deriveFont(fntStyle, fntSize);
				valueAxis.setLabelFont(fntValueAxis);
			}
		}
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