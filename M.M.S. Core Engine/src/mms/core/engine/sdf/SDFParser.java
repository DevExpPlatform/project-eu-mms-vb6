package mms.core.engine.sdf;

import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;

import org.jfree.chart.JFreeChart;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.itextpdf.awt.PdfGraphics2D;
import com.itextpdf.text.BaseColor;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.Utilities;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfTemplate;

import mms.core.engine.pdfmerger.commons.CommonUtils;
import mms.core.engine.sdf.entities.chart.ChartBar;
import mms.core.engine.sdf.entities.chart.ChartManager;
import mms.core.engine.sdf.entities.chart.ChartPie;
import mms.core.engine.sdf.entities.chart.ChartRing;
import mms.core.engine.sdf.entities.table.RoundCorners;
import mms.core.engine.sdf.entities.text.SDFText;

public class SDFParser {

	private Rectangle 	   boundingBox      = null;
	private float 		   boundingBoxH 	= 0.0f;
	private float 		   boundingBoxW		= 0.0f;
	private float 		   boundingBoxX		= 0.0f;
	private float 		   boundingBoxY		= 0.0f;
	private BaseFont 	   fieldFont		= null;
	private float 		   fieldFontSize 	= 8.0f;
	private int 		   fieldRotation	= 0;
	private String 		   fontsPath 		= "";
	private String 		   imagesPath		= "";
	private PdfContentByte pdfContentByte   = null;
	private float 		   shiftY 			= 0.0f;
	private float          shiftYTmp        = 0.0f;
	private float          tblHeight        = 0.0f;

	private PdfTemplate getChart(Element chartElement, float chartSizeW, float chartSizeH) {
		JFreeChart 	 chart 		  = null;
		PdfTemplate  chartHolder  = this.pdfContentByte.createTemplate(chartSizeW, chartSizeH);
		ChartManager chartManager = null;

		switch (chartElement.getAttributes().getNamedItem("type").getNodeValue()) {
		case "bars":
			chartManager = new ChartBar(false);
	        chartManager.setFontsPath(this.fontsPath);
	        chartManager.setFont(this.fieldFont, this.fieldFontSize);

	        chart = chartManager.getChart(chartElement);

			break;
		case "barsGrouped":
			chartManager = new ChartBar(true);
	        chartManager.setFontsPath(this.fontsPath);
	        chartManager.setFont(this.fieldFont, this.fieldFontSize);

	        chart = chartManager.getChart(chartElement);

			break;
		case "pie":
			chartManager = new ChartPie();
	        chartManager.setFontsPath(this.fontsPath);
	        chartManager.setFont(this.fieldFont, this.fieldFontSize);

	        chart = chartManager.getChart(chartElement);

			break;
		case "ringPie":
			chartManager = new ChartRing();
	        chartManager.setFontsPath(this.fontsPath);
	        chartManager.setFont(this.fieldFont, this.fieldFontSize);

	        chart = chartManager.getChart(chartElement);
		}

		Graphics2D  graphicsChart = new PdfGraphics2D(chartHolder, chartSizeW, chartSizeH, chartManager.getFontMapper());
		Rectangle2D chartRegion   = new Rectangle2D.Double(0, 0, chartSizeW, chartSizeH);

		chart.draw(graphicsChart, chartRegion);
		graphicsChart.dispose();

		return chartHolder;
	}

	private void getChartEntity(Element chartElement) {
		float chartSizeH = (this.boundingBoxY - this.boundingBoxH);

		this.pdfContentByte.addTemplate(this.getChart(chartElement, (this.boundingBoxW - this.boundingBoxX), chartSizeH), this.boundingBoxX, (this.boundingBoxY - chartSizeH));
	}

	private int getDecodedAlignment(String nodeValue) {
		int rValue = com.itextpdf.text.Element.ALIGN_LEFT;

		switch (nodeValue) {
		case "bottom":
			rValue = com.itextpdf.text.Element.ALIGN_BOTTOM;

			break;
		case "center":
			rValue = com.itextpdf.text.Element.ALIGN_CENTER;

			break;
		case "justified":
			rValue = com.itextpdf.text.Element.ALIGN_JUSTIFIED;

			break;
		case "justified_all":
			rValue = com.itextpdf.text.Element.ALIGN_JUSTIFIED_ALL;

			break;
		case "middle":
			rValue = com.itextpdf.text.Element.ALIGN_MIDDLE;

			break;
		case "right":
			rValue = com.itextpdf.text.Element.ALIGN_RIGHT;

			break;
		case "top":
			rValue = com.itextpdf.text.Element.ALIGN_TOP;

			break;
		}

		return rValue;
	}

	public void getEntities(Node elementNode) {
		this.shiftY    = 0.0f;
		this.shiftYTmp = 0.0f;
		this.tblHeight = this.boundingBoxY;

		NodeList nodelistEntities = elementNode.getChildNodes();

		for (int i = 0; i < nodelistEntities.getLength(); i++) {
			Node nodeEntity = nodelistEntities.item(i);

			switch (nodeEntity.getNodeName()) {
			case "chart":
				this.getChartEntity((Element) nodeEntity);

				break;
			case "table":
				this.getTableEntity((Element) nodeEntity, false);

				break;
			case "text":
				this.getTextEntity((Element) nodeEntity);

				break;
			}
		}
	}

	private PdfPTable getTableEntity(Element tableElement, boolean innerTable) {
		int 	  	baseAlignH		= getDecodedAlignment("left");
		int 	  	baseAlignV		= getDecodedAlignment("middle");
		int	      	baseBackColor[] = null;
		int 	 	baseBorders  	= Rectangle.NO_BORDER;
		BaseColor[] baseBordersC 	= new BaseColor[5];
		float[]  	baseBordersW 	= {-1, -1, -1, -1, -1};
		String    	baseFontName 	= "";
		float	  	baseFontSize	= -1;
		String    	baseFontStyle 	= "";
		int	     	baseForeColor[] = null;
		float	  	baseHeight      = -1;
		float     	baseLeading		= -1;
		float[]   	basePadding 	= {-1, -1, -1, -1, -1};
		PdfPTable 	pdfPTable    	= new PdfPTable(Integer.valueOf(tableElement.getAttributes().getNamedItem("columns").getNodeValue()));
		NodeList  	tableCells 		= tableElement.getChildNodes();
		float 	  	tableShiftY   	= 0;

		pdfPTable.setTotalWidth(this.boundingBoxW);
		pdfPTable.setWidthPercentage(100.0f);

		if (tableElement.hasAttribute("widths")) {
			try {
				pdfPTable.setWidths(CommonUtils.getFloatArray(tableElement.getAttribute("widths").split(",")));
			} catch (DocumentException e) {
				System.out.println(e.getMessage());
			}
		}

		if (tableElement.hasAttribute("tableshiftY"))
			tableShiftY = Float.valueOf(tableElement.getAttributes().getNamedItem("tableshiftY").getNodeValue());

		if (tableElement.hasAttribute("borders")) {
			baseBorders = Rectangle.BOX;
			baseBordersW[0] = Float.valueOf(tableElement.getAttribute("borders"));

			if (tableElement.hasAttribute("bordersColor")) {
				int[] intArray = CommonUtils.getIntegerArray(tableElement.getAttribute("bordersColor").split(","));

				baseBordersC[0] = new BaseColor(intArray[0], intArray[1], intArray[2]);
			}
		}

		if (baseBorders == Rectangle.NO_BORDER) {
			if (tableElement.hasAttribute("borderLeft")) {
				baseBorders = Rectangle.LEFT;
				baseBordersW[1] = Float.valueOf(tableElement.getAttribute("borderLeft"));

				if (tableElement.hasAttribute("borderLeftColor")) {
					int[] intArray = CommonUtils.getIntegerArray(tableElement.getAttribute("borderLeftColor").split(","));

					baseBordersC[1] = new BaseColor(intArray[0], intArray[1], intArray[2]);
				}
			}

			if (tableElement.hasAttribute("borderTop")) {
				baseBorders |= Rectangle.TOP;
				baseBordersW[2] = Float.valueOf(tableElement.getAttribute("borderTop"));

				if (tableElement.hasAttribute("borderTopColor")) {
					int[] intArray = CommonUtils.getIntegerArray(tableElement.getAttribute("borderTopColor").split(","));

					baseBordersC[2] = new BaseColor(intArray[0], intArray[1], intArray[2]);
				}
			}

			if (tableElement.hasAttribute("borderRight")) {
				baseBorders |= Rectangle.RIGHT;
				baseBordersW[3] = Float.valueOf(tableElement.getAttribute("borderRight"));

				if (tableElement.hasAttribute("borderRightColor")) {
					int[] intArray = CommonUtils.getIntegerArray(tableElement.getAttribute("borderRightColor").split(","));

					baseBordersC[3] = new BaseColor(intArray[0], intArray[1], intArray[2]);
				}
			}

			if (tableElement.hasAttribute("borderBottom")) {
				baseBorders |= Rectangle.BOTTOM;
				baseBordersW[4] = Float.valueOf(tableElement.getAttribute("borderBottom"));

				if (tableElement.hasAttribute("borderBottomColor")) {
					int[] intArray = CommonUtils.getIntegerArray(tableElement.getAttribute("borderBottomColor").split(","));

					baseBordersC[4] = new BaseColor(intArray[0], intArray[1], intArray[2]);
				}
			}
		}

		if (tableElement.hasAttribute("padding")) {
			basePadding[0] = Float.valueOf(tableElement.getAttributes().getNamedItem("padding").getNodeValue());
		} else {
			if (tableElement.hasAttribute("paddingLeft"))
				basePadding[1] = Float.valueOf(tableElement.getAttributes().getNamedItem("paddingLeft").getNodeValue());

			if (tableElement.hasAttribute("paddingTop"))
				basePadding[2] = Float.valueOf(tableElement.getAttributes().getNamedItem("paddingTop").getNodeValue());

			if (tableElement.hasAttribute("paddingRight"))
				basePadding[3] = Float.valueOf(tableElement.getAttributes().getNamedItem("paddingRight").getNodeValue());

			if (tableElement.hasAttribute("paddingBottom"))
				basePadding[4] = Float.valueOf(tableElement.getAttributes().getNamedItem("paddingBottom").getNodeValue());
		}

		if (tableElement.hasAttribute("rowheight"))
			baseHeight = Float.valueOf(tableElement.getAttributes().getNamedItem("rowheight").getNodeValue());

		if (tableElement.hasAttribute("fontname"))
			baseFontName = tableElement.getAttributes().getNamedItem("fontname").getNodeValue();

		if (tableElement.hasAttribute("fontsize"))
			baseFontSize = Float.valueOf(tableElement.getAttributes().getNamedItem("fontsize").getNodeValue());

		if (tableElement.hasAttribute("fontstyle"))
			baseFontStyle = tableElement.getAttributes().getNamedItem("fontstyle").getNodeValue();

		if (tableElement.hasAttribute("rgbforecolor"))
			baseForeColor = CommonUtils.getIntegerArray(tableElement.getAttribute("rgbforecolor").split(","));

		if (tableElement.hasAttribute("rgbbackcolor"))
			baseBackColor = CommonUtils.getIntegerArray(tableElement.getAttribute("rgbbackcolor").split(","));

		if (tableElement.hasAttribute("alignmentH"))
			baseAlignH = getDecodedAlignment(tableElement.getAttribute("alignmentH"));

		if (tableElement.hasAttribute("alignmentV"))
			baseAlignV = getDecodedAlignment(tableElement.getAttribute("alignmentV"));

		if (tableElement.hasAttribute("leading"))
			baseLeading = Float.valueOf(tableElement.getAttributes().getNamedItem("leading").getNodeValue());

		for (int tableCellIdx = 0; tableCellIdx < tableCells.getLength(); tableCellIdx++) {
			Node nodeCell = tableCells.item(tableCellIdx);

			if ((nodeCell.getNodeType() == Node.ELEMENT_NODE) && nodeCell.getNodeName().equals("cell")) {
				SDFText 	 sdfText 			= new SDFText();
				int 	 	 cellAlignH  		= baseAlignH;
				int 	 	 cellAlignV  		= baseAlignV;
				BaseColor 	 cellBackColor		= null;
				int 	 	 cellBorders  		= Rectangle.NO_BORDER;
				BaseColor[]  cellBordersC 		= new BaseColor[5];
				float[]  	 cellBordersW 		= { -1, -1, -1, -1, -1 };
				int 	 	 cellColSpan 		= 0;
				boolean 	 cellChart  		= false;
				float 		 cellChartHeight  	= 100;
				float 		 cellChartWidth		= 100;
				boolean 	 cellChunked 		= false;
				String[] 	 cellData 			= null;
				Element 	 cellElement 		= (Element) nodeCell;
				float		 cellHeight			= -1;
				boolean 	 cellImage			= false;
				String 		 cellImageName		= "";
				float 		 cellImageScale		= -1;
				float 		 cellLeading     	= -1;
				float[]  	 cellPadding 		= { -1, -1, -1, -1, -1 };
				Integer 	 cellRowSpan 		= 0;
				boolean 	 cellTable			= false;
				PdfPCell 	 pdfPCell			= null;
				RoundCorners roundedBorders 	= null;

				if (this.fontsPath.equals("")) {
					sdfText.setFont(this.fieldFont);
				} else {
					sdfText.setFontBasePath(this.fontsPath);
				}

				if (cellElement.hasAttribute("chart")) {
					cellChart = cellElement.getAttribute("chart").equals("true");

					if (cellElement.hasAttribute("chartHeight"))
						cellChartHeight = Float.valueOf(cellElement.getAttribute("chartHeight"));

					if (cellElement.hasAttribute("chartWidth"))
						cellChartWidth = Float.valueOf(cellElement.getAttribute("chartWidth"));
				} else if (cellElement.hasAttribute("chunked")) {
					cellChunked = cellElement.getAttribute("chunked").equals("true");
				} else if (cellElement.hasAttribute("image")) {
					cellImage     = true;
					cellImageName = cellElement.getAttribute("image");

					if (cellElement.hasAttribute("imagescale"))
						cellImageScale = Float.valueOf(cellElement.getAttribute("imagescale"));
				} else if (cellElement.hasAttribute("table")) {
					cellTable = cellElement.getAttribute("table").equals("true");
				}

				if (cellElement.hasAttribute("colspan"))
					cellColSpan = Integer.valueOf(cellElement.getAttribute("colspan"));

				if (cellElement.hasAttribute("rowspan"))
					cellRowSpan = Integer.valueOf(cellElement.getAttribute("rowspan"));


				if (cellElement.hasAttribute("borders")) {
					cellBorders = Rectangle.BOX;
					cellBordersW[0] = Float.valueOf(cellElement.getAttribute("borders"));

					if (cellElement.hasAttribute("bordersColor")) {
						int[] intArray = CommonUtils.getIntegerArray(cellElement.getAttribute("bordersColor").split(","));

						cellBordersC[0] = new BaseColor(intArray[0], intArray[1], intArray[2]);
					}
				} else {
					cellBorders = baseBorders;
					cellBordersW[0] = baseBordersW[0];
					cellBordersC[0] = baseBordersC[0];
				}

				if (cellBordersW[0] == -1) {
					if (cellElement.hasAttribute("borderLeft")) {
						cellBorders = Rectangle.LEFT;
						cellBordersW[1] = Float.valueOf(cellElement.getAttribute("borderLeft"));

						if (cellElement.hasAttribute("borderLeftColor")) {
							int[] intArray = CommonUtils.getIntegerArray(cellElement.getAttribute("borderLeftColor").split(","));

							cellBordersC[1] = new BaseColor(intArray[0], intArray[1], intArray[2]);
						}
					} else {
						cellBorders     = baseBorders;
						cellBordersW[1] = baseBordersW[1];
						cellBordersC[1] = baseBordersC[1];
					}

					if (cellElement.hasAttribute("borderTop")) {
						cellBorders     |= Rectangle.TOP;
						cellBordersW[2]  = Float.valueOf(cellElement.getAttribute("borderTop"));

						if (cellElement.hasAttribute("borderTopColor")) {
							int[] intArray = CommonUtils.getIntegerArray(cellElement.getAttribute("borderTopColor").split(","));

							cellBordersC[2] = new BaseColor(intArray[0], intArray[1], intArray[2]);
						}
					} else {
						cellBorders     = baseBorders;
						cellBordersW[2] = baseBordersW[2];
						cellBordersC[2] = baseBordersC[2];
					}

					if (cellElement.hasAttribute("borderRight")) {
						cellBorders |= Rectangle.RIGHT;
						cellBordersW[3] = Float.valueOf(cellElement.getAttribute("borderRight"));

						if (cellElement.hasAttribute("borderRightColor")) {
							int[] intArray = CommonUtils.getIntegerArray(cellElement.getAttribute("borderRightColor").split(","));

							cellBordersC[3] = new BaseColor(intArray[0], intArray[1], intArray[2]);
						}
					} else {
						cellBorders = baseBorders;
						cellBordersW[3] = baseBordersW[3];
						cellBordersC[3] = baseBordersC[3];
					}

					if (cellElement.hasAttribute("borderBottom")) {
						cellBorders |= Rectangle.BOTTOM;
						cellBordersW[4] = Float.valueOf(cellElement.getAttribute("borderBottom"));

						if (cellElement.hasAttribute("borderBottomColor")) {
							int[] intArray = CommonUtils.getIntegerArray(cellElement.getAttribute("borderBottomColor").split(","));

							cellBordersC[4] = new BaseColor(intArray[0], intArray[1], intArray[2]);
						}
					} else {
						cellBorders = baseBorders;
						cellBordersW[4] = baseBordersW[4];
						cellBordersC[4] = baseBordersC[4];
					}
				}

				if (cellElement.hasAttribute("padding")) {
					cellPadding[0] = Float.valueOf(cellElement.getAttributes().getNamedItem("padding").getNodeValue());
				} else {
					cellPadding[0] = basePadding[0];
				}

				if (cellPadding[0] == -1) {
					if (cellElement.hasAttribute("paddingLeft")) {
						cellPadding[1] = Float.valueOf(cellElement.getAttributes().getNamedItem("paddingLeft").getNodeValue());
					} else {
						cellPadding[1] = basePadding[1];
					}

					if (cellElement.hasAttribute("paddingTop")) {
						cellPadding[2] = Float.valueOf(cellElement.getAttributes().getNamedItem("paddingTop").getNodeValue());
					} else {
						cellPadding[2] = basePadding[2];
					}

					if (cellElement.hasAttribute("paddingRight")) {
						cellPadding[3] = Float.valueOf(cellElement.getAttributes().getNamedItem("paddingRight").getNodeValue());
					} else {
						cellPadding[3] = basePadding[3];
					}

					if (cellElement.hasAttribute("paddingBottom")) {
						cellPadding[4] = Float.valueOf(cellElement.getAttributes().getNamedItem("paddingBottom").getNodeValue());
					} else {
						cellPadding[4] = basePadding[4];
					}
				}

				if (cellElement.hasAttribute("cellheight")) {
					if (cellElement.getAttribute("cellheight").endsWith("mm")) {
						cellHeight = Utilities.millimetersToPoints(Float.valueOf(cellElement.getAttribute("cellheight").substring(0, (cellElement.getAttribute("cellheight").length() -2))));
					} else {
						cellHeight = Float.valueOf(cellElement.getAttribute("cellheight"));
					}
				} else {
					cellHeight = baseHeight;
				}

				if (cellElement.hasAttribute("fontname")) {
					sdfText.setFont(cellElement.getAttribute("fontname"));
				} else {
					if (!baseFontName.equals(""))
						sdfText.setFont(baseFontName);
				}

				if (cellElement.hasAttribute("fontsize")) {
					sdfText.setFontSize(Float.valueOf(cellElement.getAttribute("fontsize")));
				} else {
					if (baseFontSize > -1)
						sdfText.setFontSize(baseFontSize);
				}

				if (cellElement.hasAttribute("fontstyle")) {
					sdfText.setFontStyle(cellElement.getAttribute("fontstyle"));
				} else {
					if (!baseFontStyle.equals(""))
						sdfText.setFontStyle(baseFontStyle);
				}

				if (cellElement.hasAttribute("rgbforecolor")) {
					int intArray[] = CommonUtils.getIntegerArray(cellElement.getAttribute("rgbforecolor").split(","));

					sdfText.setFontColor(intArray[0], intArray[1], intArray[2]);
				} else {
					if (baseForeColor != null)
						sdfText.setFontColor(baseForeColor[0], baseForeColor[1], baseForeColor[2]);
				}

				if (cellElement.hasAttribute("rgbbackcolor")) {
					int intArray[] = CommonUtils.getIntegerArray(cellElement.getAttribute("rgbbackcolor").split(","));

					cellBackColor = new BaseColor(intArray[0], intArray[1], intArray[2]);
				} else {
					if (baseBackColor != null)
						cellBackColor = new BaseColor(baseBackColor[0], baseBackColor[1], baseBackColor[2]);
				}

				if (cellElement.hasAttribute("roundBorders")) {
					BaseColor cellBorderColor = new BaseColor(0,0,0);

					if (cellElement.hasAttribute("roundBordersColor")) {
						int[] intArray  = CommonUtils.getIntegerArray(cellElement.getAttribute("roundBordersColor").split(","));
						cellBorderColor = new BaseColor(intArray[0], intArray[1], intArray[2]);
					}

					String[] rValues = cellElement.getAttribute("roundBorders").split(",");
					roundedBorders   = new RoundCorners(cellBorderColor, cellBackColor, Float.valueOf(rValues[0]), Float.valueOf(rValues[1]), rValues[2].equals("1"), rValues[3].equals("1"), rValues[4].equals("1"), rValues[5].equals("1"));
				}

				if (cellElement.hasAttribute("alignmentH"))
					cellAlignH = getDecodedAlignment(cellElement.getAttribute("alignmentH"));

				if (cellElement.hasAttribute("alignmentV"))
					cellAlignV = getDecodedAlignment(cellElement.getAttribute("alignmentV"));

				if (cellElement.hasAttribute("leading")) {
					cellLeading = Float.valueOf(cellElement.getAttributes().getNamedItem("leading").getNodeValue());
				} else {
					cellLeading = baseLeading;
				}

				/*
				 * Write to Cell
				 */
				if (cellChart || cellChunked || cellImage ||cellTable) {
					cellData = new String[1];
					cellData[0] = "";
				} else {
					if (cellElement.getChildNodes().item(0) == null || cellElement.getChildNodes().item(0).getNodeValue().equals(" ")) {
						cellData = new String[1];
						cellData[0] = "";
					} else if (cellElement.getChildNodes().item(0).getNodeValue().contains("|")) {
						cellData = cellElement.getChildNodes().item(0).getNodeValue().split("\\|");
					} else {
						cellData = new String[1];
						cellData[0] = cellElement.getChildNodes().item(0).getNodeValue();
					}
				}

				for (String element : cellData) {
					if (cellChunked) {
						pdfPCell = new PdfPCell(this.getText((Element) cellElement.getElementsByTagName("text").item(0)).getText());
					} else if (cellChart) {
						try {
							Image image = Image.getInstance(this.getChart((Element) cellElement.getElementsByTagName("chart").item(0), cellChartWidth, cellChartHeight));
							image.setAlignment(cellAlignH);

							pdfPCell = new PdfPCell(image);
						} catch (Exception e) {
							e.printStackTrace();
						}
					} else if (cellImageName.length() > 0) {
						try {
							Image image = Image.getInstance(imagesPath + cellImageName);
							image.setAlignment(cellAlignH);

							if (cellImageScale > -1)
								image.scalePercent(cellImageScale);

							pdfPCell = new PdfPCell();
							pdfPCell.addElement(image);
						} catch (Exception e) {
							e.printStackTrace();
						}
					} else if (cellTable) {
						pdfPCell = new PdfPCell();
						pdfPCell.addElement(this.getTableEntity((Element) cellElement.getElementsByTagName("table").item(0), true));
					} else {
						if (element.equals("")) {
							pdfPCell = new PdfPCell(new Phrase(""));
						} else {
							sdfText.addChunk(element);

							pdfPCell = new PdfPCell(new Paragraph(sdfText.getText()));

							if (cellData.length > 1)
								sdfText.clear();
						}
					}

					if (cellColSpan > 0)
						pdfPCell.setColspan(cellColSpan);

					if (cellRowSpan > 0)
						pdfPCell.setRowspan(cellRowSpan);

					if (roundedBorders != null)
						pdfPCell.setCellEvent(roundedBorders);

					pdfPCell.setBorder(cellBorders);

					if (cellBordersW[0] > -1) {
						pdfPCell.setBorderWidth(cellBordersW[0]);

						if (cellBordersC[0] != null)
							pdfPCell.setBorderColor(cellBordersC[0]);
					} else {
						if (cellBordersW[1] > 0) {
							pdfPCell.setBorderWidthLeft(cellBordersW[1]);

							if (cellBordersC[1] != null)
								pdfPCell.setBorderColorLeft(cellBordersC[1]);
						}

						if (cellBordersW[2] > 0) {
							pdfPCell.setBorderWidthTop(cellBordersW[2]);

							if (cellBordersC[2] != null)
								pdfPCell.setBorderColorTop(cellBordersC[2]);
						}

						if (cellBordersW[3] > 0) {
							pdfPCell.setBorderWidthRight(cellBordersW[3]);

							if (cellBordersC[3] != null)
								pdfPCell.setBorderColorRight(cellBordersC[3]);
						}

						if (cellBordersW[4] > 0) {
							pdfPCell.setBorderWidthBottom(cellBordersW[4]);

							if (cellBordersC[4] != null)
								pdfPCell.setBorderColorBottom(cellBordersC[4]);
						}
					}

					if (cellPadding[0] > -1) {
						pdfPCell.setPadding(cellPadding[0]);
					} else {
						if (cellPadding[1] > -1)
							pdfPCell.setPaddingLeft(cellPadding[1]);

						if (cellPadding[2] > -1)
							pdfPCell.setPaddingTop(cellPadding[2]);

						if (cellPadding[3] > -1)
							pdfPCell.setPaddingRight(cellPadding[3]);

						if (cellPadding[4] > -1)
							pdfPCell.setPaddingBottom(cellPadding[4]);
					}

					if (cellHeight > 0)
						pdfPCell.setFixedHeight(cellHeight);
//						myCell.setMinimumHeight(cellHeight);

					pdfPCell.setHorizontalAlignment(cellAlignH);
					pdfPCell.setVerticalAlignment(cellAlignV);

					if ((roundedBorders == null) && (cellBackColor != null))
						pdfPCell.setBackgroundColor(cellBackColor);

					if (cellLeading > -1)
						pdfPCell.setLeading(cellLeading, 1.2f);

					if (fieldRotation != 0)
						pdfPCell.setRotation(fieldRotation);

					pdfPTable.addCell(pdfPCell);
				}
			}
		}

		if (!innerTable) {
			ColumnText columnText = new ColumnText(this.pdfContentByte);
			columnText.setSimpleColumn(this.boundingBoxX, (this.boundingBoxY - this.shiftY - tableShiftY), this.boundingBoxW, 0);
			columnText.addElement(pdfPTable);

			try {
				columnText.go();
			} catch (DocumentException e) {
				e.printStackTrace();
			}

			this.shiftY    = (this.boundingBoxY - columnText.getYLine());
			this.tblHeight = (this.shiftY - this.shiftYTmp);
			this.shiftYTmp = this.shiftY;
		}

		return pdfPTable;
	}

	public float getTableHeight() {
		return this.tblHeight;
	}

	private SDFText getText(Element textElement) {
		int	     baseColor[]   	= null;
		String   baseFontName 	= "";
		float	 baseFontSize	= -1;
		String   baseFontStyle 	= "";
		SDFText  sdfText      	= new SDFText();
		NodeList textChunks		= textElement.getElementsByTagName("chunk");

		if (this.fontsPath.equals("")) {
			sdfText.setFont(this.fieldFont);
		} else {
			sdfText.setFontBasePath(this.fontsPath);
		}

		if (textElement.hasAttribute("fontname"))
			baseFontName = textElement.getAttributes().getNamedItem("fontname").getNodeValue();

		if (textElement.hasAttribute("fontsize"))
			baseFontSize = Float.valueOf(textElement.getAttributes().getNamedItem("fontsize").getNodeValue());

		if (textElement.hasAttribute("fontstyle"))
			baseFontStyle = textElement.getAttributes().getNamedItem("fontstyle").getNodeValue();

		if (textElement.hasAttribute("rgbcolor"))
			baseColor = CommonUtils.getIntegerArray(textElement.getAttribute("rgbcolor").split(","));

		for (int textChunk = 0; textChunk < textChunks.getLength(); textChunk++) {
			Element textChunksElement = (Element) textChunks.item(textChunk);

			if (textChunksElement.hasAttribute("fontsize")) {
				sdfText.setFontSize(Float.valueOf(textChunksElement.getAttribute("fontsize")));
			} else {
				if (baseFontSize > -1)
					sdfText.setFontSize(baseFontSize);
			}

			if (textChunksElement.hasAttribute("fontstyle")) {
				sdfText.setFontStyle(textChunksElement.getAttribute("fontstyle"));
			} else {
				if (baseFontStyle.equals("")) {
					sdfText.setFontStyle("normal");
				} else {
					sdfText.setFontStyle(baseFontStyle);
				}
			}

			if (textChunksElement.hasAttribute("fontname")) {
				sdfText.setFont(textChunksElement.getAttribute("fontname"));
			} else {
				if (!baseFontName.equals(""))
					sdfText.setFont(baseFontName);
			}

			if (textChunksElement.hasAttribute("rgbcolor")) {
				int intArray[] = CommonUtils.getIntegerArray(textChunksElement.getAttribute("rgbcolor").split(","));

				sdfText.setFontColor(intArray[0], intArray[1], intArray[2]);
			} else {
				if (baseColor != null)
					sdfText.setFontColor(baseColor[0], baseColor[1], baseColor[2]);
			}

			sdfText.addChunk(textChunksElement.getChildNodes().item(0).getNodeValue());
		}

		return sdfText;
	}

	private void getTextEntity(Element textElement) {
		float      baseLeading	  = -1;
		float      baseLeadingFix = -1;
		int        baseTextAlign  = com.itextpdf.text.Element.ALIGN_LEFT;
		ColumnText columnText     = new ColumnText(this.pdfContentByte);
		SDFText    sdfText        = this.getText(textElement);

		if (textElement.hasAttribute("alignment"))
			baseTextAlign = getDecodedAlignment(textElement.getAttributes().getNamedItem("alignment").getNodeValue());

		if (textElement.hasAttribute("leading"))
			baseLeading = Float.valueOf(textElement.getAttributes().getNamedItem("leading").getNodeValue());

		if (textElement.hasAttribute("leadingfixed"))
			baseLeadingFix = Float.valueOf(textElement.getAttributes().getNamedItem("leadingfixed").getNodeValue());

		if (fieldRotation == 0) {
			columnText.setSimpleColumn(this.boundingBoxX, (this.boundingBoxY - this.shiftY), this.boundingBoxW, 0);
			columnText.setText(sdfText.getText());
			columnText.setAlignment(baseTextAlign);

			if ((baseLeading == -1) && (baseLeadingFix == -1)) {
				columnText.setLeading(sdfText.getFontSize(), 1.2f);
			} else {
				if (baseLeading != -1)
					columnText.setLeading(baseLeading, 1.2f);

				if (baseLeadingFix != -1)
					columnText.setLeading(baseLeadingFix);
			}
		} else {
			PdfPCell pdfPCell = new PdfPCell(sdfText.getText());
			pdfPCell.setBorder(0);
			pdfPCell.setFixedHeight(this.boundingBox.getHeight());
			pdfPCell.setHorizontalAlignment(baseTextAlign);
			pdfPCell.setLeading((baseLeading == -1 ? sdfText.getFontSize() : baseLeading), 1.2f);
			pdfPCell.setPadding(0);
			pdfPCell.setRotation(this.fieldRotation);

			PdfPTable pdfPTable = new PdfPTable(1);
			pdfPTable.setTotalWidth(this.boundingBoxW);
			pdfPTable.setWidthPercentage(100.0f);
			pdfPTable.addCell(pdfPCell);

			columnText.setSimpleColumn(this.boundingBoxX, this.boundingBoxY, this.boundingBoxW, this.boundingBoxH);
			columnText.addElement(pdfPTable);
		}

		try {
			columnText.go();
		} catch (DocumentException e) {
			e.printStackTrace();
		}

		if (fieldRotation == 0)
			this.shiftY = (this.boundingBoxY - columnText.getYLine());
	}

	public void setBoundingBox(Rectangle fieldPosition) {
		this.boundingBox  = fieldPosition;
		this.boundingBoxX = fieldPosition.getLeft();
		this.boundingBoxY = fieldPosition.getTop();
		this.boundingBoxW = (this.boundingBoxX + fieldPosition.getWidth());
		this.boundingBoxH = (this.boundingBoxY - fieldPosition.getHeight());
	}

	public void setFieldRotation(int fieldRotation) {
		this.fieldRotation = fieldRotation;
	}

	public void setFont(BaseFont font, float fontSize) {
		this.fieldFont     = font;
		this.fieldFontSize = fontSize;
	}

	public void setPDFContentByte(PdfContentByte pdfContentByte) {
		this.pdfContentByte = pdfContentByte;
	}

	public void setPrjBasePath(String prjBasePath) {
		this.fontsPath  = prjBasePath + "Fonts/";
		this.imagesPath = prjBasePath + "Images/";
	}

}
