package mms.core.engine.packer;

import java.awt.Color;
import java.awt.FontFormatException;
import java.awt.color.ColorSpace;
import java.awt.geom.AffineTransform;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Map.Entry;
import java.util.Random;

import javax.imageio.IIOImage;
import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.ImageWriter;
import javax.imageio.metadata.IIOMetadata;
import javax.imageio.stream.ImageInputStream;
import javax.imageio.stream.ImageOutputStream;

import com.pt.awt.font.NFontType1;
import com.pt.io.InputUni;
import com.pt.io.InputUniByteArray;
import com.pt.io.InputUniFile;
import com.pt.io.OutputUni;

import multivalent.ParseException;
import multivalent.std.adaptor.pdf.COS;
import multivalent.std.adaptor.pdf.Cmd;
import multivalent.std.adaptor.pdf.CryptFilter;
import multivalent.std.adaptor.pdf.Dict;
import multivalent.std.adaptor.pdf.Fonts;
import multivalent.std.adaptor.pdf.IRef;
import multivalent.std.adaptor.pdf.Images;
import multivalent.std.adaptor.pdf.InputStreamComposite;
import multivalent.std.adaptor.pdf.PDFReader;
import multivalent.std.adaptor.pdf.PDFWriter;
import phelps.lang.Integers;
import phelps.util.Arrayss;
import phelps.util.Version;

public class PDFPacker {

	
	public  static final String 	COPYRIGHT 		 = "Copyright (c) 2002 - 2005  Thomas A. Phelps.  All rights reserved.";
	private static final String[] 	CORE14			 = new String[] { "Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic", "Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique", "Courier", "Courier-Bold", "Courier-Oblique", "Courier-BoldOblique", "Symbol", "ZapfDingbats", "NimbusRomNo9L-Regu", "NimbusRomNo9L-ReguItal", "NimbusRomNo9L-Medi", "NimbusRomNo9L-MediItal", "NimbusSanL-Regu",  "NimbusSanL-Bold", "NimbusSanL-ReguItal", "NimbusSanL-BoldItal", "NimbusMonL-Regu", "NimbusMonL-Bold", "NimbusMonL-ReguObli", "NimbusMonL-BoldObli", "Standard Symbols L", "Dingbats" };
	private static final int 		PREDICT_OVERHEAD = "/DecodeParms<</Predictor 15/Columns10/Colors 3>>>>".length();
	private static final int[] 		TTF_KEEP		 = new int[] { 1735162214, 1819239265, 1668112752, 1751474532, 1751672161, 1752003704, 1835104368, 1668707360, 1718642541, 1886545264, 1330851634, 1886352244, 1734439792, 1751412088, 1280594760 };
	public  static final String 	VERSION 		 = "2.4 of $Date: 2005/07/26 21:13:18 $";

	static {
		Arrays.sort(CORE14);
		Arrays.sort(TTF_KEEP);
	}
	
	private static int sign(Object obj) {
		int 					rValue = 0;
		Class<? extends Object> objClass = obj.getClass();
		
		if (COS.OBJECT_NULL == obj)
			rValue = 0;
		else if (COS.CLASS_NAME == objClass)
			rValue = ((String) obj).hashCode();
		else if (COS.CLASS_STRING == objClass)
			rValue = ((StringBuffer) obj).length() * 11;
		else if (COS.CLASS_DATA == objClass)
			rValue = ((byte[]) obj).length << 8;
		else if (COS.CLASS_BOOLEAN == objClass)
			rValue = ((Boolean) obj).booleanValue() ? 11 : 7;
		else if (obj instanceof Number)
			rValue = ((Number) obj).intValue();
		else if (COS.CLASS_IREF == objClass)
			rValue = ((IRef) obj).id * 131;
		else if (COS.CLASS_DICTIONARY == objClass) {
			Dict dict = (Dict) obj;
			rValue += dict.size() * 13;
			
			for (Iterator<?> iterator = dict.entrySet().iterator(); iterator.hasNext();) {
				Entry<?, ?> entry = (Entry<?, ?>) iterator.next();
				
				rValue += entry.getKey().hashCode();
				rValue += sign(entry.getValue());
			}

			rValue *= dict.size();
		} else if (COS.CLASS_ARRAY == objClass) {
			Object aobj[]  = (Object[]) obj;
			rValue        += aobj.length * 7;
			
			int j = 0;
			
			for (int k = aobj.length; j < k; j++)
				rValue += sign(aobj[j]) * (j + 1);
		}

		return rValue;
	}
	
	private boolean 	 alt			= false;
	private boolean 	 compact		= false;
	private boolean 	 compact0		= false;
	private boolean 	 content		= false;
	private boolean 	 core14			= false;
	private boolean 	 embed			= false;
	private boolean 	 force			= false;
	private boolean 	 jpeg			= false;
	private boolean 	 jpeg2000		= false;
	private boolean 	 monitor		= false;
	private boolean 	 old			= false;
	private boolean 	 outline		= false;
	private boolean 	 pagepiece		= false;
	private PDFReader 	 pdfReader		= null;
	private boolean 	 pre			= false;
	private boolean 	 quiet			= false;
	private boolean 	 struct			= false;
	private boolean 	 subset			= false;
	private boolean 	 testable		= false;
	private boolean 	 verbose		= false;
	private boolean 	 webcap			= false;
	private List<String> wacky			= null;;

	public PDFPacker(byte abyte0[]) throws IOException, ParseException {
		this(((InputUni) (new InputUniByteArray(abyte0))));
	}

	public PDFPacker(File file) throws IOException, ParseException {
		this(((InputUni) (new InputUniFile(file))));
	}

	public PDFPacker(InputUni inputuni) throws IOException, ParseException {
		this(new PDFReader(inputuni));
	}

	public PDFPacker(PDFReader pdfreader) {
		this.alt 		= true;
		this.compact0 	= false;
		this.content 	= true;
		this.core14 	= true;
		this.embed 		= true;
		this.force 		= false; 
//		this.inplace 	= false;
		this.jpeg 		= false;
		this.jpeg2000 	= false; 
		this.monitor 	= false;
		this.old      	= false;
		this.outline 	= true;
		this.pagepiece 	= true;
		this.pdfReader	= pdfreader;
		this.pre 		= true;
		this.quiet 		= false;
		this.struct 	= true;
		this.subset 	= false;
		this.testable 	= false;
		this.verbose 	= false;
		this.webcap 	= true;
	}

	private boolean addPredictor(Dict dict, int width, int height, int bitPerComponent, byte buffer[], PDFReader pdfreader) throws IOException {
		if (bitPerComponent != 8 || buffer.length < PREDICT_OVERHEAD + width)
			return false;
		
		Object 	   objDictColorSpace = pdfreader.getObject(dict.get("ColorSpace"));
		ColorSpace colorspace        = pdfreader.getColorSpace(objDictColorSpace, null, null);
		int        numComponents     = colorspace.getNumComponents();
		
		if (COS.CLASS_ARRAY == objDictColorSpace.getClass() && "Indexed".equals(((Object[]) objDictColorSpace)[0]))
			numComponents = 1;
		
		int bitsPerPixel = (numComponents * bitPerComponent + 7) / 8;
		int dstPos1      = bitsPerPixel;
		int imageSize    = (width * height * numComponents * bitPerComponent) / 8;

		if ((buffer.length - imageSize) > 1) {
			String s = "extra bytes " + width + "x" + height + "*" + numComponents + "=" + imageSize + " < " + buffer.length;

			wacky.add(s);
			
			if (monitor)
				System.out.print(" / " + s);
		}
		
		int  length1  	  = (width * bitsPerPixel);
		byte dstBuffer1[] = new byte[(length1 + 1) * height];
		int  dstBufferLen = (length1 + dstPos1);
		byte dstBuffer2[] = new byte[dstBufferLen];
		byte dstBuffer3[] = new byte[dstBufferLen];
		byte dstBuffer4[] = new byte[dstBufferLen];
		byte dstBuffer5[] = new byte[dstBufferLen];
		byte dstBuffer6[] = new byte[dstBufferLen];
		int  dstPos2      = 0;
		
		dstBuffer1[dstPos2++] = 0;
		
		System.arraycopy(buffer, 0, dstBuffer1, dstPos2, length1);
		
		dstPos2 += length1;
		
		int k2 = 1;
		int srcPos2 = length1;
		
		for (int i3 = 0 - dstPos1; k2 < height; i3 += length1) {
			int k3 = 0;
			int l3 = 0;
			int i4 = 0;
			int j4 = 0;
			int k4 = 0;
			
			System.arraycopy(buffer, srcPos2, dstBuffer2, dstPos1, length1);
			
			for (int l4 = dstBufferLen - 1; l4 >= dstPos1; l4--)
				k3 += dstBuffer2[l4] & 0xff;

			int  i5          = k3;
			byte bufferShift = 0;
			byte dstBuffer[] = dstBuffer2;
			
			System.arraycopy(buffer, srcPos2, dstBuffer3, dstPos1, length1);
			
			int k5 = dstBufferLen - 1;
			
			for (int k6 = dstPos1 + bitsPerPixel; k5 >= k6; k5--) {
				dstBuffer3[k5] -= dstBuffer3[k5 - bitsPerPixel] & 0xff;
				l3             += dstBuffer3[k5] & 0xff;
			}

			if (l3 < i5) {
				i5          = l3;
				dstBuffer   = dstBuffer3;
				bufferShift = 1;
			}
			
			System.arraycopy(buffer, srcPos2, dstBuffer4, dstPos1, length1);
			
			for (int l5 = dstBufferLen - 1; l5 >= dstPos1; l5--) {
				dstBuffer4[l5] -= buffer[i3 + l5] & 0xff;
				i4             += dstBuffer4[l5] & 0xff;
			}

			if (i4 < i5) {
				i5          = i4;
				dstBuffer   = dstBuffer4;
				bufferShift = 2;
			}
			
			System.arraycopy(buffer, srcPos2, dstBuffer5, dstPos1, length1);
			
			for (int i6 = dstBufferLen - 1; i6 >= dstPos1; i6--) {
				dstBuffer5[i6] -= ((dstBuffer5[i6 - bitsPerPixel] & 0xff) + (buffer[i3 + i6] & 0xff)) / 2;
				j4             += dstBuffer5[i6] & 0xff;
			}

			if (j4 < i5) {
				i5          = j4;
				dstBuffer   = dstBuffer5;
				bufferShift = 3;
			}
			
			System.arraycopy(buffer, srcPos2, dstBuffer6, dstPos1, length1);
			
			for (int j6 = dstBufferLen - 1; j6 >= dstPos1; j6--) {
				int l6 = dstBuffer6[j6 - bitsPerPixel] & 0xff;
				int i7 = buffer[i3 + j6] & 0xff;
				int j7 = j6 < dstPos1 + bitsPerPixel ? 0 : buffer[(i3 + j6) - bitsPerPixel] & 0xff;
				int k7 = (l6 + i7) - j7;
				int l7 = Math.abs(k7 - l6);
				int i8 = Math.abs(k7 - i7);
				int j8 = Math.abs(k7 - j7);
				int k8 = l7 > i8 || l7 > j8 ? i8 > j8 ? j7 : i7 : l6;
				
				dstBuffer6[j6] -= (byte) k8;
				k4             += dstBuffer6[j6] & 0xff;
			}

			if (k4 < i5) {
				dstBuffer   = dstBuffer6;
				bufferShift = 4;
			}

			dstBuffer1[dstPos2++] = bufferShift;

			System.arraycopy(dstBuffer, dstPos1, dstBuffer1, dstPos2, length1);
			
			dstPos2 += length1;
			k2++;
			srcPos2 += length1;
		}

		byte deflateData[] = PDFWriter.maybeDeflateData(dstBuffer1);
		srcPos2 = deflateData.length;

		int j3 = PDFWriter.maybeDeflateData(buffer).length;
		
		boolean flag = dstBuffer1 != deflateData && srcPos2 + PREDICT_OVERHEAD + 10 + (compact ? 2048 : 0) < j3;
		
		if (flag) {
			if (monitor)
				System.out.print(" pre" + (j3 - srcPos2));
			
			dict.put("Filter", "FlateDecode");
			dict.put("DATA", deflateData);
			
			Dict dictDecodeParams = new Dict(5);
			dictDecodeParams.put("Predictor", Integers.getInteger(15));
			
			if (numComponents != 1)
				dictDecodeParams.put("Colors", Integers.getInteger(numComponents));
			
			dictDecodeParams.put("Columns", Integers.getInteger(width));
			
			if (bitPerComponent != 8)
				dictDecodeParams.put("BitsPerComponent", Integers.getInteger(8));
			
			dict.put("DecodeParms", dictDecodeParams);
		}
		
		return flag;
	}

	private void axeCore14(PDFWriter pdfwriter) throws IOException {
		if (core14 || testable)
			return;
		
		if (monitor)
			System.out.print(", axeCore14");
		
		for (int i = 1; i < pdfwriter.getObjCnt(); i++) {
			Object obj = pdfwriter.getObject(i, false);
			
			if (COS.CLASS_DICTIONARY != obj.getClass())
				continue;
			
			Dict   dict = (Dict) obj;
			Object obj1 = pdfwriter.getObject(dict.get("Type"));
			
			pdfwriter.getObject(dict.get("Subtype"));
			
			if (!"Font".equals(obj1) || !"WinAnsiEncoding".equals(pdfwriter.getObject(dict.get("Encoding"))))
				continue;
			
			String s = (String) pdfwriter.getObject(dict.get("BaseFont"));
			String s1 = Fonts.isSubset(s) ? s.substring("SIXCAP+".length()) : s;
			
			if (!core14 && s1 != null && Arrays.binarySearch(CORE14, s1) >= 0 && dict.get("FontDescriptor") != null) {
				dict.remove("FontDescriptor");
				dict.put("BaseFont", s1);
			}
		}

	}


	private Object Compress(PDFWriter pdfwriter) throws IOException, ParseException {
		long startTimeMillis = System.currentTimeMillis();
		
		PDFReader pdfreader = pdfReader;
		
		if (!pdfreader.isAuthorized())
			throw new ParseException("invalid password");
		
		Dict dictCompress = (Dict) pdfreader.getTrailer().get("Compress");
		
		if (!force && dictCompress != null) {
			String filter = (String) dictCompress.get("Filter");
		
			if (filter == null && !compact0 || "Compact".equals(filter) && compact0) {
				System.out.println("Already Compressed.  (Force reCompression with -force.)");
				
				return null;
			}
		}
		
		long srcPDFLength = pdfreader.getRA().length();
		
		if (!quiet) {
			System.out.print(pdfreader.getURI() + ", " + srcPDFLength + " bytes");
			
			if (pdfreader.getEncrypt().getStmF() != CryptFilter.IDENTITY)
				System.out.print(", encrypted");
			
			System.out.println();
			
			Dict dictInfo = pdfreader.getInfo();
			
			if (dictInfo != null)
				System.out.println("PDF " + pdfreader.getVersion() + ", producer=" + dictInfo.get("Producer") + ", creator=" + dictInfo.get("Creator"));
		}

		if (monitor)
			System.out.print(pdfreader.getObjCnt() + " objects / " + pdfreader.getPageCnt() + " pages");
		
		Compress2(pdfwriter);
		
		if (monitor)
			System.out.print("write ");
		
		Object obj          = pdfwriter.writePDF();
		long   dstPDFLength = pdfwriter.getOutputStream().getCount();
		
		pdfReader.close();
		pdfReader = null;
		
		pdfwriter.close();
		
		long endTimeMillis = System.currentTimeMillis();
		
		if (!this.quiet)
			System.out.println("=> new length = " + dstPDFLength + ", saved " + ((srcPDFLength - dstPDFLength) * 100L) / srcPDFLength + "%, elapsed time = " + (endTimeMillis - startTimeMillis) / 1000L + " sec");
		
		return obj;
	}

	private void Compress2(PDFWriter pdfwriter) throws IOException, ParseException {
		wacky = new ArrayList<String>(5);
		
		PDFReader pdfreader = pdfReader;
		Version   version   = pdfreader.getVersion();
		long      l         = pdfreader.getRA().length();
		
		pdfreader.fault();
		
		Dict dictTrailer = (Dict) pdfwriter.getTrailer().get("Compress");
		
		if (dictTrailer == null) {
			dictTrailer = new Dict(5);
			
			pdfwriter.getTrailer().put("Compress", dictTrailer);
			
			dictTrailer.put("LengthO", new Integer((int) l));
			dictTrailer.put("SpecO", version.toString());
		}
		
		compact = compact0 && l > l / 4L + 714L;
		
		if (compact)
			dictTrailer.put("Filter", "Compact");
		else
			dictTrailer.remove("Filter");
		
		if (old)
			pdfreader.setExact(true);
		else if (!compact)
			pdfwriter.getVersion().setMin(new Version(1L, 5L));
		
		int cntrASCII         = 0;
		int cntrDeleted       = 0;
		int cntrEmbeddedType1 = 0;
		int cntrImage         = 0;
		int cntrIRef 		  = 0;
		int cntrLZW 		  = 0;
		int cntrPagePiece 	  = 0;
		int cntrRawSamples 	  = 0;
		int embeddedType1Size = 0;
		int objId 			  = 1;
		int rawSamplesSize    = 0;
		
		for (int i = pdfreader.getObjCnt(); objId < i; objId++) {
			Object pdfReaderObj = pdfreader.getObject(objId);
			
			if (COS.OBJECT_DELETED == pdfReaderObj)
				cntrDeleted++;
			else if (COS.CLASS_DICTIONARY == pdfReaderObj.getClass()) {
				Dict                    dict     = (Dict) pdfReaderObj;
				Object                  oFilters = pdfreader.getObject(dict.get("Filter"));
				Class<? extends Object> classObj = oFilters == null ? null : oFilters.getClass();
				
				if (COS.CLASS_NAME == classObj) {
					if ("ASCII85Decode".equals(oFilters) || "ASCIIHexDecode".equals(oFilters))
						cntrASCII++;
					else if ("LZWDecode".equals(oFilters))
						cntrLZW++;
				} else if (COS.CLASS_ARRAY == classObj) {
					Object oFiltersArray[]     = (Object[]) oFilters;
					int    oFiltersArrayLength = oFiltersArray.length;
					
					for (int j = 0; j < oFiltersArrayLength; j++) {
						Object oFilter = oFiltersArray[j];
						
						if ("ASCII85Decode".equals(oFilter) || "ASCIIHexDecode".equals(oFilter)) {
							cntrASCII++;
							
							continue;
						}
						
						if ("LZWDecode".equals(oFilter))
							cntrLZW++;
					}

				}
				
				if (dict.get("DATA") != null && (dict.get("Length") instanceof IRef))
					cntrIRef++;
				
				Object oType    = pdfreader.getObject(dict.get("Type"));
				Object oSubType = pdfreader.getObject(dict.get("Subtype"));
				
				if (("XObject".equals(oType) || oType == null) && "Image".equals(oSubType) && dict.get("Alternates") != null)
					cntrImage++;
			}
			
			pdfwriter.getObject(objId, true);
			
			if (COS.CLASS_DICTIONARY != pdfReaderObj.getClass())
				continue;
			
			Dict dict = (Dict) pdfReaderObj;
			
			if (dict.get("PieceInfo") != null)
				cntrPagePiece++;
			
			Object oFilter  = pdfreader.getObject(dict.get("Filter"));
			Object oType    = pdfwriter.getObject(dict.get("Type"));
			Object oSubType = pdfwriter.getObject(dict.get("Subtype"));
			
			if (("XObject".equals(oType) || oType == null) && "Image".equals(oSubType)) {
				byte arrayData[] = (byte[]) dict.get("DATA");
				
				if (oFilter == null) {
					cntrRawSamples++;
					rawSamplesSize += arrayData.length;
				}
				
				recodeImage(dict, pdfwriter, pdfreader);
				
				if (oFilter == null)
					pdfwriter.deflateStream(dict, objId);
				
				continue;
			}
			
			if ("Font".equals(oType) && !embed) {
				Dict fontDescriptor = (Dict) pdfwriter.getObject(dict.get("FontDescriptor"));
				
				if (fontDescriptor != null) {
					fontDescriptor.remove("FontFile");
					fontDescriptor.remove("FontFile2");
					fontDescriptor.remove("FontFile3");
				}
				
				continue;
			}
			
			if (!"Font".equals(oType) || !"Type1".equals(oSubType))
				continue;
			
			Dict fontDescriptor = (Dict) pdfwriter.getObject(dict.get("FontDescriptor"));
			Dict fontFile       = fontDescriptor == null ? null : (Dict) pdfwriter.getObject(fontDescriptor.get("FontFile"));
			
			if (fontFile == null)
				continue;
			
			cntrEmbeddedType1++;
			
			Object fontFileLength = pdfwriter.getObject(fontFile.get("Length"));
			
			if (fontFileLength instanceof Number)
				embeddedType1Size += ((Number) fontFileLength).intValue();
		}

		if (monitor) {
			if (cntrLZW > 0)
				System.out.print(", " + cntrLZW + " LZW");
			
			if (cntrASCII > 0)
				System.out.print(", " + cntrASCII + " ASCII");
			
			if (cntrDeleted > 0)
				System.out.print(", " + cntrDeleted + " deleted");
			
			if (cntrPagePiece > 0)
				System.out.print(", " + cntrPagePiece + " pagepiece");
			
			if (cntrImage > 0)
				System.out.print(", " + cntrImage + " image /Alt");
			
			if (cntrIRef > 0)
				System.out.print(", " + cntrIRef + " /Length IRef");
			
			if (cntrRawSamples > 0)
				System.out.print(", " + cntrRawSamples + " raw samples = " + rawSamplesSize / 1024 + "K");
			
			if (cntrEmbeddedType1 > 0)
				System.out.print(", " + cntrEmbeddedType1 + " embedded Type 1 = " + embeddedType1Size / 1024 + "K");
		}
		
		Dict dictCatalog = pdfwriter.getCatalog();
		strip(dictCatalog, pdfwriter);
		
		int refCnt    = 0;
		int thumbCntr = 0;
		
		for (int i = pdfreader.getPageCnt(); refCnt < i; refCnt++) {
			IRef   iRef = pdfreader.getPageRef(refCnt + 1);
			Object oIRefObj = pdfwriter.getObject(iRef);
			
			if (COS.OBJECT_NULL == oIRefObj || oIRefObj == null)
				continue;
			
			Dict dict = (Dict) oIRefObj;
			oIRefObj  = dict.get("Thumb");
			
			if (oIRefObj != null) {
				dict.remove("Thumb");
				
				if (COS.CLASS_IREF == oIRefObj.getClass())
					pdfwriter.setObject(((IRef) oIRefObj).id, COS.OBJECT_DELETED);
				
				thumbCntr++;
			}
			
			if (!testable)
				stripLZW(dict, dict.get("Contents"), pdfreader, pdfwriter);
		}

		if (monitor && thumbCntr > 0)
			System.out.print(", " + thumbCntr + " thumb");
		
		axeCore14(pdfwriter);
		recodeFonts(pdfwriter);
		subset(pdfreader, pdfwriter);
		
		if (monitor)
			System.out.print(", liftPageTree");
		
		pdfwriter.liftPageTree();
		
		inline(pdfwriter);
		unique(pdfwriter);
		
		refCnt = pdfwriter.refcntRemove();
		
		if (refCnt > 0 && monitor)
			System.out.print(", ref cnt " + refCnt);
		
		boolean flag = !old && !compact ? pdfwriter.makeObjectStreams(0, pdfwriter.getObjCnt()) : false;
		
		if (monitor)
			System.out.println();
		
		if (verbose) {
			if (pdfreader.getLinearized() > 0)
				System.out.println("lost Linearization (aka Fast Web View)");
			
			if (pdfreader.isRepaired())
				System.out.println("repaired errors");
			
			if (cntrLZW > 0)
				System.out.println("converted LZW to Flate");
			
			if (cntrASCII > 0)
				System.out.println("stripped off verbose ASCII filters");
			
			if (cntrDeleted > 0)
				System.out.println("nulled out deleted objects");
			
			if (thumbCntr > 0)
				System.out.println("removed thumbnails (Acrobat can generate on the fly)");
			
			if (cntrPagePiece > 0 && !pagepiece)
				System.out.println("removed /PieceInfo");
			
			if (cntrImage > 0 && !alt)
				System.out.println("remove alternate images");
			
			if (!old) {
				System.out.println("cleaned and modernized");
			
				if (pdfwriter.getVersion().compareTo(version) > 0 && pdfwriter.getVersion().compareTo(1L, 5L) >= 0) {
					System.out.println("\tcross references as stream");
				
					if (flag)
						System.out.println("\tadditional Compression via object streams");
					
					System.out.println("\tnow REQUIRES Multivalent or Acrobat 6 to read (use -old for older PDF)");
				}
			}
		}
		
		if (!quiet) {
			if (compact)
				System.out.println("Compact PDF format -- requires Multivalent to read");
		
			StringBuffer sb = new StringBuffer();
			
			if (old && pdfwriter.getVersion().compareTo(1L, 5L) < 0)
				sb.append(" [omit -old]");
			
			if (!compact0)
				sb.append(" -compact");
			
			if (dictCatalog.get("StructTreeRoot") != null)
				sb.append(" -nostruct");
			
			if (cntrRawSamples > 0 && !jpeg)
				sb.append(" -jpeg");
			
			if (dictCatalog.get("SpiderInfo") != null)
				sb.append(" -nowebcap");
			
			if (cntrPagePiece > 0 && pagepiece)
				sb.append(" -nopagepiece");
			
			if (cntrImage > 0 && alt)
				sb.append(" -noalt");
			
			if (sb.length() > 0) {
				System.out.println("additional Compression may be possible with:");
				System.out.println("\t" + sb);
			}
		}
	}

	public List<String> getWacky() {
		return wacky;
	}

	@SuppressWarnings("unchecked")
	private int inline(PDFWriter pdfwriter) {
		int    i 		= 0;
		int    j 		= pdfwriter.getObjCnt();
		Object aObjects[] 	= pdfwriter.getObjects();
		int    ai[] 	= pdfwriter.refcnt();
		
		if (monitor)
			System.out.print(", inline");
		
		label0: for (int k = 1; k < j; k++) {
			Class<? extends Object> class1 = aObjects[k].getClass();
			
			if (ai[k] == 0)
				continue;
			
			if (COS.CLASS_ARRAY == class1) {
				Object aobj1[] = (Object[]) aObjects[k];
				int k1 = 0;
				int l1 = aobj1.length;
				
				do {
					if (k1 >= l1)
						continue label0;
					
					Object obj;
				
					if ((obj = inlineObj(aobj1[k1], ai, pdfwriter)) != null)
						aobj1[k1] = obj;
					
					k1++;
				} while (true);
			}
			
			if (COS.CLASS_DICTIONARY != class1)
				continue;
			
			Dict dict1 = (Dict) aObjects[k];
			Iterator<?> iterator = dict1.entrySet().iterator();
			
			do {
				Object obj1;
				Entry<?, Object> entry;
				
				do {
					if (!iterator.hasNext())
						continue label0;
					entry = (Entry<?, Object>) iterator.next();
				} while ((obj1 = inlineObj(entry.getValue(), ai, pdfwriter)) == null);
				
				entry.setValue(obj1);
			} while (true);
		}

		for (int l = 1; l < j; l++) {
			if (COS.CLASS_DICTIONARY != aObjects[l].getClass() || ai[l] == 0)
				continue;
			
			Dict dict = (Dict) aObjects[l];
			Object obj2;
			
			if (!"Page".equals(dict.get("Type")) || (obj2 = dict.get("Contents")) == null)
				continue;
			
			if (COS.CLASS_IREF == obj2.getClass()) {
				int j1 = ((IRef) obj2).id;
				obj2 = aObjects[j1];
			}
			
			if (COS.CLASS_ARRAY != obj2.getClass())
				continue;
			
			Object aobj2[] = (Object[]) obj2;
			Object aobj3[] = new Object[aobj2.length];
			
			int i2 = 0;
			byte abyte0[] = null;
			int j2 = -1;
			int k2 = 0;
			int l2 = 0;
			
			for (int i3 = aobj2.length; l2 < i3; l2++) {
				IRef iref = (IRef) aobj2[l2];
				int j3 = iref.id;
				aobj3[i2++] = iref;
				byte abyte1[] = (byte[]) ((Dict) aObjects[j3]).get("DATA");
			
				if (ai[j3] == 1 || j2 != -1 && abyte1.length < PDFWriter.PDFOBJREF_OVERHEAD * 2) {
					if (abyte0 == null) {
						abyte0 = abyte1;
						j2 = j3;
						k2 = 1;
					} else {
						int k3 = abyte0.length;
						abyte0 = Arrayss.resize(abyte0, k3 + abyte1.length + 1);
						abyte0[k3] = 32;
						
						System.arraycopy(abyte1, 0, abyte0, k3 + 1, abyte1.length);
						
						i2--;
						k2++;
					}
				
					continue;
				}
				
				if (k2 > 1)
					((Dict) aObjects[j2]).put("DATA", abyte0);
				
				abyte0 = null;
				j2 = -1;
				k2 = 0;
			}

			if (k2 > 1)
				((Dict) aObjects[j2]).put("DATA", abyte0);
			
			if (i2 < aobj2.length) {
				i += aobj2.length - i2;
				
				if (i2 == 0) {
					dict.remove("Contents");
					continue;
				}
				
				if (i2 == 1)
					dict.put("Contents", aobj3[0]);
				else
					dict.put("Contents", ((Object) (Arrayss.resize(aobj3, i2))));
			} else {
				dict.put("Contents", ((Object) (aobj2)));
			}
		}

		int i1 = pdfwriter.refcntRemove();

		if (monitor) {
			if (monitor)
				System.out.print(" " + -i1);
			
			if (i > 0)
				System.out.print(", " + i + " concat");
		}
		
		return i1;
	}

	private Object inlineObj(Object obj, int ai[], PDFWriter pdfwriter) {
		if (!(obj instanceof IRef))
			return null;
		
		int i = ((IRef) obj).id;
		obj = pdfwriter.getCache(i);
		
		Class<? extends Object> clsObj = obj.getClass();
		
		if (COS.OBJECT_NULL == obj || COS.CLASS_INTEGER == clsObj || COS.CLASS_BOOLEAN == clsObj || COS.CLASS_NAME == clsObj && ((String) obj).length() < PDFWriter.PDFOBJREF_OVERHEAD || COS.CLASS_STRING == clsObj && ((StringBuffer) obj).length() + 2 < PDFWriter.PDFOBJREF_OVERHEAD || ai[i] == 1 && (COS.CLASS_NAME == clsObj || COS.CLASS_STRING == clsObj || COS.CLASS_REAL == clsObj)) {
			ai[i]--;
			
			return obj;
		} else {
			return null;
		}
	}

	private boolean recodeDCTAsRaw(Dict dict, int i, int j, byte abyte0[]) throws IOException {
		if (testable)
			return false;
		
		int k = abyte0.length;
		
		if (i * j > k)
			return false;
		try {
			ImageReader imagereader = (ImageReader) ImageIO.getImageReadersByFormatName("JPEG").next();
			ImageIO.setUseCache(false);
			
			ImageInputStream imageinputstream = ImageIO.createImageInputStream(new ByteArrayInputStream(abyte0));
			imagereader.setInput(imageinputstream, true);
			
			BufferedImage bufferedimage = imagereader.read(0);
			imagereader.dispose();
			
			imageinputstream.close();
		
			int l = bufferedimage.getColorModel().getNumComponents();
			int i1 = l * i * j;
			
			if (i1 / 2 < k && (l == 1 || l == 3)) {
				byte abyte1[] = new byte[i1];
				int j1 = 0;
				int k1 = 0;
				
				for (; j1 < j; j1++) {
					for (int l1 = 0; l1 < i; l1++) {
						int i2 = bufferedimage.getRGB(l1, j1);
						if (l == 3) {
							abyte1[k1++] = (byte) (i2 >> 16);
							abyte1[k1++] = (byte) (i2 >> 8);
						}
						
						abyte1[k1++] = (byte) i2;
					}

				}

				dict.put("DATA", abyte1);
				dict.remove("Filter");
			}
		} catch (Exception exception) {}
		
		return true;
	}

	private void recodeFonts(PDFWriter pdfwriter) throws IOException {
	}

	private void recodeImage(Dict dict, PDFWriter pdfwriter, PDFReader pdfreader) throws IOException {
		byte   buffer[]        = (byte[]) dict.get("DATA");
		String filter          = (String) pdfwriter.getObject(dict.get("Filter"));
		int    width           = ((Number) pdfwriter.getObject(dict.get("Width"))).intValue();
		int    height   	   = ((Number) pdfwriter.getObject(dict.get("Height"))).intValue();
		int    bitPerComponent = ((pdfwriter.getObject(dict.get("ImageMask")) != Boolean.TRUE) ? ((Number) pdfwriter.getObject(dict.get("BitsPerComponent"))).intValue() : 1);
		
		Object obj;
		
		if (!alt && (obj = pdfreader.getObject(dict.get("Alternates"))) != null) {
			Object            aobj[]    = (Object[]) obj;
			ArrayList<Object> arraylist = new ArrayList<Object>(aobj.length);
			Object            aobj1[]   = aobj;
			int               l 		= aobj1.length;
			
			for (int i1 = 0; i1 < l; i1++) {
				Object obj1 = aobj1[i1];
				Dict dict1 = (Dict) pdfreader.getObject(obj1);
				
				if (dict1.get("OC") != null)
					arraylist.add(obj1);
			}

			if (arraylist.size() == 0)
				dict.remove("Alternates");
			else
				dict.put("Alternates", ((Object) (arraylist.toArray())));
		}
		
		if ("DCTDecode".equals(filter))
			recodeDCTAsRaw(dict, width, height, buffer);
		else if (!"CCITTFaxDecode".equals(filter) && filter == null) {
			boolean flag = false;
			
			if (jpeg)
				flag = recodeRawAsDCT(dict, width, height, bitPerComponent, buffer, pdfreader);
			if (!flag && pre)
				flag = addPredictor(dict, width, height, bitPerComponent, buffer, pdfreader);
		}
	}

	private boolean recodeRawAsDCT(Dict dict, int width, int height, int bitPerComponent, byte buffer[], PDFReader pdfreader) throws IOException {
        BufferedImage 			bufferedImage 			= null; 
        ByteArrayOutputStream 	byteArrayOutputStream 	= null;
        int 					deflateDataLength 		= 0;
        IIOImage 				iioimage 				= null;
        String 					imageFormat 			= null;
        ImageOutputStream 		imageoutputstream 		= null;
        ImageWriter 			imagewriter 			= null;
        
		if (testable)
			return false;

		if (bitPerComponent != 8 || width < 20 || height < 20 || buffer.length < 5420)
			return false;

		byte deflateData[] = PDFWriter.maybeDeflateData(buffer);
        
        deflateDataLength = deflateData.length;
        deflateData       = null;
        
        if (deflateDataLength < 5420)
            return false;
        
        InputStreamComposite inputStreamComposite = pdfreader.getInputStream(dict);
        
        bufferedImage = Images.createImage(dict, inputStreamComposite, new AffineTransform(), Color.WHITE, pdfreader);
        inputStreamComposite.close();
        
        if (bufferedImage == null)
            return false;
        
        byteArrayOutputStream = new ByteArrayOutputStream(width * height);
        
        imageFormat = (jpeg2000 ? "JPEG2000" : "JPEG");
        
		imagewriter       = (ImageWriter) ImageIO.getImageWritersByFormatName(imageFormat).next();
		ImageIO.setUseCache(false);
		
		imageoutputstream = ImageIO.createImageOutputStream(byteArrayOutputStream);
		imagewriter.setOutput(imageoutputstream);
        
		iioimage          = new IIOImage(bufferedImage, null, null);
        imagewriter.write((IIOMetadata)null, iioimage, null);

        imagewriter.dispose();
        imageoutputstream.close();

        byte baosBuffer[] = byteArrayOutputStream.toByteArray();
        int  imageLength  = (deflateDataLength - baosBuffer.length);
        
		if (imageLength > 5120) {
			if (monitor) {
				System.out.print(" jpeg" + imageLength);
				
				if (imageLength > 0x19000)
					System.out.print("/" + width + "x" + height);
			}
			
			dict.put("Filter", "JPEG2000".equals(imageFormat) ? "JPXDecode" : "DCTDecode");
			dict.put("DATA", baosBuffer);
			
			String imageType = (((10 != bufferedImage.getType()) && (11 != bufferedImage.getType())) ? "DeviceRGB" : "DeviceGray");
			
			dict.put("ColorSpace", imageType);
		}
		
		return true;
    }

	public void setCompact(boolean flag) {
		compact0 = flag;
	}

	public void setJPEG(boolean flag) {
		this.jpeg = flag;
	}

	public void setMax() {
		old = testable = false;
		compact0 = jpeg = subset = true;
		struct = webcap = pagepiece = alt = core14 = embed = false;
	}

	public void setMonitor(boolean bValue) {
		this.monitor = bValue;
	}

	public void setOld(boolean flag) {
		old = flag;
	}

	public boolean setPassword(String password) throws IOException {
		return pdfReader.setPassword(password);
	}

	public void setQuiet(boolean flag) {
		quiet = flag;
	}

	public void setStruct(boolean flag) {
		struct = flag;
	}

	public void setSubset(boolean flag) {
		subset = flag;
	}

	public void setTestable(boolean flag) {
		testable = flag;
	}

	public void setVerbose(boolean flag) {
		verbose = flag;
	}

	private void strip(Dict dict, PDFWriter pdfwriter) {
		if (!struct) {
			dict.remove("StructTreeRoot");
			dict.remove("MarkInfo");
		}
		if (!webcap)
			dict.remove("SpiderInfo");
		if (!pagepiece)
			dict.remove("PieceInfo");
		if (!outline)
			dict.remove("Outlines");
		if (!pagepiece) {
			int i = 1;
			for (int j = pdfwriter.getObjCnt(); i < j; i++) {
				Object obj = pdfwriter.getCache(i);
				if (COS.CLASS_DICTIONARY == obj.getClass()) {
					Dict dict1 = (Dict) obj;
					dict1.remove("PieceInfo");
				}
			}

		}
	}

	private void stripLZW(Dict dict, Object obj, PDFReader pdfreader, PDFWriter pdfwriter) throws IOException {
		if (testable || !content)
			return;
		
		Object obj1   = pdfwriter.getObject(obj);
		Object aobj[] = ((obj1 != null) ? COS.CLASS_DICTIONARY != obj1.getClass() ? (Object[]) obj1 : (new Object[] { obj }) : new Object[0]);

		int i = 0;
		
		for (int j = aobj.length; i < j; i++) {
			Dict dict3 = (Dict) pdfwriter.getObject(aobj[i]);
		
			try {
				Cmd acmd[] = pdfreader.readCommandArray(aobj[i]);
				dict3.put("DATA", pdfwriter.writeCommandArray(acmd, false));
			} catch (IOException ioexception) {}
		}

		Dict dict1 = (Dict) pdfreader.getObject(dict.get("Resources"));
		Dict dict2 = ((dict1 == null) ? null : (Dict) pdfreader.getObject(dict1.get("XObject")));
		
		if (dict2 != null) {
			Iterator<Object> iterator = dict2.values().iterator();
			
			do {
				if (!iterator.hasNext())
					break;
			
				IRef   iref = (IRef) iterator.next();
				Object obj2 = pdfreader.getObject(iref);
				
				if (COS.CLASS_DICTIONARY == obj2.getClass()) {
					Dict dict4 = (Dict) obj2;
				
					if ("Form".equals("Subtype"))
						stripLZW(dict4, iref, pdfreader, pdfwriter);
				}
			} while (true);
		}
	}

	private void subset(PDFReader pdfreader, PDFWriter pdfwriter) throws IOException {
		if (!subset || testable)
			return;
		
		if (monitor)
			System.out.print(", subset");
		
		int  i       = pdfwriter.getObjCnt();
		Dict adict[] = new Dict[i];
		int  ai[]    = new int[i];
		
		for (int j = 1; j < i; j++) {
			Object obj = pdfwriter.getObject(j, false);
			
			if (COS.CLASS_DICTIONARY != obj.getClass())
				continue;
			
			Dict   dict = (Dict) obj;
			Object obj1 = pdfwriter.getObject(dict.get("Type"));
			Object obj2 = pdfwriter.getObject(dict.get("Subtype"));
			String s    = (String) pdfwriter.getObject(dict.get("BaseFont"));
			
			if (!"Font".equals(obj1) || Fonts.isSubset(s))
				continue;
			
			Dict dict4 = (Dict) pdfwriter.getObject(dict.get("FontDescriptor"));
			
			if (dict4 == null)
				continue;
			
			Object obj4;
			
			if ((obj4 = dict4.get("FontFile")) != null && ("Type1".equals(obj2) || "MMType1".equals(obj2))) {
				adict[j] = dict;
				ai[j] = ((IRef) obj4).id;
			
				continue;
			}
			
			if ((obj4 = dict4.get("FontFile2")) != null) {
				ai[j] = ((IRef) obj4).id;
			
				continue;
			}
			
			if ((obj4 = dict4.get("FontFile3")) != null)
				ai[j] = ((IRef) obj4).id;
		}

		int ai1[][] = new int[i][];
		int l       = 0;
		
		for (int i1 = adict.length; l < i1; l++)
			if (adict[l] != null) {
				ai1[ai[l]] = new int[256];
			}

		l = 0;
		
		for (int j1 = pdfreader.getPageCnt(); l < j1; l++) {
			Dict dict1 = pdfreader.getPage(l + 1);
			Dict dict3 = (Dict) pdfwriter.getObject(dict1.get("Resources"));
		
			if (dict3 == null)
				continue;
			
			subsetCensus(dict3.get("Font"), dict1.get("Contents"), pdfreader, pdfwriter, ai1, ai);
			
			Object obj3 = dict3.get("XObject");
			
			if (obj3 == null)
				continue;
			
			Dict dict5 = (Dict) pdfwriter.getObject(obj3);
			Dict dict6 = (Dict) pdfwriter.getObject(dict5.get("Resources"));
			
			if (dict6 != null)
				subsetCensus(dict6.get("Font"), obj3, pdfreader, pdfwriter, ai1, ai);
		}

		l = 1;
		
		for (int k1 = adict.length; l < k1; l++) {
			Dict dict2 = adict[l];
		
			int ai2[] = ai1[ai[l]];
			
			if (ai2 == null)
				continue;
			
			ai1[ai[l]] = null;
			
			boolean aflag[] = new boolean[256];
			int     l1      = 1;
			
			for (int i2 = 1; i2 < 256; i2++)
				if (ai2[i2] > 0) {
					aflag[i2] = true;
					l1++;
				}

			Dict dict7 = (Dict) pdfwriter.getObject(dict2.get("FontDescriptor"));
			IRef iref = (IRef) dict7.get("FontFile");
			Dict dict8 = (Dict) pdfwriter.getObject(iref);
			
			try {
				NFontType1 nfonttype1 = new NFontType1(null, pdfwriter.getStreamData(dict8));
				System.out.print(", " + dict2.get("BaseFont") + " " + nfonttype1.getNumGlyphs() + " subset " + l1);
			
				if (monitor) {
					System.out.print(": ");
				
					for (int j2 = 32; j2 < 128; j2++)
						if (aflag[j2])
							System.out.print((char) j2);

					System.out.println();
				}
				
				String s1 = nfonttype1.getName();
				
				if (s1 != null && !Fonts.isSubset(s1)) {
					StringBuffer stringbuffer = new StringBuffer(7 + s1.length());
					Random       random       = new Random();
				
					for (int l2 = 0; l2 < 6; l2++)
						stringbuffer.append((char) (65 + random.nextInt(26)));

					stringbuffer.append('+');
					stringbuffer.append(s1);
					
					s1 = stringbuffer.toString();
				}
				
				nfonttype1 = nfonttype1.deriveFont(s1, aflag);
				
				byte abyte0[] = nfonttype1.toPFB();
				abyte0 = Arrayss.resize(abyte0, abyte0.length - NFontType1.PFB_00_LENGTH);
				dict8.put("DATA", abyte0);
				
				int k2 = NFontType1.getClen(abyte0);
				dict8.put("Length1", Integers.getInteger(k2));
				dict8.put("Length2", Integers.getInteger(abyte0.length - k2));
				dict8.put("Length3", Integers.ZERO);
				dict7.put("FontName", s1);
				dict2.put("BaseFont", s1);
			} catch (FontFormatException fontformatexception) {} catch (IOException ioexception) {}
		}
	}

	private void subsetCensus(Object obj, Object obj1, PDFReader pdfreader, PDFWriter pdfwriter, int ai[][], int ai1[])
			throws IOException {
		if (obj == null || obj1 == null)
			return;
		Dict dict = (Dict) pdfwriter.getObject(obj);
		int ai2[] = null;
		int ai3[][] = new int[100][];
		int i = 0;
		Iterator<?> iterator = dict.values().iterator();
		do {
			if (!iterator.hasNext())
				break;
			IRef iref = (IRef) iterator.next();
			int j = ai1[iref.id];
			if (j == 0)
				continue;
			multivalent.std.adaptor.pdf.InputStreamComposite inputstreamcomposite = pdfreader.getInputStream(obj1,
					true);
			label0: do {
				Cmd cmd;
				String s;
				do {
					do {
						if ((cmd = pdfreader.readCommand(inputstreamcomposite)) == null)
							break label0;
						s = cmd.op;
						if ("Tf" == s) {
							IRef iref1 = (IRef) dict.get(cmd.ops[0]);
							int k = ai1[iref1.id];
							if (ai[k] != null)
								ai2 = ai[k];
						} else if ("q" == s) {
							ai3[i++] = ai2;
						} else {
							if ("Q" != s)
								continue;
							if (--i >= 0)
								ai2 = ai3[i];
						}
						continue label0;
					} while (ai2 == null);
					if ("Tj" != s)
						continue;
					StringBuffer stringbuffer = (StringBuffer) cmd.ops[0];
					int l = 0;
					int j1 = stringbuffer.length();
					while (l < j1) {
						ai2[stringbuffer.charAt(l)]++;
						l++;
					}
					continue label0;
				} while ("TJ" != s);
				Object aobj[] = (Object[]) cmd.ops[0];
				int i1 = 0;
				int k1 = aobj.length;
				while (i1 < k1) {
					if (COS.CLASS_STRING == aobj[i1].getClass()) {
						StringBuffer stringbuffer1 = (StringBuffer) aobj[i1];
						int l1 = 0;
						for (int i2 = stringbuffer1.length(); l1 < i2; l1++)
							ai2[stringbuffer1.charAt(l1)]++;

					}
					i1++;
				}
			} while (true);
			break;
		} while (true);
	}

	private int unique(PDFWriter pdfwriter) {
		System.currentTimeMillis();
		
		int i = pdfwriter.getObjCnt();
		Object aobj[] = pdfwriter.getObjects();
		Class<?> aclass[] = new Class[i];
		
		for (int k = 1; k < i; k++)
			aclass[k] = aobj[k].getClass();

		long al[] = new long[i];
		int ai[] = new int[i];
		int i1 = 0;
		int j1 = 0;
		
		do {
			int k1 = 0;
			Object aobj1[] = pdfwriter.getObjects();
			int j = pdfwriter.getObjCnt();
			assert al[0] == 0L;
			for (int i2 = 1; i2 < j; i2++)
				al[i2] = ((long) sign(aobj1[i2]) << 32) + (long) i2;

			Arrays.sort(al, 1, j);
			for (int j2 = 1; j2 < j; j2++)
				ai[(int) (al[j2] & 0x7fffffffL)] = j2;

			int ai1[] = new int[j];
			label0: for (int k2 = 1; k2 < j; k2++) {
				ai1[k2] = k2;
				Object obj = aobj1[k2];
				Class<?> class1 = aclass[k2];
				if (COS.CLASS_DICTIONARY == class1 && "Page".equals(((Dict) obj).get("Type")))
					continue;
				int i3 = ai[k2];
				long l3 = al[i3] & 0xffffffff00000000L;
				i3--;
				do {
					if (i3 <= 0 || (al[i3] & l3) != l3)
						continue label0;
					int j3 = (int) (al[i3] & 0x7fffffffL);
					if (class1 == aclass[j3] && pdfwriter.objEquals(obj, aobj1[j3])) {
						int k3 = ai1[j3];
						ai1[k2] = k3;
						aobj1[k2] = aobj1[k3];
						k1++;
						continue label0;
					}
					i3--;
				} while (true);
			}

			if (k1 > 0) {
				int ai2[] = pdfwriter.renumberRemove(ai1);
				for (int l2 = 1; l2 < j; l2++)
					aclass[l2 - ai2[l2]] = aclass[l2];

				if (monitor)
					System.out.print(j1 != 0 ? " + " + k1 : ", " + k1 + " dups");
				i1 += k1;
				j1++;
			} else {
				System.currentTimeMillis();
				return i1;
			}
		} while (true);
	}

//	public byte[] writeBytes() throws IOException, ParseException {
//		OutputUni outputuni = OutputUni.getInstance(new byte[pdfReader.getObjCnt() * 100], null);
//		writeUni(outputuni);
//		
//		return outputuni.toByteArray();
//	}

	public byte[] writeBytes() throws IOException, ParseException {
        File      tempFile  = File.createTempFile("pdfw", ".pdf", null);
        OutputUni outputuni = OutputUni.getInstance(tempFile, tempFile.toURI());

		writeUni(outputuni);

		byte[] byteArray    = outputuni.toByteArray();
		
		tempFile.delete();
		
		return byteArray;
	}

	public void writeFile(File file) throws IOException, ParseException {
		writeUni(OutputUni.getInstance(file, null));
	}

//	public PDFReader writePipe() throws IOException, ParseException {
//		return null;
//	}

	public void writeUni(OutputUni outputuni) throws IOException, ParseException {
		Compress(new PDFWriter(outputuni, pdfReader));
	}
	
}
