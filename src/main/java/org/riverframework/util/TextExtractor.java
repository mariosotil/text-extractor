package org.riverframework.util;

import java.io.IOException;
import java.io.InputStream;

import javax.swing.text.DefaultStyledDocument;
import javax.swing.text.rtf.RTFEditorKit;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.poifs.eventfilesystem.POIFSReaderEvent;
import org.apache.poi.poifs.eventfilesystem.POIFSReaderListener;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.util.LittleEndian;
import org.apache.pdfbox.cos.COSDocument;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.util.PDFTextStripper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Extracts text only from Word, Excel, RTF and PDF files. It's based on the code
 * published on <a href="https://codezrule.wordpress.com/2010/03/24/extract-text-from-pdf-office-files-doc-ppt-xls-open-office-files-rtf-and-textplain-files-in-java/">Extract Text From pdf, office files(.doc, .ppt, .xls), open office files, .rtf, and text/plain files in Java</a>
 *
 * @author mario.sotil@gmail.com
 *
 */
public class TextExtractor {
	private static Logger log = LoggerFactory.getLogger(TextExtractor.class);

	private StringBuffer sb = new StringBuffer(8192);
	
	public String pdftotext(InputStream is) {
		PDFParser parser;
		String parsedText;
		PDFTextStripper pdfStripper;
		PDDocument pdDoc = null;
		COSDocument cosDoc = null;
		try {
			parser = new PDFParser(is);
		} catch (Exception e) {
			log.error("Unable to open PDF Parser.", e);
			return null;
		}
		try {
			parser.parse();
			cosDoc = parser.getDocument();
			pdfStripper = new PDFTextStripper();
			pdDoc = new PDDocument(cosDoc);
			parsedText = pdfStripper.getText(pdDoc);
			cosDoc.close();
			pdDoc.close();
		} catch (Exception e) {
			log.error("Exception occured in parsing the PDF Document.", e);
			try {
				if (cosDoc != null) {
					cosDoc.close();
				}
				if (pdDoc != null) {
					pdDoc.close();
				}
			} catch (Exception e1) {
				log.error("", e1);
			}
			return null;
		}
		log.info("Done");
		return parsedText;
	}

	class MyPOIFSReaderListener implements POIFSReaderListener {

		public void processPOIFSReaderEvent(POIFSReaderEvent event) {
			char ch0 = (char) 0;
			char ch11 = (char) 11;
			try {
				DocumentInputStream dis = null;
				dis = event.getStream();
				byte btoWrite[] = new byte[dis.available()];
				dis.read(btoWrite, 0, dis.available());
				for (int i = 0; i < btoWrite.length - 20; i++) {
					long type = LittleEndian.getUShort(btoWrite, i + 2);
					long size = LittleEndian.getUInt(btoWrite, i + 4);
					if (type == 4008) {
						try {
							String s = new String(btoWrite, i + 4 + 1, (int) size + 3).replace(ch0, ' ').replace(ch11, ' ');
							if (s.trim().startsWith("Click to edit") == false) {
								sb.append(s);
							}
						} catch (Exception e) {
							log.error("", e);
						}
					}
				}
			} catch (Exception ex) {
				ex.printStackTrace();
				return;
			}
		}
	}

	public String xls2text(InputStream is) throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook(is);
		XSSFExcelExtractor we = new XSSFExcelExtractor(wb); 
		String text = we.getText();
		we.close();
		return text;
	}

	public String doc2text(InputStream is) throws IOException {
		XWPFDocument doc = new XWPFDocument(is);
		XWPFWordExtractor we = new XWPFWordExtractor(doc); 
		String text = we.getText();
		we.close();
		return text;
	}

	public String rtf2text(InputStream is) throws Exception {
		DefaultStyledDocument styledDoc = new DefaultStyledDocument();
		new RTFEditorKit().read(is, styledDoc, 0);
		return styledDoc.getText(0, styledDoc.getLength());
	}

	public String parse(String fileName, InputStream is) {
		String result = "";
		String f = fileName.toLowerCase();
		
		try {
			if (f.endsWith(".pdf")) {
				result = pdftotext(is);
			} else if (f.endsWith(".doc") || f.endsWith(".docx")) {
				result = doc2text(is);
			} else if (f.endsWith(".rtf")) {
				result = rtf2text(is);
			} else if (f.endsWith(".xls") || f.endsWith(".xlsx")) {
				result = xls2text(is);
			} 
		} catch (Exception e) {
			throw new RuntimeException(e);
		}

		return result;
	}
}
