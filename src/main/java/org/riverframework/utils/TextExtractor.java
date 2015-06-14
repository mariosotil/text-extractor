package org.riverframework.utils;

import java.io.IOException;
import java.io.InputStream;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.swing.text.DefaultStyledDocument;
import javax.swing.text.rtf.RTFEditorKit;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.poifs.eventfilesystem.POIFSReaderEvent;
import org.apache.poi.poifs.eventfilesystem.POIFSReaderListener;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.util.LittleEndian;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.snowtide.PDF;
import com.snowtide.pdf.Document;
import com.snowtide.pdf.OutputTarget;

/**
 * Extracts text only from Word, Excel, RTF and PDF files. It's based on the code
 * published on <a href="https://codezrule.wordpress.com/2010/03/24/extract-text-from-pdf-office-files-doc-ppt-xls-open-office-files-rtf-and-textplain-files-in-java/">Extract Text From pdf, office files(.doc, .ppt, .xls), open office files, .rtf, and text/plain files in Java</a>
 *
 * @author mario.sotil@gmail.com
 *
 */
public class TextExtractor {
	private static Logger log = Logger.getLogger(TextExtractor.class.getName());

	private StringBuffer sb = new StringBuffer(8192);

	public enum Type { PDF, XLSX, XLS, DOC, DOCX, RTF }

	public String pdftotext(InputStream is) {
		StringBuilder parsedText = new StringBuilder(1024);
		
		try {
			Document pdf = PDF.open(is, "");
			pdf.pipe(new OutputTarget(parsedText));
			pdf.close();
		} catch (Exception e) {
			log.log(Level.WARNING, "Exception occured in parsing the PDF Document.", e);
		}

		log.info("Done");
		return parsedText.toString();
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
							log.log(Level.WARNING, "Exception occured.", e);
						}
					}
				}
			} catch (Exception ex) {
				ex.printStackTrace();
				return;
			}
		}
	}

    public String xls2text(InputStream is) throws IOException {
    	ExcelExtractor wd = new ExcelExtractor(new HSSFWorkbook(is));
        String text = wd.getText();
        wd.close();
        return text;
    }
    
	public String xlsx2text(InputStream is) throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook(is);
		XSSFExcelExtractor we = new XSSFExcelExtractor(wb); 
		String text = we.getText();
		we.close();
		return text;
	}

    public String doc2text(InputStream is) throws IOException {
        WordExtractor wd = new WordExtractor(is);
        String text = wd.getText();
        wd.close();
        return text;
    }
    
	public String docx2text(InputStream is) throws IOException {
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

	public String parse(Type type, InputStream is) {
		String result = "";

		try {
			if (type.equals(Type.PDF)) {
				result = pdftotext(is);
			} else if (type.equals(Type.DOC)) {
				result = doc2text(is);
			} else if (type.equals(Type.DOCX)) {
				result = docx2text(is);
			} else if (type.equals(Type.RTF)) {
				result = rtf2text(is);
			} else if (type.equals(Type.XLS)) {
				result = xls2text(is);
			} else if (type.equals(Type.XLSX)) {
				result = xlsx2text(is);
			} 
		} catch (Exception e) {
			throw new RuntimeException(e);
		}

		return result;
	}

	public String parse(String fileName, InputStream is) {
		String result = "";
		String f = fileName.toLowerCase();

		try {
			if (f.endsWith(".pdf") || f.equals("application/pdf")) {
				result = parse(Type.PDF, is);
			} else if (f.endsWith(".doc") || f.equals("application/msword")) {
				result = parse(Type.DOC, is);
			} else if (f.endsWith(".docx") || f.equals("application/vnd.openxmlformats-officedocument.wordprocessingml.document")) {
				result = parse(Type.DOCX, is);
			} else if (f.endsWith(".rtf") || f.equals("application/rtf")) {
				result = parse(Type.RTF, is);
			} else if (f.endsWith(".xls") || f.equals("application/vnd.ms-excel")) {
				result = parse(Type.XLS, is);
			} else if (f.endsWith(".xlsx") || f.equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
				result = parse(Type.XLSX, is);
			} 
		} catch (Exception e) {
			throw new RuntimeException(e);
		}

		return result;
	}

}
