package org.riverframework.utils;

import com.snowtide.PDF;
import com.snowtide.pdf.Document;
import com.snowtide.pdf.OutputTarget;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.swing.text.DefaultStyledDocument;
import javax.swing.text.rtf.RTFEditorKit;
import java.io.IOException;
import java.io.InputStream;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * Extracts text only from Word, Excel, RTF and PDF files. It's based on the code
 * published on <a href="https://codezrule.wordpress.com/2010/03/24/extract-text-from-pdf-office-files-doc-ppt-xls-open-office-files-rtf-and-textplain-files-in-java/">Extract Text From pdf, office files(.doc, .ppt, .xls), open office files, .rtf, and text/plain files in Java</a>
 *
 * @author mario.sotil@gmail.com
 */
@SuppressWarnings("unused")
public class TextExtractor {
    private static Logger log = Logger.getLogger(TextExtractor.class.getName());

    public enum Type {PDF, XLSX, XLS, DOC, DOCX, RTF}

    public String parse(String reference, InputStream is) {
        return parse(referenceToType(reference), is);
    }

    private Type referenceToType(String reference) {
        Type type = null;
        String referenceInLowerCase = reference.toLowerCase();

        if (referenceInLowerCase.endsWith(".pdf") || referenceInLowerCase.equals("application/pdf")) {
            type = Type.PDF;
        } else if (referenceInLowerCase.endsWith(".doc") || referenceInLowerCase.equals("application/msword")) {
            type = Type.DOC;
        } else if (referenceInLowerCase.endsWith(".docx") || referenceInLowerCase.equals("application/vnd.openxmlformats-officedocument.wordprocessingml.document")) {
            type = Type.DOCX;
        } else if (referenceInLowerCase.endsWith(".rtf") || referenceInLowerCase.equals("application/rtf")) {
            type = Type.RTF;
        } else if (referenceInLowerCase.endsWith(".xls") || referenceInLowerCase.equals("application/vnd.ms-excel")) {
            type = Type.XLS;
        } else if (referenceInLowerCase.endsWith(".xlsx") || referenceInLowerCase.equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
            type = Type.XLSX;
        }

        if (type == null)
            throw new RuntimeException("There's no a valid reference to detect the type of the input stream.");

        return type;
    }

    public String parse(Type type, InputStream is) {
        String result = "";

        try {
            if (type.equals(Type.PDF)) {
                result = pdf2text(is);
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

    public String pdf2text(InputStream is) {
        StringBuilder parsedText = new StringBuilder(1024);

        try {
            Document pdf = PDF.open(is, "");
            pdf.pipe(new OutputTarget(parsedText));
            pdf.close();
        } catch (Exception e) {
            log.log(Level.WARNING, "An exception has occurred while parsing the PDF Document.", e);
        }

        log.info("Done");
        return parsedText.toString();
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
}
