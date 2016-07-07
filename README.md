# text-extractor
Extracts text from Office and PDFs files, using POI and PDFxStream, as a very, very tiny alternative to [Apache Tika](https://tika.apache.org/)

This library, obviously, NO replaces Apache Tika. Only extracts text from Word, Excel, RTF and PDF files. It's based on the code found on the blog article [Extract Text From pdf, office files(.doc, .ppt, .xls), open office files, .rtf, and text/plain files in Java](https://codezrule.wordpress.com/2010/03/24/extract-text-from-pdf-office-files-doc-ppt-xls-open-office-files-rtf-and-textplain-files-in-java/)  but using the last Apache POI and PDFxStream versions (06/10/2015).

- org.apache.poi, 3.12
- com.snowtide.pdfxstream, 3.1.2

