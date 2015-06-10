# text-extractor
Extracts text from Office and PDFs files, using POI and PDFBox, as a very, very tiny alternative to [Apache Tika](https://tika.apache.org/)

This library, obviously, NO replaces Apache Tika. Only extracts text from Word, Excel, RTF and PDF files. It's based on the code found on the blog article [Extract Text From pdf, office files(.doc, .ppt, .xls), open office files, .rtf, and text/plain files in Java](https://codezrule.wordpress.com/2010/03/24/extract-text-from-pdf-office-files-doc-ppt-xls-open-office-files-rtf-and-textplain-files-in-java/)  but using the last Apache POI and PDFxStream versions (06/10/2015).

- org.apache.poi, 3.12
- com.snowtide.pdfxstream, 3.1.2

## Why I created this library?

I wrote a Crawler that extracts text from Office and PDF files. This crawler was written in Java, but it had to work on an IBM Notes database. For some reason that I still don't know, Tika stops to work. Maybe some library that I changed, maybe some configuration that I moved considering that the crawler uses [jcifs](https://jcifs.samba.org/), [Apache HTTP Components](http://hc.apache.org/), [jsoup](http://jsoup.org/) and my own [framework to connect to IBM Notes](https://github.com/mariosotil/river-framework). Anyway, to not waste time, looking for an alternative to Tika, I found that code in the blog CodezRule, but it was written in 2010. I updated parts of the code using the "extractor" classes from POI and I got a simpler library than Tika, with less JAR files, but able to do what I need. It worked for me, and maybe it will be useful for you.

## ToDo

- Convert it to a Singleton or a static class
- Add log entries
- Continue to make improvements! (but without change the scope :-)  This is just the 0.0.1 version ;-)
