# linkedin-pdf-scraper
 Python script to convert LinkedIn profiles in PDF to an Excel xslx file

linkedin pdf parsing

![alt tag](https://github.com/btphan95/linkedin-pdf-scraper/blob/master/preview.png?raw=true)
            
Requirements
============
python
pdfplumber
pandas

Usage
============
<pre> scrape.py -i inputfile -o outputfile
</pre>
Script will search for 'inputfile', a PDF file, and will create an Excel xlsx 'outputfile'.

Example usage:
<pre>
python scrape.py -i profiles.pdf -o profiles.xlsx
</pre>
