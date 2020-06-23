# linkedin-pdf-scraper
Python script to convert LinkedIn profiles in PDF to an Excel xslx file

[<img src="https://img.shields.io/badge/live-demo-brightgreen?style=for-the-badge&logo=appveyor?">](https://colab.research.google.com/drive/11Y05qje8OhoB36qS7Mo5oEo3vdJygPGT?usp=sharing)

Click the badge above for a live demo on Google Colab! Simply "Save a copy" of the notebook to run it yourself :)


![alt tag](https://github.com/btphan95/linkedin-pdf-scraper/blob/master/preview.png?raw=true)
            
Requirements
============
python

pdfplumber

xlsxwriter

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
