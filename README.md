DocCleaner
==========

A python utility to apply xsl files on docx and odt files. Primarely intended for automating some copyediting tasks, but I tried to design it to be easily extended.

Usage :
docCleaner.py -i "input file" -o "output file" -t "xslt file"

example:
python docCleaner.py -i "c:\MyDocToProcess.docx" -o "c:\dest\ProcessedDoc.docx" -t "c:\MyTransformationFile.xsl"
