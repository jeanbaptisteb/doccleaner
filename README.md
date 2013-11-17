DocCleaner
==========


A python utility to edit zipped, XML-based files (for instance, docx or odt files). It can be rather easily extended with xsl stylesheets.

It is primarely intended for automating some copyediting tasks (removing local formatting, correcting non-breaking spaces according to the language's typography rules, checking the use of smart quotes, etc.). 

This is often a more efficient and reliable method than using VBA or OOBASIC macros, particularly on large documents.
Moreover, it is independant from the OS and the word processor you use.


##SHORT DOCUMENTATION

You need to create a xsl for each processing you want to make to the document. 

The xsl will be applied to all xml "subfiles" contained in the zipped file. A set of these subfiles are defined in the ".path" file (in the "docx" or "odt" subdirectory, you will find a docx.path or odt.path, which contains a list of subfiles pathes).

You can also chose to apply the XSL file only on xml subfiles of your choice, with the -s argument.

###Usage :

 docCleaner.py 
    -i "input file" 
    -o "output file" 
    -t "xslt file" 
    (optional) -s "subfile to process, contained in the inputfile"

###Example:

 python docCleaner.py -i "c:\MyDocToProcess.docx" -o "c:\dest\ProcessedDoc.docx" -t "c:\MyTransformationFile.xsl"
 
    -> Will apply the XSL to document.xml, footnotes.xml, endnotes.xml (all contained in the "MyDocToProcess.docx" document)
 
 python docCleaner.py -i "c:\MyDocToProcess.docx" -o "c:\dest\ProcessedDoc.docx" -t "c:\MyTransformationFile.xsl" -s "word/endnotes.xml"
 
    -> Will apply the XSL only to endnotes.xml, that is to say it will process only the endnotes of the docx document.
