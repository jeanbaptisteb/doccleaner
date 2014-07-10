DocCleaner 0.1.2
==========


A python command-line utility which uses XSLT 1.0 to edit zipped, XML-based files (for instance, docx or odt files). It can be rather easily extended with xsl stylesheets.

It is primarely intended for automating some copyediting tasks (removing local formatting, correcting non-breaking spaces according to the language's typography rules, checking the use of smart quotes, etc.). 

This is often a more efficient and reliable method than using VBA or OOBASIC macros, particularly on large documents.
Moreover, it is independant from the OS and the word processor you use.

###A word of caution
About **XML files as input** ("-i" argument): defusedxml implementation still pending. Not sure if untrusted documents should be allowed as input via the "-i" argument.

About **XSL files as input** ("-t" argument): If you want to use this script on a public server open to everyone, you should forbid the use of untrusted XSL file as input, and may use a whitelist for such files.

In a nutshell: if you use this script with documents you trust, on your personal computer -> no security issue. If you plan to use it as a public service available on a server, you should add a security layer for preventing malicious inputs.

##SHORT DOCUMENTATION

You need to create a xsl for each processing you want to make to the document. 

The xsl will be applied to all xml "subfiles" contained in the zipped file. A set of these subfiles is defined in the ".path" file (in the "docx" or "odt" subdirectory, you will find a docx.path or odt.path, which contains a list of subfiles pathes).

You can also apply the XSL file on xml subfiles of your choice, with the -s argument.

###Usage :

docCleaner.py  
 -i "input file"   
 -o "output file"   
 -t "xslt file"  
 -s "subfile(s) to process, contained in the inputfile" (optional)  
 -p "parameter(s) to pass to the XSL stylesheet (optional)"
 
###Examples:
 To apply the XSL to document.xml, footnotes.xml, endnotes.xml (all contained in the "MyDocToProcess.docx" document):

    python docCleaner.py -i "c:\MyDocToProcess.docx" -o "c:\dest\ProcessedDoc.docx" -t "c:\MyTransformationFile.xsl"
 
To apply the XSL only to endnotes.xml, that is to say it will process only the endnotes of the docx document:
 
    python docCleaner.py -i "c:\MyDocToProcess.docx" -o "c:\dest\ProcessedDoc.docx" -t "c:\MyTransformationFile.xsl" -s "word/endnotes.xml"
    
To apply parameters to the XSL stylesheet, e.g. $foo=True, $foo2="blue", and $foo3=24 :

    python docCleaner.py -i "c:\MyDocToProcess.docx" -o "c:\dest\ProcessedDoc.docx" -t "c:\MyTransformationFile.xsl" -p "foo=True,foo2='blue',foo3=24"
 

