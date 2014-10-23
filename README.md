DocCleaner 0.2
==========


A python command-line utility which uses XSLT 1.0 to edit zipped, XML-based files (for instance, docx or odt files). It can be rather easily extended with xsl stylesheets.

It is primarely intended for automating some copyediting tasks (removing local formatting, correcting non-breaking spaces according to the language's typography rules, checking the use of smart quotes, etc.). 

This is often a more efficient and reliable method than using VBA or OOBASIC macros, particularly on large documents.
Moreover, it is independant from the OS and the word processor you use.

Plugins for text processing softwares (Microsoft Word, LibreOffice) are developped at  **https://github.com/jbber/doccleaner-plugins**. 

Pull requests welcomed!


##SHORT DOCUMENTATION
Compatible with Python 2.7 and 3.4.

You need to create a xsl for each processing you want to make to the document. 

The xsl will be applied to all xml "subfiles" contained in the zipped file. A set of these subfiles is defined in the ".path" file (e.g. in the "docx" or "odt" subdirectory, you will find a docx.path or odt.path, which contains a list of subfiles pathes). 

You can also apply the XSL file on xml subfiles of your choice, with the -s argument. If you don't define subfile to process, all the subfiles listed in the .path file will be processed by default.


###Usage

doccleaner.py  
 -i "input file"  
 -o "output file"  
 -t "xslt file"  
 -s "path to a json configuration file defining which subfile(s) to process" (optional)  
 -p "parameter(s) to pass to the XSL stylesheet" (optional) 
 
**Alternative syntax**  
doccleaner.py  
 --input "input file"  
 --output "output file"  
 --transform "xslt file"  
 --subfile "path to a json configuration file defining which subfile(s) to process" (optional)  
 --XSLparameter "parameter(s) to pass" (optional) 


####$tempdir variable
The python script also create a "$tempdir" variable, which can be passed to the XSL sheet. This variable contains the absolute path of the temporary folder to which the XML files are extracted. It can be useful if you need to process multiple XML files at the same time with the same XSL sheet.
 
#### Json conf file specifications
When using the "subfile" parameter, you need to define a Json configuration file, containing a list of input subfiles from which you want to retrieve contents, and a list of output subfiles where to put the result of the XSL transformation. For an example, see the docx.json file in the docx folder, or the examples below.
 
###Examples
####From command line
 To apply the XSL to document.xml, footnotes.xml, endnotes.xml (all contained in the "MyDocToProcess.docx" document):

    python doccleaner.py -i "c:\MyDocToProcess.docx" -o "c:\dest\ProcessedDoc.docx" -t "c:\MyTransformationFile.xsl"

To apply parameters to the XSL stylesheet, e.g. $foo=True, $foo2="blue", and $foo3=24 :

    python doccleaner.py -i "c:\MyDocToProcess.docx" -o "c:\dest\ProcessedDoc.docx" -t "c:\MyTransformationFile.xsl" -p "foo=True,foo2='blue',foo3=24"
	
To apply the XSL only to endnotes.xml, that is to say it will process only the endnotes of the docx document:

    python doccleaner.py  -i "c:\MyDocToProcess.docx" -o "c:\dest\ProcessedDoc.docx" -t "c:\MyTransformationFile.xsl" -s "path/to/confFile.json"
	
with the following Json conf file :
```
{ 

    "subfile_input" : ["word/endnotes.xml"], 
    "subfile_output": ["word/endnotes.xml"]
   
}
```

To get the contents of document.xml, and then output the result of the XSL processing to a comments.xml subfile :

    python doccleaner.py -i "c:\MyDocToProcess.docx" -o "c:\dest\ProcessedDoc.docx" -t "c:\MyTransformationFile.xsl" -s "path/to/confFile.json"

with the following Json conf file :
```
{ 

    "subfile_input" : ["word/document.xml"], 
    "subfile_output": ["word/comments.xml"]
   
}
```

If you need to process several subfiles, the Json may look like this : 
```
{ 

    "subfile_input" : ["word/document.xml", "word/footnotes.xml"], 
    "subfile_output": ["word/comments.xml", "word/footnotes.xml"]
   
}
```
The example above will output the result of the XSL transformation :
* from word/document.xml to word/comments.xml
* from word/footnotes.xml to word/footnotes.xml

You will probably have to use the [$tempdir variable](#tempdir-variable) in your XSL for making the above example work.
    
####From a script
#####Python 2 or 3
Applying a MyTransformationFile.xsl transformation sheet to the endnotes of the document MyDocToProcess.docx, with XSL parameters $foo=True, $foo2="blue", and $foo3=24:
```
from doccleaner import doccleaner

inputDoc = 'c:\\MyDocToProcess.docx'
outputDoc = 'c:\\dest\\ProcessedDoc.docx'
xslDoc = 'c:\\MyTransformationFile.xsl'
subFiles = "path/to/conf.json"
params = "foo=True,foo2='blue',foo3=24"

doccleaner.main(['--input', str(inputDoc),
                 '--output', str(outputDoc),
                 '--transform', str(xslDoc),
                 '--subfile', "path/to/conf.json",
                 '--XSLparameter', str(params)
                 ])
```
#####Python 2
Same as above, with a syntax which will work only with Python 2 :
```
import doccleaner #won't work with Python 3

inputDoc = 'c:\\MyDocToProcess.docx'
outputDoc = 'c:\\dest\\ProcessedDoc.docx'
xslDoc = 'c:\\MyTransformationFile.xsl'
subFiles = 'path/to/conf.json'
params = "foo=True,foo2='blue',foo3=24"

doccleaner.doccleaner.main(['--input', str(inputDoc),
                 '--output', str(outputDoc),
                 '--transform', str(xslDoc),
                 '--subfile', "path/to/conf.json",
                 '--XSLparameter', str(params)
                 ])
```
###A word of caution about server side processings
About **XML files as input** ("-i" argument): defusedxml implementation still pending. Untrusted documents should not be allowed as input via the "-i" argument. If you want to allow untrusted documents as input, you should add a layer a security before processing them with doccleaner.

About **XSL files as input** ("-t" argument): If you want to use this script on a public server open to everyone, you should forbid the use of untrusted XSL file as input, and may use a whitelist for such files.

In a nutshell: if you use this script with documents you trust, on your personal computer -> no security issue. If you plan to use it as a public service available on a server, you should add a security layer for preventing malicious inputs.
