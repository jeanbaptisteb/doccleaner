# -*- coding: utf-8 -*-


#-------------------------------------------------------------------------------
# Name: DocCleaner.py
# Purpose: Opens zip documents (docx, odt), apply an xslt to the xml they contain, and saves
#
# Author: Jean-Baptiste Bertrand
#
# Created: 14/11/2013
# Copyright: LGPL (c) Jean-Baptiste Bertrand
#-------------------------------------------------------------------------------
#Inspired by EditDocx of John Sutton : https://github.com/jdsutton/EditDocx/blob/master/editDocx.py


import shutil
import zipfile
#from lxml import etree
from defusedxml import lxml
#import lxml
import os
import sys, getopt
import gettext
import locale
import tempfile

def init_localization():
    '''prepare l10n'''
    #print locale.setlocale(locale.LC_ALL,"")
    locale.setlocale(locale.LC_ALL, '') # use user's preferred locale

    # take first two characters of country code
    loc = locale.getlocale()
    
    #filename = os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), "lang", "messages_%s.mo" % locale.getlocale()[0][0:2])
    filename = os.path.join("lang", "messages_%s.mo") % locale.getlocale()[0][0:2]
    try:
        print "Opening message file %s for locale %s" % (filename, loc[0])
        #If the .mo file is badly generated, this line will return an error message: "LookupError: unknown encoding: CHARSET"
        trans = gettext.GNUTranslations(open(filename, "rb"))

    except IOError:
        print "Locale not found. Using default messages"
        trans = gettext.NullTranslations()

    trans.install()

def createDocument(sourceFile, destFile):
    #Creating a copy of the source document
    shutil.copyfile(sourceFile, destFile)

def openDocument(fileName, subFileName):
    #opens zip file and getting the subfile (for instance, in a docx, word/document.xml)
    
    mydoc = zipfile.ZipFile(fileName)
    xmlcontent = mydoc.read(subFileName)
    document = lxml._etree.fromstring(xmlcontent)
    return document

def saveElement(fileName, element):
    #Save a .xml element
    f = open(fileName, 'w')
    text = lxml._etree.tostring(element, pretty_print = True)
    f.write(text)
    f.close()


def usage():
    print _("Some arguments are missing!")
    print _("Usage :")
    print _(" -i <inputFile.docx>")
    print _(" -o <outputFile.docx>")
    print _(" -t <transformFile.xsl>")
    print _(" -p <XSLparameter=value>")

def checkCommandline(inputArg):
    if os.path.isfile(inputArg) == True:
        return os.path.isfile(inputArg)
    else:
        print _("%s is an invalid file name, or does not exist") % inputArg
        print _("You must define a valid file name")
        usage()
        return os.path.isfile(inputArg)

def checkIfFileExists(fileToCheck):

    try: 
        test = open(fileToCheck)
        test.close()
        return True
    
    except IOError:
        print  ("%s does not exist!") % fileToCheck
        return False
#MAIN---------------------------------------------------------------------------

def main(argv):

    try:                                
        opts, args = getopt.getopt(argv, "i:o:t:s:p:", ["input=", "output=", "transform=", "subfile=", "XSLparameter="])
        
    except getopt.GetoptError:           
        usage()                          
        sys.exit(2)          

    inputFile = None
    outputFile = None
    transformFile = None
    subFile = None
    XSLparameter = None
    for opt, arg in opts:
        if opt in ("-i", "--input"): 
            inputFile = arg   
        elif opt in ("-o", "--output"):
            outputFile = arg
        elif opt in ("-t", "--transform"):
            #Will have to forbid the use of a path to a xsl
            transformFile = arg
        elif opt in ("-s", "--subfile"):
            subFile = arg
        elif opt in ("-p", "--XSLparameter"):
            XSLparameter = arg
    
    if inputFile == None:
        sys.exit(2)
    elif outputFile == None:
        sys.exit(2)
    elif transformFile == None:
        sys.exit(2)
        
    if checkIfFileExists(inputFile) == False:
        sys.exit(2)


    #Retrieving the file extension, to know which kind of document we are processing
    inputFile_Name, inputFile_Extension = os.path.splitext(inputFile)
    fileType = inputFile_Extension[1:]
    
    #Retrieving the path of the script's folder
    script_directory = os.path.dirname(os.path.realpath(__file__)) #Will have to replace __file__ by sys.argv[0]) when freezing with py2exe

    #Retrieving the path containing xsl files for the current format
    #(for instance, for docx processing, xls are in the ./docx/ subdirectory)
   
    xslFilesPath = os.path.join(script_directory, fileType)  
    print "1 : " + xslFilesPath
    xslFilePath = os.path.join(xslFilesPath, transformFile)
    print "2 : " + xslFilePath
    #Check if the path to the xsl file is an authorized path
    if checkIfFileExists(xslFilePath) == True:
        transformFile = xslFilePath
    else:
        print _("The XSL %s does not exist!") % transformFile
        print _("The following XSL are available :")
        for xslfile in os.listdir(xslFilesPath):
            if xslfile.endswith(".xsl"):
                print "- " + xslfile
        sys.exit(2)

    
    #Function to make a xsl transformation with the xsl defined in command line
    transform = lxml._etree.XSLT(lxml._etree.parse(transformFile))


    
    #To retrieve the data file listing the path of the zip subfiles
    pathfile = os.path.join(script_directory,
                            fileType, 
                            fileType + '.path')
    
    
    #Creating a copy of the sourceFile
    createDocument(inputFile, outputFile)
    
    #Creating a temp folder
    folder = tempfile.mkdtemp()

    #Fill the folder with all files zipped in the input file
    f = zipfile.ZipFile(outputFile, mode='r', compression=zipfile.ZIP_DEFLATED)
    for name in f.namelist():
        f.extract(name, folder)
    f.close()    
    
    #declaring a "lines" list, which is intended to contain a list of original doc subfiles 
    #lines = ['']
    
    #Check if a parameter has been passed for the xsl
    if XSLparameter != None:
        
        #Convert the XSL parameter into a list, splitting with a semi-colon (;)
        #NB : 1 list element = one argument to use on the next XSL processing, not in the current XSL processing
        #To use parameters simulteanously, split them with a comma (,) inside a larger split with semi-colon
        #Example : i want to use simulteanously a "foo" and "foo2" parameters, and then make a second XSL processing with "foo3" and "foo4" parameters :
            #--XSLparameter foo1='my value 1',foo2='my value 2';foo3='my value 3', foo4='my value 4'
        XSLparameter = XSLparameter.split(",") 


    else:        
        XSLparameter = None
        
    #Check if a subFile has been passed as an argument (-s "subfile name")
    if subFile == None:
        #If no subFile has been passed as an argument, retrieving the list of subfiles from the .path file    
        lines = [line.strip() for line in open(pathfile)]
        
    elif subFile == "":
        lines = [line.strip() for line in open(pathfile)]
        
    else:
    #If a subfile has been passed as an argument, process it exclusively
        #The argument can contain a list of files, separated by a comma
        lines = [line.strip() for line in subFile.split(",")]
        
    #For each file listed in the "lines" list    
    for line in lines:
        #TODO : each line in the .path file contains a string representing a path, with a "/" path separator.
        #This separator is only valid in Windows, so we will need to replace any "/" found in the "line" var, with the OS path separator (os.sep).
        #Otherwise, the script won't be usable on Mac or Linux.
           
        #DEBUG
        #check if each listed file exists...     
        try:
            #Retrieving the document
            document = openDocument(inputFile, line)
            #Process it with the xsl file defined as transformFile (with the -t commandline)
            
            #If some parameters have been passed, they are in a list "XSLparameter"
            
            if XSLparameter != None:
                
                paramDict = {}
                #for each element separated by a coma (for instance "fooA=value A,fooB=value B", we'll make 2 processings: one for "fooA=value A", another one for "fooB=value B"
                for element in XSLparameter:      
                    try:
                        #If we have several parameters, separated by a comma -> generator comprehension                        
                        paramDict = dict(item.split(",") for item in element.split('='))

                    except:
                        #if we have only one parameter
                        paramDict[str(element.split("=")[0])] = (str(element.split("=")[1]))
                        
                #Pass all the parameters simulteanously
                document = transform(document, **paramDict)#, **paramDict)#, paramDict)


            elif XSLparameter == "":
                #Empty parameter, don't pass it
                document = transform(document)
            else:
                #Any other case (e.g. no parameter), don't pass it
                document = transform(document)

           

            #Extract it to the temp "files" folder
            #print os.path.join("files", line)
            saveElement(os.path.join(folder, line), document)
        #if file doesn't exist, do not try to process it
        except Exception,e:
            print _("line %s does not exist!") %  line
            pass
    #Unzip the outputFIle
    z= zipfile.ZipFile(outputFile, mode='w', compression=zipfile.ZIP_DEFLATED)
    
    #Copy all the "files" in the outputFile
    #...and overwrite existing files
    
    os.chdir(folder)
    for root, dirs, files in os.walk("."):
        for f in files:
            z.write(os.path.join(root, f))
    os.chdir("..")
    z.close() 
    #Deleting the temp folder

    try:
        shutil.rmtree(folder)
        print folder + " deleted"
    except:
        pass
if __name__ == '__main__':
    

    init_localization()
    main(sys.argv[1:])

