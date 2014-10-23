# -*- coding: utf-8 -*-


#-------------------------------------------------------------------------------
# Name: doccleaner.py
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
import simplejson

class FileResolver(lxml._etree.Resolver):
    def resolve(self, url, pubid, context):
        return self.resolve_filename(url, context)
        
def load_json(filename):
    f = open(filename, "r")
    data = f.read()
    f.close()
    return simplejson.loads(data)
    
def init_localization():
    '''prepare l10n'''
    #print locale.setlocale(locale.LC_ALL,"")
    locale.setlocale(locale.LC_ALL, '') # use user's preferred locale

    # take first two characters of country code
    loc = locale.getlocale()

    #filename = os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), "lang", "messages_%s.mo" % locale.getlocale()[0][0:2])
    filename = os.path.join("lang", "messages_%s.mo") % locale.getlocale()[0][0:2]
    try:
        print("Opening message file %s for locale %s" % (filename, loc[0]))
        #If the .mo file is badly generated, this line will return an error message: "LookupError: unknown encoding: CHARSET"
        trans = gettext.GNUTranslations(open(filename, "rb"))

    except IOError:
        print("Locale not found. Using default messages")
        trans = gettext.NullTranslations()
    trans.install()

def createDocument(sourceFile, destFile):
    #Creating a copy of the source document
    shutil.copyfile(sourceFile, destFile)


def openDocument(fileName, subFileName, parser):
    #opens zip file and getting the subfile (for instance, in a docx, word/document.xml)
    mydoc = zipfile.ZipFile(fileName)
    xmlcontent = mydoc.read(subFileName)
    document = lxml._etree.fromstring(xmlcontent, parser)
    return document

def saveElement(fileName, element):
    #Save a .xml element
    f = open(fileName, 'wb')
    text = lxml._etree.tostring(element, pretty_print = True)
    f.write(text)
    f.close()

def usage():
    #TODO: updating this part, new parameters are available
    print("Some arguments are missing!")
    print("Usage :")
    print(" -i <inputFile.docx>")
    print(" -o <outputFile.docx>")
    print(" -t <transformFile.xsl>")
    print(" -p <XSLparameter=value>")

def checkIfFileExists(fileToCheck):
    try:
        test = open(fileToCheck)
        test.close()
        return True
    except IOError:
        print(("%s does not exist!") % fileToCheck)
        return False

def main(argv):
    try:
        opts, args = getopt.getopt(argv, "i:o:t:s:p:g", ["input=", "output=", "transform=", "subfile=", "XSLparameter="])

    except:# getopt.GetoptError:
        usage()
        sys.exit(2)

    inputFile = None
    outputFile = None
    transformFile = None
    subFile = None
    XSLparameter = None
    tempdir = None
    #Creating a temp folder
    folder = tempfile.mkdtemp()
    print(folder + "created")
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
    
    
    #Variable to pass the absolute path of the temporary folder to the XSL sheet, if needed
    tempdir = "=".join([ "tempdir", "\"{0}\"".format(str(folder)) ])

    #If "-g" or "--get_tempdir" is used, append tempdir="path/to/temporary folder" to the XSLparameter string
    #It will pass the absolute path of the temporary folder to the XSL sheet, in a $tempdir parameter
    if tempdir != None and XSLparameter != None:
        XSLparameter = ",".join([ XSLparameter, tempdir ])
    elif tempdir != None and XSLparameter == None:
        XSLparameter = tempdir

    #If no input file, output file, nor XSL sheet have been defined, we need to exit the code
    if inputFile == None:
        sys.exit(2)
    elif outputFile == None:
        sys.exit(2)
    elif transformFile == None:
        sys.exit(2)

    if checkIfFileExists(inputFile) == False:
        sys.exit(2)
    if checkIfFileExists(transformFile) == False:
        sys.exit(2)

    #defining a parser
    # parser = lxml._etree.XMLParser()
    parser = lxml._etree.XMLParser(encoding='utf-8', recover=True)
    parser.resolvers.add(FileResolver())

    #Function to make a xsl transformation with the xsl defined in command line
    transform = lxml._etree.XSLT(lxml._etree.parse(open(transformFile, "r", encoding="utf8"), parser))
    
    #To retrieve the data file listing the path of the zip subfiles
    inputFile_Name, inputFile_Extension = os.path.splitext(inputFile)
    fileType = inputFile_Extension[1:]
    script_directory = os.path.dirname(os.path.realpath(__file__))
    
    if subFile == None:
        subFile = os.path.join(script_directory,
                            fileType,
                            fileType + '.json')
    
    subFileConf = load_json(subFile)

    #Creating a copy of the sourceFile
    createDocument(inputFile, outputFile)

    #Fill the folder with all files zipped in the input file
    f = zipfile.ZipFile(outputFile, mode='r', compression=zipfile.ZIP_DEFLATED)
    for name in f.namelist():
        print(folder)
        f.extract(name, folder)
    f.close()

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

    #For each input file listed in the json file
    subfileNumber = 0
    for subfile_input in subFileConf["subfile_input"]:
        #check if each listed file exists...
        try:
            #Retrieving the document
            document = openDocument(inputFile, subFileConf["subfile_input"][subfileNumber], parser)
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
            saveElement(os.path.join(folder, subFileConf["subfile_output"][subfileNumber]), document)
            subfileNumber += 1
        #if file doesn't exist, do not try to process it
        except Exception as e:
            print("Error : " + str(e))
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
        print(folder + " deleted")
    except:
        pass
if __name__ == '__main__':
    init_localization()
    main(sys.argv[1:])
