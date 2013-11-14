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


#Beware, this is not ready for prime time, some lines still need debug!


#Some comments are in french, i'll translate them soon

import shutil
import zipfile
from lxml import etree
import os
import sys, getopt

def createDocument(sourceFile, destFile):
    #Creating a copy of the source document
    shutil.copyfile(sourceFile, destFile)

def openDocument(fileName, subFileName):
    #opens zip file and getting the subfile (for instance, in a docx, word/document.xml)
    
    mydoc = zipfile.ZipFile(fileName)
    xmlcontent = mydoc.read(subFileName)
    document = etree.fromstring(xmlcontent)
    return document

def createTempFolder(folder):
    #Create a temp folder
    if os.path.exists(folder):
        shutil.rmtree(folder)
    os.mkdir(folder)

def saveElement(fileName, element):
    #Save a .xml element
    f = open(fileName, 'w')
    text = etree.tostring(element, pretty_print = True)
    f.write(text)
    f.close()


def usage():
    print "Some arguments are missing!"
    print "Usage :"
    print " -i <inputFile.docx>"
    print " -o <outputFile.docx>"
    print " -t <transformFile.xsl>"    
#MAIN---------------------------------------------------------------------------

def main(argv):
    
    try:                                
        opts, args = getopt.getopt(argv, "i:o:t:", ["input=", "output=", "transform="])
        
    except getopt.GetoptError:           
        usage()                          
        sys.exit(2)          
       
    for opt, arg in opts:
        if opt in ("-i", "--input"): 
            inputFile = arg                     
        elif opt in ("-o", "--output"):
            outputFile = arg
        elif opt in ("-t", "--transform"):
            transformFile = arg

    #For debugging purposes            
    print inputFile
    print outputFile
    print transformFile            

    #Function to make a xsl transformation with the xsl defined in command line
    transform = etree.XSLT(etree.parse(transformFile))
    
    #Retrieving the file extension, to know which kind of document we are processing
    inputFile_Name, inputFile_Extension = os.path.splitext(inputFile)
    fileType = inputFile_Extension[1:]
    
    #Retrieving the path of the script's folder
    script_directory = os.path.dirname(sys.executable)
    print script_directory
    
    #To retrieving the data file listing the path of the zip subfiles
    pathfile = os.path.join(script_directory,
                            fileType, 
                            fileType + '.path')
    
    
    #Creating a copy of the sourceFile
    createDocument(inputFile, outputFile)
    
    #Creating a temp "files" folder
    folder = "files"
    createTempFolder(folder)
    
    #Fill it with all the files of the outputFile
    f = zipfile.ZipFile(outputFile, mode='r', compression=zipfile.ZIP_DEFLATED)
    for name in f.namelist():
        f.extract(name, folder)
    f.close()    
    
    #For each file listed in the .path file
    lines = [line.strip() for line in open(pathfile)]
    for line in lines:
        #We have to retrieve the file we are interested in...
        try:        
            document = openDocument(inputFile, line)
        #then, we have to process it with the xsl defined in the command line
            document = transform(document)
        
        #at last, we extract it in the temporary "files" folder
        
            print os.path.join("files", line)
            saveElement(os.path.join("files", line), document)
        except:
            print "No " + line

    #Unzipping the outputFile
    z= zipfile.ZipFile(outputFile, mode='w', compression=zipfile.ZIP_DEFLATED)
    
    #Copying the whole "files" folder in the outputFile,
    #...and overwriting existing files
       
    os.chdir("files")
    for root, dirs, files in os.walk("."):
        for f in files:
            z.write(os.path.join(root, f))
    os.chdir("..")
    z.close() 

if __name__ == '__main__':
    main(sys.argv[1:])
