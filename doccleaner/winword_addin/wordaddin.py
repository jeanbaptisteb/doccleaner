#coding: utf-8 -*-

#This is just a proof of concept with probably a lot of bugs, use at your own risks!

#It creates a new tab in the MS Word ribbon, with buttons calling docCleaner scripts from inside Word
#Currently, the only action available is calling a XSL which removes all direct formatting of the current document, except italics, bold, underline, superscript & subscript
#You have to put the docCleaner.py script in the same directory
#Launch the script, and it will automatically install the new tab in MS Word
#To uninstall it, launch it with the --unregister argument
#You can also remove it from Word (in the "Developer" tab, look for COM Addins > Remove)

#Inspired by the Excel addin provided in the win32com module demos, and the "JJ Word Addin" (I don't remember where I get it, but thanks!)
#!python2.7.8
import sys
sys.path.insert(0, 'pkgs')

import win32com
win32com.__path__
from win32com import universal
from win32com.server.exception import COMException
from win32com.client import gencache, DispatchWithEvents
import winerror
import pythoncom
from win32com.client import constants, Dispatch
import sys
import win32com.client
import doccleaner
import os
import win32ui
import win32con
import locale
import gettext
import ConfigParser
import doccleaner.localization
import tempfile
import shutil
import mimetypes

#win32com.client.gencache.is_readonly=False
#win32com.client.gencache.GetGeneratePath()
# Support for COM objects we use.
gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 1, bForDemand=True) # Office 9
gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 5, bForDemand=True)

# The TLB defiining the interfaces we implement
universal.RegisterInterfaces('{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}', 0, 1, 0, ["_IDTExtensibility2"])
universal.RegisterInterfaces('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 5, ["IRibbonExtensibility", "IRibbonControl"])

def checkIfDocx(filepath):
    if mimetypes.guess_type(filepath)[0] == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        return True
    else:
        return False
        
#TODO : localization             
def init_localization():
    '''prepare l10n'''
    print locale.setlocale(locale.LC_ALL,"")
    locale.setlocale(locale.LC_ALL, '') # use user's preferred locale

    # take first two characters of country code
    loc = locale.getlocale()
    
    #filename = os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), "lang", "messages_%s.mo" % locale.getlocale()[0][0:2])
    filename = os.path.join("lang", "messages_{0}.mo").format(locale.getlocale()[0][0:2])
    try:
        print "Opening message file {0} for locale {1}".format(filename, loc[0])
        #If the .mo file is badly generated, this line will return an error message: "LookupError: unknown encoding: CHARSET"
        trans = gettext.GNUTranslations(open(filename, "rb"))

    except IOError:
        print "Locale not found. Using default messages"
        trans = gettext.NullTranslations()

    trans.install()
class WordAddin:
    config = ConfigParser.ConfigParser()
    
    _com_interfaces_ = ['_IDTExtensibility2', 'IRibbonExtensibility']
    _public_methods_ = ['clean', 'do','GetImage']
    _reg_clsctx_ = pythoncom.CLSCTX_INPROC_SERVER
    _reg_clsid_ = "{C5482ECA-F559-45A0-B078-B2036E6F011A}"
    _reg_progid_ = "Python.DocCleaner.WordAddin"
    _reg_policy_spec_ = "win32com.server.policy.EventHandlerPolicy"

    def __init__(self):
        self.appHostApp = None    

        
    def do(self,ctrl):
    #This is the core of the Word addin : manipulates docs and calls docCleaner
    #The ctrl argument is a callback for the button the user made an action on (e.g. clicking on it)
        
            
        #Creating a word object inside a wd variable
        wd = win32com.client.Dispatch("Word.Application")
        

        try:
            #Check if the file is not a new one (unsaved)
            if os.path.isfile(wd.ActiveDocument.FullName) == True:
                #Before processing the doc, let's save the user's last modifications
                
                wd.ActiveDocument.Save                   
                    

                    
                    
                originDoc = wd.ActiveDocument.FullName #:Puts the path of the current file in a variable
                tmp_dir = tempfile.mkdtemp() #:Creates a temp folder, which will contain the temp docx files necessary for processing        
                
                #TODO: If the document is in another format than docx, convert it temporarily to docx
                #At the processing's end, we'll have to convert it back to its original format, so we need to store this information
                  
                    
                
                transitionalDoc = originDoc #:Creates a temp transitional doc, which will be used if we need to make consecutive XSLT processings. #E.g..: original doc -> xslt processing -> transitional doc -> xslt processing -> final doc -> copying to original doc                
                newDoc = os.path.join(tmp_dir, "~" + wd.ActiveDocument.Name) #:Creates a temporary file (newDoc), which will be the docCleaner output
                
                
                jj = 0 #:This variable will be increased by one for each XSL parameter defined in the wordAddin_xx.ini file (separated by a semi-colon ;). Used for handling temp docx filenames, and for subfiles consecutive processing (vs. simulteanous processing, which is already handled in the docCleaner script)
                
                #Then, we take the current active document as input, the temp doc as output
                #+ the XSL file passed as argument ("ctrl. Tag" variable, which is a callback for the ribbon button tag)
                
                #Check if the XSLparameter contains a semi-colon, which means we have to make several XSL processing
                try:                
                    XSLparameters = self.config.get(str(ctrl. Tag), 'XSLparameter').split(";")
                except:
                    XSLparameters = ""
                
                if XSLparameters != None:
                
                    for XSLparameter in XSLparameters:

                        #Check if there are subfiles to process consecutively instead of simulteanously (separated by a semi-colon instead of a comma)
                        #NB : the script implies that in the ini file, we can have:
                        # 1) one XSL parameter, and a single subfiles processing
                        # 2) multiple XSL parameters, and the exact same number of consecutive subfiles processings
                        # 3) multiple XSL parameters, and a single subfiles processing
                        #We can never have multiple subfiles and a single XSL processing, because this use case is handled separately by the docCleaner script. If we're in this case, simply split subfiles with commas (",") instead of semi-colon (";")
                        try:
                            subFileArg = str(self.config.get(str(ctrl. Tag), 'subfile')).split(";")[jj]
                            
                        except:
                            #Probably a "out of range" error, which means there is a single subfiles string to process
                            subFileArg = self.config.get(str(ctrl. Tag), 'subfile')

                        if jj > 0:
                           
                            #If there is more than one XSL parameter, we'll have to make consecutive processings
                            newDocName, newDocExtension = os.path.splitext(newDoc)
                            transitionalDoc = newDoc
                            newDoc =  newDocName + str(jj)+ newDocExtension 
                 
                                         
                        doccleaner.main(['--input', str(transitionalDoc), 
                                     '--output', str(newDoc), 
                                     '--transform', os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])),
                                                                 "docx", str(ctrl. Tag) + ".xsl"),
                                    '--subfile', subFileArg,
                                     '--XSLparameter', XSLparameter
                                     ]) 


                        jj +=1
                            
                #Opening the temp file
                wd.Documents.Open(newDoc)
            
                #Copying the temp file content to the original doc                
                #To do this, never use the MSO Content.Copy() and Content.Paste() methods, because :
                # 1) It would overwrite important data the user may have copied to the clipboard.
                # 2) Other programs, like antiviruses, may use simulteanously the clipboard, which would result in a big mess for the user.
                #Instead, use the Content.FormattedText function, it's simple, and takes just one line of code:
                wd.Documents(originDoc).Content.FormattedText = wd.Documents(newDoc).Content.FormattedText

                #Closing and removing the temp document                
                wd.Documents(newDoc).Close()
                os.remove(newDoc) 
      
                #Saving the changes
                wd.Documents(originDoc).Save
                
                #Removing the whole temp folder
                try:
                    shutil.rmtree(tmp_dir)
                except:
                    #TODO: What kind of error would be possible when removing the temp folder? How to handle it?
                    pass
                
            else:
                win32ui.MessageBox("You need to save the file before launching this script!"
                ,"Error",win32con.MB_OK)

        except Exception, e:
            
            tb = sys.exc_info()[2]
            
            win32ui.MessageBox(str(e) + "\n" +
            str(tb.tb_lineno)+ "\n" +
            str(newDoc)
            ,"Error",win32con.MB_OKCANCEL)

            
           
    def GetImage(self,ctrl):
        #TODO : Is this function actually useful?
        #TODO : Retrieving the path from the conf file
        from gdiplus import LoadImage
        i = LoadImage( 'path/to/image.png' )
        return i

    def GetCustomUI(self,control):
        #Getting the button variables from the localized ini file
        #TODO : 
        self.config.read(os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), 'wordAddin_fr.ini'))
        #Constructing the Word ribbon XML                                          
        ribbonHeader = '''<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
                            <ribbon startFromScratch="false">
                                   <tabs>
                                          <tab id="CustomTab" label="Sample tab">
                                                 <group id="MainGroup" label="Sample group">
                       '''
        
        ribbonFooter = '''</group>       
                     </tab>        
                     </tabs>        
                     </ribbon>        
                     </customUI>        
                 '''       
        
        #Initializing the ribbon body
        ribbonBody = ""
        buttonsNumber = 0

        #Generating dynamically the buttons of the ribbon, according to the available XSL sheets for the docx format
        for path, subdirs, files in os.walk(os.path.join('..', 'docx')):#os.walk(os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), "docx")):      
            for filename in files:       
                if filename.endswith(".xsl"):      
                    buttonsNumber += 1
                    
                    ribbonBody += '''<button id="{0}" label="{1}" imageMso="{2}"
                                size="{3}" onAction="{4}" tag="{5}"/>'''.format(
                                "Button" + str(buttonsNumber),              #variable {0} : id 
                                self.config.get(filename[:-4], 'label'),    #variable {1} : label, from the ini file
                                self.config.get(filename[:-4], 'imageMso'), #variable {2} : button icon from the ini file
                                self.config.get(filename[:-4], 'size'),     #variable {3} : button size from ini file
                                self.config.get(filename[:-4], 'onAction'), #variable {4} : action from ini file. Call the function do() most of time (onAction=do)
                                str(filename[:-4])                          #variable {5} : button tag 
                                )
        #Generating the final XML for the ribbon
        s = ribbonHeader + ribbonBody + ribbonFooter
        return s

        
    def OnConnection(self, application, connectMode, addin, custom):
        print "OnConnection", application, connectMode, addin, custom
        try:
            self.appHostApp = application
        except pythoncom.com_error, (hr, msg, exc, arg):
            print "The Word call failed with code {0}: {1}".format(unicode(hr), msg)
            if exc is None:
                print "There is no extended error information"
            else:
                wcode, source, text, helpFile, helpId, scode = exc
                print "The source of the error is", source
                print "The error message is", text
                print "More info can be found in {0} (id={1})".format(unicode(helpFile), helpId)

    def OnDisconnection(self, mode, custom):
        print "OnDisconnection"
        #self.appHostApp.CommandBars("PythonBar").Delete
        self.appHostApp=None
        
        
    def OnAddInsUpdate(self, custom):
        print "OnAddInsUpdate", custom
        
    def OnStartupComplete(self, custom):
        print "OnStartupComplete", custom
        
    def OnBeginShutdown(self, custom):
        print "OnBeginShutdown", custom
        
        

def RegisterAddin(klass):
    import _winreg
    key = _winreg.CreateKey(_winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Office\\Word\\Addins")
    subkey = _winreg.CreateKey(key, klass._reg_progid_)
    _winreg.SetValueEx(subkey, "CommandLineSafe", 0, _winreg.REG_DWORD, 0)
    _winreg.SetValueEx(subkey, "LoadBehavior", 0, _winreg.REG_DWORD, 3)
    _winreg.SetValueEx(subkey, "Description", 0, _winreg.REG_SZ, "DocCleaner Word Addin")
    _winreg.SetValueEx(subkey, "FriendlyName", 0, _winreg.REG_SZ, "DocCleaner Word Addin")
    
    word = gencache.EnsureDispatch("Word.Application")
    mod = sys.modules[word.__module__]
    print "The module hosting the object is", mod


def UnregisterAddin(klass):
    import _winreg
    try:
        _winreg.DeleteKey(_winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Office\\Word\\Addins\\" + klass._reg_progid_)
    except WindowsError:
        pass
def main():
    init_localization()
    
    import win32com.server.register
    win32com.server.register.UseCommandLine( WordAddin )
    if "--unregister" in sys.argv:
        UnregisterAddin( WordAddin )
    else:
        RegisterAddin( WordAddin )
if __name__ == '__main__':
    main()