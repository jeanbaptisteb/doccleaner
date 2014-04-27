
#coding: utf-8 -*-

#This is just a proof of concept with probably a lot of bugs, use at your own risks!

#It creates a new tab in the MS Word ribbon, with buttons calling docCleaner scripts from inside Word
#Currently, the only action available is calling a XSL which removes all direct formatting of the current document, except italics, bold, underline, superscript & subscript
#You have to put the docCleaner.py script in the same directory
#Launch the script, and it will automatically install the new tab in MS Word
#To uninstall it, launch it with the --unregister argument
#You can also remove it from Word (in the "Developer" tab, look for COM Addins > Remove)

#Inspired by the Excel addin provided in the win32com module demos, and the "JJ Word Addin" (I don't remember where I get it, but thanks!)

from win32com import universal
from win32com.server.exception import COMException
from win32com.client import gencache, DispatchWithEvents
import winerror
import pythoncom
from win32com.client import constants, Dispatch
import sys
import win32com.client
import docCleaner
import os
import win32ui
import win32con
# Support for COM objects we use.
gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 1, bForDemand=True) # Office 9
gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 5, bForDemand=True)

# The TLB defiining the interfaces we implement
universal.RegisterInterfaces('{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}', 0, 1, 0, ["_IDTExtensibility2"])
universal.RegisterInterfaces('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 5, ["IRibbonExtensibility", "IRibbonControl"])

#TODO : localization
       
class WordAddin:
    _com_interfaces_ = ['_IDTExtensibility2', 'IRibbonExtensibility']
    _public_methods_ = ['clean', 'do','GetImage']
    _reg_clsctx_ = pythoncom.CLSCTX_INPROC_SERVER
    _reg_clsid_ = "{C5482ECA-F559-45A0-B078-B2036E6F011A}"
    _reg_progid_ = "Python.DocCleaner.WordAddin"
    _reg_policy_spec_ = "win32com.server.policy.EventHandlerPolicy"

    def __init__(self):
        self.appHostApp = None    
        
    def clean(self,action):
        #This is the core of the Word addin : manipulates docs and calls docCleaner
        wd = win32com.client.Dispatch("Word.Application")
        
        try:
            #Check if the file is not a new one (unsaved)
            if os.path.isfile(wd.ActiveDocument.FullName) == True:
                #Before processing the doc, let's save the user's last modifications
                wd.ActiveDocument.Save      
                
                #Put the path of the current file in a variable
                originDoc = wd.ActiveDocument.FullName
                
                #Create a temporary file (newDoc), which will be the docCleaner output
                #TODO : check if the output file does not exist yet; if it does, will have to add another prefix
                newDoc = os.path.join(wd.ActiveDocument.Path,
                                      '~' + wd.ActiveDocument.Name)
                
                #Then, we take the current active document as input, the temp doc as output
                #+ the XSL file passed as argument ("action" variable)
                docCleaner.main(['--input', str(originDoc), 
                             '--output', newDoc, 
                             '--transform', os.path.join(os.path.dirname(os.path.realpath(__file__)),
                                                         "docx", action + ".xsl")
                             ])
                

                #Opening the temp file, make it invisible
                wd.Documents.Open(newDoc).Visible = 0
            
                #Copying the temp file content to the original doc                
                #To do this, NEVER use the MSO Content.Copy() and Content.Paste() methods, because :
                # 1) It would overwrite important data the user may have copied to the clipboard.
                # 2) Other programs, like antiviruses, may use simulteanously the clipboard, which would result in a big mess for the user.
                # 3) Why does it even exist?
                #Instead, use the Content.FormattedText function, it's simple, elegant, and takes just one line of code:
                wd.Documents(originDoc).Content.FormattedText = wd.Documents(newDoc).Content.FormattedText

             
                #Closing and removing the temp document                
                wd.Documents(newDoc).Close()
                os.remove(newDoc) 
      
                #Saving the changes
                wd.Documents(originDoc).Save
                
            else:
                win32ui.MessageBox("You need to save the file before launching this script!"
                ,"PythonTest",win32con.MB_OK)

        except Exception, e:
            win32ui.MessageBox(str(e),"PythonTest",win32con.MB_OKCANCEL)

    def do(self,ctrl):
        
        self.clean("cleanDirectFormatting")
        #win32ui.MessageBox(wd.ActiveDocument.FullName,"PythonTest",win32con.MB_OKCANCEL)
        

            
           
    def GetImage(self,ctrl):
        from gdiplus import LoadImage
        i = LoadImage( 'c:/path.png' )
        print i, 'ddd'
        return i

    def GetCustomUI(self,control):
        s = '''
            <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
              <ribbon startFromScratch="false">
                <tabs>
                  <tab id="CustomTab" label="Sample tab">
                    <group id="MainGroup" label="Sample group">
                      <button id="Button" label="Sample button 1" imageMso="HappyFace" 
                        size="large" onAction="do" />
                        
                      <button id="Button2" label="Sample button 2" getImage='GetImage' 
                        size="large" onAction="do" />

                    </group>

                  </tab>
                </tabs>
              </ribbon>
            </customUI>
        '''
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
    _winreg.SetValueEx(subkey, "Description", 0, _winreg.REG_SZ, "Word Addin")
    _winreg.SetValueEx(subkey, "FriendlyName", 0, _winreg.REG_SZ, "DocCleaner Word Addin")

def UnregisterAddin(klass):
    import _winreg
    try:
        _winreg.DeleteKey(_winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Office\\Word\\Addins\\" + klass._reg_progid_)
    except WindowsError:
        pass

if __name__ == '__main__':
    import win32com.server.register
    win32com.server.register.UseCommandLine( WordAddin )
    if "--unregister" in sys.argv:
        UnregisterAddin( WordAddin )
    else:
        RegisterAddin( WordAddin )

