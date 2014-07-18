# -*- coding: utf-8 -*-
"""
Created on Tue Apr 29 10:37:33 2014

@author: Bertrand
"""
import gettext
import locale
import os

def init_localization():
    '''prepare l10n'''
    print(locale.setlocale(locale.LC_ALL,""))
    locale.setlocale(locale.LC_ALL, '') # use user's preferred locale

    # take first two characters of country code
    loc = locale.getlocale()
    
    #filename = os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), "lang", "messages_%s.mo" % locale.getlocale()[0][0:2])
    filename = os.path.join("lang", "messages_{0}.mo").format(locale.getlocale()[0][0:2])
    try:
        print("Opening message file {0} for locale {1}".format(filename, loc[0]))
        #If the .mo file is badly generated, this line will return an error message: "LookupError: unknown encoding: CHARSET"
        trans = gettext.GNUTranslations(open(filename, "rb"))

    except IOError:
        print("Locale not found. Using default messages")
        trans = gettext.NullTranslations()

    trans.install()
    
class Addin():

    def WordIniGen(self):
        #function to generate the conf/ini file for the word Addin
        cdf = '''
[cleanDirectFormatting]
label={0} 
imageMso=HappyFace
size=large
onAction=do
subFile=""
XSLparameter={2}
        '''.format(
                _("Nettoyer les mises en forme locales"), #0 = label macro
                (""),                                     #1 = subFiles to process
                ("bla='underline'"),                      #2 = XSL parameter(s) to use. If > 1, separate by a comma
                ) 
                
                
        dummy = '''
[dummy]
label={0} 
imageMso=HappyFace
size=large
onAction=do
subFilesToProcess=""
XSLparameter=""
        '''.format(
                _("Dummy"), #label macro 0
                                                
                                        
                ) 
        wordConf = cdf + dummy
        return wordConf

def main():
    try:
        addin = Addin()
        f = open(os.path.join(os.path.dirname(os.path.realpath(__file__)), 
                            "wordAddin_"+ locale.getlocale()[0][0:2]+ ".ini")
                , 'w')
                
        f.write(addin.WordIniGen())
        
    except Exception as e:
        print(str(e))
    

if __name__ == '__main__':
    init_localization()
    main()