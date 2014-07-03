import os
from setuptools import setup

setup(name='doccleaner',
      version='0.1.0',
      description='A python command-line utility to edit zipped, XML-based files (e.g. docx, odt, or epub). Can be rather easily extended with xsl stylesheets. Intended for automating some copyediting tasks',
      url='http://github.com/jeanbaptisteb/DocCleaner',
      author='Jean-Baptiste Bertrand',
      author_email='jean-baptiste.bertrand@openedition.org',
      license='LICENSE',
      packages=['doccleaner'],
      install_requires=['defusedxml'],
      zip_safe=False)
      

