from setuptools import setup


setup(name='doccleaner',
      version='0.1.0',
      description='A python command-line utility to edit zipped, XML-based files (e.g. docx, odt, or epub). Can be rather easily extended with xsl stylesheets. Intended for automating some copyediting tasks',
      url='https://github.com/jbber/DocCleaner',
      download_url='https://github.com/jbber/DocCleaner/tarball/0.1',
      author='Jean-Baptiste Bertrand',
      author_email='jean-baptiste.bertrand@openedition.org',
      license='LICENSE',
	  include_package_data=True,
      packages=['doccleaner'],
      install_requires=['defusedxml'],	
      package_data = {
      'doccleaner': ['docx/*.*', 'lang/*.*'],
      },
      keywords = ['xsl', 'docx'],
      zip_safe=False)
      

