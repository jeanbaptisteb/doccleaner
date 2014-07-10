from setuptools import setup, find_packages


setup(name='doccleaner',
      version='0.1.2',
      description='A python command-line utility to edit zipped, XML-based files (e.g. docx, odt, or epub). Can be rather easily extended with xsl stylesheets. Intended for automating some copyediting tasks',
      url='',
      download_url='',
      author='Jean-Baptiste Bertrand',
      author_email='jean-baptiste.bertrand@openedition.org',
      license='LGPL 3.0',
	  include_package_data=True,
      packages=find_packages(),
      install_requires=['defusedxml'],	
      zip_safe=False)
      

