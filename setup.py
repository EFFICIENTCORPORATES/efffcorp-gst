import pathlib
from setuptools import setup


#the Directory Contatining ths file

HERE=pathlib.Path(__file__).parent

#The Text of the Readme FIle

README=(HERE/"README.md").read_text()


#this call to setup() does all the work


setup(
	name = 'effcorp-gst',
	version = '1.0.0',
	py_modules = ['effcorp-gst'],
	packages=["gst"],
	include_package_data=True,
	author = 'efficient_corporates',
	author_email = 'efficientcorporates.info@gmail.com',
	install_requires=['pandas','numpy','openpyxl'],
	url = 'https://github.com/EFFICIENTCORPORATES/efffcorp-gst',
	description = 'A python module to merge GSTr2A and also reconcile the GSTR2A vs the Purchase Register',
	long_description=README,
	long_description_content_type="text/markdown",
	license="GNU GP License",
	classifiers=[
        "License :: OSI Approved :: GNU General Public License v3 or later (GPLv3+)",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
    ],
    entry_points={"console_scripts":["gst=gst.__main__:main",]},)