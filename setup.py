__author__ = 'ashwin'
"""A python module to convert spreadsheets to dictionary and filter output data

"""

from setuptools import setup
from codecs import open
from os import path

projectpath = path.abspath(path.dirname(__file__))

# Get the long description from the relevant file
with open(path.join(projectpath, 'README.rst')) as f:
    long_description = f.read()

setup(
    name='xl2dict',
    version='0.1.1',
    description='Spreadsheet to dictionary converter and data explorer',
    long_description=long_description,
    url='https://github.com/gettalent/xl2dict',
    author='Ashwin Kondapalli',
    author_email='ashwin@gettalent.com',
    license='MIT',
    classifiers=[

        'Development Status :: 5 - Production/Stable',
        'Intended Audience :: Developers',
        'Topic :: Software Development :: Quality Assurance',
        'Programming Language :: Python',
        'License :: OSI Approved :: MIT License',

    ],

    keywords='data driven testing, webdriver, selenium, excel, xls to dict converter, xlsx to dict converter',

    packages=['xl2dict'],

    install_requires=['openpyxl','xlrd']
)
