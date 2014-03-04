# -*- coding: utf-8 -*-

import sys
python_version = sys.version_info[:2]

if python_version < (3, 3):
    raise Exception("excelpy requires Python 3.3 or above.")

try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

from src.excelpy import __version__
setup(
    name='excelpy',
    version=__version__,
    packages=['excelpy', 'templates'],
    package_dir={'excelpy': 'src/excelpy', 'templates': 'src/excelpy/templates'},
    package_data={'excelpy': ['templates/xl/worksheets/*']},
    install_requires=['lxml'],
    license='MIT License',
    author='dokenzy',
    author_email='dokenzy@gmail.com',
    url='https://github.com/dokenzy/excelpy',
    description='Minimal Microsoft Excel 2010 library for Python 3.3. ',
    long_description='Excelpy can add new sheets, copy sheets, delete sheets, and edit string and number type datas.',
    keywords=['xlsx', 'excel', 'spreadsheet'],
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
        'Operating System :: OS Independent',
        'Topic :: Database',
        'Topic :: Office/Business',
        'Topic :: Software Development :: Libraries :: Python Modules',
    ],
)
