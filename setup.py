#!/usr/bin/env python

from distutils.core import setup


setup(name='docx2csv',
      version='1.0',
      description='Table extractor from .docx files',
      author='Ivan Begtin',
      author_email='ibegtin@gmail.com',
      url='https://github.com/ivbeg/docx2csv',
      packages=['docx2csv'],
      scripts=['bin/docx2csv'],
      install_requires=[
          'click',
          'python-docx',
          'xlwt',
      ],
      )
