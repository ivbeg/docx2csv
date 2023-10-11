import re
import io
from setuptools import setup, find_packages

open_as_utf = lambda x: io.open(x, encoding='utf-8')

(__version__, ) = re.findall("__version__.*\s*=\s*[']([^']+)[']",
                             open('docx2csv/__init__.py').read())

readme = re.sub(r':members:.+|..\sautomodule::.+|:class:|:func:', '', open_as_utf('README.rst').read())
readme = re.sub(r'`Settings`_', '`Settings`', readme)
readme = re.sub(r'`Contributing`_', '`Contributing`', readme)
history = re.sub(r':mod:|:class:|:func:', '', open_as_utf('HISTORY.rst').read())



setup(
    name='docx2csv',
    version=__version__,
    description="Extracts tables from .docx files and saves them as csv or xlsx",
    long_description=readme + '\n\n' + history,
    author='Ivan Begtin',
    author_email='ivan@begtin.tech',
    url='https://github.com/ivbeg/docx2csv',
    packages=find_packages(exclude=('tests', 'tests.*')),
    include_package_data=True,
    install_requires=[
        'click',
		'python-docx>=0.2.4',
		'openpyxl>=3.0.5',
		'xlwt>=1.3.0'
    ],
    python_requires='>=3.8',
    entry_points={
        'console_scripts': [
            'docx2csv = docx2csv.__main__:main',
        ],
    },
    license="BSD",
    zip_safe=False,
    keywords='docx converter tables',
    classifiers=[
        'Development Status :: 2 - Pre-Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: BSD License',
        'Natural Language :: English',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: Implementation :: CPython',
        'Programming Language :: Python :: Implementation :: PyPy'
    ]
)
