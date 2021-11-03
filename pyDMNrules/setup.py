import setuptools

with open('README.md', 'r') as fh:
    long_description = fh.read()

setuptools.setup(
    name='pyDMNrules',
    version='1.2.5',
    author='Russell McDonell',
    author_email='russell.mcdonell@c-cost.com',
    description='An implementation of DMN in Python. DMN rules are read from an Excel workbook',
    long_description=long_description,
    long_description_content_type='text/markdown',
    url='https://github.com/russellmcdonell/pyDMNrules',
    packages=setuptools.find_packages(),
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: GNU General Public License v3 (GPLv3)',
        'Operating System :: OS Independent',
    ],
    keywords='DMN,pandas',
    python_requires='>=3.6',
    install_requires=['re', 'csv', 'datetime', 'pySFeel','openpyxl', 'pandas'],
    extra_requires={
        'docs': ['sphinx==4.2.0',
        'sphinx_rtd_theme>0.2.4',
        'recommonmark'],
    }
)

