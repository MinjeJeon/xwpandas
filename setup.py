from setuptools import setup, find_packages

setup(
    name='xwpandas',
    version='0.1.0',
    url='https://github.com/Minjejeon/xwpandas',
    license='BSD 3-clause',
    author='Minje Jeon',
    author_email='i@minje.kr',
    classifiers=[
        'Development Status :: 3 - Alpha',
        'License :: OSI Approved :: BSD License',
        'Operating System :: MacOS :: MacOS X',
        'Operating System :: Microsoft :: Windows',
        'Programming Language :: Python :: 3 :: Only',
        'Topic :: Office/Business :: Financial :: Spreadsheet'
    ],
    keywords=['excel', 'pandas', 'DataFrame', 'xls', 'xlsx'],
    description='High performance Excel IO tools for DataFrame',
    packages=find_packages(exclude=['tests']),
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    setup_requires=['pandas>=0.23.0',
                    'xlwings>=0.11.0']
)
