from setuptools import setup, find_packages
from os import path

here = path.abspath(path.dirname(__file__))
with open(path.join(here, 'README.md'),encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='pdSharepoint',
    version='1.0.0',
    packages=find_packages(),
    url='https://github.com/danielsioli/pdSharepoint',
    author='Daniel Oliveira',
    author_email='danielsioli@gmail.com',
    license='Creative Commons Attribution-Noncommercial-Share Alike license',
    description='Pacote para obter arquivos excel ou listas do Sharepoint em um DataFrame pandas',
    long_description=long_description,
    long_description_content_type="text/markdown",
    keywords='Sharepoint,pandas',
    install_requires=['O365', 'pandas', 'openpyxl', 'xlrd'],
    include_package_data=True,
    classifiers=[
        # How mature is this project? Common values are
        #   3 - Alpha
        #   4 - Beta
        #   5 - Production/Stable
        'Development Status :: 4 - Beta',

        # Indicate who your project is intended for
        'Intended Audience :: Developers',
        'Topic :: Internet :: WWW/HTTP',

        # Specify the Python versions you support here. In particular, ensure
        # that you indicate whether you support Python 2, Python 3 or both.
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
    ],
    python_requires='>=3.6'
)