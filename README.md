# automated-ncms
automated-ncms automatically extracts NCMS information and formats it into a spreadsheet detailing each open NCR. This project aims to semi-automate the process of providing a visual aid of the current status of each NCR by periodically pulling NCMS information and updating the board accordingly.
 
## Table of Contents
1. [Requirements and Setup](#setup)
2. [Usage](#usage)
3. [Troubleshooting](#troubleshooting)
 
## Requirements and Setup <a name="setup"></a>
 
This project requires Python 3.9 and VS Code to be installed on the device. Python can be downloaded [here](https://www.python.org/downloads/) and will require admin credentials that can be obtained by contacting IT services. VS Code, which is a code editor that will be used to run the code, can be downloaded [here](https://code.visualstudio.com/) and will not require any credentials.
 
This repository also requires the installation of the following packages and modules:
* [Pandas](https://pandas.pydata.org)
* [Selenium WebDriver](https://www.selenium.dev)
* [Glob](https://docs.python.org/3/library/glob.html)
* [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
* [pywin32](https://pypi.org/project/pywin32/) (imported as win32com)
 
When setting up this software on a new device, these packages will need to be installed for the first time. To install packages without a certificate (as required when using `pip install`), use the [Python Package Index](https://pypi.org) to search for the package and download the most recent .whl file for Python 3.9 and Windows 64-bit. Navigate to the Downloads folder and copy the name of the file. Paste this into the following command in the terminal as the package name:
 
```
py -m pip install <package name>
```
If an error appears reading `certificate verify failed: unable to get local issuer certificate`, this means the package needs another dependency, which will be specified at the top of the error. For example, in the command prompt window below when attempting to install `Tensorflow` it's looking for `astunparse`, which needs to be installed before `Tensorflow`. Repeat the previous steps for each required dependency.
 
![example](cmd.png)
 
*NCMS excel file
*editing path
 
## Usage <a name="usage"></a>
 
To stop the software, navigate to the terminal in VS Code and type `Ctrl+C` and close the terminal.
 
## Troubleshooting <a name="troubleshooting"></a>