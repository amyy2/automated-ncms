# automated-ncms

automated-ncms automatically extracts NCMS information and formats it into a spreadsheet detailing each open NCR. This project aims to semi-automate the process of providing a visual aid of the current status of each NCR by periodically pulling NCMS information and updating the board accordingly.

## Requirements

This project requires Python 3.9 to be installed on the device. This can be downloaded from [python.org](https://www.python.org/downloads/) which will require admin credentials that can be obtained by contacting IT services.

This repository also requires the installation of the following packages and modules:
* [Pandas](https://pandas.pydata.org)
* [Selenium WebDriver](https://www.selenium.dev)
* [Glob](https://docs.python.org/3/library/glob.html)

When setting up this software on a new device, these packages will need to be installed for the first time. To install packages without a certificate (as required when using `pip install`), use the following command in the terminal:

```
py -m pip install <package name>
```
If an error appears reading `certificate verify failed: unable to get local issuer certificate`, this means the package needs another dependency, which will be specified at the top of the error. Use the [Python Package Index](https://pypi.org) to search for the required dependency and download the most recent .whl file for Python 3.9 and Windows 64-bit. Navigate to the Downloads folder and copy the name of the file. Paste this into the above command as the package name.

## Usage

## Troubleshooting