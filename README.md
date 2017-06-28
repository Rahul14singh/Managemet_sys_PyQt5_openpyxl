# Managemet_sys_PyQt5_openpyxl

## Usage:

GUI based application which is built to manage the data of a company on various grounds like Management Systems (Storage, Students, Employees Anything).

Pyinstaller can be used to make an executable file of the application and then can be used on any system without any Python or it's supporting libraries.

## Requirements:

1. Python 3 or later
2. openpyxl installed  " pip3 install openpyxl " command on cmd to install library
3. PyQt5 installed " pip3 install PyQt5 " command on cmd to install library
4. Some other necessary supporting libraries.

Install  [Python](https://www.python.org/downloads/) . Do install Python3 or later.

if facing difficulty in installing libraries here is the link for the HELP:

1. [openpyxl](https://pypi.python.org/pypi/openpyxl)

2. [PyQt5](https://pypi.python.org/pypi/PyQt5)

> Do change the Image URLs given in the code for a Window Icon and a Background Image for the GUI application.

## Features:

- The excel sheets will be created and saved in C:\Management_sys_excels folder which will be created automatically once this code runs.
- The Excel will be automatically created if the name mentioned in the GUI, not exists which could be changed in Gui.
- The Excel would get updated on the same sheet with new entries if excel with the same name already exists and details saved without changing the name of the excel in Gui.
- The top three rows of details are necessary fields so if not mentioned correct popups will come up accordingly.
- The validity of all these top three fields would be checked. 
- With a single click on the Clear button, all the entries would get Initialised.
- The Sno. and token no. gets updated automatically but if in excel updated with new value will get updated to according to this new value.
- With a single click on Save Button entries would get saved.
- If Excel is already open a popup alert will come up as Excel can't be saved once it's open.

## Run:

```
  python Management_sys.py
```
