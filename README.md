# Managemet_sys_PyQt5_openpyxl

## Usage:

GuI based application which is built to manage the data of a company on various grounds like Management systems (Storage, Students, Employees Anything).

Pyinstaller can be used to make an executable file of the application and then can be used on any system without any Python or it's supporting library.

## Requirements:

1. Python 3 or later
2. openpyxl installed  " pip3 install openpyxl " command for cmd to install library
3. PyQt5 installed " pip3 install PyQt5 " command for cmd to install library

Install  [Python](https://www.python.org/downloads/) . Do install Python3 or later .

if facing difficulty in installing libraries here is the link for the guide:

1. [openpyxl](https://pypi.python.org/pypi/openpyxl)

2. [PyQt5](https://pypi.python.org/pypi/PyQt5)

> Do change the Image Urls given in the code for a Window Icon and a Background Image .

## Features:

- The excel sheets will be created in C:\Management_sys_excels folder which will be created automatically once this code runs.
- The excel will automatically created if the name mentioned in the Gui not exisists.
- The excel would get updated on same sheet with new entries if excel with same name already exists.
- The top three rows of details are necessary fields so if not mentioned correct pops will come up accordingly.
- With a single click on Clear button all the entries would get Initialised.
- The sno. and token no. gets updated automatically but if in excel updated with new value will get updated to according to this new value.
- With a single click on save Button entries would get saved.
- If excel is already open a popup alert will come up as excel cant be saved once it's open.

## Run:

```
  python Management_sys.py
```
