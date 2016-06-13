# General

This code is just a wrapper around the xlrd library for reading excel
files. I was using xlrd recently and I found the api way too verbose,
so this is my solution. Pyxl allows indexing into excel files using
Python list and dictionary index notation. Check below for examples.

# Running

The only thing you should need from pyxl.py is the Pyxl class. That
exposes all the nice indexing capabilities.

    from pyxl import Pyxl

    # To load an excel file into a Pyxl object
    xlfile = Pyxl("TestFile.xlsx")

    # To access an individual sheet in the excel file
    sheet1 = xlfile[0]

    # To access an entire row in the sheet
    print sheet1[0]

    # To access an entire column by name
    print sheet1["Name"]

    # To access an individual cell
    print sheet1[0, 2]