# SA Python library

## Introduction

This Python library interacts with the Spatial Analyzer SDK DLL and handles all communications, allowing for true python code scripting capabilities for Spatial Analyzer.
It uses the Visual Studio Python.NET (pythonnet) package as the communication layer to the DLL.

## Dependencies

The SA Python library depends on the 'pythonnet' package, it can be installed with pip. Use the pip package manager on the command line and execute the following command:

    pip install pythonnet

## Installation of the library

Clone or download the repository into "C:\Analyzer Data\Scripts\" and use the folder name "SAPython", such that the final path looks like: "C:\Analyzer Data\Scripts\SAPython\".

If you want to deviate from this then make sure you update the line below in the file "lib/SAPLib.py":

    basepath = r"C:\Analyzer Data\Scripts\SAPython" 

You also need to update each script such that the following line in the import section points to the new path:

    sys.path.append("C:/Analyzer Data/Scripts/SAPython/lib")

## SA Python examples

In the examples folder are several 'basic' examples for kick starting your scripts.

## History

This project is inspired on the work of user oshea00. He made the first version of this library, back then in python 2.7. The library interaction was done via the 'IronPython' package which wasn't developed to support Python V3.x. Therefore, I made the switch to 'pythonnet'.
