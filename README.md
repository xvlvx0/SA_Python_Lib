# SA Python library

This Python library interacts with the Spatial Analyzer SDK DLL and handles all communications, allowing for true python code scripting capabilities for Spatial Analyzer.
It uses the Visual Studio Python.NET (pythonnet) package as communication layer to the DLL. The 'pythonnet' package can be installed with pip.
Use the pip package manager on the command line console and execute the following command:

    pip install pythonnet

# SA Python Demo

The demo folder holds several 'basic' examples for kickstarting your scripts.

# History

This project is inspired on the work of user oshea00. He made the first version of this library, back then in python 2.7. The library interaction was done via the 'IronPython' package which wasn't developed to support Python V3.x. Therefore, I made the switch to 'pythonnet'.

