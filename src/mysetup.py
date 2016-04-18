# -*- coding: utf-8 -*-

from distutils.core import setup
import py2exe
import sys
import os


#setup(windows=["E:\\src\\mainframe.py"],options = { "py2exe":{"dll_excludes":["MSVCP90.dll"]}})

includes = ["encodings", "encodings.*","xlwt"]

options = {"py2exe":
            {"compressed": 1,
             "optimize": 2,
             "ascii": 1,
             "includes":includes,
             "bundle_files": 1
            }}
setup(
    #options=options,
    zipfile=None,
    #console=[{"script": "HelloPy2exe.py", "icon_resources": [(1, "pc.ico")]}],
    windows=[{"script": "mainframe.py", "icon_resources": [(1, "pc.ico")]}],
    options = { "py2exe":{"dll_excludes":["MSVCP90.dll"]}}

)
