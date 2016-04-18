# -*- coding: utf-8 -*-
__author__ = 'Po'

from distutils.core import setup
import py2exe
import sys
import rpyc

sys.argv.append('py2exe')

py2exe_options = {
	'includes' : ['sip'],
	'dll_excludes' : ['MSVCP90.dll'],
	'compressed' : 1,
	'optimize' : 0,
	'ascii' : 0,
	'bundle_files' : 1,
}


def build(name, version, entry_py):
	setup(
		name=name,
		version=version,
		windows=[entry_py,],
		zipfile=None,
		options={'py2exe' : py2exe_options}
	)

