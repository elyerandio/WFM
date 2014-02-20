from distutils.core import setup
import py2exe

setup(windows=['WFM_Interface.py'],
		options = {"py2exe": {'includes':'decimal'}})
