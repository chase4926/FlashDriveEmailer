
from distutils.core import setup
import py2exe, sys, os, requests.certs

sys.argv.append('py2exe')

setup(windows=[{'script': 'app.py',
                'icon_resources': [(1, "icon.ico")]}],
            options={"py2exe": {"includes": ["time", "ssl", "random",
                                             "re", "smtplib",
                                             "webbrowser", "requests",
                                             "lxml", "gzip",
                                             "email.mime.text",
                                             "Tkinter", "ttk",
                                             "requests", "ssl"],
                     "packages": ["gzip", "lxml", "ssl"],
                     "dll_excludes": ["w9xpopen.exe"],
                     "bundle_files": 3,
                     "compressed": False}},
            zipfile = None)
