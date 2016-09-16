import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"include_files": ["template.txt"]}

# GUI applications require a different base on Windows (the default is for a
# console application).
base = None
if sys.platform == "win32":
  base = "Win32GUI"

target = Executable(
  script = "app.py",
  base = base,
  compress = True,
  copyDependentFiles = True,
  appendScriptToExe = True,
  appendScriptToLibrary = False,
  icon = "icon.ico")

setup(  name = "FlashDriveEmailer",
  version = "1.0",
  description = "Delta College Flash Drive Emailer",
  author = "Chase Arnold",
  options = {"build_exe": build_exe_options},
  executables = [target])

