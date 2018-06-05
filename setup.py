from cx_Freeze import setup, Executable
import os.path
import sys

# Dependencies are automatically detected, but it might need
# fine tuning.

# Build using python setup.py build

PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')

#added_files = [os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
#               os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'),
#               ]

buildOptions = dict(packages=['numpy'], excludes=[], build_exe='Sleep Numbers')

base = 'Win32GUI' if sys.platform == 'win32' else None

executables = [
    Executable('sleep_numbers.py', base=base, icon='python_icon.ico')
]

setup(name='Sleep Numbers',
      version='0.1',
      description='Calculates sleep study numbers',
      options=dict(build_exe=buildOptions),
      executables=executables)
