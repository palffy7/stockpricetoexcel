#!C:\Users\ralph\PycharmProjects\stockpricestoexcel\venv\Scripts\python.exe
# EASY-INSTALL-ENTRY-SCRIPT: 'pyinstaller==4.2','console_scripts','pyi-set_version'
__requires__ = 'pyinstaller==4.2'
import re
import sys
from pkg_resources import load_entry_point

if __name__ == '__main__':
    sys.argv[0] = re.sub(r'(-script\.pyw?|\.exe)?$', '', sys.argv[0])
    sys.exit(
        load_entry_point('pyinstaller==4.2', 'console_scripts', 'pyi-set_version')()
    )
