""" Install Coolprop Excel wrapper on Mac OS Machines """

import requests
import pathlib
import shutil
import zipfile

BASEURL = 'https://raw.githubusercontent.com/CAPP-TESTS/coolprop-mac-excel/refs/heads'

ADDIN = pathlib.Path('~/Library/Group Containers/UBF8T346G9.Office/').expanduser()
DYLIB = ADDIN / 'User Content.localized/Add-Ins.localized/'
DESKTOP = pathlib.Path('~/Desktop').expanduser()

def download(stem, filename, destination):
    print('Downloading', filename)
    r = requests.get(BASEURL + stem.format(filename))
    r.raise_for_status()
    with open(destination / filename, 'wb') as f:
        f.write(r.content)

print('Installing CoolProp 32/64bit for MacOS')
download('/main/{}', 'CoolProp_RST.xlam', DYLIB)
download('/main/{}', 'libCoolProp_arm_64.dylib', DYLIB)
download('/main/{}', 'libCoolProp_x86_64.dylib', DYLIB)
download('/main/{}', 'libCoolProp_x86_32.dylib', DYLIB)
download('/main/{}', 'Excel.with.CoolProp.app.zip', DESKTOP)

print('Add-in location:', DYLIB)
