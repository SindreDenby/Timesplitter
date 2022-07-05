:: Install guide:
:: 1. Edit directories for own computer
:: 2. Run this file (Python 3.9 with Pyinstaller and inno compiler must be installed!)
:: 3. Installer file is in ./Output

set ISCCInstall="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"

py -3.9-64 -m PyInstaller -y -i "adigo icon.png" --distpath ./pyinstoutput/dist --workpath ./pyinstoutput/build --clean -w hub.py

%ISCCInstall% ".\inno setup v 1.iss"