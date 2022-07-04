:: Install guide:
:: 1. Run this file (Pyinstaller and python 3.9 must be installed)
:: 2. Compile the inno setup file (inno setup compiler is required)
:: 3. Installer file is in ./Output

py -3.9-64 -m PyInstaller --distpath ./pyinstoutput/dist --workpath ./pyinstoutput/build --clean -w hub.py