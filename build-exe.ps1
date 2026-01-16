pip install pyinstaller openpyxl

pyinstaller --onefile --windowed --add-data "dfir-installer-selector.xlsx;." --add-data "gui_template.xml;." -n "dfir-installer-selector.exe" main.py