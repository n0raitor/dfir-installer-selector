pip install pyinstaller openpyxl

pyinstaller --onefile --windowed --add-data "dfir-installer-selector.xlsx;." --add-data "gui_template.xml;." main.py
