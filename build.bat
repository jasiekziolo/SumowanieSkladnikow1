@echo off
py -m pip install --user openpyxl pyinstaller
pyinstaller --noconfirm --clean --onefile --windowed --name SumowanieSkladnikow sumowanie_skladnikow.py
pause
