from PyInstaller.utils.hooks import collect_all

# Collect all things

datas, binaries, hiddenimports = collect_all("openpyxl")