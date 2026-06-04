2 работает перзаписало 
4 good!!!

pyinstaller --onefile --windowed --name "KitchenEditor" --distpath "D:\KitchenEditor_build" --add-data "gui.py;." --add-data "autocad_handler.py;." --hidden-import=win32com --hidden-import=win32com.client --hidden-import=pythoncom --hidden-import=gui --hidden-import=autocad_handler main.py


last
генерация exe.file запустить в директории с main:
pyinstaller kitchen_editor.spec --clean
