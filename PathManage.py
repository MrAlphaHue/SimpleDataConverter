import os;
from sys import platform


xlsx_path = os.getcwd() + "/Input_Xlsxs"
binary_path = os.getcwd() + "/Output_Binaries"
script_path = os.getcwd() + "/Output_Scripts"

##WINDOW
if platform == "win32":
    project_path = ""
##darwin means 'Mac'
elif platform == "darwin":
    project_path = ""
else:
    project_path = "SomethingWrong : no OS"

project_binary_path = project_path + "/Assets/Resources/BinaryData"
project_script_path = project_path + "/Assets/01_Scripts/Utils/TableManage"