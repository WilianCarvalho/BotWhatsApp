from cx_Freeze import setup, Executable
setup(
    name = "WhatsBot",
    version = "1.5.0",
    options = {"build_exe": {
        'packages': ["numpy","selenium","pyodbc","pandas","time","urllib","datetime","PySimpleGUI","cx_Freeze","openpyxl"],
        'include_msvcr': True,
        'include_files': ['instrucao.txt'],
    }},
    executables = [Executable("whatsappbot.py",base="Win32GUI")]
    )