from cx_Freeze import setup, Executable
import sys

# Inclua os pacotes e arquivos necessários
build_exe_options = {
    "packages": [
        "pandas", "reportlab", "tkinter", "xlrd", "os", "shutil", 
        "tkinterdnd2", "tkcalendar", "multiprocessing", "PyPDF2",
        "unicodedata", "collections", "xml", "re", "logging",
        "pyautogui", "threading", "subprocess", "pytesseract",
        "pyperclip", "xlrd", "threading", "PIL", "pyautogui", "PyPDF2", "win32gui"
    ],
    "includes": [
        "GAuto", "PromobAuto", "embananador", "G2Auto", "Moveu", "contar_chapas"
    ],
    "include_files": [
        "img/", "2021646.ico", "tesseract-ocr-w64-setup-5.5.0.20241111.exe"
    ],
    "include_msvcr": True,  # Inclua o runtime do Visual C++ se necessário
}

# Define o tipo de executável (console ou GUI)
base = None
if sys.platform == "win32":
    base = "Win32GUI"  # "Win32GUI" para GUI, "None" para console

# Configuração do setup
setup(
    name="FaztudoBotzin",
    version="0.1",
    description="BotdQuatro",
    options={"build_exe": build_exe_options},
    executables=[Executable("Trio.py", base=base, icon="2021646.ico")]
)
