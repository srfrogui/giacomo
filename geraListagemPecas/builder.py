from cx_Freeze import setup, Executable
import sys

# Inclua os pacotes que seu script precisa
build_exe_options = {
    "packages": ["pandas", "reportlab", "tkinter", "xlrd", "os"],  # Inclui pacotes necessários
    "excludes": [],  # Exclua pacotes desnecessários
    "include_files": [("arial.ttf", "arial.ttf")],  # Inclua arquivos adicionais se necessário
    "include_msvcr": True,  # Inclua runtime C++ se necessário
}

# Define o ícone do executável
base = None
if sys.platform == "win32":
    base = "Win32GUI"  # Use "Win32GUI" para aplicações GUI

setup(
    name="Embananador",
    version="0.1",
    description="Criar Relatorio de Pecas",
    options={"build_exe": build_exe_options},
    executables=[Executable("embananador.py", base=base, icon="banana.ico")]
)
