from cx_Freeze import setup, Executable
import sys

# Dependências adicionais podem ser adicionadas aqui, se necessário
build_exe_options = {
    "packages": ["tkinter", "os", "shutil", "reportlab", "pandas", "PIL", "win32com.client", "glob", ],
    "excludes": [],  # Exclua pacotes desnecessários
    "include_files": [],
    "optimize": 2,
}

# Define o ícone do executável
base = None
if sys.platform == "win32":
    base = "Win32GUI"  # Use "Win32GUI" para aplicações GUI

setup(
    name="Relatorizador de peça faltante",
    version="0.1",
    description="Script de backup automatizado",
    options={"build_exe": build_exe_options},
    executables=[Executable("GERARELATORIO_FALTANTES.py", base=base, icon="banana.ico")]
)
