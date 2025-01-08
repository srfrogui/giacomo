from cx_Freeze import setup, Executable
import sys

# Dependências adicionais podem ser adicionadas aqui, se necessário
build_exe_options = {
    "packages": ["tkinter", "os", "shutil"],  # Adicione pacotes adicionais aqui, se necessário
    "excludes": [],  # Exclua pacotes desnecessários
    "include_files": ["README.md"],  # Inclua arquivos adicionais no executável
}

# Define o ícone do executável
base = None
if sys.platform == "win32":
    base = "Win32GUI"  # Use "Win32GUI" para aplicações GUI

setup(
    name="Arrastador de Banana",
    version="0.1",
    description="Script de backup automatizado",
    options={"build_exe": build_exe_options},
    executables=[Executable("arrasta_banana.py", base=base, icon="banana.ico")]
)
