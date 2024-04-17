import sys
from cx_Freeze import setup, Executable

# definir quais dependências meu projeto possui
build_exe_options = {'packages': ['os'],'includes': ['tkinter']}

base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

configuracao = Executable(
    script='app.py',
    icon='robo.ico',
    base=base
)

setup(
    name='Conversor',
    version='1.0',
    description='Conversor de planilha',
    options={'build_exe':build_exe_options},
    executables=[configuracao]
)