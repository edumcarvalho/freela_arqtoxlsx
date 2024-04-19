import sys
from cx_Freeze import setup, Executable

base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

# Saida de arquivos
configuracao = Executable(
    script='app1.py',
    icon='robo.ico',
    base=base
)
# Configurar o executável
setup(
    name='Automatizador de login',
    version='1.0',
    description='Este programa automatizar o login deste site',
    author='Jhonatan de Souza',
    options={'build_exe':{    
        'include_msvcr': True # para windows!!
    }},
    executables=[configuracao]
)

#ABRIR O TERMINAL E DIGITAR
# python3 setup.py build