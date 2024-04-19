import PySimpleGUI as sg
import openpyxl
from time import sleep


########### Configuração da Planilha ############
workbook = openpyxl.Workbook()
del workbook['Sheet']
workbook.create_sheet('Conversao')
newSheet = workbook['Conversao']
# Cabeçalho
newSheet.append(['RG', 'MATR.', 'AG.', 'TP.', 'CONTRATO', 'ET.', 'DT. CONTRATO', 'DT. PREST. IN.', 
                 'DT. ATER.', 'VALOR BASE DFI', 'SALDO DEVEDOR', 'DFI', 'MIP', 'ATRS DFI', 'ATRS MIP', 'ST']) 
########### Configuração da Planilha ############
#################################################
############ Configuração do Layout #############
sg.theme('Reddit')

layout = [    
    [sg.FileBrowse('Selecionar arquivo', target='caminho_arquivo', file_types=(
        ('Arquivos ARQ', '*.ARQ'), ('Arquivos de texto', '*.txt'), ('Todos arquivos', '*')))],
    [sg.Input(key='caminho_arquivo')],    
]

layout_output = [
    [sg.Output(size=(45, 10))]
]

layout_principal = [
    [sg.Frame('Conversor de arquivo', layout)],
    [sg.Frame('Log de atividades', layout_output)],
    [sg.Button('Converter arquivo')]
]

window = sg.Window('Automação de sites', layout=layout_principal)
############ Configuração do Layout #############
#################################################
################### Execução ####################
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    elif event == 'Converter arquivo':
        with open(values['caminho_arquivo'], 'r') as arquivo:
            i = 1
            print('Inicio da conversão')            
            print('#'*30)
            print('Lendo arquivo')
            for linha in arquivo:
                print(f'Linha {i}')                
                rg        = linha[0:2]
                matricula = linha[2:7]
                ag        = linha[7:9]
                tp        = linha[9:10]
                contrato  = linha[10:22]
                et        = linha[22:24]
                dt_contra = linha[24:26]+'/'+linha[26:28]+'/'+linha[28:32]
                dt_presti = linha[32:34]+'/'+linha[34:36]+'/'+linha[36:40]
                dt_altera = linha[40:42]+'/'+linha[42:44]+'/'+linha[44:48]                
                vlbasedif = float(linha[83:95]+'.'+linha[95:97])
                saldo_dev = float(linha[111:123]+'.'+linha[123:125])
                dfi       = float(linha[125:134]+'.'+linha[134:136])
                mip       = float(linha[136:145]+'.'+linha[145:147])
                atrs_dfi  = float(linha[158:169]+'.'+linha[169:171])
                atrs_mip  = float(linha[171:182]+'.'+linha[182:184])
                st        = linha[609:610]               
                newSheet.append([rg,matricula,ag,tp,contrato,et,dt_contra,dt_presti,dt_altera,vlbasedif,saldo_dev,dfi,mip,atrs_dfi,atrs_mip,st])
                i += 1
            print('Fim da leitura')
            print('#'*30)
            print('Gravando planilha')
            workbook.save('arquivo.xlsx')
            print('Planilha gravada')
            print('#'*30)
            print('Fim da Execução!')
# ################### Execução ####################                



