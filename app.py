import openpyxl
from pathlib import Path
import os

def gerar_planilhas():  
    workbook = openpyxl.Workbook()    
    del workbook['Sheet']
    
    path = Path(os.getcwd())        
    arquivos = path.glob('*.ARQ')
    ################ Lendo arquivos #################
    j = 1              
    print('Inicio da conversão')

    for arquivo in arquivos:          
        
        nome_arquivo = os.path.splitext(arquivo)[0]
        nome_arquivo = nome_arquivo.split('/')
        nome_arquivo = nome_arquivo[-1]
        print('nome_arquivo')
        print(nome_arquivo)


        with open(arquivo, 'r') as arq:            
    
            workbook.create_sheet('Conversao')
            newSheet = workbook['Conversao']
            # Cabeçalho
            newSheet.append(['RG', 'MATR.', 'AG.', 'TP.', 'CONTRATO', 'ET.', 'DT. CONTRATO', 'DT. PREST. IN.', 
                            'DT. ATER.', 'VALOR BASE DFI', 'SALDO DEVEDOR', 'DFI', 'MIP', 'ATRS DFI', 'ATRS MIP', 'ST']) 
            i = 1            
            print('#'*30)
            print('Lendo arquivo:')
            print(arquivo)            
            for linha in arq:
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
            print('Gravando planilha')            
            workbook.save(f'{nome_arquivo}.xlsx')                        
            del workbook['Conversao']
            print('Planilha gravada')                        
        j +=1
    print('Fim da Execução!')

################### Execução ####################
if __name__ == '__main__':
    gerar_planilhas()


    
        
################### Execução ####################                