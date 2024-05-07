from tkinter import *
from tkinter import ttk
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
    log.insert(INSERT,'Inicio da conversão\n')

    for arquivo in arquivos:          
        
        nome_arquivo = os.path.splitext(arquivo)[0]
        nome_arquivo = nome_arquivo.split('/')
        nome_arquivo = nome_arquivo[-1]
        log.insert(INSERT,'Arquivo:\n')
        log.insert(INSERT,nome_arquivo+'\n')

        with open(arquivo, 'r') as arq:            
    
            workbook.create_sheet('Conversao')
            newSheet = workbook['Conversao']
            # Cabeçalho
            newSheet.append(['RG', 'MATR.', 'AG.', 'TP.', 'CONTRATO', 'ET.', 'DT. CONTRATO', 'DT. PREST. IN.', 
                            'DT. ATER.', 'VALOR BASE DFI', 'SALDO DEVEDOR', 'DFI', 'MIP', 'ATRS DFI', 'ATRS MIP', 'ST']) 
            i = 1            
            log.insert(INSERT,'#'*30+'\n')
            log.insert(INSERT,'Lendo arquivo:\n')
            log.insert(INSERT,arquivo)           
            log.insert(INSERT,'\n') 
            for linha in arq:
                log.insert(INSERT,f'Linha {i}\n')                
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
            log.insert(INSERT,'Fim da leitura\n')            
            log.insert(INSERT,'Gravando planilha\n')            
            workbook.save(f'{nome_arquivo}.xlsx')                        
            del workbook['Conversao']
            log.insert(INSERT,'Planilha gravada\n')
        j +=1
    log.insert(END,'Fim da Execução!\n')


################### Execução ####################
if __name__ == '__main__':
    root = Tk()
    root.title('Gerador v1.0')    
    root.geometry("400x300")        
    
    log = Text(root, height = 14, width = 50) 
    btt_gerar = Button(root, text="Gerar Excel", command=gerar_planilhas)
    label = Label(root, text ="Gerar planilha:", fg ="red")    
    
    label.pack() 
    log.pack()    
    btt_gerar.pack()

    
    # frm = ttk.Frame(root, padding=10)
    # frm.grid()
    # ttk.Label(frm, text="Gerar Planilhas").grid(column=0, row=0)
    # # ttk.Label(frm, text="").grid(column=0, row=1)
    # ttk.Button(frm, text="Gerar Excel", command=gerar_planilhas).grid(column=0, row=2)
    # # ttk.Label(frm, text="  ").grid(column=1, row=2)
    # log = Text(root, height = 6, width = 53).grid(column=0, row=3)     
    # ttk.Button(frm, text="Quit", command=root.destroy).grid(column=0, row=4)    
    root.mainloop()


    
        
################### Execução ####################                

