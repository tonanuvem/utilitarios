#from openpyxl import load_workbook
#import textract
import docx2txt
import os
import re

# Leitura de arquivos WORD DOCX e escrita no formato CSV para importar no CARNE LEAO (IRPF)

# Directory for files
file_dir = 'C:\\Users\\apsampaio\\Documents\\Python Scripts\\word\\2019'
#file_dir = 'C:\\Users\\Andre\\Dropbox\\Andre\\IR\\2018 ano base\\Charliana\\recibosreferenteaosatendimentosde2018' # "."
# Caracter separador de colunas
delimitador = "\n-----------------------------------\n"
CPFDRA = ''
ANO = '2019'
file_name = CPFDRA+'-'+ANO+'.txt'

#filtro = {"Recebemos do (a) Sr.(a)"; "Portadora do CPF.:"; "Endereço: "; "Quantia de:"}
filtro = ['Data,', 'R$', 'do CPF.:', 'Recebemos']
campos_com_base_nos_filtros = {}
# campos = {'data': '05/01/'+ANO, 'conta': '', 'valor': '0,00', 'CPFPAGOU': '', 'CPFUSOU': '', 'Historico': ''}
# data = DD/MM/AAAA ou DD/MM/AA
# contas = 1000 (recebimentos) ; 4001 (aluguel) ; 4017 (propaganda)
# CPF do paciente
# CPF da Dra
# Historico = Campo texto com ate 250 chars = usar Nome do Paciente ; GOOGLE ADS ; ALUGUEL

with open(file_name, 'w') as fileout:
    # logs
    with open('log-'+file_name, 'w') as logfileout:
        # criar os gastos mensais com propaganda e aluguel
        for mes in range(1, 13):
            campos = {'data': '05/01/'+ANO, 'conta': '', 'valor': '0,00', 'CPFPAGOU': '', 'CPFUSOU': '', 'Historico': ''}
            for conta in ['4001', '4017']:
                campos['conta'] = conta
                if mes < 10:
                    campos['data'] = '05/0'+str(mes)+'/'+ANO
                else:
                    campos['data'] = '05/'+str(mes)+'/'+ANO
                if conta == '4001':
                    campos['valor'] ='1000'
                    campos['Historico'] ='GASTO COM ALUGUEL'
                else : #  conta == '4017':
                    campos['valor'] ='350'
                    campos['Historico'] =' GASTO COM GOOGLE ADS'
                campos['CPFPAGOU'] = CPFDRA
                # campos para gravar no arquivo
                fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
        # criar os recebimentos dos pacientes
        total = 0
        erro = 0
        # Get all word files
        for f in os.listdir(file_dir):
            campos = {'data': '05/01/'+ANO, 'conta': '', 'valor': '0,00', 'CPFPAGOU': '', 'CPFUSOU': '', 'Historico': ''}
            if f.endswith(".docx"):
                total += 1
                path = os.path.join(file_dir, f)
                print("Parsing file: {}".format(path))
                # Get file name without extension
                ff = os.path.splitext(f)[0]
                logfileout.write( ff + delimitador )
                text = docx2txt.process(path)
                # Filtrar
                for line in text.splitlines():
                    for word in filtro:
                        if word in line: #palavras:
                            #print('string contains a word from the word list: %s' % (word))
                            #print(line)
                            logfileout.write('\t' + line.lstrip().rstrip() + '\n' )
                            campos_com_base_nos_filtros[word]=line.lstrip().rstrip()
                # write to file
                # escrevendo os campos
                # pega os valores do dict: campos_com_base_nos_filtros[word]
                campos['data'] = campos_com_base_nos_filtros[filtro[0]].replace('Data, ','').replace('.','') # filtro = 'Data,'
                campos['conta'] = '1000'
                campos['valor'] = campos_com_base_nos_filtros[filtro[1]].replace('R$ ','').replace('.','') # filtro = 'R$'
                if filtro[2] in campos_com_base_nos_filtros.keys(): # verificar se consta o CPF do paciente
                    patternCPF = r"\d{3}\.?\d{3}\.?\d{3}.?\d{2}"
                    texto = campos_com_base_nos_filtros[filtro[2]] # filtro = 'do CPF.:'
                    campos['CPFPAGOU'] = str(re.findall(patternCPF, texto)[0]).replace('.','').replace('-','')
                    campos['CPFUSOU'] = campos['CPFPAGOU']
                else: # nao achou CPF do paciente
                    campos['CPF'] = ''
                    print('\tERRO COM O CPF DO PACIENTE - NAO ENCONTRADO')
                    logfileout.write('\tERRO COM O CPF DO PACIENTE - NAO ENCONTRADO\n')
                    erro += 1
                campos['Historico'] = 'Arquivo fonte: ' + ff # info repetida + ' : ' + campos_com_base_nos_filtros['Recebemos']
                # campos para gravar no arquivo
                fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )

                #print("Escrito no arquivo: {}".format(file_name))
                            #fileout.write( ff + delimitador + str(text) + delimitador )
                logfileout.write( delimitador )
        print(delimitador + "\nTotal de arquivos analisados = " + str(total) + '\n Total com ERROS (sem CPF) = ' + str(erro))
    logfileout.close()
fileout.close()
