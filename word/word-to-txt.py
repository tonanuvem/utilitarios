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
filtro = ['Data,', 'R$', 'CPF.:', 'Recebemos']
campos_com_base_nos_filtros = {}
# campos = {'data': '05/01/'+ANO, 'conta': '', 'valor': '0,00', 'CPFPAGOU': '', 'CPFUSOU': '', 'Historico': ''}
# data = DD/MM/AAAA ou DD/MM/AA
# contas = 1000 (recebimentos) ; 4001 (aluguel) ; 4017 (propaganda)
# CPF do paciente
# CPF da Dra
# Historico = Campo texto com ate 250 chars = usar Nome do Paciente ; GOOGLE ADS ; ALUGUEL

def getCPF(texto):
    patternCPF = r"\d{3}\.?\d{3}\.?\d{3}.?\d{2}"
    listCPF = re.findall(patternCPF, texto)
    if listCPF:
        return str(listCPF[0]).replace('.','').replace('-','')
    else:
        return ''

if __name__ == "__main__":
    with open(file_name, 'w') as fileout:
        # logs
        with open('log-'+file_name, 'w') as logfileout:
            # criar gasto relacionados a: Material de Escritório
            campos = {'data': '29/01/'+ANO, 'conta': '4012', 'valor': '84,80', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '','Historico': 'Fonte Cartao Credito: KALUNGA'}
            fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
            # criar gasto relacionados a: Material Conservação e Limpeza de Escritório
            campos = {'data': '16/02/'+ANO, 'conta': '4011', 'valor': '171,06', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '','Historico': 'Fonte Cartao Credito: MASTER SUPERMERCADO'}
            fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
            # criar gasto relacionados a: Anuidade do Conselho
            campos = {'data': '05/01/'+ANO, 'conta': '4004', 'valor': '772,50', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '','Historico': 'CREMESP'}
            fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
            campos = {'data': '09/09/'+ANO, 'conta': '4004', 'valor': '725,00', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '','Historico': 'SOBED'}
            fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
            # criar gasto relacionados a: Livros
            campos = {'data': '19/07/'+ANO, 'conta': '4015', 'valor': '881,79', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '','Historico': 'Fonte Cartao Credito: LIVRARIA LUANA'}
            fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
            # criar gasto relacionados a: Congressos
            campos = {'data': '24/05/'+ANO, 'conta': '4018', 'valor': '1580,00', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '', 'Historico': 'CONGRESSO CAMPINAS 2019 : SOC MED CIRURGIA CAMPINAS BR'}
            fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
            campos = {'data': '22/10/'+ANO, 'conta': '4018', 'valor': '1210,00', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '', 'Historico': 'CONGRESSO SBAD : CCM SBAD 2019'}
            fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
            campos = {'data': '22/10/'+ANO, 'conta': '4018', 'valor': '450,00', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '', 'Historico': 'CONGRESSO SBAD : CCM SBAD 2019'}
            fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
            # criar os gastos mensais com propaganda e aluguel
            for mes in range(1, 13):
                campos = {'data': '05/01/'+ANO, 'conta': '', 'valor': '0,00', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '', 'Historico': ''}
                for conta in ['4001', '4017']:
                    campos['conta'] = conta
                    if mes < 10:
                        campos['data'] = '05/0'+str(mes)+'/'+ANO
                    else:
                        campos['data'] = '05/'+str(mes)+'/'+ANO
                    if conta == '4001':
                        campos['valor'] ='1100'
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
                                if getCPF(line) == getCPF(CPFDRA): # pular o CPFDRA
                                    continue
                                logfileout.write('\t' + line.lstrip().rstrip() + '\n' )
                                campos_com_base_nos_filtros[word]=line.lstrip().rstrip()
                    # write to file
                    # escrevendo os campos
                    # pega os valores do dict: campos_com_base_nos_filtros[word]
                    campos['data'] = campos_com_base_nos_filtros[filtro[0]].replace('Data, ','').replace('.','') # filtro = 'Data,'
                    campos['conta'] = '1000'
                    campos['valor'] = campos_com_base_nos_filtros[filtro[1]].replace('R$ ','').replace('.','') # filtro = 'R$'
                    if filtro[2] in campos_com_base_nos_filtros.keys(): # verificar se consta o CPF do paciente
                        texto = campos_com_base_nos_filtros[filtro[2]] # filtro = 'CPF.:'
                        if getCPF(texto) == getCPF(CPFDRA): # pular o CPFDRA
                            continue
                        campos['CPFPAGOU'] = getCPF(texto)
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
