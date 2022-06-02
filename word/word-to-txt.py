import docx2txt
import os
import re

# Leitura de arquivos WORD DOCX e escrita no formato CSV para importar no CARNE LEAO (IRPF)

# Directory for files
ANO = '202X'
file_dir = './Recibos_' + ANO
# Caracter separador de colunas
delimitador = "\n-----------------------------------\n"
CPFDRA = ''

file_name = CPFDRA+'-'+ANO+'.txt'
file_name_CSV = CPFDRA+'-'+ANO+'.csv'

#filtro = {"Recebemos do (a) Sr.(a)"; "Portadora do CPF.:"; "Endereço: "; "Quantia de:"}
filtro = ['Data,', 'R$', 'CPF.:', 'Recebemos']
campos_com_base_nos_filtros = {}


def getCPF(texto):
    patternCPF = r"\d{3}\.?\d{3}\.?\d{3}.?\d{2}"
    listCPF = re.findall(patternCPF, texto)
    if listCPF:
        return str(listCPF[0]).replace('.','').replace('-','')
    else:
        return ''

def rendimento_linha_arquivo(fileout, data, valor, CPFPAGOU, historico):
    # Exemplo Modelo de arquivo para rendimentos do Trabalho não Assalariado
    # 99/99/9999;R01.001.001;999;999999999,99;;Modelo de linha para rendimento do trabalho não assalariado recebido de Pessoa Física quando o CPF do beneficiario foi informado;PF;99999999999;99999999999
    campos = {}
    # 1) Data de recebimento - substituir 99/99/9999 pela data do recebimento.
    campos['data'] = data
    # 2) Código do rendimento - não há necessidade de substituição. O código corresponde ao tipo de rendimento descrito no campo histórico. Escolha a linha que se encaixa a sua necessidade.
    # R01.001.001 : Trabalho Não Assalariado
    campos['codrendimento'] = "R01.001.001"
    # 3) Código de ocupação - substituir pelo código correspondente a sua ocupação. Para saber qual utilizar, consulte a tabela de códigos de ocupação.
    # 225 : Médico
    campos['codocupacao'] = "225"
    # 4) Valor recebido - substituir 9999999,99 pelo valor recebido.
    campos['valor'] = valor
    # 5) Valor de dedução - vazio (mas posição do campo deve constar do arquivo).
    campos['deducao'] = ""
    # 6) Histórico - substituir o texto do modelo de acordo com seus dados.
    campos['historico'] = historico
    # 7) Recebido de - utilizar PF (pessoa física), PJ (pessoa jurídica), EX (exterior).
    campos['recebido_de'] = "PF"
    # 8) CPF do titular do pagamento - quando rendimento recebido de pessoa física, substituir 99999999999 pelo CPF correspondente.
    campos['cpf_pagou'] = CPFPAGOU
    # 9) CPF do beneficiário do serviço - quando rendimento recebido de pessoa física, substituir 99999999999 pelo CPF correspondente.
    campos['cpf_usou'] = CPFPAGOU
    # 10) Indicador de CPF não informado - preenchido somente nos casos em que houver a exigência do CPF do beneficiário e esse não foi informado. Preencher com S.
    # 11) CNPJ - quando rendimento recebido de pessoa jurídica, substituir 99999999999 pelo CNPJ correspondente.
    # escrever no arquivo
    fileout.write( campos['data']+';'+campos['codrendimento']+';'+campos['codocupacao']+';'+campos['valor']+';'+campos['deducao']+';'+campos['historico']+';'+campos['recebido_de']+';'+campos['cpf_pagou']+';'+campos['cpf_usou']+'\n' )

def pagamento_linha_arquivo (fileout, data, conta, valor, CPFPAGOU, historico):
    # Exemplo Modelos de Arquivo para Pagamentos do Plano de Contas padrão
    # 99/99/9999;P10.01.00002;999999999,99;Modelo de linha para pagamento de Aluguel do escritório/consultório
    campos = {}
    # 1) Data do pagamento - substituir 99/99/9999 pela data do pagamento.
    campos['data'] = data
    # 2) Código do pagamento - não há necessidade de substituição. O código corresponde ao tipo de pagamento descrito no campo histórico.
    # Para despesas dedutíveis:     P10 + . + código da conta
    campos['codpagamento'] = conta
    # 3) Valor pago - substituir 9999999,99 pelo valor pago.
    campos['valor'] = valor
    # 4) Histórico - substituir o texto do modelo de acordo com seus dados.
    campos['historico'] = historico
    # 5) Valor da Multa - para os pagamentos do tipo Imposto Pago e Previdência Oficial, quando aplicável, substituir 9999999,99 pelo valor da multa.
    campos['multa'] = ""
    # 6) Valor dos Juros - para os pagamentos do tipo Imposto Pago e Previdência Oficial, quando aplicável, substituir 9999999,99 pelo valor dos juros.
    campos['juros'] = ""
    # 7) Competência - para os pagamentos do tipo Imposto Pago e Previdência Oficial, substituir 99/9999 pelo mês e ano correspondente.
    campos['competencia'] = ""
    # escrever no arquivo
    fileout.write( campos['data']+';'+campos['codpagamento']+';'+campos['valor']+';'+campos['historico']+campos['multa']+campos['juros']+campos['competencia']+'\n' )

if __name__ == "__main__":
    with open('log-'+file_name, 'w') as logfileout:
        with open('pagamentos-'+file_name_CSV, 'w') as fileout:
            # criar gasto relacionados a: Material de Escritório
            # P10.01.00012 : Material de escritório
            pagamento_linha_arquivo(fileout, data='09/03/'+ANO, conta='P10.01.00012', valor='320,00', CPFPAGOU=CPFDRA, historico='Fonte PIX: GRAFICA MEDIMPRES')
            pagamento_linha_arquivo(fileout, data='07/02/'+ANO, conta='P10.01.00012', valor='29,00', CPFPAGOU=CPFDRA, historico='Fonte NFP DOC 788922: KALUNGA SA')
            pagamento_linha_arquivo(fileout, data='07/02/'+ANO, conta='P10.01.00012', valor='135,30', CPFPAGOU=CPFDRA, historico='Fonte NFP DOC 788919: KALUNGA SA')
            pagamento_linha_arquivo(fileout, data='15/11/'+ANO, conta='P10.01.00012', valor='122,50', CPFPAGOU=CPFDRA, historico='Fonte NFP DOC 218469: KALUNGA SA')
            pagamento_linha_arquivo(fileout, data='16/11/'+ANO, conta='P10.01.00012', valor='212,40', CPFPAGOU=CPFDRA, historico='Fonte NFP DOC 218760: KALUNGA SA')
            #campos = {'data': '09/03/'+ANO, 'conta': 'P10.01.00012', 'valor': '320,00', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '','Historico': 'Fonte PIX: GRAFICA MEDIMPRES'}
            #fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
            # criar gasto relacionados a: Material Conservação e Limpeza de Escritório
            # P10.01.00011 : Material de conservação e limpeza do escritório/consultório
            pagamento_linha_arquivo(fileout, data='16/02/'+ANO, conta='P10.01.00011', valor='600,00', CPFPAGOU=CPFDRA, historico='Fonte Cartao Credito: MASTER SUPERMERCADO')
            #campos = {'data': '16/02/'+ANO, 'conta': '4011', 'valor': '171,06', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '','Historico': 'Fonte Cartao Credito: MASTER SUPERMERCADO'}
            #fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
            # criar gasto relacionados a: Anuidade do Conselho
            # P10.01.00004 : Contribuições obrigatórias a entidades de classe
            pagamento_linha_arquivo(fileout, data='05/01/'+ANO, conta='P10.01.00004', valor='733,40', CPFPAGOU=CPFDRA, historico='CREMESP')
            #campos = {'data': '05/01/'+ANO, 'conta': '4004', 'valor': '733,40', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '','Historico': 'CREMESP'}
            #fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
            pagamento_linha_arquivo(fileout, data='16/02/'+ANO, conta='P10.01.00004', valor='575,00', CPFPAGOU=CPFDRA, historico='FBG')
            #campos = {'data': '09/04/'+ANO, 'conta': '4004', 'valor': '575,00', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '','Historico': 'FBG'}
            #fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
            #campos = {'data': '09/09/'+ANO, 'conta': '4004', 'valor': '725,00', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '','Historico': 'SOBED'}
            #fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
            # criar gasto relacionados a: Livros
            #campos = {'data': '19/07/'+ANO, 'conta': '4015', 'valor': '881,79', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '','Historico': 'Fonte Cartao Credito: LIVRARIA LUANA'}
            #fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
            # criar gasto relacionados a: Congressos
            # Nao achei codigo especifico
            # P10.01.00006 : Emolumentos pagos a terceiros
            pagamento_linha_arquivo(fileout, data='12/05/'+ANO, conta='P10.01.00006', valor='310,00', CPFPAGOU=CPFDRA, historico='SIMPOSIO SOBED')
            #campos = {'data': '12/05/'+ANO, 'conta': '4018', 'valor': '310,00', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '', 'Historico': 'SIMPOSIO SOBED'}
            #fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
            pagamento_linha_arquivo(fileout, data='12/05/'+ANO, conta='P10.01.00006', valor='100,00', CPFPAGOU=CPFDRA, historico='SIMPOSIO SOBED')
            #campos = {'data': '12/05/'+ANO, 'conta': '4018', 'valor': '100,00', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '', 'Historico': 'SIMPOSIO SOBED'}
            #fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
            pagamento_linha_arquivo(fileout, data='23/11/'+ANO, conta='P10.01.00006', valor='100,00', CPFPAGOU=CPFDRA, historico='SIMPOSIO SOBED')
            #campos = {'data': '23/11/'+ANO, 'conta': '4018', 'valor': '550,00', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '', 'Historico': 'CONGRESSO SBAD'}
            #fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
            # criar os gastos mensais aluguel e tambem com propaganda e aluguel
            # P10.01.00002 : Aluguel do escritório/consultório
            # Nao achei codigo especifico com propaganda e aluguel, usando o abaixo
            # P10.01.00006 : Emolumentos pagos a terceiros
            for mes in range(1, 13):
                campos = {}
                #campos = {'data': '05/01/'+ANO, 'conta': '', 'valor': '0,00', 'CPFPAGOU': CPFDRA, 'CPFUSOU': '', 'Historico': ''}
                #for conta in ['4001', '4017']:
                for conta in ['P10.01.00002', 'P10.01.00006']:
                    campos['conta'] = conta
                    if mes < 10:
                        campos['data'] = '05/0'+str(mes)+'/'+ANO
                    else:
                        campos['data'] = '05/'+str(mes)+'/'+ANO
                    if conta == 'P10.01.00002':
                        campos['valor'] ='1257,10'
                        campos['Historico'] ='GASTO COM ALUGUEL'
                    else : #  conta == '4017':
                        campos['valor'] ='350'
                        campos['Historico'] =' GASTO COM GOOGLE ADS'
                    campos['CPFPAGOU'] = CPFDRA
                    # campos para gravar no arquivo
                    pagamento_linha_arquivo(fileout, data=campos['data'], conta=campos['conta'], valor=campos['valor'], CPFPAGOU=campos['CPFPAGOU'], historico=campos['Historico'])
                    #fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )
        # fechar arquivo pagamento
        fileout.close()
        with open('recebimentos-'+file_name_CSV, 'w') as fileout:
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
                    campos['data'] = campos_com_base_nos_filtros[filtro[0]].replace('Data, ','').replace('.','') # filtro = 'Data, '
                    campos['data'] = campos['data'].replace('Data,','').replace('.','') # filtro = 'Data,'
                    campos['conta'] = '1000'
                    campos['valor'] = campos_com_base_nos_filtros[filtro[1]].replace('R$ ','').replace('.','') # filtro = 'R$ '
                    campos['valor'] = campos['valor'].replace('R$','').replace('.','') # filtro = 'R$'
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
                    rendimento_linha_arquivo(fileout, data=campos['data'], valor=campos['valor'], CPFPAGOU=campos['CPFPAGOU'], historico=campos['Historico'])
                    #fileout.write( campos['data']+';'+campos['conta']+';'+campos['valor']+';'+campos['CPFPAGOU']+';'+campos['CPFUSOU']+';'+campos['Historico']+'\n' )

                    #print("Escrito no arquivo: {}".format(file_name))
                                #fileout.write( ff + delimitador + str(text) + delimitador )
                    logfileout.write( delimitador )
            print(delimitador + "\nTotal de arquivos analisados = " + str(total) + '\n Total com ERROS (sem CPF) = ' + str(erro))
        fileout.close()
    logfileout.close()
