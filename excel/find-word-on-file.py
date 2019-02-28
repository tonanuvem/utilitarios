#http://stackabuse.com/reading-and-writing-lists-to-a-file-in-python/

import textwrap
import csv

def ajustar_arquivo(outputFile, maxcarac):
    output2 = "ajustado-" + outputFile
    with open(outputFile,'r', encoding="utf-8", errors='ignore') as file:
        text = file.read()
    with open(output2, 'w') as ajustado:
        ajustado.write(textwrap.fill(text, width=maxcarac))
    return output2

def search_for_lines(filename, fileoutput, words_list):
    words_found = []
    with open(filename,'r', encoding="utf8", errors='ignore') as db_file:
        for line_no, line in enumerate(db_file):
            for word in words_list:
                if word in line:
                    with open(fileoutput, 'a+') as fileout:
                        text = str(line_no) + ';' + word + ';' + str(line)
                        fileout.write(text)
                        words_found.append(word)
    return words_found

def words_load(words_file):
    with open(words_file,'r') as filehandle:
        words_list = []
        for line in filehandle:
            # remove linebreak which is the last character of the string
            currentPlace = line[:-1]
            # add item to the list
            words_list.append(currentPlace)
    return words_list

def danfe_space(danfe):
    # inclui um espaço a cada 4 caracteres, exemplo abaixo
    # IN:  "31130400085144000198550010000046891004689437"
    # OUT: 3113 0400 0851 4400 0198 5500 1000 0046 8910 0468 9437
    return  " ".join(danfe[i:i+4] for i in range(0, len(danfe), 4))

def resultado(nome_arq):
    return "resultado-" + nome_arq

def ajustado(nome_arq):
    return "ajustado-" + nome_arq

def pesquisar_nas_prox_linhas(filecsv, texto_procurado, qtdlinhas):
    with open(filecsv) as csvfile:
        readCSV = csv.reader(csvfile, delimiter=';')
        linha_info = -1
        words_set = set()
        for row in readCSV:
            if linha_info == -1:
                linha_info =0
                continue;
            #print(row)
            linha_atual = int(row[0])
            if row[1] in texto_procurado:
                print("achou INFORMAES COMPLEMENTARES na linha "+str(linha_atual))
                linha_info = linha_atual
                continue
            if (linha_atual-linha_info)<qtdlinhas:
                print("achou DANFE na linha "+str(linha_atual))
                words_set.add(row[1])
    return words_set

#### definindo parametros

arq_danfe = "DANFE.txt"
arq_danfe_outras_ufs = "DANFE_OUTRAS_UFs.txt"
arq_danfe2 = "DANFE2.txt"
max_caracteres_por_linha = 150
arq_num_doc_fiscal = "nf-sp.txt"
arq_num_doc_fiscal_outra_uf = "nf-outra-uf.txt"
cabecalho_csv = 'numero_linha_encontrado;num_doc_fiscal_encontrado;texto_encontrado\n'
chave_retorno_sp = "chave_retorno_sp.txt"
chave_retorno_outras_ufs = "chave_retorno_outra_uf.txt"

#### ajustar tamanho de cada linha

#ajustar_arquivo(arq_danfe,max_caracteres_por_linha)
#ajustar_arquivo(arq_danfe_outras_ufs,max_caracteres_por_linha)

#### exec das funcoes



## NUMERO DO DOC FISCAL

#words_list_sp = ["INFORMAES COMPLEMENTARES","INFORMAÇÕES COMPLEMENTARES"]
words_list_sp = words_load(arq_num_doc_fiscal)
with open(resultado(arq_danfe), 'w') as fileout:
    fileout.write(cabecalho_csv)
    fileout.close()
    words_found_sp = search_for_lines(arq_danfe2, resultado(arq_danfe), words_list_sp)
'''
#words_list_outra_uf = ["INFORMAES COMPLEMENTARES", "@"]
words_list_outra_uf = words_load(arq_num_doc_fiscal_outra_uf)
with open(resultado(arq_danfe_outras_ufs), 'w') as fileout:
    fileout.write(cabecalho_csv)
    fileout.close()
    words_found_outra_uf = search_for_lines("ajustado-"+arq_danfe_outras_ufs, resultado_outras_ufs, words_list_outra_uf)
'''

## CHAVE RETORNO: Pesquisar os numeros das NFEs para verificar as que fizeram parte do relatório BO e as que ficaram de fora...
'''
chave_retorno_list_sp = words_load(chave_retorno_sp)
for i in range(0,len(chave_retorno_list_sp)): chave_retorno_list_sp[i] = danfe_space(chave_retorno_list_sp[i])
with open(resultado(chave_retorno_sp), 'w') as fileout:
    fileout.write(cabecalho_csv)
    fileout.close()
    chave_retorno_found_sp = search_for_lines(arq_danfe2, resultado(chave_retorno_sp), chave_retorno_list_sp)

chave_retorno_list_outra_uf = words_load(chave_retorno_outras_ufs)
for i in range(0,len(chave_retorno_list_outra_uf)): chave_retorno_list_outra_uf[i] = danfe_space(chave_retorno_list_outra_uf[i])
with open(resultado(chave_retorno_outras_ufs), 'w') as fileout:
    fileout.write(cabecalho_csv)
    fileout.close()
    chave_retorno_found_outra_uf = search_for_lines(ajustado(arq_danfe_outras_ufs), resultado(chave_retorno_outras_ufs), chave_retorno_list_outra_uf)
'''

## Pesquisar somente as 5 proximas linhas depois que encontrar a string: INFORMAES COMPLEMENTARES ou INFORMAÇÕES COMPLEMENTARES

with open(resultado("prox_linhas_info_compl_sp.txt"), 'w') as fileout:
    pesquisa = pesquisar_nas_prox_linhas(resultado(arq_danfe), ["INFORMAES COMPLEMENTARES","INFORMAÇÕES COMPLEMENTARES"], 5)
    for item in pesquisa:
        fileout.write(item+"\n")
    fileout.close()
'''
with open(resultado("prox_linhas_info_compl_outras_ufs.txt"), 'w') as fileout:
    pesquisa = pesquisar_nas_prox_linhas('resultado-DANFE_OUTRAS_UFs.txt',5)
    for item in pesquisa:
        fileout.write(item+"\n")
    fileout.close()
'''
#### print dos resultados

print(str(len(words_found_sp)) + " : Total de palavras encontradas EM SP\n\n------------\n")
#print(str(len(words_found_outra_uf)) + " : Total de palavras encontradas EM OUTRAS UFs")
#print(str(len(chave_retorno_found_sp)) + " : Total de CHAVES DE RETORNO encontradas EM SP")
#print(str(len(chave_retorno_found_outra_uf)) + " : Total de CHAVES DE RETORNO encontradas EM OUTRAS UFs")



