# URLs
# https://pymupdf.readthedocs.io/en/latest/document/
# https://pymupdf.readthedocs.io/en/latest/page/
# https://pymupdf.readthedocs.io/en/latest/tutorial/#opening-a-document
# https://github.com/JorjMcKie/PyMuPDF-Utilities/blob/master/README.md
# Pyinstaller: https://datatofish.com/executable-pyinstaller/

# pip3 install --upgrade pip
# python3 -m venv pdf_venv
# cd pdf_venv/
# source bin/activate                       CONDA: activate pdf
# python -m pip install --upgrade pip
# python -m pip install --upgrade pymupdf
# pip3 install pyinstaller                  CONDA: conda install -c conda-forge pyinstaller
# pyinstaller --onefile pdfhighlight.py

from operator import itemgetter
from itertools import groupby
import fitz
import json
import sys

#==============================================================================
# Funções auxiliares
#==============================================================================
def words_load(words_file):
    with open(words_file,'r') as filehandle:
        words_list = []
        for line in filehandle:
            # remove linebreak which is the last character of the string
            currentPlace = line[:-1]
            # add item to the list
            words_list.append(currentPlace)
    return words_list

def textfrombox(page, palavra_inicial, palavra_final):
    """
    -------------------------------------------------------------------------------
    Identify the rectangle. We use the text search function here. The two
    search strings are chosen to be unique, to make our case work.
    The two returned rectangle lists both have only one item.
    -------------------------------------------------------------------------------
    """
    rl1 = page.searchFor(palavra_inicial) # rect list one
    rl2 = page.searchFor(palavra_final)   # rect list two
    #print("rl1 = ", rl1, "rl2 = ", rl2)
    rect = rl1[0] | rl2[0]       # union rectangle
    # Now we have the rectangle ---------------------------------------------------

    """
    Get all words on page in a list of lists. Each word is represented by:
    [x0, y0, x1, y1, word, bno, lno, wno]
    The first 4 entries are the word's rectangle coordinates, the last 3 are just
    technical info (block number, line number, word number).
    """
    words = page.getTextWords()
    #print("words = ", words) # We subselect from above list.

    # Case 2: select the words which at least intersect the rect
    #------------------------------------------------------------------------------
    palavras_do_retangulo = []
    mywords = [w for w in words if fitz.Rect(w[:4]).intersects(rect)]
    mywords.sort(key = itemgetter(3, 0))
    group = groupby(mywords, key = itemgetter(3))
    #print("\nSelect the words intersecting the rectangle")
    #print("-------------------------------------------")
    for y1, gwords in group:
        palavra = str(" ".join(w[4] for w in gwords))
        palavras_do_retangulo.append(palavra)
    return palavras_do_retangulo

def getJSON(file):
    with open(file) as jsonFile:
        dados = json.load(jsonFile)
        jsonFile.close()
    return dados

#==============================================================================
# Main Program
#==============================================================================
# inicializar
arquivo_config = "config.json"
if len(sys.argv) > 1:
    arquivo_config = sys.argv[1]
dados = getJSON(arquivo_config)
ifile = dados['nome_arquivo_entrada']
ofile = dados['nome_arquivo_saida']
arq_num_doc_fiscal = dados['nome_arquivo_destacar_texto']
palavra_inicial = dados['delimitador_palavra_inicial']
palavra_final = dados['delimitador_palavra_final']
# carregar
words_list = words_load(arq_num_doc_fiscal)

doc = fitz.open(ifile)
pages = len(doc)

# exemplo de leitura de texto do documento
#for page in doc:
#    text = page.getText()

print("\tUtilizando arquivo de configuracao: '%s'" % (arquivo_config))
print("\tDocumento '%s' tem %i paginas." % (doc.name, len(doc)))

# page number should start from 0
num_pag=11
pn = int(num_pag)-1

#gerando novo pdf
docout = fitz.open()
qtd_highlight = 0

# get the page
for i in range(doc.pageCount):
    pagina = doc[i]
    try :
        palavras_pdf = textfrombox(pagina, palavra_inicial, palavra_final)
    except Exception as e1: #falha na 1a tentativa
        print("\t\t\tTentando novo  processamento da pagina ",str(i+1), " com mensagem de erro: ",str(e1))
        try :
            palavras_pdf = textfrombox(pagina, "Nr. Doc.", palavra_final)
        except Exception as e2: #falha na 2a tentativa :
            print("\t\t\t\tErro no processamento da pagina ",str(i+1), " com mensagem de erro: ",str(e2))
            continue;
    
    #print("\tPesquisando página "+str(i+1)+" com as palavras encontradas: "+str(palavras_pdf)) #utilizado para testes
    #flag de controle para inserir nova pagina
    nova_pagina = True
    for word in words_list:
      for k in range(len(palavras_pdf)):
        if word in palavras_pdf[k]:
            if (word != palavras_pdf[k]):
                continue
            #print("\tPesquisando palavras do TXT "+ str(words_list) + " na página "+str(i+1)+" com as palavras encontradas no PDF: "+str(palavras_pdf)) #utilizado para testes
            # verifica se a palavra foi encontrada no fim da lista (esta entre as ultimas 2 palavras), caso positivo, devem ser inseridas a pagina que contem a palavra e tb a seguinte, evitando quebra das pags
            encontrado_fim_pag_pdf = 0 #if palavras_pdf.index(word) < (len(palavras_pdf)-2) else 1
            # insere pdf no doc de saida, depois faz o highlight
            if nova_pagina :
                docout.insertPDF(doc, from_page = i, to_page = i+encontrado_fim_pag_pdf)  # inserindo página
                nova_pagina = False
            lastpage = docout[docout.pageCount-1 - encontrado_fim_pag_pdf]
            pesquisa = lastpage.searchFor(word, hit_max = 16)
            print("\t\tEncontrado texto '%s' on the page %i" % (word, i))
            for p in pesquisa:
                # inserindo highlights
                #print("p = %s e word = %s" % (palavras_pdf[k],word))
                lastpage.addHighlightAnnot(p)
                qtd_highlight+=1

# now we are ready for search, with max hit count limited to 16
# the return result is a list of hit box rectangles

#salvando com highlights
if docout.pageCount <= 0:
    print("\n\narquivo de saida vazio")
    raise SystemExit('%s has %d pages only' % (docout, docout.pageCount))
docout.save(ofile)
print("arquivo gerado com ",str(docout.pageCount)," paginas. Foram feitos ",str(qtd_highlight),"highlights; de um total de ",str(len(words_list))," itens procurados")
