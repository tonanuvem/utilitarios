#from openpyxl import load_workbook
#import textract
import docx2txt
import os

# Directory for files
file_dir = 'C:\\Users\\apsampaio\\Documents\\Python Scripts\\word\\2019'
#file_dir = 'C:\\Users\\Andre\\Dropbox\\Andre\\IR\\2018 ano base\\Charliana\\recibosreferenteaosatendimentosde2018' # "."
# Caracter separador de colunas
delimitador = "\n-----------------------------------\n"
file_name = 'word-to-txt.txt'

#filtro = {"Recebemos do (a) Sr.(a)"; "Portadora do CPF.:"; "Endereço: "; "Quantia de:"}
filtro = ['R$', 'CPF.:', 'Data,']

with open(file_name, 'w') as fileout:

    total = 0
    # Get all word files
    for f in os.listdir(file_dir):
        if f.endswith(".docx"):
            total += 1
            filtrotemporario = set(filtro)
            path = os.path.join(file_dir, f)
            print("Parsing file: {}".format(path))
            # Get file name without extension
            ff = os.path.splitext(f)[0]
            fileout.write( ff + delimitador )
            text = docx2txt.process(path)
            # Filtrar
            for line in text.splitlines():
                Achou = False
                for word in filtrotemporario:
                    palavras = line.split()
                    if word in palavras:
                        #print('string contains a word from the word list: %s' % (word))
                        print(line)
                        fileout.write('\t' + line.lstrip().rstrip() + '\n' )
                        #filtrotemporario.remove(word)
                        Achou = True
                #if Achou:
                #    filtrotemporario.remove(word)
            # write to file
            #fileout.write( ff + delimitador + str(text) + delimitador )
            fileout.write( delimitador )
            print("Writed to file: {}".format(file_name))
    print(delimitador + "\nTotal de arquivos analisados = " + str(total))
fileout.close()
