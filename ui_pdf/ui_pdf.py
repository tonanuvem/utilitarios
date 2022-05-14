#!/usr/bin/env python

# TODOs:
# - inserir ProgressDialog ao clicar em cada um dos botoes (ex: juntar PDFs ou separar PDFs).
# - fazer um interface gráfica para o PDF Highlight
# OK - inserir a opcao de excluir paginas de um PDF
# OK - inserir visualizador de PDF
# OK - inserir a configuracao em json(ex: ultimo diretorio escolhido) 
# OK - inserir os botões (mover para cima ou para baixo) para reajustar a ordem dos arquivos exibidos

import os, wx, fitz, time, math
from natsort import natsorted
import json, sys
#from pdf_split_tool import pdf_splitter

import wx.lib.sized_controls as sc
from wx.lib.pdfviewer import pdfViewer, pdfButtonPanel

###
# Funcionalidade de PDF VIEWER dentro da UI
###
class PDFViewer(sc.SizedFrame):
    def __init__(self, parent, **kwds):
        super(PDFViewer, self).__init__(parent, **kwds)

        paneCont = self.GetContentsPane()
        self.buttonpanel = pdfButtonPanel(paneCont, wx.NewId(),
                                wx.DefaultPosition, wx.DefaultSize, 0)
        self.buttonpanel.SetSizerProps(expand=True)
        self.viewer = pdfViewer(paneCont, wx.NewId(), wx.DefaultPosition,
                                wx.DefaultSize,
                                wx.HSCROLL|wx.VSCROLL|wx.SUNKEN_BORDER)
        self.viewer.UsePrintDirect = False
        self.viewer.SetSizerProps(expand=True, proportion=1)

        # introduce buttonpanel and viewer to each other
        self.buttonpanel.viewer = self.viewer
        self.viewer.buttonpanel = self.buttonpanel

###
# Funcionalidade de MERGE
###
class PDF():

    @staticmethod
    def merge(dir, files, ofile, maxsize=10):
        pdfs_count = 1
        if len(files) > 0:
            # files = [f for f in os.listdir(dir) if f.endswith(".pdf")]
            # pip install natsort
            # files = natsorted(files)
            # files = sorted(files)
            print("merge : Diretorio: "+str(dir)+"\tTamanho maximo dos arquivos de saida: "+str(maxsize))
            print("\tArquivos ordenados: "+str(files))
            result = fitz.open()
            output_pdf = dir+'/'+ofile
            for pdf in files:
                pdfs_count += 1
                with fitz.open(dir+'/'+pdf) as mfile:
                    result.insert_pdf(mfile)
            result.save(output_pdf)
            print("Arquivos juntados!")
            output_pdf_size = os.path.getsize(output_pdf)/(1024*1024)
            if output_pdf_size > 0:
                print("Tamanho do novo arquivo gerado: %.2f " %(output_pdf_size) )
                return pdfs_count
        return 0

###
# Funcionalidade de SPLIT :  POR TAMANHO ou POR PAGINAS
###
    @staticmethod
    def sizeSplit(filename, max_size):
        #filename = 'C:\\Users\\apsampaio\\Downloads\\4145916_1.pdf'
        print("sizeSplit : arquivo = "+ str(filename) +" , tamanho ="+ str(max_size))
        doc = fitz.open(filename)
        # Verificar tamanho o PDF em MB
        pdf_size = os.path.getsize(filename)/(1024*1024)
        total_pages = len(doc)
        avg_size = pdf_size / total_pages
        #Transformando em MB:
        #avg_size = avg_size/(1024*1024)
        print("\tDocumento '%s' tem %i paginas com tamanho total %.2f MB" % (doc.name,total_pages,pdf_size))
        print("\t Páginas de tamanho em torno de %.2f MB." % (avg_size))
        # exemplo de leitura do documento
        #for page in doc:
        #    text = page.getText()
        if pdf_size > max_size:
            avg_step = int(max_size / avg_size)
            pdfs_count = 1
            current_page = 0
            end_page = current_page + avg_step
            out_pdf_size = 0
            while current_page != total_pages:
                if end_page > total_pages:
                    end_page = total_pages
                incluir_paginas = str(current_page+1) + "-" + str(end_page)
                print("\tProcessando Documento '%s' : Parte %i : pag_inicial = %i : pag_final = %i" % (doc.name,pdfs_count,current_page, end_page))
                pages = PDF.auxiliar(incluir_paginas) # ajustar os indices começando em zero
                print("\tProcessando Documento : intervalo = %s" % (pages))
                # tratamento para quando há somente 1 página e ainda é maior que o tamanho máximo
                if current_page == end_page:
                    print("\t\t\t\tErro no tamanho máximo %i, pois a pagina %i é maior que isso : %.2f MB" %(max_size,end_page, out_pdf_size))
                    return 0
                #gerando novo pdf
                docout = fitz.open()
                # get the pages
                try : # inserir paginas
                    outdoc = fitz.open(filename)
                    outdoc.select(pages)                   # delete all others
                    ofile = filename.replace(".pdf", "-{}.pdf".format(pdfs_count))
                    outdoc.ez_save(ofile)                     # save and clean new PDF
                    outdoc.close()
                    # verificar se o tamanho do arquivo ainda ficou menor, e pegar menos páginas
                    out_pdf_size = os.path.getsize(ofile)/(1024*1024)
                    print("\tProcessando Documento : Parte %i : ficou com tamanho de %.2f MB" % (pdfs_count, out_pdf_size))
                    if out_pdf_size > max_size:
                        # refazer essa parte do arquivo com menos páginas (metade da qtd de páginas)
                        # qtd_digitos = len(str(end_page))
                        # tratamento para arquivos com muitas paginas
                        #end_page = current_page + int(avg_step/2)
                        end_page = int(end_page * 0.95)
                        # diminui mais do que deveria
                        if end_page < current_page:
                            end_page = current_page
                        print(end_page)
                        #return
                        continue
                except Exception as e1: #falha na tentativa
                    print("\t\t\t\tErro no processamento com mensagem de erro: ",str(e1))
                    break
                #continuar o loop
                current_page = end_page
                end_page = current_page + avg_step
                pdfs_count += 1
            return pdfs_count
        return 0


    #==============================================================================
    # Delimitadores, função auxiliar
    # Input: recebe as páginas iniciando em 1 (como o usuario ve o doc)
    # Retorno: lista de paginas iniciando em 0 (como o fitz ve o doc)
    #==============================================================================
    @staticmethod
    def auxiliar(data):
    #def auxiliar(self, data):
        delimiters = "\n\r\t,;"
        for d in delimiters:
            data = data.replace(d, ' ')
        lines = data.split(' ')
        numbers = []
        i = 0
        for line in lines:
            if line == '':
                continue
            elif '-' in line:
                t = line.split('-')
                #numbers += range(int(t[0]), int(t[1]) + 1)
                numbers += range(int(t[0])-1, int(t[1]))
            else:
                #numbers.append(int(line))
                numbers.append(int(line)-1)
        return numbers

    @staticmethod
    def pageSplit(filename, pages, ofile):
    #def pageSplit(self, filename, pages):
        #dir="C:/Users/apsampaio/OneDrive - Secretaria da Fazenda e Planejamento do Estado de São Paulo/DRTC-III/ICMS/_Operação/MONITORAMENTO/142.086.828.111 - BIOEXTRA INDUSTRIA E COMERCIO EIRELI/AIIMs/2 AIIM 4.145.916-7 ref FOX/"
        #dir="H:/drtc-3/drtciii-nf2/03- CCQ - EQUIPE 23/AIIM 4.145.916-7_BIOEXTRA_ref_FOX/"

        #fileInput="prova7_SFPPRC202104143V01-parte2.pdf"

        # exemplos:
        #incluir_paginas = "1-40; 43-44; 46-49"
        #excluir_paginas = "1-3; 5-33" 

        #ifile=dir+fileInput
        #ofile=dir+"selecao_"+fileInput
        
        # paths longos dao problema -> copiar para o M: ou H:
        #ofile=dir+"paginas_"+str(paginaInicial)+"_"+str(paginaFinal)+"_"+fileInput
        print("pageSplit : arquivo = "+filename+"\n paginas a serem excluidas: " + pages+"\n arquivo de saida : "+ofile)

        doc = fitz.open(filename)
        excluir_paginas = PDF.auxiliar(pages) # ajustar os indices começando em zero
        incluir_paginas = range(len(doc))
        pages = set(incluir_paginas) - set(excluir_paginas)

        # exemplo de leitura de texto do documento
        #for page in doc:
        #    text = page.getText()

        print("\tDocumento '%s' tem %i paginass." % (doc.name, len(doc)))
        print("\t %i Paginas que serao EXCLUIDAS no documento: %s" % (len(excluir_paginas), str(excluir_paginas)))
        print("\t %i Paginas que VAO FICAR no documento: %s" % (len(pages), str(pages)))

        #gerando novo pdf
        docout = fitz.open()

        # get the pages
        try : # inserir paginas
            doc.select(list(pages))                   # delete all others
            doc.ez_save(ofile)                     # save and clean new PDF
            doc.close()
            return len(excluir_paginas)
        except Exception as e1: #falha na tentativa
            print("\t\tErro no processamento com mensagem de erro: ",str(e1))
            return 0

###
# Interface grafica
###
class AppFrame(wx.Frame):    
    def __init__(self):
        # Init de componentes
        super().__init__(parent=None, title='SEFAZ PDF Utils', size=(530,550))
        panel = wx.Panel(self)        
        vbox = wx.BoxSizer(wx.VERTICAL)

        #------------
        # Lista de arquivos
        #------------
        self.list = wx.ListCtrl(panel, size=(300,180), style=wx.LC_REPORT)
        self.list.EnableCheckBoxes()
        # Evento: Botao direito na Lista abre o PDF Viewer
        self.Bind(wx.EVT_LIST_ITEM_RIGHT_CLICK, self.on_doubleclick_list)
        # Adicionar componente na UI
        vbox.Add(self.list, 0, wx.EXPAND)

        #------------
        # Texto que descreve o tamanho total
        #------------
        self.text_total_size = wx.TextCtrl(panel, style=wx.TE_READONLY | wx.TE_CENTRE)
        #self.text_ctrl_size.AppendText("Tamanho estimado (soma dos arquivos) = "+ str(f'{tamanhoTotal/(1024*1024):.2f}') +" MBs")
        vbox.Add(self.text_total_size, 0, wx.ALL | wx.EXPAND, 5)

        #------------
        # Botao para mudar arquivos de posição na lista : Up ou Down + Select All ou Deselect All
        #------------
        hbox_pos = wx.BoxSizer(wx.HORIZONTAL)
        btn_pos_up = wx.Button(panel, label='Up')
        btn_pos_up.Bind(wx.EVT_BUTTON, self.on_press_up)
        hbox_pos.Add(btn_pos_up)
        btn_pos_down = wx.Button(panel, label='Down')
        btn_pos_down.Bind(wx.EVT_BUTTON, self.on_press_down)
        hbox_pos.Add(btn_pos_down)
        btn_pos_select_all = wx.Button(panel, label='Selecionar Todos')
        btn_pos_select_all.Bind(wx.EVT_BUTTON, self.on_press_select_all)
        hbox_pos.Add(btn_pos_select_all)
        btn_pos_select_clear = wx.Button(panel, label='Limpar seleção')
        btn_pos_select_clear.Bind(wx.EVT_BUTTON, self.on_press_select_clear)
        hbox_pos.Add(btn_pos_select_clear)
        vbox.Add(hbox_pos, flag=wx.CENTER)

        #------------
        # Linha separadora
        #------------

        #------------
        # Diretorio selecionado
        #------------
        # Atualiza a Lista de arquivos e o campo com o tamanho total
        # ler configuracao do arquivo json
        self.arquivo_config = "config.json"
        self.dados = dict()
        try:
            with open(self.arquivo_config) as jsonFile:
                self.dados = json.load(jsonFile)
                jsonFile.close()
            if len(self.dados) > 0:
                idir = str(self.dados['diretorio_selecionado'])
                self.atualizarFileListCtrl(diretorio=idir)
            else:
                self.atualizarFileListCtrl(diretorio='.')
        except Exception as e1: #falha
            print("\tErro o carregar arquivo config.json: ",str(e1))
            self.dados['diretorio_selecionado'] = '.' 
            self.atualizarFileListCtrl(diretorio='.')
        # Texto do Diretorio selecionado
        self.text_ctrl_dir = wx.TextCtrl(panel, style=wx.TE_READONLY)
        if len(self.dados) > 0:
            idir = str(self.dados['diretorio_selecionado'])
            self.text_ctrl_dir.AppendText(idir)
        else:
            self.text_ctrl_dir.AppendText("Clique no botão abaixo para selecionar um diretorio")
        vbox.Add(self.text_ctrl_dir, 0, wx.ALL | wx.EXPAND, 5)

        #------------
        # Botao para selecionar Dir
        #------------
        my_btn_dir = wx.Button(panel, label='Mudar diretorio')
        my_btn_dir.Bind(wx.EVT_BUTTON, self.on_press_dir)
        vbox.Add(my_btn_dir, 0, wx.ALL | wx.CENTER, 5)

        #------------
        # Nome do Arquivo de saida
        #------------
        self.text_ctrl_ofile = wx.TextCtrl(panel)
        self.text_ctrl_ofile.AppendText("_output.pdf")
        vbox.Add(self.text_ctrl_ofile, 0, wx.ALL | wx.EXPAND, 5)

        #------------
        # Botao para processar
        #------------        
        my_btn_merge = wx.Button(panel, label='Fazer o merge dos arquivos')
        my_btn_merge.Bind(wx.EVT_BUTTON, self.on_press_merge)
        vbox.Add(my_btn_merge, 0, wx.ALL | wx.CENTER, 5)

        #------------
        # Tamanho do Arquivo de saida em Mb
        #------------
        self.text_ctrl_size = wx.TextCtrl(panel)
        self.text_ctrl_size.AppendText("10")
        vbox.Add(self.text_ctrl_size, 0, wx.ALL | wx.EXPAND, 5)

        #------------
        # Botao para processar
        #------------        
        my_btn_size = wx.Button(panel, label='Gerar arquivos com tamanhos máximos em MB.')
        my_btn_size.Bind(wx.EVT_BUTTON, self.on_press_size)
        vbox.Add(my_btn_size, 0, wx.ALL | wx.CENTER, 5)

        #------------
        # Excluir páginas do PDF
        #------------
        self.text_ctrl_delete = wx.TextCtrl(panel)
        self.text_ctrl_delete.AppendText("2-7")
        vbox.Add(self.text_ctrl_delete, 0, wx.ALL | wx.EXPAND, 5)
        #------------
        # Botao para processar a Exclusão de páginas
        #------------    
        my_btn_delete = wx.Button(panel, label='Deletar páginas do intervalo selecionado')
        my_btn_delete.Bind(wx.EVT_BUTTON, self.on_press_delete)
        vbox.Add(my_btn_delete, 0, wx.ALL | wx.CENTER, 5)
        
        #------------
        # Atualizar tela
        #------------
        panel.SetSizer(vbox)        
        self.Show()

###
# Classe: AppFrame
# Método para atualizar os arquivos do Diretorio selecionado
###
    def atualizarFileListCtrl(self, diretorio):
        j = 0
        totalsize = 0
        self.list.ClearAll()
        # Criar campos:
        self.list.InsertColumn(0, 'Nome')
        self.list.InsertColumn(1, 'Extensão')
        self.list.InsertColumn(2, 'Tamanho', wx.LIST_FORMAT_RIGHT)
        self.list.InsertColumn(3, 'Data motificação')
        self.list.SetColumnWidth(0, 250)
        self.list.SetColumnWidth(1, 60)
        self.list.SetColumnWidth(2, 70)
        self.list.SetColumnWidth(3, 110)
        #self.list.InsertItem(0, '..')
        # Exibir arquivos
        files = os.listdir(diretorio)
        files = natsorted(files)
        for i in files:
            (name, ext) = os.path.splitext(i)
            ex = ext[1:]
            # continua somente para inserir aquivo pdf
            if ex not in ["pdf", "PDF"] :
                continue
            size = os.path.getsize(diretorio+"/"+i)
            sec = os.path.getmtime(diretorio+"/"+i)

            self.list.InsertItem(j, name)
            self.list.SetItem(j, 1, ex)
            self.list.SetItem(j, 2, str(f'{size/(1024*1024):.2f}') + ' MB')
            self.list.SetItem(j, 3, time.strftime('%Y-%m-%d %H:%M', time.localtime(sec)))

            if (j % 2) == 0:
                self.list.SetItemBackgroundColour(j, '#e6f1f5')
            j = j + 1
            totalsize += size
        if totalsize > 0:
            self.text_total_size.SetValue("Tamanho estimado (soma dos arquivos) = "+ str(f'{totalsize/(1024*1024):.2f}') +" MBs")

###
# Classe: AppFrame
# Eventos ao clicar com botão direito na lista de arquivos da Interface grafica
###
    def on_doubleclick_list(self, event):
        select = self.list.GetItemText(event.Index)
        print("Capturado duplo click : "+str(select))
        #wx.MessageBox('Selecionado item %s' % str(select))
        pdfV = PDFViewer(self, size=(800, 600))
        #pdfV.viewer.UsePrintDirect = False
        dir = self.text_ctrl_dir.GetValue()
        filepath = str(dir)+'/'+str(select)+'.pdf'
        print("Tentando abrir PDF Viewer para : "+str(filepath))
        #pdfV.viewer.LoadFile(r'a path to a .pdf file')
        pdfV.viewer.LoadFile(filepath)
        pdfV.Show()

###
# Classe: AppFrame
# Eventos relacionados ao Drag and Drop da lista de arquivos da Interface grafica
# Código fonte: https://wiki.wxpython.org/How%20to%20create%20a%20list%20control%20with%20drag%20and%20drop%20%28Phoenix%29
###
    def on_press_up(self, event):
        itemcount = self.list.GetItemCount()
        itemschecked = [i for i in range(itemcount) if self.list.IsItemChecked(item=i)]
        print('Processando : Mudar posição UP de arquivo selecionado.')
        #print('Qtd de arquivos disponiveis na lista: ' + str(self.list.GetItemCount())
        print('Qtd de arquivos selecionados: '+str(len(itemschecked)))
        print("ItemsChecked: " + str(itemschecked))
        for i in itemschecked:
            if i == 0: 
                print("\tPrimeiro item não pode subir mais")
                continue
            else:
                print("\tTrocar a posição "+str(i)+" com o anterior "+str(i-1))
                # salva os valores do elemento anterior
                nome1 = self.list.GetItemText(i-1)
                ex1 = self.list.GetItemText(i-1, col=1)
                tamanho1 = self.list.GetItemText(i-1, col=2)
                data1 = self.list.GetItemText(i-1, col=3)
                # salva os valores do elemento posterior
                nome2 = self.list.GetItemText(i)
                ex2 = self.list.GetItemText(i, col=1)
                tamanho2 = self.list.GetItemText(i, col=2)
                data2 = self.list.GetItemText(i, col=3)
                # reescreve anterior
                self.list.SetItem(i-1, 0, nome2)
                self.list.SetItem(i-1, 1, ex2)
                self.list.SetItem(i-1, 2, tamanho2)
                self.list.SetItem(i-1, 3, data2)
                # reescreve posterior
                self.list.SetItem(i, 0, nome1)
                self.list.SetItem(i, 1, ex1)
                self.list.SetItem(i, 2, tamanho1)
                self.list.SetItem(i, 3, data1)
                # ajusta itens selecionados
                self.list.CheckItem(i-1,True)
                self.list.CheckItem(i,False)

    def on_press_down(self, event):
        itemcount = self.list.GetItemCount()
        itemschecked = [i for i in range(itemcount) if self.list.IsItemChecked(item=i)]
        print('Processando : Mudar posição DOWN de arquivo selecionado.')
        #print('Qtd de arquivos disponiveis na lista: ' + str(self.list.GetItemCount())
        print('Qtd de arquivos selecionados: '+str(len(itemschecked)))
        print("ItemsChecked: " + str(itemschecked))
        for i in itemschecked:
            if i == itemcount-1: 
                print("\tUltimo item não pode descer mais")
                continue
            else:
                print("\tTrocar a posição "+str(i)+" com o posterior "+str(i+1))
                # salva os valores do elemento anterior
                nome1 = self.list.GetItemText(i+1)
                ex1 = self.list.GetItemText(i+1, col=1)
                tamanho1 = self.list.GetItemText(i+1, col=2)
                data1 = self.list.GetItemText(i+1, col=3)
                # salva os valores do elemento posterior
                nome2 = self.list.GetItemText(i)
                ex2 = self.list.GetItemText(i, col=1)
                tamanho2 = self.list.GetItemText(i, col=2)
                data2 = self.list.GetItemText(i, col=3)
                # reescreve anterior
                self.list.SetItem(i+1, 0, nome2)
                self.list.SetItem(i+1, 1, ex2)
                self.list.SetItem(i+1, 2, tamanho2)
                self.list.SetItem(i+1, 3, data2)
                # reescreve posterior
                self.list.SetItem(i, 0, nome1)
                self.list.SetItem(i, 1, ex1)
                self.list.SetItem(i, 2, tamanho1)
                self.list.SetItem(i, 3, data1)
                # ajusta itens selecionados
                self.list.CheckItem(i+1,True)
                self.list.CheckItem(i,False)

    def on_press_select_all(self, event):
        itemcount = self.list.GetItemCount()
        itemschecked = [i for i in range(itemcount) if self.list.IsItemChecked(item=i)]
        print('Processando : Selecionar todos os arquivo.')
        #print('Qtd de arquivos disponiveis na lista: ' + str(self.list.GetItemCount())
        print('Qtd de arquivos selecionados: '+str(len(itemschecked)))
        print("ItemsChecked: " + str(itemschecked))
        for i in range(itemcount):
            self.list.CheckItem(i,True)

    def on_press_select_clear(self, event):
        itemcount = self.list.GetItemCount()
        itemschecked = [i for i in range(itemcount) if self.list.IsItemChecked(item=i)]
        print('Processando : Limpar seleção de todos os arquivo.')
        #print('Qtd de arquivos disponiveis na lista: ' + str(self.list.GetItemCount())
        print('Qtd de arquivos selecionados: '+str(len(itemschecked)))
        print("ItemsChecked: " + str(itemschecked))
        for i in range(itemcount):
            self.list.CheckItem(i,False)

###
# Classe: AppFrame
# Eventos ao clicar nos botoes da Interface grafica
###
    def on_press_dir(self, event):
        # In this case we include a "New directory" button.
        dlg = wx.DirDialog(self, "Escolha o diretorio:",
                          style=wx.DD_DEFAULT_STYLE
                           #| wx.DD_DIR_MUST_EXIST
                           #| wx.DD_CHANGE_DIR
                           )
        # If the user selects OK, then we process the dialog's data.
        # This is done by getting the path data from the dialog - BEFORE
        # we destroy it.
        if dlg.ShowModal() == wx.ID_OK:
            #print('You selected: %s\n' % dlg.GetPath())
            self.text_ctrl_dir.ChangeValue(dlg.GetPath())
        # Only destroy a dialog after you're done with it.
        dir = self.text_ctrl_dir.GetValue()
        self.atualizarFileListCtrl(dir)
        # salvar novo diretorio no arquivo de configuracao json
        self.dados['diretorio_selecionado'] = dir
        print(self.dados)
        with open(self.arquivo_config, "w") as jsonFile:
            configJSON = json.dumps(self.dados)
            jsonFile.write(configJSON)
            jsonFile.close()
        # retirar janela que pergunta o dir
        dlg.Destroy()

    def on_press_merge(self, event):
        dir = self.text_ctrl_dir.GetValue()
        if not os.path.isdir(dir):
            print("Nenhum diretorio foi selecionado! \n: "+str(dir))
            return
        itemcount = self.list.GetItemCount()
        itemschecked = [i for i in range(itemcount) if self.list.IsItemChecked(item=i)]
        print('Processando : Merge de arquivos selecionados.')
        #print('Qtd de arquivos disponiveis na lista: ' + str(self.list.GetItemCount())
        print('Qtd de arquivos selecionados: '+str(len(itemschecked)))
        print("ItemsChecked: " + str(itemschecked))
        arquivos = []
        for index in itemschecked:
            item = self.list.GetItemText(index)
            ex = self.list.GetItemText(index, col=1)
            file = str(item)+'.'+str(ex)
            print("\tAnalisando item: "+str(file))          
            # Processar cada arquivo selecionado
            arquivos.append(file)
        if len(arquivos) <= 0:
            print("Nenhum arquivo foi selecionado! \n: "+str(arquivos))
            return
        # Nome do arquivo de saida
        ofile = self.text_ctrl_ofile.GetValue()
        if len(ofile) <= 0:
            print("Nome do arquivo de saida nao foi definido! \n: "+str(arquivos))
            return
        # Tamanho deve ser int
        tamanho = int(self.text_ctrl_size.GetValue())
        if tamanho < 1:
            print("Tamanho dos arquivos deve ser maior que 1! \n: "+str(tamanho))
            return
        print(f'Diretorio: "{dir}" | Processando arquivos: "{arquivos}"')
        print(f'Arquivo de saida: "{ofile}" | Tamanho Máximo: "{tamanho}"')
        resposta = PDF.merge(dir, arquivos, ofile,tamanho)
        # Dialog de resposta
        if resposta > 0:
            dlg = wx.MessageDialog(self, 'Arquivos juntados com sucesso: '+str(resposta),
                        'Juntar Arquivos',
                        wx.OK | wx.ICON_INFORMATION
                        #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                        )
            dlg.ShowModal()
            dlg.Destroy()
        self.atualizarFileListCtrl(dir)

    def on_press_size(self, event):
        dir = self.text_ctrl_dir.GetValue()
        itemcount = self.list.GetItemCount()
        itemschecked = [i for i in range(itemcount) if self.list.IsItemChecked(item=i)]
        print('Processando baseado no tamanho.')
        #print('Qtd de arquivos disponiveis na lista: ' + str(self.list.GetItemCount())
        print('Qtd de arquivos selecionados: '+str(len(itemschecked)))
        print("ItemsChecked: " + str(itemschecked))
        for index in itemschecked:
            item = self.list.GetItemText(index)
            ex = self.list.GetItemText(index, col=1)
            file = str(item)+'.'+str(ex)
            print("\tAnalisando item: "+str(file))          
            # Processar cada arquivo selecionado
            arquivo = dir+"/"+file
            # Tamanho deve ser int
            tamanho = int(self.text_ctrl_size.GetValue())
            if not ( os.path.isdir(dir) or os.path.isfile(arquivo) ):
                print("Diretorio ou arquivo invalido! \n: "+str(arquivo))
            else:
                print(f'Processando arquivo: "{arquivo}"')
                resposta = PDF.sizeSplit(arquivo, tamanho)
                # Dialog de resposta
                if resposta > 0:
                    dlg = wx.MessageDialog(self, 'Arquivos separados com sucesso: '+str(resposta),
                                'Separar Arquivos',
                                wx.OK | wx.ICON_INFORMATION
                                #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                                )
                    dlg.ShowModal()
                    dlg.Destroy()
        self.atualizarFileListCtrl(dir)

    def on_press_delete(self, event):
        dir = self.text_ctrl_dir.GetValue()
        itemcount = self.list.GetItemCount()
        itemschecked = [i for i in range(itemcount) if self.list.IsItemChecked(item=i)]
        print('Processando exclusão de páginas do intervalo.')
        #print('Qtd de arquivos disponiveis na lista: ' + str(self.list.GetItemCount())
        print('Qtd de arquivos selecionados: '+str(len(itemschecked)))
        print("ItemsChecked: " + str(itemschecked))
        for index in itemschecked:
            item = self.list.GetItemText(index)
            ex = self.list.GetItemText(index, col=1)
            file = str(item)+'.'+str(ex)
            print("\tAnalisando item: "+str(file))          
            # Processar cada arquivo selecionado
            arquivo = dir+"/"+file
            # Excluindo paginas do intervalo
            excluir_paginas = self.text_ctrl_delete.GetValue()
            if len(excluir_paginas) <= 0:
                print("Intervalo das páginas a serem excluidas dos arquivos nao foi definido! \n: "+str(excluir_paginas))
                return
            # Nome do arquivo de saida
            ofile = self.text_ctrl_ofile.GetValue()
            output_arquivo = dir+"/"+ofile
            if len(ofile) <= 0:
                print("Nome do arquivo de saida nao foi definido! \n: "+str(ofile))
                return
            resposta = PDF.pageSplit(arquivo, excluir_paginas, output_arquivo)
            # Dialog de resposta
            if resposta > 0:
                dlg = wx.MessageDialog(self, 'Paginas excluidas com sucesso: '+str(resposta)+ '\nCliquer para continuar e verificar o tamanho maximo',
                            'Excluir paginas',
                            wx.OK | wx.ICON_INFORMATION
                            #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                            )
                dlg.ShowModal()
                dlg.Destroy()
            # Depois processar para verificar tamanho maximo
            # Tamanho deve ser int
            '''
            tamanho = int(self.text_ctrl_size.GetValue())
            if not ( os.path.isdir(dir) or os.path.isfile(arquivo) ):
                print("Diretorio ou arquivo invalido! \n: "+str(arquivo))
            else:
                print(f'Processando arquivo: "{arquivo}"')
                resposta = PDF.sizeSplit(arquivo, tamanho)
                # Dialog de resposta
                if resposta > 0:
                    dlg = wx.MessageDialog(self, 'Arquivos separados com sucesso: '+str(resposta),
                                'Separar Arquivos',
                                wx.OK | wx.ICON_INFORMATION
                                #wx.YES_NO | wx.NO_DEFAULT | wx.CANCEL | wx.ICON_INFORMATION
                                )
                    dlg.ShowModal()
                    dlg.Destroy()
            '''
        self.atualizarFileListCtrl(dir)

if __name__ == '__main__':
    app = wx.App()
    frame = AppFrame()
    app.MainLoop()
