#!/usr/bin/env python

import os, wx, fitz, time, math
from natsort import natsorted
from pdf_split_tool import pdf_splitter

###
# Funcionalidade de MERGE
###
class pdfMerge():

    @staticmethod
    def merge(dir, ofile, maxsize=0):
        files = [f for f in os.listdir(dir) if f.endswith(".pdf")]
        # pip install natsort
        files = natsorted(files)
        # files = sorted(files)
        print("Diretorio: "+str(dir))
        print("Tamanho maximo dos arquivos de saida: "+str(maxsize))
        print("Arquivos ordenados: "+str(files))

        # Calculando o espaço total
        totalsize = 0
        parte = 1
        for pdf in files:
            print(f'Processando arquivo: "{pdf}"')
            # iniciando arquivo que vai receber o resultado do merge
            result = fitz.open()
            while totalsize < maxsize:
                filepath = dir+"/"+pdf
                size = os.path.getsize(filepath)
                # convertendo o tamanho para MB
                size = math.trunc(size/(1024*1024))
                if size < totalsize:
                    print(f'Incluindo arquivo com Tamanho do arquivo: "{size}"')
                    # adicionar arquivos
                    with fitz.open(filepath) as mfile:
                        result.insertPDF(mfile)
                    #total_pages = self.input_pdf.getNumPages()
                    #size = os.path.getsize(filename)
                    #avg_size = self.size / self.total_pages
                    #print(
                    #    "File: {}\nFile size: {}\nTotal pages: {}\nAverage size: {}".format(
                    #        filepath, self.size, self.total_pages, self.avg_size
                    #    )
                    #)
                    totalsize += size
                else :
                    # quebrar o proprio arquivo (recursivo)
                    result.save(dir+"/"+parte+"-"+ofile)
                    parte += 1
                    totalsize = 0
                    continue
            # atingiu o tamanho total
            # salvar esse arquivo e continuar (outras partes)
            # outra opção seria juntar tudo e depois quebrar chamando o pdfSplit()
            return
            result.save(dir+"/"+parte+"-"+ofile)
            parte += 1
            totalsize = 0
        print("Arquivos juntados!")

###
# Funcionalidade de SPLIT : POR PAGINAS ou POR TAMANHO
###
class pdfSplit():

    @staticmethod
    def sizeSplit(filename, max_size):
        filename = 'C:\\Users\\apsampaio\\Downloads\\4145916_1.pdf'
        print("sizeSplit : arquivo = "+ str(filename) +" , tamanho ="+ str(max_size))
        splitter = pdf_splitter.PdfSplitter(filename)
        splitter.split_max_size(max_size)

    #==============================================================================
    # Delimitadores, função auxiliar
    #==============================================================================
    def auxiliar(self, data):
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

    def pageSplit(self, filename, pages):
        #dir="C:/Users/apsampaio/OneDrive - Secretaria da Fazenda e Planejamento do Estado de São Paulo/DRTC-III/ICMS/_Operação/MONITORAMENTO/142.086.828.111 - BIOEXTRA INDUSTRIA E COMERCIO EIRELI/AIIMs/2 AIIM 4.145.916-7 ref FOX/"
        dir="H:/drtc-3/drtciii-nf2/03- CCQ - EQUIPE 23/AIIM 4.145.916-7_BIOEXTRA_ref_FOX/"

        fileInput="prova7_SFPPRC202104143V01-parte2.pdf"

        #incluir_paginas = "1-40; 43-44; 46-49"
        incluir_paginas = "1-3; 5-33"

        ifile=dir+fileInput
        ofile=dir+"selecao_"+fileInput
        # paths longos dao problema -> copiar para o M: ou H:
        #ofile=dir+"paginas_"+str(paginaInicial)+"_"+str(paginaFinal)+"_"+fileInput

        doc = fitz.open(ifile)
        pages = auxiliar(incluir_paginas) # ajustar os indices começando em zero

        # exemplo de leitura de texto do documento
        #for page in doc:
        #    text = page.getText()

        print("\tDocumento '%s' tem %i paginass." % (doc.name, len(doc)))

        print("\t %i Paginas que vao ficar no documento: %s" % (len(pages), str(pages)))

        #gerando novo pdf
        docout = fitz.open()

        # get the pages
        try : # inserir paginas
            doc.select(pages)                   # delete all others
            doc.save(ofile)                     # save and clean new PDF
            doc.close()
        except Exception as e1: #falha na tentativa
            print("\t\t\t\tErro no processamento com mensagem de erro: ",str(e1))

###
# Interface grafica
###
class MyFrame(wx.Frame):    
    def __init__(self):
        # Init de componentes
        super().__init__(parent=None, title='SEFAZ PDF Utils')
        panel = wx.Panel(self)        
        my_sizer = wx.BoxSizer(wx.VERTICAL)
        # Lista de arquivos
        #self.list = CheckListCtrl(panel)
        self.list = wx.ListCtrl(panel, size=(300,180), style=wx.LC_REPORT)
        self.list.EnableCheckBoxes()
        my_sizer.Add(self.list, 0, wx.EXPAND)
        # Descreve o tamanho total
        self.text_total_size = wx.TextCtrl(panel, style=wx.TE_READONLY | wx.TE_CENTRE)
        #self.text_ctrl_size.AppendText("Tamanho estimado (soma dos arquivos) = "+ str(f'{tamanhoTotal/(1024*1024):.2f}') +" MBs")
        my_sizer.Add(self.text_total_size, 0, wx.ALL | wx.EXPAND, 5)
        # Atualiza a Lista de arquivos e o campo com o tamanho total
        self.atualizarFileListCtrl(diretorio='.')
        #self.splitter = wx.SplitterWindow(self, ID_SPLITTER, style=wx.SP_BORDER)
        # Diretorio selecionado
        self.text_ctrl_dir = wx.TextCtrl(panel, style=wx.TE_READONLY)
        self.text_ctrl_dir.AppendText("Clique no botão abaixo para selecionar um diretorio")
        my_sizer.Add(self.text_ctrl_dir, 0, wx.ALL | wx.EXPAND, 5)
        # Botao para selecionar Dir
        my_btn_dir = wx.Button(panel, label='Mudar diretorio')
        my_btn_dir.Bind(wx.EVT_BUTTON, self.on_press_dir)
        my_sizer.Add(my_btn_dir, 0, wx.ALL | wx.CENTER, 5)
        # Nome do Arquivo de saida
        self.text_ctrl_ofile = wx.TextCtrl(panel)
        self.text_ctrl_ofile.AppendText("_output.pdf")
        my_sizer.Add(self.text_ctrl_ofile, 0, wx.ALL | wx.EXPAND, 5)
        # Botao para processar        
        my_btn_merge = wx.Button(panel, label='Fazer o merge dos arquivos')
        my_btn_merge.Bind(wx.EVT_BUTTON, self.on_press_process)
        my_sizer.Add(my_btn_merge, 0, wx.ALL | wx.CENTER, 5)
        # Tamanho do Arquivo de saida em Mb
        self.text_ctrl_size = wx.TextCtrl(panel)
        self.text_ctrl_size.AppendText("10")
        my_sizer.Add(self.text_ctrl_size, 0, wx.ALL | wx.EXPAND, 5)
        # Botao para processar        
        my_btn_size = wx.Button(panel, label='Gerar arquivos com tamanhos máximos em MB.')
        my_btn_size.Bind(wx.EVT_BUTTON, self.on_press_size)
        my_sizer.Add(my_btn_size, 0, wx.ALL | wx.CENTER, 5)
        # Atualizar tela
        #size = wx.DisplaySize()
        #self.SetSize(size)
        panel.SetSizer(my_sizer)        
        self.Show()

    def atualizarFileListCtrl(self, diretorio):
        j = 0
        totalsize = 0
        self.list.ClearAll()
        # Criar campos:
        self.list.InsertColumn(0, 'Nome')
        self.list.InsertColumn(1, 'Extensão')
        self.list.InsertColumn(2, 'Tamanho', wx.LIST_FORMAT_RIGHT)
        self.list.InsertColumn(3, 'Data motificação')
        self.list.SetColumnWidth(0, 120)
        self.list.SetColumnWidth(1, 60)
        self.list.SetColumnWidth(2, 70)
        self.list.SetColumnWidth(3, 110)
        #self.list.InsertItem(0, '..')
        # Exibir arquivos
        files = os.listdir(diretorio)
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
# Classe: MyFrame
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
            #self.log.WriteText('You selected: %s\n' % dlg.GetPath())
            self.text_ctrl_dir.ChangeValue(dlg.GetPath())
        # Only destroy a dialog after you're done with it.
        dir = self.text_ctrl_dir.GetValue()
        self.atualizarFileListCtrl(dir)
        dlg.Destroy()

    def on_press_process(self, event):
        dir = self.text_ctrl_dir.GetValue()
        tamanhoMax = int(self.text_ctrl_size.GetValue())
        if not os.path.isdir(dir):
            print("Nenhum diretorio foi selecionado! \n: "+str(dir))
        else:
            ofile = self.text_ctrl_ofile.GetValue()
            #print(f'Processando diretorio: "{dir}"')
            #print(f'Arquivo de saida: "{ofile}"')
            pdfMerge.merge(dir, ofile, tamanhoMax)

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
                pdfSplit.sizeSplit(arquivo, tamanho)

if __name__ == '__main__':
    app = wx.App()
    frame = MyFrame()
    app.MainLoop()
