from PyQt6.QtWidgets import QMainWindow, QApplication, QFileDialog
from PyQt6 import QtWidgets, QtCore
from PyQt6 import QtGui
from PyQt6.QtCore import Qt
from openpyxl import load_workbook
from openpyxl import Workbook
import datetime
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *

import sys
import os
import pandas as pd
from menu_ui_ui import Ui_MainWindow
from SplashScreen_ui import Ui_SplashScreen

############################################################################################
##################################   FONTE DA FUNÇÃO ABAIXO    #############################
############################################################################################

# #https://stackoverflow.com/questions/31836104/pyinstaller-and-onefile-how-to-include-an-image-in-the-exe-file
# def resource_path(relative_path):
#     """ Get absolute path to resource, works for dev and for PyInstaller """
#     try:
#         # PyInstaller creates a temp folder and stores path in _MEIPASS
#         base_path = sys._MEIPASS2 #LEMBRAR DE ACRESCENTAR ESSE 2 NO FINLA
#     except Exception:
#         base_path = os.path.abspath(".")

#     return os.path.join(base_path, relative_path)


#DONE: MENU PARA SALVAR
#DONE: splash screen com o DOM do IAOp
#DONE: Deixar responsivo
#DONE: Passar para a próxima linha do painel da esquerda, assim a om for escolhida
#DONE: O QUE ACONTECE SE O USUÁRIO COLOCAR UMA OM E DEPOIS TROCAR? O CÓDIGO VAI FICAR TROCANDO AS VAGAS E EXISTENTES TODA A VEZ? MELHOR REVER COMO LEVANTAR AS VAGAS, TALVEZ SEJA MELHOR FAZER UM LEVANTAMENTO SOMANDO COM A COLUNA PLAMOV

#DONE: alterar a OM ao clicar no painel esquerdo
#????: (VALE A PENA?, JA QUE A LINHA ATIVA VAI PASSAR PARA A PRÓXIMA E O PAINEL VAI SER ATUALIZADO) TAFERA: mudar as vagas e taxa de ocupação assim que o usuário selecionar a OM de destino
#DONE: design gráfico
#????: Dividir os menus em "Páginas" e "Carregar"
#DONE: Colorir as OM das localidades escolhidas e a OM atual
#DONE: Reduzir a quantidade de colunas do painel esquerdo
#TODO: apagar as linhas que não tem TP no painel direito
#DONE: carregar os dois arquivos ao mesmo tempo
#DONE: Checar o tempo de execução de cada função e escovar para diminuir
#TODO: Apagar as linhas em branco do df_plamov (linhas que tem NaT)

caminho_atual = os.getcwd()
status_painel = ""
linha_alterada = -1
coluna_alterada = -1


def classificar (dataframe: pd.DataFrame):
    return dataframe.sort_values(by=['MELHOR PRIO', 'TEMPO LOC', 'ANTIGUIDADE'], ascending=[True, False, True], inplace=True)
    
def classificar_ordem_original (dataframe: pd.DataFrame):
    return dataframe.sort_values(by=['ordem original'], inplace=True)

def pegar_quadro(linha):
    global df_plamov_compilado
    quadro = df_plamov_compilado["QUADRO"][int(linha)]
    return quadro
def pegar_especialidade(linha):
    especialidade = df_plamov_compilado["ESP"][int(linha)]
    return especialidade
def pegar_subespecialidade(linha):
    try:
        sub = df_plamov_compilado["SUB ESP"][int(linha)]
        return str(sub).strip() # Remove espaços extras por segurança
    except:
        return ""
def pegar_posto(linha):
    if df_plamov_compilado["POSTO"][int(linha)] == "1S"\
        or df_plamov_compilado["POSTO"][int(linha)] == "2S"\
        or df_plamov_compilado["POSTO"][int(linha)] == "3S"\
        or df_plamov_compilado["POSTO"][int(linha)] == "SO":
        posto = "SGT"
    elif df_plamov_compilado["POSTO"][int(linha)] == "1T"\
        or df_plamov_compilado["POSTO"][int(linha)] == "2T":
        posto = "TN"
    else:
        posto = df_plamov_compilado["POSTO"][int(linha)]
    return posto
def pegar_LOC1(linha):
    loc1 = df_plamov_compilado["LOC 1"][int(linha)]
    return loc1
def pegar_LOC2(linha):
    loc2 = df_plamov_compilado["LOC 2"][int(linha)]
    return loc2
def pegar_LOC3(linha):
    loc3 = df_plamov_compilado["LOC 3"][int(linha)]
    return loc3
def pegar_LOC_atual(linha):
    loc_atual = df_plamov_compilado["LOC ATUAL"][int(linha)]
    return loc_atual

        
def pegar_OMs_do_COMPREP():
    global df_relatorio_tp
    global df_OMs
    df_relatorio_tp = pd.read_excel(endereco_do_arquivo, sheet_name="RELATÓRIO TP")
    df_OMs = df_relatorio_tp['Unidade']
    df_OMs.drop_duplicates(inplace=True)
    df_OMs.dropna(inplace=True)
    df_OMs.reset_index(drop=True, inplace=True)
    df_OMs = df_OMs.to_frame(name="OMs")
    df_OMs["Taxa de Ocup."] = ""
    df_OMs["Vagas"] = ""
    df_OMs["Localidade"] = ""
    return df_OMs
counter = 0
class SplashScreen (QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.ui = Ui_SplashScreen()
        self.ui.setupUi(self)

        self.setWindowFlags(QtCore.Qt.WindowType.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_TranslucentBackground)
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update)
        self.timer.start(25)

        self.show()
    
    def update(self):
        global counter
        self.ui.progressBar.setValue(counter)
        if counter >= 100:
            self.timer.stop()
            self.main = UI()
            self.main.show()

            self.close()
        counter += 1


class UI(QMainWindow):
    global df_plamov_compilado

    def __init__(self):
        QMainWindow.__init__(self)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
    
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        
        self.ui.stackedWidget.setCurrentIndex(0) #para inicializar na página dos militares

        self.ui.actionMilitares.triggered.connect(lambda: self.Pag_Militares())
        self.ui.actionQuadros_Especialidades.triggered.connect(lambda: self.Pag_Quadros_Especialidades())
        self.ui.actionRelat_rio_TP.triggered.connect(lambda: self.Pag_Relat_rio_TP())
        self.ui.actionMapa.triggered.connect(lambda: self.Pag_Mapa())

        self.ui.actionDados_dos_militares.triggered.connect(lambda: self.Carregar_Dados_dos_militares())
        self.ui.actionRelat_rio_TP_2.triggered.connect(lambda: self.carregar_Relat_rio_TP())
        self.ui.actionSALVAR.triggered.connect(lambda: self.salvar())
        self.ui.tableWidget.cellClicked.connect(lambda: self.linha_ativa_dados_militares())
        self.ui.tableWidget.cellClicked.connect(lambda: self.coluna_ativa_dados_militares())
        self.ui.tableWidget.cellClicked.connect(lambda: self.atualizar_Painel_Direita())
        self.ui.tableWidget.cellChanged.connect(self.celula_alterada)
        self.ui.tableWidget_2.cellDoubleClicked.connect(lambda: self.escolher_OM_no_painel_direito())


        self.show()

    def salvar (self):
        #TODO: Esta função cria uma arquivo novo para salvar a relação de oms escolhidas para cada militar durante a execução do código. MUDAR PARA ESCREVER DIRETAMENTE NO ARQUIVO EXCEL DO PLAMOV.
        df_plamov_compilado.sort_values(by=['ordem original'], ascending=[True], inplace=True)
        lista = df_plamov_compilado["PLAMOV"].values.tolist()
        arquivo_excel = Workbook()
        planilha = arquivo_excel.active
        data_hora_atual = datetime.datetime.now()
        data_hora_formatada = data_hora_atual.strftime('%d-%Y-%m %H.%M.%S')
        endereco_do_arquivo_novo = os.path.dirname(endereco_do_arquivo)
        nome_completo_arquivo_novo = f"{endereco_do_arquivo_novo}/TESSARATO (SALVO EM) {data_hora_formatada}.xlsx"
        for i in range(len(lista)):
            planilha[F"B{i+1}"] = lista[i]

        arquivo_excel.save(filename=nome_completo_arquivo_novo)
        arquivo_excel.close()

    def celula_alterada(self, linha, coluna):
        global linha_alterada
        global coluna_alterada
        
        if status_painel == "carregado":
            linha_alterada = linha
            coluna_alterada = coluna
            if coluna_alterada == 12:
                df_plamov_compilado.loc[linha_alterada, "PLAMOV"] = self.ui.tableWidget.item(linha_alterada, coluna_alterada).text()   
            
    #passar as páginas
    def Pag_Militares(self):
        self.ui.stackedWidget.setCurrentIndex(0)
    def Pag_Quadros_Especialidades(self):
        self.ui.stackedWidget.setCurrentIndex(1)
    def Pag_Relat_rio_TP(self):
        self.ui.stackedWidget.setCurrentIndex(2)
    def Pag_Mapa(self):
        self.ui.stackedWidget.setCurrentIndex(3)
   
    # def atualizar_Painel_Direita (self):
    #     global df_OMs
    #     linha = self.linha_ativa_dados_militares()
    #     posto = pegar_posto(linha)
    #     quadro = pegar_quadro(linha)
    #     especialidade = pegar_especialidade(linha)
    #     loc1 = pegar_LOC1(linha)
    #     loc2 = pegar_LOC2(linha)
    #     loc3 = pegar_LOC3(linha)
    #     loc_atual = pegar_LOC_atual(linha)
    #     self.ui.tableWidget_2.setColumnCount(3)
    #     self.ui.tableWidget_2.setRowCount(df_OMs.shape[0]) 
    #     self.ui.tableWidget_2.setHorizontalHeaderLabels(["OM", "Taxa de Ocup.", "Vagas"])

    def atualizar_Painel_Direita (self):
        global df_OMs
        global df_TP_BMA
        global df_plamov_compilado
        
        linha = self.linha_ativa_dados_militares()
        
        # --- DADOS DO MILITAR (Vêm do Painel Esquerdo / PLAMOV COMPILADO) ---
        posto = pegar_posto(linha)
        quadro = pegar_quadro(linha)
        especialidade = pegar_especialidade(linha)
        subespecialidade = pegar_subespecialidade(linha) # Fundamental para BMA
        
        loc1 = pegar_LOC1(linha)
        loc2 = pegar_LOC2(linha)
        loc3 = pegar_LOC3(linha)
        loc_atual = pegar_LOC_atual(linha)

        # Configura tabela visual
        self.ui.tableWidget_2.setColumnCount(3)
        self.ui.tableWidget_2.setRowCount(df_OMs.shape[0]) 
        self.ui.tableWidget_2.setHorizontalHeaderLabels(["OM", "Taxa de Ocup.", "Vagas"])

        for k in range(df_OMs.shape[0]):
            chegando = 0
            saindo = 0
            
            # ==============================================================================
            # LÓGICA BMA (CRUZAMENTO PLAMOV + TP BMA)
            # ==============================================================================
# ==============================================================================
            # LÓGICA BMA (CRUZAMENTO PLAMOV + TP BMA)
            # ==============================================================================
            if especialidade == "BMA":
                # 1. SEGURANÇA: Garante que a coluna existe
                if 'Subespecialidade' not in df_TP_BMA.columns:
                    print(f"ERRO: Coluna 'Subespecialidade' não encontrada. Colunas: {df_TP_BMA.columns.tolist()}")
                    return

                # 2. FILTRO: Busca a vaga na TP BMA usando a nova coluna Subespecialidade
                filtro_bma = (
                    (df_TP_BMA['Unidade'] == df_OMs.iloc[k,0]) & 
                    (df_TP_BMA['Posto'] == posto) & 
                    (df_TP_BMA['Quadro'] == quadro) & 
                    (df_TP_BMA['Subespecialidade'] == subespecialidade)
                )
                vagas_OM_selecionada = df_TP_BMA[filtro_bma]
                
                # 3. SE ACHAR DADOS NA BMA
                if not vagas_OM_selecionada.empty:
                    # Pega os valores usando índices numéricos (conforme seu print do terminal)
                    # 10=TLP, 11=Existentes, 12=Vagas
                    try:
                        TP = vagas_OM_selecionada.iloc[0, 12]        # Coluna Vagas
                        existentes = vagas_OM_selecionada.iloc[0, 11] # Coluna Existentes
                        tlp_valor = vagas_OM_selecionada.iloc[0, 10]  # Coluna TLP
                        
                        df_OMs.loc[k, "Existentes"] = existentes
                        df_OMs.loc[k, "Vagas"] = TP
                        
                        # CÁLCULO DA TAXA (FEITO AQUI DENTRO PARA NÃO DAR ERRO)
                        if tlp_valor > 0:
                            df_OMs.loc[k, "Taxa de Ocup."] = round(float(existentes) / float(tlp_valor), 4) * 100
                        else:
                            df_OMs.loc[k, "Taxa de Ocup."] = 0.0
                            
                    except Exception as e:
                        print(f"Erro de índice BMA: {e}")
                        df_OMs.loc[k, "Taxa de Ocup."] = 0.0

                # 4. SE NÃO ACHAR DADOS (OM sem previsão para essa subespecialidade)
                else:
                    df_OMs.loc[k, "Existentes"] = 0
                    df_OMs.loc[k, "Vagas"] = 0
                    df_OMs.loc[k, "Taxa de Ocup."] = 0.0

            # ==============================================================================
            # LÓGICA PARA OUTROS QUADROS (Mantida original)
            # ==============================================================================
            else:
                if posto == "CP":
                    vagas_OM_selecionada = df_TP.query(f"Unidade == '{df_OMs.iloc[k,0]}' & ((Posto == 'CP/TN') | (Posto == 'CP')) & Quadro == '{quadro}' & Especialidade == '{especialidade}'")
                elif posto == "TN":
                    vagas_OM_selecionada = df_TP.query(f"Unidade == '{df_OMs.iloc[k,0]}' & ((Posto == 'CP/TN') | (Posto == 'TN')) & Quadro == '{quadro}' & Especialidade == '{especialidade}'")
                else:
                    vagas_OM_selecionada = df_TP.query(f"Unidade == '{df_OMs.iloc[k,0]}' & Posto == '{posto}' & Quadro == '{quadro}' & Especialidade == '{especialidade}'")
                
                if not vagas_OM_selecionada.empty:
                    df_OMs.loc[k,"Localidade"] = vagas_OM_selecionada.iloc[0,0] # Pega localidade da TP Geral

                    if posto == "CP" or posto == "TN":
                        chegando = df_plamov_compilado.query(f"PLAMOV == '{df_OMs.iloc[k,0]}' & POSTO == '{posto}' & QUADRO == '{quadro}' & ESP == '{especialidade}'").shape[0]
                        saindo = df_plamov_compilado.query(f"`OM ATUAL` == '{df_OMs.iloc[k,0]}' & POSTO == '{posto}' & QUADRO == '{quadro}' & ESP == '{especialidade}' & PLAMOV != ''").shape[0]
                    else:
                        chegando = df_plamov_compilado.query(f"PLAMOV == '{df_OMs.iloc[k,0]}' & POSTO == '{posto}' & QUADRO == '{quadro}' & ESP == '{especialidade}'").shape[0]
                        saindo = df_plamov_compilado.query(f"`OM ATUAL` == '{df_OMs.iloc[k,0]}' & POSTO == '{posto}' & QUADRO == '{quadro}' & ESP == '{especialidade}' & PLAMOV != ''").shape[0]

                    # Índices TP Geral (conforme seu código original)
                    TP = vagas_OM_selecionada.iloc[0,15] 
                    # Em vez de confiar que a coluna 11 é Existentes:
                    try:
                        existentes_na_TP = vagas_OM_selecionada.iloc[0]['Existentes']
                    except:
                        # Fallback apenas se der erro no nome
                        existentes_na_TP = vagas_OM_selecionada.iloc[0, 11]
                        existentes = existentes_na_TP + chegando - saindo

                    df_OMs.loc[k,"Vagas"] = TP + saindo - chegando
                    existentes = existentes_na_TP + chegando - saindo

                    if vagas_OM_selecionada.iloc[0,10] != 0:    
                        df_OMs.loc[k,"Taxa de Ocup."] = round(float(existentes)/float(vagas_OM_selecionada.iloc[0,10]), 4) * 100
                    else:
                        df_OMs.loc[k,"Taxa de Ocup."] = "Sem TP"
                        df_OMs.loc[k,"Vagas"] = ""
                else:
                    df_OMs.loc[k,"Taxa de Ocup."] = "Sem TP"
                    df_OMs.loc[k,"Vagas"] = ""

        # ==============================================================================
        # PREENCHIMENTO VISUAL (Comum a todos)
        # ==============================================================================
        df_OMs.sort_values(by=['Taxa de Ocup.', 'Vagas'], ascending=[True, False], inplace=True)
        df_OMs.reset_index(drop=True, inplace=True)

        for i in range(df_OMs.shape[0]):
            for j in range(3):
                item = QtWidgets.QTableWidgetItem(str(df_OMs.iloc[i,j]))
                self.ui.tableWidget_2.setItem(i,j, item)
                
                if i%2:
                    self.ui.tableWidget_2.item(i, j).setBackground(QtGui.QColor(100, 139, 245))
                
                om_loc = str(df_OMs.iloc[i,3]).strip().upper()
                l1 = str(loc1).strip().upper()
                l2 = str(loc2).strip().upper()
                l3 = str(loc3).strip().upper()
                
                if om_loc == l3 and l3 != "":
                    self.ui.tableWidget_2.item(i, j).setBackground(QtGui.QColor(255, 0, 255))
                if om_loc == l2 and l2 != "":
                    self.ui.tableWidget_2.item(i, j).setBackground(QtGui.QColor(255, 243, 8))
                if om_loc == l1 and l1 != "":
                    self.ui.tableWidget_2.item(i, j).setBackground(QtGui.QColor(29, 181, 2))
            
            if str(df_OMs.iloc[i,0]).strip().upper() == str(loc_atual).strip().upper():
                item = QtWidgets.QTableWidgetItem(str(df_OMs.iloc[i,0]))
                self.ui.tableWidget_2.setItem(i,0, item)
                self.ui.tableWidget_2.item(i, 0).setBackground(QtGui.QColor(107, 107, 106))

        df_OMs["Taxa de Ocup."] = ""
        df_OMs["Vagas"] = ""

        
        for k in range(df_OMs.shape[0]):
            chegando = 0
            saindo = 0
            if posto == "CP":
                vagas_OM_selecionada  = df_TP.query(f"Unidade == '{df_OMs.iloc[k,0]}' & ((Posto == 'CP/TN') | (Posto == 'CP')) & Quadro == '{quadro}' & Especialidade == '{especialidade}'")
                if not vagas_OM_selecionada.empty:
                    
                    chegando = df_plamov_compilado.query(f"PLAMOV == '{df_OMs.iloc[k,0]}' & POSTO == 'CP' & QUADRO == '{quadro}' & ESP == '{especialidade}'").shape[0]
                    saindo = df_plamov_compilado.query(f"`OM ATUAL` == '{df_OMs.iloc[k,0]}' & POSTO == 'CP' & QUADRO == '{quadro}' & ESP == '{especialidade}' & PLAMOV != ''").shape[0]
                    TP = vagas_OM_selecionada.iloc[0,15]
                    df_OMs.loc[k,"Vagas"] = TP + saindo - chegando
                    # Em vez de confiar que a coluna 11 é Existentes:
                    try:
                        existentes_na_TP = vagas_OM_selecionada.iloc[0]['Existentes']
                    except:
                        # Fallback apenas se der erro no nome
                        existentes_na_TP = vagas_OM_selecionada.iloc[0, 11]
                        existentes = existentes_na_TP + chegando - saindo
                    existentes = existentes_na_TP + chegando - saindo
                    

                    #PEGA A LOCALIDADE DA OM
                    df_OMs.loc[k,"Localidade"] = vagas_OM_selecionada.iloc[0,0]
                    
                    #SE A TP PARA ESSAS 3 DIMENSÕES NÃO ESTIVER ZERADA, É FEITO O CÁLCULO DA TAXA DE OCUP.
                    if vagas_OM_selecionada.iloc[0,10] != 0:     
                        df_OMs.loc[k,"Taxa de Ocup."] = round(float(existentes)/float(vagas_OM_selecionada.iloc[0,10]), 4) * 100
                        
                    #SE FOR ZERADA, É APRESENTADO "SEM TP"
                    else:
                        df_OMs.loc[k,"Taxa de Ocup."] = "Sem TP"
                        df_OMs.loc[k,"Vagas"] = ""
                else:
                    #TRABALHA A CONDIÇÃO DE A QUERY NÃO RETORNAR NADA, OU SEJA, NÃO EXISTE ESSA COMBINAÇÃO NA TABELA TP
                    df_OMs.loc[k,"Taxa de Ocup."] = "Sem TP"
                    df_OMs.loc[k,"Vagas"] = ""
                    for i in range(3):
                        item = QtWidgets.QTableWidgetItem(str(df_OMs.iloc[k,i]))
                        self.ui.tableWidget_2.setItem(k,i, item)
            elif posto == "TN":
                vagas_OM_selecionada  = df_TP.query(f"Unidade == '{df_OMs.iloc[k,0]}' & ((Posto == 'CP/TN') | (Posto == 'TN')) & Quadro == '{quadro}' & Especialidade == '{especialidade}'")
                if not vagas_OM_selecionada.empty:

                    chegando = df_plamov_compilado.query(f"PLAMOV == '{df_OMs.iloc[k,0]}' & (POSTO == 'TN') & QUADRO == '{quadro}' & ESP == '{especialidade}'").shape[0]
                    saindo = df_plamov_compilado.query(f"`OM ATUAL` == '{df_OMs.iloc[k,0]}' & (POSTO == 'TN') & QUADRO == '{quadro}' & ESP == '{especialidade}' & PLAMOV != ''").shape[0]
                    TP = vagas_OM_selecionada.iloc[0,15]
                    df_OMs.loc[k,"Vagas"] = TP + saindo - chegando
                    # Em vez de confiar que a coluna 11 é Existentes:
                    try:
                        existentes_na_TP = vagas_OM_selecionada.iloc[0]['Existentes']
                    except:
                        # Fallback apenas se der erro no nome
                        existentes_na_TP = vagas_OM_selecionada.iloc[0, 11]
                        existentes = existentes_na_TP + chegando - saindo
                    existentes = existentes_na_TP + chegando - saindo
                    #PEGA A LOCALIDADE DA OM
                    df_OMs.loc[k,"Localidade"] = vagas_OM_selecionada.iloc[0,0]
                    
                    #SE A TP PARA ESSAS 3 DIMENSÕES NÃO ESTIVER ZERADA, É FEITO O CÁLCULO DA TAXA DE OCUP.
                    if vagas_OM_selecionada.iloc[0,10] != 0:     
                        df_OMs.loc[k,"Taxa de Ocup."] = round(float(existentes)/float(vagas_OM_selecionada.iloc[0,10]), 4) * 100
                    #SE FOR ZERADA, É APRESENTADO "SEM TP"
                    else:
                        df_OMs.loc[k,"Taxa de Ocup."] = "Sem TP"
                        df_OMs.loc[k,"Vagas"] = ""
                else:
                    #TRABALHA A CONDIÇÃO DE A QUERY NÃO RETORNAR NADA, OU SEJA, NÃO EXISTE ESSA COMBINAÇÃO NA TABELA TP
                    df_OMs.loc[k,"Taxa de Ocup."] = "Sem TP"
                    df_OMs.loc[k,"Vagas"] = ""
                    for i in range(3):
                        item = QtWidgets.QTableWidgetItem(str(df_OMs.iloc[k,i]))
                        self.ui.tableWidget_2.setItem(k,i, item)
            else:
                vagas_OM_selecionada  = df_TP.query(f"Unidade == '{df_OMs.iloc[k,0]}' & Posto == '{posto}' & Quadro == '{quadro}' & Especialidade == '{especialidade}'")
                if not vagas_OM_selecionada.empty:

                    chegando = df_plamov_compilado.query(f"PLAMOV == '{df_OMs.iloc[k,0]}' & POSTO == '{posto}' & QUADRO == '{quadro}' & ESP == '{especialidade}'").shape[0]
                    saindo = df_plamov_compilado.query(f"`OM ATUAL` == '{df_OMs.iloc[k,0]}' & POSTO == '{posto}' & QUADRO == '{quadro}' & ESP == '{especialidade}' & PLAMOV != ''").shape[0]
                    TP = vagas_OM_selecionada.iloc[0,15]
                    df_OMs.loc[k,"Vagas"] = TP + saindo - chegando
                    # Em vez de confiar que a coluna 11 é Existentes:
                    try:
                        existentes_na_TP = vagas_OM_selecionada.iloc[0]['Existentes']
                    except:
                        # Fallback apenas se der erro no nome
                        existentes_na_TP = vagas_OM_selecionada.iloc[0, 11]
                        existentes = existentes_na_TP + chegando - saindo

                    #PEGA A LOCALIDADE DA OM
                    df_OMs.loc[k,"Localidade"] = vagas_OM_selecionada.iloc[0,0]
                    
                    #SE A TP PARA ESSAS 3 DIMENSÕES NÃO ESTIVER ZERADA, É FEITO O CÁLCULO DA TAXA DE OCUP.
                    if vagas_OM_selecionada.iloc[0,10] != 0:     
                        df_OMs.loc[k,"Taxa de Ocup."] = round(float(existentes)/float(vagas_OM_selecionada.iloc[0,10]), 4) * 100
                    #SE FOR ZERADA, É APRESENTADO "SEM TP"
                    else:
                        df_OMs.loc[k,"Taxa de Ocup."] = "Sem TP"
                        df_OMs.loc[k,"Vagas"] = ""
                else:
                    #TRABALHA A CONDIÇÃO DE A QUERY NÃO RETORNAR NADA, OU SEJA, NÃO EXISTE ESSA COMBINAÇÃO NA TABELA TP
                    df_OMs.loc[k,"Taxa de Ocup."] = "Sem TP"
                    df_OMs.loc[k,"Vagas"] = ""
                    for i in range(3):
                        item = QtWidgets.QTableWidgetItem(str(df_OMs.iloc[k,i]))
                        self.ui.tableWidget_2.setItem(k,i, item)


        #DESCRIÇÃO: ORDENAÇÃO
        df_OMs.sort_values(by=['Taxa de Ocup.', 'Vagas'], ascending=[True, False], inplace=True)
        df_OMs.reset_index(drop=True, inplace=True)

        #DESCRIÇÃO: COLOCA OS VALORES NA TABELA DA DIREIRA
        for i in range(df_OMs.shape[0]):
            for j in range(3):
                item = QtWidgets.QTableWidgetItem(str(df_OMs.iloc[i,j]))
                self.ui.tableWidget_2.setItem(i,j, item)
                #DESCRIÇÃO: COLORE AS LOCALIDADES E AS LINHAS PARES
                if i%2:
                    self.ui.tableWidget_2.item(i, j).setBackground(QtGui.QColor(100, 139, 245))
                if (df_OMs.iloc[i,3]) == loc3:
                    self.ui.tableWidget_2.item(i, j).setBackground(QtGui.QColor(255, 0, 255))
                if (df_OMs.iloc[i,3]) == loc2:
                    self.ui.tableWidget_2.item(i, j).setBackground(QtGui.QColor(255, 243, 8))
                if (df_OMs.iloc[i,3]) == loc1:
                    self.ui.tableWidget_2.item(i, j).setBackground(QtGui.QColor(29, 181, 2))
            if (df_OMs.iloc[i,3]) == loc_atual:
                    item = QtWidgets.QTableWidgetItem(str(df_OMs.iloc[i,0]))
                    self.ui.tableWidget_2.setItem(i,0, item)
                    self.ui.tableWidget_2.item(i, 0).setBackground(QtGui.QColor(107, 107, 106))
        df_OMs["Taxa de Ocup."] = ""
        df_OMs["'Vagas'"] = ""
        df_OMs["Localidade"] = ""

    ####################################

    def Abrir_Dialogo_Carregar_Dados(self):
        resultado = QFileDialog.getOpenFileName(self, "Qual arquivo gostaria de carregar?", caminho_atual, 'Excel files (*.xlsx)')
        endereco_do_arquivo = resultado[0]  # obtém o endereço do arquivo do resultado
        if endereco_do_arquivo:  # verifica se o endereço do arquivo não é vazio
            self.Carregar_Dados_dos_militares()  # chama a função para carregar os dados

    # def Carregar_Dados_dos_militares(self):
    #     global endereco_do_arquivo
    #     global df_OMs
    #     global df_plamov_compilado
    #     global status_painel 
    #     endereco_do_arquivo = QFileDialog.getOpenFileName(self, "Qual arquivo gostaria de carregar?", caminho_atual, 'Excel files (*.xlsx)')[0]
    #     if endereco_do_arquivo:
    #         df_plamov_compilado = pd.read_excel(endereco_do_arquivo, sheet_name="PLAMOV COMPILADO")
    #         df_plamov_compilado = df_plamov_compilado.fillna("") #Troca os "NaN" valor vazio

    #         df_plamov_compilado['ordem original'] = df_plamov_compilado.index
            
    #         self.ui.tableWidget.setColumnCount(12) #quantidade de colunas
    #         self.ui.tableWidget.setRowCount(df_plamov_compilado.shape[0]) #quantidade de linhas
    #         # self.ui.tableWidget.setHorizontalHeaderLabels(df_plamov_compilado.columns.to_list())

    #         df_plamov_compilado = df_plamov_compilado.sort_values(by=['MELHOR PRIO', 'TEMPO LOC', 'ANTIGUIDADE'], ascending=[True, False, True])
    #         df_plamov_compilado = df_plamov_compilado.reset_index(drop=True)

    #         #rodar um for para pegar os header e montar uma lista
    #         lista_nome_colunas_painel_esquerda = []
    #         for j in range(1, 7):
    #             lista_nome_colunas_painel_esquerda.append(df_plamov_compilado.columns[j])
    #         for j in range(13, 16):
    #             lista_nome_colunas_painel_esquerda.append(df_plamov_compilado.columns[j]) 
    #         for j in range(18, 20):
    #             lista_nome_colunas_painel_esquerda.append(df_plamov_compilado.columns[j])
    #         lista_nome_colunas_painel_esquerda.append(df_plamov_compilado.columns[41])

    #         coluna_tableWidget_esquerda = 0
    #         self.ui.tableWidget.setHorizontalHeaderLabels(lista_nome_colunas_painel_esquerda)
    #         for i in range(df_plamov_compilado.shape[0]):
    #             for j in range(1, 7):
    #                 item = QtWidgets.QTableWidgetItem(str(df_plamov_compilado.iloc[i,j]))
    #                 self.ui.tableWidget.setItem(i,coluna_tableWidget_esquerda, item)
    #                 if i%2:
    #                     self.ui.tableWidget.item(i, coluna_tableWidget_esquerda).setBackground(QtGui.QColor(100, 139, 245))
    #                 coluna_tableWidget_esquerda += 1
    #             for j in range(13, 16):
    #                 item = QtWidgets.QTableWidgetItem(str(df_plamov_compilado.iloc[i,j]))
    #                 self.ui.tableWidget.setItem(i,coluna_tableWidget_esquerda, item)
    #                 if i%2:
    #                     self.ui.tableWidget.item(i, coluna_tableWidget_esquerda).setBackground(QtGui.QColor(100, 139, 245))
    #                 coluna_tableWidget_esquerda += 1
    #             for j in range(18, 20):
    #                 item = QtWidgets.QTableWidgetItem(str(df_plamov_compilado.iloc[i,j]))
    #                 self.ui.tableWidget.setItem(i,coluna_tableWidget_esquerda, item)
    #                 if i%2:
    #                     self.ui.tableWidget.item(i, coluna_tableWidget_esquerda).setBackground(QtGui.QColor(100, 139, 245))
    #                 coluna_tableWidget_esquerda += 1
    #             item = QtWidgets.QTableWidgetItem(str(df_plamov_compilado.iloc[i,41]))
    #             self.ui.tableWidget.setItem(i,coluna_tableWidget_esquerda, item)
    #             if i%2:
    #                 self.ui.tableWidget.item(i, coluna_tableWidget_esquerda).setBackground(QtGui.QColor(100, 139, 245))
    #             coluna_tableWidget_esquerda = 0
    #         status_painel = "carregado"
    #     else:
    #         pass

    #     df_OMs = pegar_OMs_do_COMPREP()

    #     #TODO apagar essa linha antes de entregar
    #     self.ui.tableWidget.setCurrentCell(5,41)
    #     self.carregar_Relat_rio_TP()
        ####################################

    def Carregar_Dados_dos_militares(self):
        global endereco_do_arquivo
        global df_OMs
        global df_plamov_compilado
        global status_painel 
        
        # 1. Tenta pegar o endereço do arquivo
        try:
            # Pega apenas a string do caminho (índice [0])
            endereco_do_arquivo = QFileDialog.getOpenFileName(self, "Qual arquivo gostaria de carregar?", caminho_atual, 'Excel files (*.xlsx)')[0]
        except:
            endereco_do_arquivo = ""

        # 2. Só executa o carregamento SE o endereço não estiver vazio
        if endereco_do_arquivo:
            # --- Carrega a aba PLAMOV COMPILADO ---
            df_plamov_compilado = pd.read_excel(endereco_do_arquivo, sheet_name="PLAMOV COMPILADO")
            df_plamov_compilado = df_plamov_compilado.fillna("") 
            df_plamov_compilado['ordem original'] = df_plamov_compilado.index
            
            # --- Configuração das Colunas (Sua lógica nova) ---
            COLUNAS_DESEJADAS = [
                "LOC ATUAL", "OM ATUAL", "SARAM", "POSTO", "QUADRO", "ESP", "SUB ESP", "LOC 1", "LOC 2", "LOC 3", "CÔNJUGE DA FAB?", "DADOS CÔNJUGE", "PLAMOV"
            ]

            colunas_existentes = [col for col in COLUNAS_DESEJADAS if col in df_plamov_compilado.columns]
            
            try:
                mapa_indices = {nome: df_plamov_compilado.columns.get_loc(nome) for nome in colunas_existentes}
                indices_a_exibir = [mapa_indices[nome] for nome in colunas_existentes]
            except KeyError as e:
                print(f"ERRO CRÍTICO: Coluna não encontrada: {e}")
                return 

            self.ui.tableWidget.setColumnCount(len(colunas_existentes))
            self.ui.tableWidget.setRowCount(df_plamov_compilado.shape[0])
            self.ui.tableWidget.setHorizontalHeaderLabels(colunas_existentes)

            # --- Ordenação ---
            cols_ordenacao = ['MELHOR PRIO', 'TEMPO LOC', 'ANTIGUIDADE']
            cols_presentes = [c for c in cols_ordenacao if c in df_plamov_compilado.columns]
            if cols_presentes:
                asc_order = [True, False, True][:len(cols_presentes)]
                df_plamov_compilado = df_plamov_compilado.sort_values(by=cols_presentes, ascending=asc_order)
                df_plamov_compilado = df_plamov_compilado.reset_index(drop=True)

            # --- Preenchimento da Tabela Visual ---
            coluna_tableWidget_esquerda = 0
            for i in range(df_plamov_compilado.shape[0]): 
                for df_index in indices_a_exibir: 
                    valor_celula = str(df_plamov_compilado.iloc[i, df_index])
                    item = QtWidgets.QTableWidgetItem(valor_celula)
                    self.ui.tableWidget.setItem(i, coluna_tableWidget_esquerda, item)
                    
                    if i % 2:
                        self.ui.tableWidget.item(i, coluna_tableWidget_esquerda).setBackground(QtGui.QColor(100, 139, 245))
                        
                    coluna_tableWidget_esquerda += 1
                coluna_tableWidget_esquerda = 0 
            
            status_painel = "carregado"

            # -------------------------------------------------------------------------
            # CORREÇÃO: Estas funções agora estão DENTRO do 'if endereco_do_arquivo:'
            # Elas só rodam se o arquivo tiver sido carregado com sucesso.
            # -------------------------------------------------------------------------
            df_OMs = pegar_OMs_do_COMPREP() # Carrega a lista de OMs
            self.carregar_Relat_rio_TP()    # Carrega as tabelas TP e TP BMA
            
            # Tenta selecionar uma célula inicial (cosmético)
            try:
                self.ui.tableWidget.setCurrentCell(5, len(COLUNAS_DESEJADAS)-1)
            except:
                pass
        
        else:
            # Se o usuário cancelar ou o arquivo for inválido, não faz nada
            print("Nenhum arquivo selecionado.")
            pass

    def carregar_Relat_rio_TP(self):
        global df_TP
        global df_TP_BMA 
        
        # Carrega a TP Padrão
        try:
            df_TP = pd.read_excel(endereco_do_arquivo, sheet_name="RELATÓRIO TP")
        except:
            pass

        # --- CARREGAMENTO DA TP BMA ---
        try:
            df_TP_BMA = pd.read_excel(endereco_do_arquivo, sheet_name="RELATÓRIO TP BMA")
            df_TP_BMA.fillna(0, inplace=True)
            
            # 1. Remove espaços em branco antes e depois dos nomes das colunas
            df_TP_BMA.columns = df_TP_BMA.columns.str.strip()

            # --- DEBUG: Verifique no terminal o que está sendo carregado ---
            print("Colunas encontradas no Excel (TP BMA):", df_TP_BMA.columns.tolist())

            # 2. PADRONIZAÇÃO DE NOMES
            # O código espera "Subespecialidade", mas o Excel pode ter variações.
            # Adicione aqui qualquer outra variação que seu Excel possa ter.
            mapa_correcao = {
                "Sub Especialidade": "Subespecialidade",
                "SUB ESP": "Subespecialidade",
                "Sub Esp": "Subespecialidade",
                "Sub-Especialidade": "Subespecialidade",
                "Subespecialidade ": "Subespecialidade" # Caso tenha espaço no final
            }
            df_TP_BMA.rename(columns=mapa_correcao, inplace=True)
            
        except Exception as e:
            print(f"Erro ao carregar aba RELATÓRIO TP BMA: {e}")
            df_TP_BMA = pd.DataFrame()   
        
    def linha_ativa_dados_militares (self): 
        global linha_selecionada_painel_esquerda
        linha_selecionada_painel_esquerda = self.ui.tableWidget.currentRow()
       
        return linha_selecionada_painel_esquerda
       
    def coluna_ativa_dados_militares (self):
        #nem sempre a coluna ativa no df_plamov_compilado vai ser a coluna ativa no tablewidget
    #depois que a célula da coluna "PLAMOV" checa se o militar foi movimentado e ajusta a quantidade de vagas na TP, dimunuindo a quantidade da "OM de destino" e aumentando da "OM ATUAL"
    #essa função vai precisar saber as dimensões do militar selecionado que foi obtida quando o usuário clicou na linha militare a linha ativa também.
    #parto do princípio que não existe mais de uma linha com a mesma combinação de OM,posto,quadro e especilidade
    #regras para ativar a função que atualiza as vagas na tabela TP
    #1- checar se o militar está sendo transferido realmente, pq pode acontecer de colocar a unidade de destino igual à unidade atual
    #2- checar se a coluna alterada é a coluna "PLAMOV"
    #3- checar se a célula foi feita pelo usuário, caso contrário a função seria ativada quando o relatório fosse carregado.
    #
        global coluna_ativa_painel_esquerda
        coluna_ativa_painel_esquerda = self.ui.tableWidget.currentColumn()
        return coluna_ativa_painel_esquerda
    
    def vaga_liberada_e_preenchida(self):
        global df_plamov_compilado
        global df_TP
        
        global linha_selecionada_painel_esquerda
        linha_ativa = int(self.linha_ativa_dados_militares())
        coluna_ativa = int(self.coluna_ativa_dados_militares())
       
        if status_painel == "carregado":
            global df_TP
            #nessa fase preciso saber qual a linha ativa que o usuário editou
            #nessa etapa preciso saber a OM_destino e OM_origem, isso vai ser buscado no df_plamov_compilado
            OM_atual = df_plamov_compilado.loc[linha_ativa , "OM ATUAL"]

            # Obtenha o novo valor da célula editada

            OM_Destino = self.ui.tableWidget.item(linha_alterada, 11).text()

            global  posto
            posto = pegar_posto(linha_ativa)

            global  quadro
            quadro = df_plamov_compilado["QUADRO"][linha_ativa]

            global  especialidade
            especialidade = df_plamov_compilado["ESP"][linha_ativa]


            #nessa fase preciso achar duas linhas no df_TP
            #1-linha da combinação entre a OM_destino e as três dimensões - dataframe.query("nome da coluna == 'valor da condição'").index[0])

            ###Está funcionando mas tem que colocar um tratamento para quando não achar uma combinação.
            # a melhor opção é criar uma coluna com as pessoas "chegando" e "saindo" de cada OM
            # uma outra coluna com as "vagas dinâmicas" que refletem o existente, vagas na tp, chegando e saindo.
            # Se colocar o destino de alguém para alguma OM que não tenha TP prevista, vai ser criada linha com a combinação e uma unidade somada à coluna "chegando", dessa forma é possivel manter o controle de quantas pessoas estão chegando em cada unidade.
            # TODO idéia de gráfico, colocar um gráfico para cada OM uma quantidade de pessoas saíndo e chegando, talvez uma indicação de estão perdendo gente, ou seja, com uma quantidade maior de pessoas saindo do que chegando, ou o contrário. 
            ###O que fazer nesse caso, criar uma e deixar uma flag dizendo que não tem TP
            ###Ver como está o tratamento no painel superior

            #se a OM inserida não estiver na relação, mostrar um popup com um warning
            #Se for do COMPREP  mas não tiver TP, mostrar um popup
            if posto == "TN":
                linha_OM_destino = df_TP.query(f"Unidade == '{OM_Destino}' & (Posto == 'CP/TN' | Posto == 'TN') & Quadro == '{quadro}' & Especialidade == '{especialidade}'")
            elif posto == "CP":
                linha_OM_destino = df_TP.query(f"Unidade == '{OM_Destino}' & (Posto == 'CP/TN' | Posto == 'CP') & Quadro == '{quadro}' & Especialidade == '{especialidade}'")
            else:
                linha_OM_destino = df_TP.query(f"Unidade == '{OM_Destino}' & Posto == '{posto}' & Quadro == '{quadro}' & Especialidade == '{especialidade}'")


            if linha_OM_destino.empty:
                #DESCRIÇÃO: ESSE CASO CRIA UMA LINHA COM A COMBINAÇÃO DAS TRÊS DIMENSÕES DO MILITAR CASO ELE SEJA ALOCADO EM UMA OM QUE NÃO EXISTA A PREVISÃO PARA AS SUAS 3 DIMENSÕES NA TABELA DE TP
                #AQUI eu devo criar uma nova linha com a combinação da query acima, inserir no DF_TP e colocar os valores de vagas nas respectivas colunas.
                nova_linha = pd.DataFrame({'Unidade': [OM_Destino],'Posto': [posto],'Quadro': [quadro],'Especialidade': [especialidade],'TLP Ano Corrente': [0],'Existentes': [1], 'Vagas': [-1]})
                df_TP = pd.concat([df_TP, nova_linha], axis=0, ignore_index=True)
                df_TP.fillna("", inplace=True)

            ####UNIDADE QUE O MILITAR ESTÁ CHEGANDO
            if posto == "CP":
                #TIRA UMA VAGA NA UNIDADE QUE O MILITAR ESTÁ CHEGANDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_Destino}") & ((df_TP['Posto'] == "CP/TN") | (df_TP['Posto'] == "CP")) & (df_TP['Quadro'] == f"{quadro}") &(df_TP['Especialidade'] == f"{especialidade}"), 'Vagas'] += -1
                #COLOCA UM EXISTENTE NA UNIDADE QUE O MILITAR ESTÁ CHEGANDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_Destino}") & ((df_TP['Posto'] == "CP/TN") | (df_TP['Posto'] == "CP")) & (df_TP['Quadro'] == f"{quadro}") &(df_TP['Especialidade'] == f"{especialidade}"), 'Existentes'] += 1
            elif posto == "TN":
                #TIRA UMA VAGA NA UNIDADE QUE O MILITAR ESTÁ CHEGANDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_Destino}") & ((df_TP['Posto'] == "CP/TN") | (df_TP['Posto'] == "TN")) & (df_TP['Quadro'] == f"{quadro}") &(df_TP['Especialidade'] == f"{especialidade}"), 'Vagas'] += -1
                #COLOCA UM EXISTENTE NA UNIDADE QUE O MILITAR ESTÁ CHEGANDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_Destino}") & ((df_TP['Posto'] == "CP/TN") | (df_TP['Posto'] == "TN")) & (df_TP['Quadro'] == f"{quadro}") &(df_TP['Especialidade'] == f"{especialidade}"), 'Existentes'] += 1
            else:
                #TIRA UMA VAGA NA UNIDADE QUE O MILITAR ESTÁ CHEGANDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_Destino}") & (df_TP['Posto'] == f"{posto}") & (df_TP['Quadro'] == f"{quadro}") &(df_TP['Especialidade'] == f"{especialidade}"), 'Vagas'] += -1
                #COLOCA UM EXISTENTE NA UNIDADE QUE O MILITAR ESTÁ CHEGANDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_Destino}") & (df_TP['Posto'] == f"{posto}") & (df_TP['Quadro'] == f"{quadro}") &(df_TP['Especialidade'] == f"{especialidade}"), 'Existentes'] += 1

            ####UNIDADE QUE O MILITAR ESTÁ SAINDO
            if posto == "CP":
                #COLOCA UMA VAGA NA UNIDADE QUE O MILITAR ESTÁ SAINDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_atual}") & ((df_TP['Posto'] == "CP/TN") | (df_TP['Posto'] == "CP")) & (df_TP['Quadro'] == f"{quadro}") &(df_TP['Especialidade'] == f"{especialidade}"), 'Vagas'] += -1
                #TIRA UM EXISTENTE NA UNIDADE QUE O MILITAR ESTÁ SAINDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_atual}") & ((df_TP['Posto'] == "CP/TN") | (df_TP['Posto'] == "CP")) & (df_TP['Quadro'] == f"{quadro}") &(df_TP['Especialidade'] == f"{especialidade}"), 'Existentes'] += 1
            elif posto == "TN":
                #COLOCA UMA VAGA NA UNIDADE QUE O MILITAR ESTÁ SAINDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_atual}") & ((df_TP['Posto'] == "CP/TN") | (df_TP['Posto'] == "TN")) & (df_TP['Quadro'] == f"{quadro}") &(df_TP['Especialidade'] == f"{especialidade}"), 'Vagas'] += -1
                #TIRA UM EXISTENTE NA UNIDADE QUE O MILITAR ESTÁ SAINDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_atual}") & ((df_TP['Posto'] == "CP/TN") | (df_TP['Posto'] == "TN")) & (df_TP['Quadro'] == f"{quadro}") &(df_TP['Especialidade'] == f"{especialidade}"), 'Existentes'] += 1
            else:
                #COLOCA UMA VAGA NA UNIDADE QUE O MILITAR ESTÁ SAINDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_atual}") & (df_TP['Posto'] == f"{posto}") & (df_TP['Quadro'] == f"{quadro}") &(df_TP['Especialidade'] == f"{especialidade}"), 'Vagas'] += -1
                #TIRA UM EXISTENTE NA UNIDADE QUE O MILITAR ESTÁ SAINDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_atual}") & (df_TP['Posto'] == f"{posto}") & (df_TP['Quadro'] == f"{quadro}") &(df_TP['Especialidade'] == f"{especialidade}"), 'Existentes'] += 1

    def escolher_OM_no_painel_direito(self):
        coluna_ativa_painel_direita = self.ui.tableWidget_2.currentColumn()
        linha_ativa_painel_direita = self.ui.tableWidget_2.currentRow()
        nome_coluna_ativa_painel_direita = df_OMs.columns[coluna_ativa_painel_direita]
        if (nome_coluna_ativa_painel_direita == "OMs"):
            #ir na linha ativa esquerda e coluna PLAMOV 
            #pegar o valor da célula doubleclicked no painel da direita
            #igualar os dois
            ###############################Parei aqui###############################################
            OM_selecionada_painel_direita = QtWidgets.QTableWidgetItem(self.ui.tableWidget_2.item(linha_ativa_painel_direita, coluna_ativa_painel_direita))
            if (linha_selecionada_painel_esquerda%2):
                #colorir de azul
                OM_selecionada_painel_direita.setBackground(QtGui.QColor(100, 139, 245))
            else:
                #colorir de branco
                OM_selecionada_painel_direita.setBackground(QtGui.QColor(255,255,255))
                
            self.ui.tableWidget.setItem(linha_selecionada_painel_esquerda, 12, OM_selecionada_painel_direita)
            df_plamov_compilado.loc[linha_selecionada_painel_esquerda, "PLAMOV"] = self.ui.tableWidget_2.item(linha_ativa_painel_direita, coluna_ativa_painel_direita).text()
            linha_ativa_painel_esquerda = self.linha_ativa_dados_militares()
            coluna_ativa_painel_esquerda = self.coluna_ativa_dados_militares()

            self.ui.tableWidget.setCurrentCell(linha_ativa_painel_esquerda + 1, coluna_ativa_painel_esquerda)
            self.atualizar_Painel_Direita()
        #     self.ui.tableWidget_2.setItem(linha_selecionada_painel_esquerda, coluna_ativa_painel_esquerda)
        #     item = QtWidgets.QTableWidgetItem(str(df_OMs.iloc[k,i]))
        #     self.ui.tableWidget_2.setItem(k,i, item)
    


app = QApplication(sys.argv)
UIWindow = SplashScreen()
app.exec()