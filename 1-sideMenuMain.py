from PyQt6.QtWidgets import QMainWindow, QApplication, QFileDialog
from PyQt6 import QtWidgets, QtCore
from PyQt6 import QtGui
from PyQt6.QtCore import Qt
from openpyxl import load_workbook
from openpyxl import Workbook
import datetime
import sqlite3
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *

import sys
import os
import pandas as pd
from menu_ui_ui import Ui_MainWindow
from SplashScreen_ui import Ui_SplashScreen
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import numpy as np

from pagina_voluntarios_cb import install as install_vol_cb

# TODO : Ao carregar um arquivo excel (Menu >> Carregar arquivos >> Dados dos militares [e selecionar um arquivo excel compatível]) com dados já carregados, demora muito para o segundo carregamento

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


############################################################################################

# ????: (VALE A PENA?, JA QUE A LINHA ATIVA VAI PASSAR PARA A PRÓXIMA E O PAINEL VAI SER ATUALIZADO) TAFERA: mudar as vagas e taxa de ocupação assim que o usuário selecionar a OM de destino
# DONE: design gráfico
# ????: Dividir os menus em "Páginas" e "Carregar"
# DONE: Colorir as OM das localidades escolhidas e a OM atual
# DONE: Reduzir a quantidade de colunas do painel esquerdo
# TODO: apagar as linhas que não tem TP no painel direito
# DONE: carregar os dois arquivos ao mesmo tempo
# DONE: Checar o tempo de execução de cada função e escovar para diminuir
# TODO: Apagar as linhas em branco do df_plamov (linhas que tem NaT)
# TODO: Perguntar para o Itamar quais gráfico/indicadores seriam úteis
# TODO: Terminar o filtro de colunas (efetivar a filtragem com o botão "Aplicar Filtro")
# TODO: Colocar o filtro no painel da direita tb
# TODO: Desassociar o tema do computador das cores do programa (modo claro/escuro)
# TODO: Inserir um aviso de quantos militares com a mesma qualificação ainda querem ir pra localidade do militar selecionado, dessa forma o usuário consegue saber se vai ter que entubar alguém pra localidade do militar selecionado se esse sair.

caminho_atual = os.getcwd()
status_painel = ""
linha_alterada = -1
coluna_alterada = -1
counter = 0

# Inicialização de variáveis globais para evitar avisos de 'name not found' no IDE
df_plamov_compilado = pd.DataFrame()
df_TP = pd.DataFrame()
df_TP_BMA = pd.DataFrame()
df_OMs = pd.DataFrame()
endereco_do_arquivo = ""
linha_selecionada_painel_esquerda = -1
coluna_ativa_painel_esquerda = -1
posto = ""
quadro = ""
especialidade = ""


def classificar(dataframe: pd.DataFrame):
    return dataframe.sort_values(by=['MELHOR PRIO', 'TEMPO LOC', 'ANTIGUIDADE'], ascending=[True, False, True], inplace=True)


def classificar_ordem_original(dataframe: pd.DataFrame):
    return dataframe.sort_values(by=['ordem original'], inplace=True)


def pegar_quadro(linha):
    global df_plamov_compilado
    quadro = df_plamov_compilado["QUADRO"][int(linha)]
    return quadro


def pegar_especialidade(linha):
    especialidade = df_plamov_compilado["ESP"][int(linha)]
    return especialidade


def pegar_Projeto(linha):
    try:
        sub = df_plamov_compilado["PROJETO"][int(linha)]
        return str(sub).strip()  # Remove espaços extras por segurança
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
    global df_TP_BMA
    global df_TP
    global df_OMs

    # Tenta usar a TP BMA primeiro (que já estará na memória), se não existir, usa a TP Geral
    if 'df_TP_BMA' in globals() and not df_TP_BMA.empty:
        df_referencia = df_TP_BMA.copy()
    elif 'df_TP' in globals() and not df_TP.empty:
        df_referencia = df_TP.copy()
    else:
        # Se não tiver nada carregado (nem banco nem excel), retorna vazio
        return pd.DataFrame(columns=["OMs", "Localidade", "Taxa de Ocup.", "Vagas"])

    # 2. Cria a lista de OMs únicas a partir dos dados em memória
    if 'Unidade' in df_referencia.columns:
        df_OMs = df_referencia['Unidade'].drop_duplicates()
    else:
        df_OMs = pd.DataFrame(columns=["Unidade"])

    df_OMs.dropna(inplace=True)
    df_OMs = df_OMs.to_frame(name="OMs")
    df_OMs.reset_index(drop=True, inplace=True)

    # 3. Inicializa colunas com dtype object para aceitar str e numéricos
    _n = len(df_OMs)
    df_OMs["Taxa de Ocup."] = np.empty(_n, dtype=object)
    df_OMs["Taxa de Ocup."][:] = ""
    df_OMs["Vagas"] = np.empty(_n, dtype=object)
    df_OMs["Vagas"][:] = ""

    # 4. MAPEAMENTO DE LOCALIDADE
    try:
        if 'Localidade' in df_referencia.columns:
            dict_localidades = df_referencia.set_index(
                'Unidade')['Localidade'].to_dict()
            df_OMs["Localidade"] = df_OMs["OMs"].map(dict_localidades)
        else:
            temp_df = df_referencia.iloc[:, [0, 1]]
            temp_df.columns = ['Unidade', 'Localidade']
            dict_localidades = temp_df.set_index(
                'Unidade')['Localidade'].to_dict()
            df_OMs["Localidade"] = df_OMs["OMs"].map(dict_localidades)
    except Exception as e:
        print(f"Erro ao mapear localidades: {e}")
        df_OMs["Localidade"] = ""

    return df_OMs


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
        if counter >= 30:
            self.timer.stop()
            self.main = UI()
            self.main.show()

            self.close()
        counter += 1


class FilterMenu(QtWidgets.QMenu):
    filterApplied = QtCore.pyqtSignal()

    def __init__(self, values, title, parent=None, active_filter=None, enable_numeric=True):
        super().__init__(title, parent)
        self.values = sorted(list(set(str(v)
                             for v in values if v is not None)))
        self.check_boxes = []

        self.state = {'selecionados': self.values, 'maior': '', 'menor': ''}

        if active_filter:
            if isinstance(active_filter, list):
                self.state['selecionados'] = active_filter
            elif isinstance(active_filter, dict):
                self.state.update(active_filter)

        self.setStyleSheet(
            "QMenu { background-color: white; border: 1px solid gray; }")

        # --- 0. CAMPOS NUMÉRICOS (Maior/Menor) ---
        if enable_numeric:
            self.widget_numerico = QtWidgets.QWidget()
            layout_num = QtWidgets.QGridLayout(self.widget_numerico)
            layout_num.setContentsMargins(5, 5, 5, 5)

            lbl_maior = QtWidgets.QLabel("Maior ou igual (>=):")
            self.edt_maior = QtWidgets.QLineEdit()
            self.edt_maior.setPlaceholderText("Ex: 70")
            self.edt_maior.setText(self.state['maior'])

            lbl_menor = QtWidgets.QLabel("Menor ou igual (<=):")
            self.edt_menor = QtWidgets.QLineEdit()
            self.edt_menor.setPlaceholderText("Ex: 100")
            self.edt_menor.setText(self.state['menor'])

            layout_num.addWidget(lbl_maior, 0, 0)
            layout_num.addWidget(self.edt_maior, 0, 1)
            layout_num.addWidget(lbl_menor, 1, 0)
            layout_num.addWidget(self.edt_menor, 1, 1)

            act_num = QtWidgets.QWidgetAction(self)
            act_num.setDefaultWidget(self.widget_numerico)
            self.addAction(act_num)
            self.addSeparator()

        # =================================================================
        # --- 0.5 CAMPO DE BUSCA (TEXTO) --- NOVIDADE AQUI
        # =================================================================
        self.widget_busca = QtWidgets.QWidget()
        layout_busca = QtWidgets.QVBoxLayout(self.widget_busca)
        layout_busca.setContentsMargins(5, 5, 5, 5)

        self.edt_busca = QtWidgets.QLineEdit()
        self.edt_busca.setPlaceholderText("Pesquisar...")
        # Ícone de lupa (opcional, só visual)
        self.edt_busca.setClearButtonEnabled(True)

        # Conecta o sinal de digitação à função de filtragem
        self.edt_busca.textChanged.connect(self.filtrar_lista_checkbox)

        layout_busca.addWidget(self.edt_busca)

        act_busca = QtWidgets.QWidgetAction(self)
        act_busca.setDefaultWidget(self.widget_busca)
        self.addAction(act_busca)
        # =================================================================

        # --- 1. ÁREA DE ROLAGEM E CHECKBOXES ---
        self.widget_conteudo = QtWidgets.QWidget()
        self.layout_conteudo = QtWidgets.QVBoxLayout(self.widget_conteudo)
        self.layout_conteudo.setContentsMargins(5, 5, 5, 5)
        self.layout_conteudo.setSpacing(2)

        self.cb_all = QtWidgets.QCheckBox(
            " (Selecionar Tudo)", self.widget_conteudo)

        # Lógica inicial do Selecionar Tudo
        lista_selecionados = self.state['selecionados']
        if len(lista_selecionados) == len(self.values):
            self.cb_all.setChecked(True)
        else:
            self.cb_all.setChecked(False)

        self.cb_all.stateChanged.connect(self.toggle_all)
        self.layout_conteudo.addWidget(self.cb_all)

        line = QtWidgets.QFrame()
        line.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        line.setFrameShadow(QtWidgets.QFrame.Shadow.Sunken)
        self.layout_conteudo.addWidget(line)

        for val in self.values:
            cb = QtWidgets.QCheckBox(str(val), self.widget_conteudo)
            if str(val) in lista_selecionados:
                cb.setChecked(True)
            else:
                cb.setChecked(False)

            cb.stateChanged.connect(self.atualizar_estado_selecionar_tudo)
            self.layout_conteudo.addWidget(cb)
            self.check_boxes.append(cb)

        self.layout_conteudo.addStretch()

        self.scroll_area = QtWidgets.QScrollArea()
        self.scroll_area.setWidget(self.widget_conteudo)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setMinimumHeight(150)
        self.scroll_area.setMaximumHeight(300)
        self.scroll_area.setMinimumWidth(220)

        item_action = QtWidgets.QWidgetAction(self)
        item_action.setDefaultWidget(self.scroll_area)
        self.addAction(item_action)

        self.addSeparator()

        btn_apply = QtWidgets.QPushButton("Aplicar Filtro")
        btn_apply.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        btn_apply.setStyleSheet("""
            QPushButton {
                background-color: #0078d7; color: white; border: none; padding: 5px; font-weight: bold;
            }
            QPushButton:hover { background-color: #005a9e; }
        """)
        btn_apply.clicked.connect(self.emitir_e_fechar)

        action_btn = QtWidgets.QWidgetAction(self)
        action_btn.setDefaultWidget(btn_apply)
        self.addAction(action_btn)

    # --- NOVA FUNÇÃO DE FILTRAGEM ---
    def filtrar_lista_checkbox(self, texto):
        """Esconde ou mostra os checkboxes baseado no texto digitado."""
        texto = texto.lower()
        visiveis = 0

        for cb in self.check_boxes:
            if texto in cb.text().lower():
                cb.setVisible(True)
                visiveis += 1
            else:
                cb.setVisible(False)

        # Opcional: Se a busca não retornar nada, desabilita o "Selecionar Tudo"
        self.cb_all.setEnabled(visiveis > 0)
    # --------------------------------

    def emitir_e_fechar(self):
        self.filterApplied.emit()
        self.close()

    def toggle_all(self, state):
        """
        Marca/Desmarca. 
        ATENÇÃO: Agora só afeta os itens VISÍVEIS (filtrados).
        Isso permite buscar "SGT", clicar em selecionar tudo, e marcar apenas os SGTs.
        """
        is_checked = (state == QtCore.Qt.CheckState.Checked.value)

        for cb in self.check_boxes:
            # Só altera o estado se o checkbox estiver visível (passou na busca)
            if cb.isVisible():
                cb.blockSignals(True)
                cb.setChecked(is_checked)
                cb.blockSignals(False)

        # Após alterar os visíveis, precisamos verificar se isso afetou o estado global
        # para manter a consistência interna (opcional, mas bom pra UX)
        # self.atualizar_estado_selecionar_tudo() # Pode causar loop visual, melhor deixar sem por enquanto.

    def atualizar_estado_selecionar_tudo(self):
        # Verifica apenas checkboxes visíveis? Não, verifica todos para saber se "TUDO" está marcado.
        # Mas para a lógica visual do checkbox pai, verificamos se todos estão True.
        todos_marcados = all(cb.isChecked() for cb in self.check_boxes)

        self.cb_all.blockSignals(True)
        self.cb_all.setChecked(todos_marcados)
        self.cb_all.blockSignals(False)

    def get_filter_state(self):
        # Pega todos que estão marcados (mesmo os ocultos pela busca)
        selecionados = [cb.text() for cb in self.check_boxes if cb.isChecked()]

        if hasattr(self, 'edt_maior'):
            val_maior = self.edt_maior.text().strip()
            val_menor = self.edt_menor.text().strip()
        else:
            val_maior = ""
            val_menor = ""

        return {
            'selecionados': selecionados,
            'maior': val_maior,
            'menor': val_menor,
            'all_checked': self.cb_all.isChecked()
        }

    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key.Key_Return or event.key() == QtCore.Qt.Key.Key_Enter:
            self.emitir_e_fechar()
        else:
            super().keyPressEvent(event)


class GraficoCanvas(FigureCanvas):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        # Configuração da Figura
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        super(GraficoCanvas, self).__init__(self.fig)

# --- CLASSE PARA CORRIGIR A COR DA SELEÇÃO ---


class UI(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # 1. Obriga a tabela a selecionar a LINHA INTEIRA ao clicar, não só a célula
        self.ui.tableWidget.setSelectionBehavior(
            QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)

        # 2. (Opcional) Permite selecionar apenas uma linha por vez (evita bagunça)
        self.ui.tableWidget.setSelectionMode(
            QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)

        # =================================================================
        # 1. CRIAÇÃO DA PÁGINA DE GRÁFICOS (VIA CÓDIGO)
        # =================================================================
        self.page_graficos = QtWidgets.QWidget()
        self.layout_graficos = QtWidgets.QVBoxLayout(self.page_graficos)
        # Adiciona essa nova página ao seu StackedWidget existente
        self.ui.stackedWidget.addWidget(self.page_graficos)

        # =================================================================
        # 2. CRIAÇÃO DO BOTÃO NO MENU (VIA CÓDIGO)
        # =================================================================
        # Criamos uma ação nova chamada "Dashboard / Gráficos"
        self.actionGraficos = QtGui.QAction("Dashboard / Gráficos", self)
        # Adicionamos ao menu existente "menuMenu"
        self.ui.menuMenu.addAction(self.actionGraficos)

        # Conectamos o botão à função de abrir a página
        self.actionGraficos.triggered.connect(lambda: self.Pag_Graficos())

        # =================================================================
        # 3. CRIAÇÃO DA PÁGINA DE MILITARES PRIORITÁRIOS (BLOCOS A, B, C, D)
        # =================================================================
        self.page_prioritarios = QtWidgets.QWidget()
        self.layout_prioritarios = QtWidgets.QVBoxLayout(
            self.page_prioritarios)
        self.layout_prioritarios.setContentsMargins(2, 2, 2, 2)
        self.layout_prioritarios.setSpacing(0)

        # Linha 1: Título + Resumo lado a lado (altura fixa)
        widget_titulo_resumo = QtWidgets.QWidget()
        widget_titulo_resumo.setFixedHeight(24)
        layout_titulo_resumo = QtWidgets.QHBoxLayout(widget_titulo_resumo)
        layout_titulo_resumo.setContentsMargins(5, 0, 5, 0)
        layout_titulo_resumo.setSpacing(15)

        lbl_titulo_prio = QtWidgets.QLabel(
            "Militares Prioritários (Blocos A, B, C e D)")
        lbl_titulo_prio.setStyleSheet("font-size: 13px; font-weight: bold;")
        layout_titulo_resumo.addWidget(lbl_titulo_prio)

        self.lbl_resumo_prio = QtWidgets.QLabel("")
        self.lbl_resumo_prio.setStyleSheet("font-size: 11px; color: #555;")
        layout_titulo_resumo.addWidget(self.lbl_resumo_prio)
        layout_titulo_resumo.addStretch()
        self.layout_prioritarios.addWidget(widget_titulo_resumo)

        # Linha 2: Legenda de cores dos blocos (altura fixa)
        widget_legenda = QtWidgets.QWidget()
        widget_legenda.setFixedHeight(20)
        layout_legenda = QtWidgets.QHBoxLayout(widget_legenda)
        layout_legenda.setContentsMargins(5, 0, 5, 0)
        layout_legenda.setSpacing(8)
        legendas = [
            ("A — Loc. Difíceis (≥2a)", "#DC3545"),
            ("B — BV/PV (≥4a)", "#FF9900"),
            ("C — MN/BE (≥5a)", "#FFC107"),
            ("D — Santa Cruz (≥6a)", "#007BFF"),
        ]
        for texto_leg, cor_leg in legendas:
            lbl = QtWidgets.QLabel(f"■ {texto_leg}")
            lbl.setStyleSheet(
                f"color: {cor_leg}; font-weight: bold; font-size: 10px;")
            layout_legenda.addWidget(lbl)
        layout_legenda.addStretch()
        self.layout_prioritarios.addWidget(widget_legenda)

        # --- SPLITTER HORIZONTAL: TABELA ESQUERDA + PAINEL DIREITO ---
        self.splitter_prioritarios = QtWidgets.QSplitter(
            QtCore.Qt.Orientation.Horizontal)

        # Tabela de prioritários (ESQUERDA)
        self.tableWidget_prioritarios = QtWidgets.QTableWidget()
        self.tableWidget_prioritarios.setSelectionBehavior(
            QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.tableWidget_prioritarios.setSelectionMode(
            QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        self.tableWidget_prioritarios.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tableWidget_prioritarios.setStyleSheet("""
            QTableWidget::item:selected {
                background-color: #7f807c;
            }
            QTableWidget::item:selected:focus {
                outline: none;
            }
            QHeaderView::section {
                padding-right: 5px;
                padding-left: 5px;
            }
        """)
        self.splitter_prioritarios.addWidget(self.tableWidget_prioritarios)

        # Painel Direito (DIREITA) — container com título + legenda + tabela de OMs
        self.widget_painel_dir_prio = QtWidgets.QWidget()
        layout_dir_prio = QtWidgets.QVBoxLayout(self.widget_painel_dir_prio)
        layout_dir_prio.setContentsMargins(0, 0, 0, 0)
        layout_dir_prio.setSpacing(0)

        # Título do painel direito
        lbl_titulo_dir = QtWidgets.QLabel("OMs Disponíveis")
        lbl_titulo_dir.setFixedHeight(22)
        lbl_titulo_dir.setStyleSheet("font-size: 12px; font-weight: bold;")
        lbl_titulo_dir.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        layout_dir_prio.addWidget(lbl_titulo_dir)

        # Legenda de cores das opções
        widget_legenda_dir = QtWidgets.QWidget()
        widget_legenda_dir.setFixedHeight(18)
        layout_legenda_dir = QtWidgets.QHBoxLayout(widget_legenda_dir)
        layout_legenda_dir.setContentsMargins(5, 0, 5, 0)
        layout_legenda_dir.setSpacing(6)
        legendas_opcoes = [
            ("1ª Opção", "#1DB502"),
            ("2ª Opção", "#FFF308"),
            ("3ª Opção", "#FF00FF"),
        ]
        for texto_op, cor_op in legendas_opcoes:
            lbl_op = QtWidgets.QLabel(f"■ {texto_op}")
            lbl_op.setStyleSheet(
                f"color: {cor_op}; font-weight: bold; font-size: 9px;")
            layout_legenda_dir.addWidget(lbl_op)
        layout_legenda_dir.addStretch()
        layout_dir_prio.addWidget(widget_legenda_dir)

        # Tabela de OMs do painel direito dos prioritários
        self.tableWidget_prio_direita = QtWidgets.QTableWidget()
        self.tableWidget_prio_direita.setSelectionBehavior(
            QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.tableWidget_prio_direita.setSelectionMode(
            QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        self.tableWidget_prio_direita.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tableWidget_prio_direita.setStyleSheet("""
            QTableWidget::item:selected {
                background-color: #7f807c;
            }
            QTableWidget::item:selected:focus {
                outline: none;
            }
            QHeaderView::section {
                padding-right: 5px;
                padding-left: 5px;
            }
        """)
        layout_dir_prio.addWidget(self.tableWidget_prio_direita)
        self.splitter_prioritarios.addWidget(self.widget_painel_dir_prio)

        # Proporção inicial do Splitter (70% esquerda, 30% direita)
        self.splitter_prioritarios.setStretchFactor(0, 7)
        self.splitter_prioritarios.setStretchFactor(1, 3)
        self.layout_prioritarios.addWidget(self.splitter_prioritarios, 1)

        self.ui.stackedWidget.addWidget(self.page_prioritarios)

        # Botão no menu
        self.actionPrioritarios = QtGui.QAction("Militares Prioritários", self)
        self.ui.menuMenu.addAction(self.actionPrioritarios)
        self.actionPrioritarios.triggered.connect(
            lambda: self.Pag_Prioritarios())

        # Conexão: clique na tabela de prioritários atualiza o painel direito
        self.tableWidget_prioritarios.cellClicked.connect(
            lambda: self.atualizar_Painel_Direita_Prioritarios())

        # Armazena o DataFrame filtrado para uso no painel direito
        self.df_prioritarios_filtrado = pd.DataFrame()

        # =================================================================
        # 3B. CRIAÇÃO DA PÁGINA "QUEREM LOC. DIFÍCEIS" (cond_quer_localidades_d / Score 9)
        # =================================================================
        self.page_quer_loc_dificeis = QtWidgets.QWidget()
        self.layout_quer_loc_dificeis = QtWidgets.QVBoxLayout(
            self.page_quer_loc_dificeis)
        self.layout_quer_loc_dificeis.setContentsMargins(2, 2, 2, 2)
        self.layout_quer_loc_dificeis.setSpacing(0)

        # Linha 1: Título + Resumo lado a lado (altura fixa)
        widget_titulo_qld = QtWidgets.QWidget()
        widget_titulo_qld.setFixedHeight(24)
        layout_titulo_qld = QtWidgets.QHBoxLayout(widget_titulo_qld)
        layout_titulo_qld.setContentsMargins(5, 0, 5, 0)
        layout_titulo_qld.setSpacing(15)

        lbl_titulo_qld = QtWidgets.QLabel(
            "Querem Loc. Difíceis (Guarantã, Eirunepê, S. Gabriel, Vilhena)")
        lbl_titulo_qld.setStyleSheet("font-size: 13px; font-weight: bold;")
        layout_titulo_qld.addWidget(lbl_titulo_qld)

        self.lbl_resumo_qld = QtWidgets.QLabel("")
        self.lbl_resumo_qld.setStyleSheet("font-size: 11px; color: #555;")
        layout_titulo_qld.addWidget(self.lbl_resumo_qld)
        layout_titulo_qld.addStretch()
        self.layout_quer_loc_dificeis.addWidget(widget_titulo_qld)

        # Linha 2: Descrição curta (altura fixa)
        widget_desc_qld = QtWidgets.QWidget()
        widget_desc_qld.setFixedHeight(18)
        layout_desc_qld = QtWidgets.QHBoxLayout(widget_desc_qld)
        layout_desc_qld.setContentsMargins(5, 0, 5, 0)
        lbl_desc_qld = QtWidgets.QLabel(
            "Militares que escolheram localidades difíceis em pelo menos uma das 3 opções")
        lbl_desc_qld.setStyleSheet("font-size: 10px; color: #888;")
        layout_desc_qld.addWidget(lbl_desc_qld)
        layout_desc_qld.addStretch()
        self.layout_quer_loc_dificeis.addWidget(widget_desc_qld)

        # --- SPLITTER HORIZONTAL: TABELA ESQUERDA + PAINEL DIREITO ---
        self.splitter_qld = QtWidgets.QSplitter(
            QtCore.Qt.Orientation.Horizontal)

        # Tabela esquerda
        self.tableWidget_qld = QtWidgets.QTableWidget()
        self.tableWidget_qld.setSelectionBehavior(
            QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.tableWidget_qld.setSelectionMode(
            QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        self.tableWidget_qld.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tableWidget_qld.setStyleSheet("""
            QTableWidget::item:selected {
                background-color: #7f807c;
            }
            QTableWidget::item:selected:focus {
                outline: none;
            }
            QHeaderView::section {
                padding-right: 5px;
                padding-left: 5px;
            }
        """)
        self.splitter_qld.addWidget(self.tableWidget_qld)

        # Painel Direito — container com título + legenda + tabela de OMs
        self.widget_painel_dir_qld = QtWidgets.QWidget()
        layout_dir_qld = QtWidgets.QVBoxLayout(self.widget_painel_dir_qld)
        layout_dir_qld.setContentsMargins(0, 0, 0, 0)
        layout_dir_qld.setSpacing(0)

        lbl_titulo_dir_qld = QtWidgets.QLabel("OMs Disponíveis")
        lbl_titulo_dir_qld.setFixedHeight(22)
        lbl_titulo_dir_qld.setStyleSheet("font-size: 12px; font-weight: bold;")
        lbl_titulo_dir_qld.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        layout_dir_qld.addWidget(lbl_titulo_dir_qld)

        widget_legenda_dir_qld = QtWidgets.QWidget()
        widget_legenda_dir_qld.setFixedHeight(18)
        layout_legenda_dir_qld = QtWidgets.QHBoxLayout(widget_legenda_dir_qld)
        layout_legenda_dir_qld.setContentsMargins(5, 0, 5, 0)
        layout_legenda_dir_qld.setSpacing(6)
        for texto_op, cor_op in [("1ª Opção", "#1DB502"), ("2ª Opção", "#FFF308"), ("3ª Opção", "#FF00FF")]:
            lbl_op = QtWidgets.QLabel(f"■ {texto_op}")
            lbl_op.setStyleSheet(
                f"color: {cor_op}; font-weight: bold; font-size: 9px;")
            layout_legenda_dir_qld.addWidget(lbl_op)
        layout_legenda_dir_qld.addStretch()
        layout_dir_qld.addWidget(widget_legenda_dir_qld)

        self.tableWidget_qld_direita = QtWidgets.QTableWidget()
        self.tableWidget_qld_direita.setSelectionBehavior(
            QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.tableWidget_qld_direita.setSelectionMode(
            QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        self.tableWidget_qld_direita.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tableWidget_qld_direita.setStyleSheet("""
            QTableWidget::item:selected {
                background-color: #7f807c;
            }
            QTableWidget::item:selected:focus {
                outline: none;
            }
            QHeaderView::section {
                padding-right: 5px;
                padding-left: 5px;
            }
        """)
        layout_dir_qld.addWidget(self.tableWidget_qld_direita)
        self.splitter_qld.addWidget(self.widget_painel_dir_qld)

        self.splitter_qld.setStretchFactor(0, 7)
        self.splitter_qld.setStretchFactor(1, 3)
        self.layout_quer_loc_dificeis.addWidget(self.splitter_qld, 1)

        self.ui.stackedWidget.addWidget(self.page_quer_loc_dificeis)

        # Botão no menu
        self.actionQuerLocDificeis = QtGui.QAction(
            "Querem Loc. Difíceis", self)
        self.ui.menuMenu.addAction(self.actionQuerLocDificeis)
        self.actionQuerLocDificeis.triggered.connect(
            lambda: self.Pag_QuerLocDificeis())

        # Conexão: clique atualiza o painel direito
        self.tableWidget_qld.cellClicked.connect(
            lambda: self.atualizar_Painel_Direita_QLD())

        # DataFrame filtrado para uso no painel direito
        self.df_qld_filtrado = pd.DataFrame()

        # Cria a página "Voluntários C e B" e adiciona o item ao menu
        install_vol_cb(self)

        # 4. Define a cor do destaque (Amarelo com letra preta) usando CSS (QSS)
        # O 'outline: none' remove aquele pontilhado em volta da célula
        self.ui.tableWidget.setStyleSheet("""
            QTableWidget::item:selected {
                background-color: #7f807c;
                /*color: #000000;*/
                
            }
            QTableWidget::item:selected:focus {
                outline: none;
            }
            /* Adicione isso para o ícone do cabeçalho ficar bonito */
            QHeaderView::section {
                padding-right: 5px; 
                padding-left: 5px;
            }
        """)

        # para inicializar na página dos militares
        self.ui.stackedWidget.setCurrentIndex(0)

        self.ui.actionMilitares.triggered.connect(lambda: self.Pag_Militares())
        self.ui.actionQuadros_Especialidades.triggered.connect(
            lambda: self.Pag_Quadros_Especialidades())
        self.ui.actionRelat_rio_TP.triggered.connect(
            lambda: self.Pag_Relat_rio_TP())
        self.ui.actionMapa.triggered.connect(lambda: self.Pag_Mapa())

        self.ui.actionDados_dos_militares.triggered.connect(
            lambda: self.Carregar_Dados_dos_militares())
        self.ui.actionRelat_rio_TP_2.triggered.connect(
            lambda: self.carregar_Relat_rio_TP())
        self.ui.actionSALVAR.triggered.connect(lambda: self.salvar())
        self.ui.tableWidget.cellClicked.connect(
            lambda: self.linha_ativa_dados_militares())
        self.ui.tableWidget.cellClicked.connect(
            lambda: self.coluna_ativa_dados_militares())
        self.ui.tableWidget.cellClicked.connect(
            lambda: self.atualizar_Painel_Direita())
        self.ui.tableWidget.cellChanged.connect(self.celula_alterada)
        self.ui.tableWidget_2.cellDoubleClicked.connect(
            lambda: self.escolher_OM_no_painel_direito())

        # --- CONFIGURAÇÃO DE FILTRO: PAINEL ESQUERDO (tableWidget) ---
        self.ui.tableWidget.horizontalHeader().setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu)
        # Note que agora passamos a tabela e o dicionário específicos via lambda
        self.ui.tableWidget.horizontalHeader().customContextMenuRequested.connect(
            lambda pos: self.abrir_menu_filtro(
                pos, self.ui.tableWidget, self.filtros_ativos_esquerda)
        )
        self.filtros_ativos_esquerda = {}  # Renomeei para ficar claro

        # --- CONFIGURAÇÃO DE FILTRO: PAINEL DIREITO (tableWidget_2) ---
        self.ui.tableWidget_2.horizontalHeader().setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu)
        self.ui.tableWidget_2.horizontalHeader().customContextMenuRequested.connect(
            lambda pos: self.abrir_menu_filtro(
                pos, self.ui.tableWidget_2, self.filtros_ativos_direita)
        )
        self.filtros_ativos_direita = {}  # Dicionário novo para a direita

        # --- CRIAÇÃO DO ÍCONE DE FILTRO NA MEMÓRIA ---
        self.icone_filtro = QtGui.QPixmap(20, 20)
        self.icone_filtro.fill(QtCore.Qt.GlobalColor.transparent)
        painter = QtGui.QPainter(self.icone_filtro)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)
        painter.setBrush(QtGui.QColor("#4a4a4a"))  # Cor Cinza Escuro
        painter.setPen(QtCore.Qt.PenStyle.NoPen)
        # Desenha um triângulo/funil
        path = QtGui.QPainterPath()
        path.moveTo(4, 5)
        path.lineTo(16, 5)
        path.lineTo(10, 12)
        path.closeSubpath()
        painter.drawPath(path)
        painter.end()
        self.icone_filtro = QtGui.QIcon(self.icone_filtro)
        # ---------------------------------------------

        self.carregar_tudo_do_banco()

        self.show()

    def abrir_menu_filtro(self, position, tabela_alvo, dic_filtros):
        col_clicada = tabela_alvo.horizontalHeader().logicalIndexAt(position)

        # Coleta valores (mantido igual)
        valores_coluna = []
        for row in range(tabela_alvo.rowCount()):
            linha_valida_pelo_contexto = True
            for col_filtro, estado_filtro in dic_filtros.items():
                if col_filtro == col_clicada:
                    continue

                if isinstance(estado_filtro, dict):
                    valores_permitidos = estado_filtro.get('selecionados', [])
                    f_maior = estado_filtro.get('maior', '')
                    f_menor = estado_filtro.get('menor', '')
                else:
                    valores_permitidos = estado_filtro
                    f_maior = ''
                    f_menor = ''

                item_teste = tabela_alvo.item(row, col_filtro)
                valor_teste = item_teste.text() if item_teste else ""

                if valor_teste not in valores_permitidos:
                    linha_valida_pelo_contexto = False

                if linha_valida_pelo_contexto and (f_maior or f_menor):
                    val_num = self.converter_para_float(valor_teste)
                    if val_num is not None:
                        if f_maior and not (val_num >= float(f_maior)):
                            linha_valida_pelo_contexto = False
                        if f_menor and not (val_num <= float(f_menor)):
                            linha_valida_pelo_contexto = False
                    else:
                        linha_valida_pelo_contexto = False

                if not linha_valida_pelo_contexto:
                    break

            if linha_valida_pelo_contexto:
                item = tabela_alvo.item(row, col_clicada)
                valores_coluna.append(item.text() if item else "")

        filtro_atual = dic_filtros.get(col_clicada)

        # --- LÓGICA CORRIGIDA AQUI ---

        # 1. Definimos como False por padrão (Assim o painel da esquerda nunca terá)
        mostrar_numerico = False

        # 2. Só habilitamos SE for o Painel da Direita (tableWidget_2)
        #    E a coluna for diferente de 0 (OM)
        if tabela_alvo == self.ui.tableWidget_2 and col_clicada != 0:
            mostrar_numerico = True

        # -----------------------------

        menu = FilterMenu(valores_coluna, f"Filtro", self,
                          active_filter=filtro_atual, enable_numeric=mostrar_numerico)

        menu.filterApplied.connect(lambda: self.aplicar_e_guardar_filtros(
            col_clicada, menu, tabela_alvo, dic_filtros))
        menu.exec(tabela_alvo.horizontalHeader().mapToGlobal(position))

    def aplicar_e_guardar_filtros(self, col, menu, tabela_alvo, dic_filtros):
        # Pega o estado completo do menu
        estado = menu.get_filter_state()

        # Lógica: O filtro está "limpo" se "Selecionar Tudo" estiver marcado E os campos numéricos vazios
        filtro_limpo = estado['all_checked'] and estado['maior'] == "" and estado['menor'] == ""

        if filtro_limpo:
            if col in dic_filtros:
                del dic_filtros[col]

            item_header = tabela_alvo.horizontalHeaderItem(col)
            if item_header:
                item_header.setIcon(QtGui.QIcon())
        else:
            # Salva o estado completo no dicionário
            dic_filtros[col] = estado

            item_header = tabela_alvo.horizontalHeaderItem(col)
            if item_header:
                item_header.setIcon(self.icone_filtro)

        self.executar_filtros_combinados(tabela_alvo, dic_filtros)

    def executar_filtros_combinados(self, tabela_alvo, dic_filtros):
        tabela_alvo.setUpdatesEnabled(False)

        for row in range(tabela_alvo.rowCount()):
            exibir_linha = True

            for col, estado_filtro in dic_filtros.items():
                if isinstance(estado_filtro, dict):
                    valores_permitidos = estado_filtro.get('selecionados', [])
                    f_maior = estado_filtro.get('maior', '')
                    f_menor = estado_filtro.get('menor', '')
                else:
                    valores_permitidos = estado_filtro
                    f_maior = ''
                    f_menor = ''

                item = tabela_alvo.item(row, col)
                texto_celula = item.text() if item else ""

                if texto_celula not in valores_permitidos:
                    exibir_linha = False

                if exibir_linha and (f_maior or f_menor):
                    val_num = self.converter_para_float(texto_celula)

                    if val_num is not None:
                        # --- LÓGICA ATUALIZADA (>= e <=) ---
                        if f_maior:
                            try:
                                limit_maior = float(f_maior)
                                # Se NÃO for maior OU IGUAL, esconde
                                if not (val_num >= limit_maior):
                                    exibir_linha = False
                            except ValueError:
                                pass

                        if f_menor and exibir_linha:
                            try:
                                limit_menor = float(f_menor)
                                # Se NÃO for menor OU IGUAL, esconde
                                if not (val_num <= limit_menor):
                                    exibir_linha = False
                            except ValueError:
                                pass
                    else:
                        exibir_linha = False

                if not exibir_linha:
                    break

            tabela_alvo.setRowHidden(row, not exibir_linha)

        tabela_alvo.setUpdatesEnabled(True)

    def converter_para_float(self, valor_str):
        try:
            # Remove % e espaços
            limpo = valor_str.replace('%', '').strip()
            return float(limpo)
        except ValueError:
            return None

    def salvar(self):
        # TODO: Esta função cria uma arquivo novo para salvar a relação de oms escolhidas para cada militar durante a execução do código. MUDAR PARA ESCREVER DIRETAMENTE NO ARQUIVO EXCEL DO PLAMOV.
        global endereco_do_arquivo
        if 'endereco_do_arquivo' not in globals() or not endereco_do_arquivo:
            print("Nenhum arquivo carregado. Não é possível salvar.")
            return
        if 'df_plamov_compilado' not in globals() or df_plamov_compilado.empty:
            return
        df_plamov_compilado.sort_values(
            by=['ordem original'], ascending=[True], inplace=True)
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
            if coluna_alterada == 15:
                df_plamov_compilado.loc[linha_alterada, "PLAMOV"] = self.ui.tableWidget.item(
                    linha_alterada, coluna_alterada).text()
            self.salvar_tudo_no_banco()

    # --- FUNÇÃO PARA FORMATAR DATAS (DD/MM/AAAA) ---
    def formatar_datas_brasileiras(self):
        global df_plamov_compilado

        # Lista das colunas que são datas
        colunas_de_data = ["APRESENTAÇÃO NA LOC", "DATA DE PRAÇA"]

        if 'df_plamov_compilado' in globals() and not df_plamov_compilado.empty:
            for col in colunas_de_data:
                if col in df_plamov_compilado.columns:
                    # 1. Converte para objeto de data (garante que o Pandas entenda)
                    # errors='coerce' transforma textos ruins em NaT (Not a Time)
                    datas_convertidas = pd.to_datetime(
                        df_plamov_compilado[col], errors='coerce')

                    # 2. Converte para String no formato DD/MM/AAAA
                    # .dt.strftime faz a mágica. Onde for NaT, vira NaN, e o fillna limpa.
                    df_plamov_compilado[col] = datas_convertidas.dt.strftime(
                        '%d/%m/%Y').fillna("")

    # ---  FUNÇÃO DEDICADA À ORDENAÇÃO ---
    # OBS: Existe um critério especial para contar o tempo de localidade, pois o tempo não é contado quando está sendo feito o plamov, mas considera o mês de janeiro do ano seguinte. Por enquanto o tempo de localidade está sendo capturado da planilha excel, mas essa lógica deve ser implementada no tesserato no futuro.
    # --- FUNÇÃO DE ORDENAÇÃO POR LOCALIDADE (BLOCOS A, B, C, D) ---
    def aplicar_ordenacao_militares(self):
        global df_plamov_compilado

        # Verifica se o DataFrame existe e não está vazio
        if 'df_plamov_compilado' not in globals() or df_plamov_compilado.empty:
            return

        # --- CONJUNTO DE BLOCOS DE MILITARES QUE QUEREM SAIR DA SELVA + SANTA CRUZ ---
        # Dica: Colocamos variações com e sem acento para garantir
        loc_bloco_a = ["GUARANTÃ DO NORTE", "EIRUNEPÊ", "EIRUNEPE",
                       "SÃO GABRIEL DA CACHOEIRA", "SAO GABRIEL DA CACHOEIRA", "VILHENA"]
        loc_bloco_b = ["BOA VISTA", "PORTO VELHO"]
        loc_bloco_c = ["MANAUS", "BELÉM", "BELEM"]
        loc_bloco_d = ["1 GAVCA", "1/7 GAV", "3/8 GAV",
                       "BASC", "GLOG-SC", "GSAU-SC", "GSD-SC"]

        if "LOC 1" in df_plamov_compilado.columns:
            # Normaliza as 3 opções para comparação
            s_l1 = df_plamov_compilado["LOC 1"].astype(
                str).str.strip().str.upper()
            s_l2 = df_plamov_compilado["LOC 2"].astype(
                str).str.strip().str.upper()
            s_l3 = df_plamov_compilado["LOC 3"].astype(
                str).str.strip().str.upper()

            # Regra: Se a localidade estiver em QUALQUER uma das 3 opções
            cond_quer_localidades_d = (s_l1.isin(loc_bloco_a) | s_l2.isin(
                loc_bloco_a) | s_l3.isin(loc_bloco_a))
            cond_quer_localidades_c_e_b = (s_l1.isin(loc_bloco_b) | s_l2.isin(loc_bloco_b) | s_l3.isin(
                loc_bloco_b)) | (s_l1.isin(loc_bloco_c) | s_l2.isin(loc_bloco_c) | s_l3.isin(loc_bloco_c))
            cond_quer_OMs_em_Santa_Cruz = (s_l1.isin(loc_bloco_d) | s_l2.isin(
                loc_bloco_d) | s_l3.isin(loc_bloco_d))

        # 2. Garante que as colunas necessárias existem
        # ATENÇÃO: Agora olhamos para "LOC ATUAL" em vez de "OM ATUAL"
        if "LOC ATUAL" in df_plamov_compilado.columns and "TEMPO LOC" in df_plamov_compilado.columns:

            # --- TRATAMENTO DE DADOS ---
            # Converte TEMPO LOC para número (trata vírgula e ponto)
            df_plamov_compilado["TEMPO LOC"] = (
                df_plamov_compilado["TEMPO LOC"]
                .astype(str)
                .str.replace(',', '.')
                .apply(pd.to_numeric, errors='coerce')
                .fillna(0)
            )

            # Normaliza o texto da LOCALIDADE (Maiúsculo e sem espaços extras)
            serie_loc_atual = df_plamov_compilado["LOC ATUAL"].astype(
                str).str.strip().str.upper()
            # Adicione esta linha onde você normaliza os dados
            serie_om_atual = df_plamov_compilado["OM ATUAL"].astype(
                str).str.strip().str.upper()
            serie_tempo = df_plamov_compilado["TEMPO LOC"]

            # --- DEFINIÇÃO DAS REGRAS (AGORA POR LOCALIDADE) ---

            # Bloco A: Localidades Difíceis (>= 2 anos)
            cond_a = serie_loc_atual.isin(
                [x.upper() for x in loc_bloco_a]) & (serie_tempo >= 2)

            # Bloco B: Boa Vista / Porto Velho (>= 4 anos)
            cond_b = serie_loc_atual.isin(
                [x.upper() for x in loc_bloco_b]) & (serie_tempo >= 4)

            # Bloco C: Manaus / Belém (>= 5 anos)
            cond_c = serie_loc_atual.isin(
                [x.upper() for x in loc_bloco_c]) & (serie_tempo >= 5)

            # Bloco D: Santa Cruz (>= 6 anos)
            cond_d = serie_om_atual.isin(
                [x.upper() for x in loc_bloco_d]) & (serie_tempo >= 6)

            # --- SISTEMA DE PONTUAÇÃO (HIERARQUIA) ---
            # A função np.select respeita a ordem: Se for A, ganha 40 e sai. Se não, testa B...
            # Isso impede que alguém de Boa Vista com 10 anos caia no Bloco D (10 pts).
            # Como B vem antes, ele garante os 30 pts.

            condicoes = [cond_a, cond_b, cond_c, cond_d, cond_quer_localidades_d,
                         cond_quer_localidades_c_e_b, cond_quer_OMs_em_Santa_Cruz]
            pontos = [40,     30,     20,     10, 9, 8, 7]

            # Cria coluna auxiliar de Score
            df_plamov_compilado['SCORE_PRIORIDADE'] = np.select(
                condicoes, pontos, default=0)

            # 4. Define a ordem final
            # 1º: SCORE (Decrescente -> 40, 30, 20, 10, 9, 8, 7, 0)
            # 2º: MELHOR PRIO (Crescente -> 1ª Opção melhor que 2ª)
            # 3º: TEMPO LOC (Decrescente -> Quanto mais tempo, mais no topo dentro do mesmo bloco)
            # 4º: ANTIGUIDADE (Crescente -> Mais antigo primeiro)
            cols_ordenacao = ['SCORE_PRIORIDADE',
                              'MELHOR PRIO', 'TEMPO LOC', 'ANTIGUIDADE']
            asc_order = [False, True, False, True]

        else:
            # Fallback
            cols_ordenacao = ['MELHOR PRIO', 'TEMPO LOC', 'ANTIGUIDADE']
            asc_order = [True, False, True]

        # 5. Aplica a ordenação
        cols_finais = [
            c for c in cols_ordenacao if c in df_plamov_compilado.columns]
        asc_finais = [asc_order[i] for i, c in enumerate(
            cols_ordenacao) if c in df_plamov_compilado.columns]

        if cols_finais:
            df_plamov_compilado = df_plamov_compilado.sort_values(
                by=cols_finais, ascending=asc_finais)
            df_plamov_compilado = df_plamov_compilado.reset_index(drop=True)

        # --- CRIAÇÃO DOS DATAFRAMES POR GRUPO ---
        # Filtramos o DataFrame principal com base nos scores definidos anteriormente
        self.df_grupo_a = df_plamov_compilado[df_plamov_compilado['SCORE_PRIORIDADE'] == 40].copy(
        )
        self.df_grupo_b = df_plamov_compilado[df_plamov_compilado['SCORE_PRIORIDADE'] == 30].copy(
        )
        self.df_grupo_c = df_plamov_compilado[df_plamov_compilado['SCORE_PRIORIDADE'] == 20].copy(
        )
        self.df_grupo_d = df_plamov_compilado[df_plamov_compilado['SCORE_PRIORIDADE'] == 10].copy(
        )
        self.df_grupo_e = df_plamov_compilado[df_plamov_compilado['SCORE_PRIORIDADE'] == 9].copy(
        )
        self.df_grupo_f = df_plamov_compilado[df_plamov_compilado['SCORE_PRIORIDADE'] == 8].copy(
        )
        self.df_grupo_g = df_plamov_compilado[df_plamov_compilado['SCORE_PRIORIDADE'] == 7].copy(
        )
        # Grupo dos que não entraram em nenhuma regra (Score 0)
        self.df_sem_prioridade = df_plamov_compilado[df_plamov_compilado['SCORE_PRIORIDADE'] == 0].copy(
        )

        # Exemplo de Log para conferência no terminal
        print(f"--- Divisão de Blocos Concluída ---")
        print(f"A: {len(self.df_grupo_a)} | B: {len(self.df_grupo_b)} | C: {len(self.df_grupo_c)} | D: {len(self.df_grupo_d)}")
        print(f"E: {len(self.df_grupo_e)} | F: {len(self.df_grupo_f)} | G: {len(self.df_grupo_g)} | Restante: {len(self.df_sem_prioridade)}")

    # ---  FUNÇÃO PARA COLORIR O SARAM DOS MILITARES COM PRIORIDADE ESPECIAL DOS BLOCOS A, B, C, D ---
    # --- FUNÇÃO PARA COLORIR O SARAM E ADICIONAR DICA (TOOLTIP) ---
    def destacar_saram_prioritarios(self):
        global df_plamov_compilado

        mapa_blocos = {
            40: "Bloco A: quer sair de CACHIMBO, EIRUNEPÊ, SÃO GABRIEL, VILHENA",
            30: "Bloco B: quer sair de Boa Vista / Porto Velho",
            20: "Bloco C: quer sair de Manaus / Belém",
            10: "Bloco D: quer sair de Santa Cruz",
            9: "Bloco E: quer ir para CACHIMBO, EIRUNEPÊ, SÃO GABRIEL, VILHENA",
            8: "Bloco F: quer ir para Boa Vista / Porto Velho / Manaus / Belém",
            7: "Bloco G: quer ir para OMs em Santa Cruz"
        }

        for row in range(self.ui.tableWidget.rowCount()):
            if row < len(df_plamov_compilado):
                score = df_plamov_compilado.at[row, 'SCORE_PRIORIDADE']

                if score > 0:
                    # Aplica o tooltip em todas as células daquela linha
                    for col in range(self.ui.tableWidget.columnCount()):
                        item = self.ui.tableWidget.item(row, col)
                        if item:
                            # Define o aviso que aparece ao parar o mouse
                            item.setToolTip(mapa_blocos.get(
                                score, "Prioridade Especial"))

                            # Mantém o SARAM em vermelho (coluna 2)
                            if col == 2:
                                item.setForeground(QtGui.QColor("red"))

    # passar as páginas
    def Pag_Militares(self):
        self.ui.stackedWidget.setCurrentIndex(0)

    def Pag_Quadros_Especialidades(self):
        self.ui.stackedWidget.setCurrentIndex(1)

    def Pag_Relat_rio_TP(self):
        self.ui.stackedWidget.setCurrentIndex(2)

    def Pag_Mapa(self):
        self.ui.stackedWidget.setCurrentIndex(3)

    def Pag_Prioritarios(self):
        """Navega para a página de militares prioritários e atualiza a tabela."""
        indice = self.ui.stackedWidget.indexOf(self.page_prioritarios)
        self.ui.stackedWidget.setCurrentIndex(indice)
        self.popular_tabela_prioritarios()

    def Pag_QuerLocDificeis(self):
        """Navega para a página 'Querem Loc. Difíceis' e atualiza a tabela."""
        indice = self.ui.stackedWidget.indexOf(self.page_quer_loc_dificeis)
        self.ui.stackedWidget.setCurrentIndex(indice)
        self.popular_tabela_qld()

    def popular_tabela_prioritarios(self):
        """Preenche a tabela com os militares dos Blocos A, B, C e D."""
        global df_plamov_compilado

        if 'df_plamov_compilado' not in globals() or df_plamov_compilado.empty:
            self.tableWidget_prioritarios.setRowCount(0)
            self.lbl_resumo_prio.setText("Nenhum dado carregado.")
            return

        if 'SCORE_PRIORIDADE' not in df_plamov_compilado.columns:
            self.tableWidget_prioritarios.setRowCount(0)
            self.lbl_resumo_prio.setText(
                "Execute a ordenação primeiro (carregue os dados dos militares).")
            return

        # Filtra apenas blocos A (40), B (30), C (20) e D (10)
        scores_alvo = [40, 30, 20, 10]
        df_filtrado = df_plamov_compilado[
            df_plamov_compilado['SCORE_PRIORIDADE'].isin(scores_alvo)
        ].copy()

        if df_filtrado.empty:
            self.tableWidget_prioritarios.setRowCount(0)
            self.lbl_resumo_prio.setText(
                "Nenhum militar se enquadra nos blocos A, B, C ou D.")
            return

        # Mapeamento de scores para nomes de blocos
        mapa_blocos = {40: "A", 30: "B", 20: "C", 10: "D"}
        df_filtrado['BLOCO'] = df_filtrado['SCORE_PRIORIDADE'].map(mapa_blocos)

        # Resumo com contagem por bloco
        contagem = df_filtrado['BLOCO'].value_counts().sort_index()
        resumo_parts = []
        for bloco_letra in ['A', 'B', 'C', 'D']:
            qtd = contagem.get(bloco_letra, 0)
            resumo_parts.append(f"Bloco {bloco_letra}: {qtd}")
        total = len(df_filtrado)
        self.lbl_resumo_prio.setText(
            f"Total: {total} militares  |  " + "  |  ".join(resumo_parts))

        # Colunas a exibir (BLOCO na frente para destaque visual)
        COLUNAS_DESEJADAS = [
            "BLOCO", "LOC ATUAL", "OM ATUAL", "SARAM", "POSTO", "QUADRO",
            "ESP", "PROJETO", "TEMPO LOC", "LOC 1", "LOC 2", "LOC 3", "PLAMOV"
        ]
        colunas_existentes = [
            col for col in COLUNAS_DESEJADAS if col in df_filtrado.columns]

        self.tableWidget_prioritarios.setColumnCount(len(colunas_existentes))
        self.tableWidget_prioritarios.setRowCount(len(df_filtrado))
        self.tableWidget_prioritarios.setHorizontalHeaderLabels(
            colunas_existentes)

        # Cores por bloco
        cores_bloco = {
            "A": QtGui.QColor(220, 53, 69),    # Vermelho
            "B": QtGui.QColor(255, 153, 0),    # Laranja
            "C": QtGui.QColor(255, 193, 7),    # Amarelo
            "D": QtGui.QColor(0, 123, 255),    # Azul
        }

        # Tooltip descritivo por bloco
        dicas_bloco = {
            "A": "Bloco A: Guarantã do Norte, Eirunepê, São Gabriel, Vilhena (≥ 2 anos)",
            "B": "Bloco B: Boa Vista / Porto Velho (≥ 4 anos)",
            "C": "Bloco C: Manaus / Belém (≥ 5 anos)",
            "D": "Bloco D: Santa Cruz (≥ 6 anos)",
        }

        # Preenchimento visual
        for i, (idx_original, row_data) in enumerate(df_filtrado.iterrows()):
            bloco = row_data.get('BLOCO', '')
            cor_bloco = cores_bloco.get(bloco, QtGui.QColor(255, 255, 255))
            dica = dicas_bloco.get(bloco, "")

            for j, col_name in enumerate(colunas_existentes):
                valor = str(row_data.get(col_name, ''))
                item = QtWidgets.QTableWidgetItem(valor)
                item.setToolTip(dica)

                # Coluna BLOCO recebe cor forte do bloco
                if col_name == "BLOCO":
                    item.setBackground(cor_bloco)
                    item.setForeground(QtGui.QColor(255, 255, 255))
                    item.setTextAlignment(
                        QtCore.Qt.AlignmentFlag.AlignCenter)
                else:
                    # Coloração alternada para as demais colunas
                    if i % 2:
                        item.setBackground(QtGui.QColor(100, 139, 245))

                self.tableWidget_prioritarios.setItem(i, j, item)

        # Ajusta largura das colunas ao conteúdo
        self.tableWidget_prioritarios.resizeColumnsToContents()

        # Armazena o DataFrame filtrado para uso no painel direito
        self.df_prioritarios_filtrado = df_filtrado.copy()

    def atualizar_Painel_Direita_Prioritarios(self):
        # TODO (dúvida- Gerente) Eu não sei se devo mostrar as OMs que possuem vaga para a capacitação dos
        """Atualiza o painel direito da aba Militares Prioritários
        com as OMs disponíveis para o militar selecionado.
        Mostra primeiro as 3 opções do militar (LOC 1, 2, 3) destacadas,
        depois um separador, e todas as demais OMs por taxa de ocupação."""
        global df_OMs
        global df_plamov_compilado

        # Guard: dados devem estar carregados
        if 'df_plamov_compilado' not in globals() or df_plamov_compilado.empty:
            return
        if 'df_OMs' not in globals() or df_OMs.empty:
            return
        if self.df_prioritarios_filtrado.empty:
            return

        # Identifica a linha selecionada na tabela de prioritários
        linha_visual = self.tableWidget_prioritarios.currentRow()
        if linha_visual < 0 or linha_visual >= len(self.df_prioritarios_filtrado):
            return

        # Pega o índice original do DataFrame filtrado
        row_data = self.df_prioritarios_filtrado.iloc[linha_visual]
        idx_original = self.df_prioritarios_filtrado.index[linha_visual]

        # Dados do militar selecionado
        posto = str(pegar_posto(idx_original)).strip()
        quadro = str(pegar_quadro(idx_original)).strip()
        especialidade = str(pegar_especialidade(idx_original)).strip()
        Projeto = str(pegar_Projeto(idx_original)).strip()

        loc1 = str(row_data.get('LOC 1', '')).strip().upper()
        loc2 = str(row_data.get('LOC 2', '')).strip().upper()
        loc3 = str(row_data.get('LOC 3', '')).strip().upper()

        # --- CALCULA TAXA E VAGAS PARA CADA OM ---
        n_oms = df_OMs.shape[0]
        lista_taxa = [""] * n_oms
        lista_vagas = [""] * n_oms
        lista_loc = list(df_OMs["Localidade"].astype(
            str)) if "Localidade" in df_OMs.columns else [""] * n_oms

        usar_bma = (
            especialidade == "BMA"
            and 'df_TP_BMA' in globals()
            and not df_TP_BMA.empty
            and 'Unidade' in df_TP_BMA.columns
        )

        for k in range(n_oms):
            chegando = 0
            saindo = 0
            om_k = str(df_OMs.iloc[k, 0]).strip()

            if usar_bma:
                filtro_bma = (
                    (df_TP_BMA['Unidade'].astype(str).str.strip() == om_k) &
                    (df_TP_BMA['Posto'].astype(str).str.strip() == posto) &
                    (df_TP_BMA['Quadro'].astype(str).str.strip() == quadro) &
                    (df_TP_BMA['Projeto'].astype(str).str.strip() == Projeto)
                )
                vagas_OM_selecionada = df_TP_BMA[filtro_bma]

                if not vagas_OM_selecionada.empty:
                    if posto == "SGT":
                        query_posto = "POSTO in ['1S', '2S', '3S', 'SO']"
                    elif posto == "TN":
                        query_posto = "POSTO in ['1T', '2T']"
                    else:
                        query_posto = f"POSTO == '{posto}'"

                    chegando = df_plamov_compilado.query(
                        f"PLAMOV == '{om_k}' & {query_posto} & QUADRO == '{quadro}' & ESP == 'BMA' & `PROJETO` == '{Projeto}'"
                    ).shape[0]
                    saindo = df_plamov_compilado.query(
                        f"`OM ATUAL` == '{om_k}' & {query_posto} & QUADRO == '{quadro}' & ESP == 'BMA' & `PROJETO` == '{Projeto}' & PLAMOV != ''"
                    ).shape[0]

                    try:
                        TP = int(
                            vagas_OM_selecionada.iloc[0]['TLP Ano Corrente'])
                        existentes_na_TP = int(
                            vagas_OM_selecionada.iloc[0]['Existentes'])
                        if 'Localidade' in vagas_OM_selecionada.columns:
                            lista_loc[k] = str(
                                vagas_OM_selecionada.iloc[0]['Localidade'])
                    except (KeyError, ValueError):
                        TP = 0
                        existentes_na_TP = 0

                    if TP == 0:
                        lista_taxa[k] = "Sem TP"
                    else:
                        existentes_futuro = existentes_na_TP + chegando - saindo
                        lista_vagas[k] = int(
                            TP - existentes_na_TP + saindo - chegando)
                        lista_taxa[k] = round(
                            (float(existentes_futuro) / float(TP)) * 100, 2)
            else:
                if posto == "CP":
                    vagas_OM_selecionada = df_TP.query(
                        f"Unidade == '{om_k}' & ((Posto == 'CP/TN') | (Posto == 'CP')) & Quadro == '{quadro}' & Especialidade == '{especialidade}'")
                elif posto == "TN":
                    vagas_OM_selecionada = df_TP.query(
                        f"Unidade == '{om_k}' & ((Posto == 'CP/TN') | (Posto == 'TN')) & Quadro == '{quadro}' & Especialidade == '{especialidade}'")
                else:
                    vagas_OM_selecionada = df_TP.query(
                        f"Unidade == '{om_k}' & Posto == '{posto}' & Quadro == '{quadro}' & Especialidade == '{especialidade}'")

                if not vagas_OM_selecionada.empty:
                    if "Localidade" in vagas_OM_selecionada.columns:
                        lista_loc[k] = str(
                            vagas_OM_selecionada.iloc[0]["Localidade"])

                    chegando = df_plamov_compilado.query(
                        f"PLAMOV == '{om_k}' & POSTO == '{posto}' & QUADRO == '{quadro}' & ESP == '{especialidade}'").shape[0]
                    saindo = df_plamov_compilado.query(
                        f"`OM ATUAL` == '{om_k}' & POSTO == '{posto}' & QUADRO == '{quadro}' & ESP == '{especialidade}' & PLAMOV != ''").shape[0]

                    try:
                        TP = float(
                            vagas_OM_selecionada.iloc[0]["TLP Ano Corrente"])
                        existentes_na_TP = float(
                            vagas_OM_selecionada.iloc[0]["Existentes"])
                    except (KeyError, ValueError):
                        TP = 0
                        existentes_na_TP = 0

                    if TP != 0:
                        existentes = existentes_na_TP + chegando - saindo
                        lista_vagas[k] = int(TP + saindo - chegando)
                        lista_taxa[k] = round(
                            (float(existentes) / float(TP)) * 100, 2)
                    else:
                        lista_taxa[k] = "Sem TP"
                else:
                    lista_taxa[k] = "Sem TP"

        # Atualiza df_OMs com resultados calculados
        df_OMs["Taxa de Ocup."] = np.array(lista_taxa, dtype=object)
        df_OMs["Vagas"] = np.array(lista_vagas, dtype=object)
        df_OMs["Localidade"] = np.array(lista_loc, dtype=object)

        # --- CONSTRUÇÃO DAS LISTAS: OMs ESCOLHIDAS + DEMAIS ---
        # Identifica OMs das localidades escolhidas (LOC 1, 2, 3)
        oms_loc1 = []  # OMs cuja Localidade == LOC 1
        oms_loc2 = []  # OMs cuja Localidade == LOC 2
        oms_loc3 = []  # OMs cuja Localidade == LOC 3
        indices_escolhidos = set()

        for k in range(n_oms):
            om_loc = str(df_OMs.iloc[k, 3]).strip().upper(
            ) if "Localidade" in df_OMs.columns else ""
            if om_loc == loc1 and loc1 != "":
                oms_loc1.append(k)
                indices_escolhidos.add(k)
            elif om_loc == loc2 and loc2 != "":
                oms_loc2.append(k)
                indices_escolhidos.add(k)
            elif om_loc == loc3 and loc3 != "":
                oms_loc3.append(k)
                indices_escolhidos.add(k)

        # OMs restantes (não escolhidas), ordena por taxa de ocup. crescente / vagas decrescente
        demais = []
        for k in range(n_oms):
            if k not in indices_escolhidos:
                taxa_val = lista_taxa[k]
                vaga_val = lista_vagas[k]
                # Converte para números para ordenação
                try:
                    taxa_num = float(
                        taxa_val) if taxa_val != "" and taxa_val != "Sem TP" else 99999
                except (ValueError, TypeError):
                    taxa_num = 99999
                try:
                    vaga_num = int(vaga_val) if vaga_val != "" else -99999
                except (ValueError, TypeError):
                    vaga_num = -99999
                demais.append((k, taxa_num, vaga_num))

        # Ordena: taxa crescente, vagas decrescente (desempate)
        demais.sort(key=lambda x: (x[1], -x[2]))

        # --- MONTAGEM DA TABELA ---
        # Total de linhas: OMs LOC1 + OMs LOC2 + OMs LOC3 + 1 separador + demais
        n_escolhidas = len(oms_loc1) + len(oms_loc2) + len(oms_loc3)
        tem_escolhidas = n_escolhidas > 0
        n_separador = 1 if tem_escolhidas else 0
        total_linhas = n_escolhidas + n_separador + len(demais)

        self.tableWidget_prio_direita.setColumnCount(3)
        self.tableWidget_prio_direita.setRowCount(total_linhas)
        self.tableWidget_prio_direita.setHorizontalHeaderLabels(
            ["OM", "Taxa de Ocup.", "Vagas"])

        # Cores das opções
        cor_loc1 = QtGui.QColor(29, 181, 2)     # Verde (1ª Opção)
        cor_loc2 = QtGui.QColor(255, 243, 8)    # Amarelo (2ª Opção)
        cor_loc3 = QtGui.QColor(255, 0, 255)    # Magenta (3ª Opção)
        cor_separador = QtGui.QColor(80, 80, 80)  # Cinza escuro

        linha_tabela = 0

        def preencher_linha(linha_t, idx_om, cor_fundo):
            """Preenche uma linha da tabela com dados da OM e cor de fundo."""
            om_nome = str(df_OMs.iloc[idx_om, 0])
            taxa_val = lista_taxa[idx_om]
            vaga_val = lista_vagas[idx_om]

            if isinstance(taxa_val, (int, float)):
                taxa_str = "{:.2f}".format(taxa_val)
            else:
                taxa_str = str(taxa_val)

            valores = [om_nome, taxa_str, str(vaga_val)]
            for col, val in enumerate(valores):
                item = QtWidgets.QTableWidgetItem(val)
                item.setBackground(cor_fundo)
                # Para cores claras (amarelo), usa texto preto para legibilidade
                if cor_fundo == cor_loc2:
                    item.setForeground(QtGui.QColor(0, 0, 0))
                else:
                    item.setForeground(QtGui.QColor(255, 255, 255))
                self.tableWidget_prio_direita.setItem(linha_t, col, item)

        # 1ª Opção (Verde)
        for idx_om in oms_loc1:
            preencher_linha(linha_tabela, idx_om, cor_loc1)
            linha_tabela += 1

        # 2ª Opção (Amarelo)
        for idx_om in oms_loc2:
            preencher_linha(linha_tabela, idx_om, cor_loc2)
            linha_tabela += 1

        # 3ª Opção (Magenta)
        for idx_om in oms_loc3:
            preencher_linha(linha_tabela, idx_om, cor_loc3)
            linha_tabela += 1

        # Linha separadora
        if tem_escolhidas:
            for col in range(3):
                item_sep = QtWidgets.QTableWidgetItem(
                    "━━━━━━━━━━" if col == 0 else "━━━━━")
                item_sep.setBackground(cor_separador)
                item_sep.setForeground(QtGui.QColor(200, 200, 200))
                item_sep.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
                self.tableWidget_prio_direita.setItem(
                    linha_tabela, col, item_sep)
            linha_tabela += 1

        # Demais OMs (alternância de cor)
        for pos, (idx_om, _, _) in enumerate(demais):
            om_nome = str(df_OMs.iloc[idx_om, 0])
            taxa_val = lista_taxa[idx_om]
            vaga_val = lista_vagas[idx_om]

            if isinstance(taxa_val, (int, float)):
                taxa_str = "{:.2f}".format(taxa_val)
            else:
                taxa_str = str(taxa_val)

            valores = [om_nome, taxa_str, str(vaga_val)]
            for col, val in enumerate(valores):
                item = QtWidgets.QTableWidgetItem(val)
                # Coloração alternada
                if pos % 2:
                    item.setBackground(QtGui.QColor(100, 139, 245))
                self.tableWidget_prio_direita.setItem(linha_tabela, col, item)
            linha_tabela += 1

        # Ajusta largura das colunas
        self.tableWidget_prio_direita.resizeColumnsToContents()

    def alerta_deficit(self):
        pass

    def popular_tabela_qld(self):
        """Preenche a tabela com os militares que querem ir para localidades difíceis (Score 9)."""
        global df_plamov_compilado

        if 'df_plamov_compilado' not in globals() or df_plamov_compilado.empty:
            self.tableWidget_qld.setRowCount(0)
            self.lbl_resumo_qld.setText("Nenhum dado carregado.")
            return

        if 'SCORE_PRIORIDADE' not in df_plamov_compilado.columns:
            self.tableWidget_qld.setRowCount(0)
            self.lbl_resumo_qld.setText(
                "Execute a ordenação primeiro (carregue os dados dos militares).")
            return

        # Filtra apenas Score 9 (cond_quer_localidades_d)
        df_filtrado = df_plamov_compilado[
            df_plamov_compilado['SCORE_PRIORIDADE'] == 9
        ].copy()

        if df_filtrado.empty:
            self.tableWidget_qld.setRowCount(0)
            self.lbl_resumo_qld.setText(
                "Nenhum militar quer ir para localidades difíceis.")
            return

        total = len(df_filtrado)
        self.lbl_resumo_qld.setText(f"Total: {total} militares")

        # Colunas a exibir
        COLUNAS_DESEJADAS = [
            "LOC ATUAL", "OM ATUAL", "SARAM", "POSTO", "QUADRO",
            "ESP", "PROJETO", "TEMPO LOC", "LOC 1", "LOC 2", "LOC 3", "PLAMOV"
        ]
        colunas_existentes = [
            col for col in COLUNAS_DESEJADAS if col in df_filtrado.columns]

        self.tableWidget_qld.setColumnCount(len(colunas_existentes))
        self.tableWidget_qld.setRowCount(len(df_filtrado))
        self.tableWidget_qld.setHorizontalHeaderLabels(colunas_existentes)

        # Preenchimento visual
        for i, (idx_original, row_data) in enumerate(df_filtrado.iterrows()):
            for j, col_name in enumerate(colunas_existentes):
                valor = str(row_data.get(col_name, ''))
                item = QtWidgets.QTableWidgetItem(valor)
                item.setToolTip(
                    "Quer ir para localidade difícil (Guarantã, Eirunepê, S. Gabriel, Vilhena)")

                # Coloração alternada
                if i % 2:
                    item.setBackground(QtGui.QColor(100, 139, 245))

                self.tableWidget_qld.setItem(i, j, item)

        self.tableWidget_qld.resizeColumnsToContents()

        # Armazena o DataFrame filtrado para uso no painel direito
        self.df_qld_filtrado = df_filtrado.copy()

    def atualizar_Painel_Direita_QLD(self):
        """Atualiza o painel direito da aba 'Querem Loc. Difíceis'.
        Mostra primeiro as OMs de localidades difíceis (loc_bloco_a) que o
        militar escolheu, na ordem de preferência dele, destacadas em cores.
        Localidades que NÃO são difíceis (ex: Brasília) são ignoradas no destaque.
        Abaixo, todas as demais OMs por taxa de ocupação crescente /
        vagas crescente."""
        global df_OMs
        global df_plamov_compilado

        # Guard: dados devem estar carregados
        if 'df_plamov_compilado' not in globals() or df_plamov_compilado.empty:
            return
        if 'df_OMs' not in globals() or df_OMs.empty:
            return
        if self.df_qld_filtrado.empty:
            return

        # Identifica a linha selecionada na tabela
        linha_visual = self.tableWidget_qld.currentRow()
        if linha_visual < 0 or linha_visual >= len(self.df_qld_filtrado):
            return

        # Pega o índice original do DataFrame filtrado
        row_data = self.df_qld_filtrado.iloc[linha_visual]
        idx_original = self.df_qld_filtrado.index[linha_visual]

        # Dados do militar selecionado
        posto = str(pegar_posto(idx_original)).strip()
        quadro = str(pegar_quadro(idx_original)).strip()
        especialidade = str(pegar_especialidade(idx_original)).strip()
        Projeto = str(pegar_Projeto(idx_original)).strip()

        loc1 = str(row_data.get('LOC 1', '')).strip().upper()
        loc2 = str(row_data.get('LOC 2', '')).strip().upper()
        loc3 = str(row_data.get('LOC 3', '')).strip().upper()

        # --- CALCULA TAXA E VAGAS PARA CADA OM ---
        n_oms = df_OMs.shape[0]
        lista_taxa = [""] * n_oms
        lista_vagas = [""] * n_oms
        lista_loc = list(df_OMs["Localidade"].astype(
            str)) if "Localidade" in df_OMs.columns else [""] * n_oms

        usar_bma = (
            especialidade == "BMA"
            and 'df_TP_BMA' in globals()
            and not df_TP_BMA.empty
            and 'Unidade' in df_TP_BMA.columns
        )

        for k in range(n_oms):
            chegando = 0
            saindo = 0
            om_k = str(df_OMs.iloc[k, 0]).strip()

            if usar_bma:
                filtro_bma = (
                    (df_TP_BMA['Unidade'].astype(str).str.strip() == om_k) &
                    (df_TP_BMA['Posto'].astype(str).str.strip() == posto) &
                    (df_TP_BMA['Quadro'].astype(str).str.strip() == quadro) &
                    (df_TP_BMA['Projeto'].astype(str).str.strip() == Projeto)
                )
                vagas_OM_selecionada = df_TP_BMA[filtro_bma]

                if not vagas_OM_selecionada.empty:
                    if posto == "SGT":
                        query_posto = "POSTO in ['1S', '2S', '3S', 'SO']"
                    elif posto == "TN":
                        query_posto = "POSTO in ['1T', '2T']"
                    else:
                        query_posto = f"POSTO == '{posto}'"

                    chegando = df_plamov_compilado.query(
                        f"PLAMOV == '{om_k}' & {query_posto} & QUADRO == '{quadro}' & ESP == 'BMA' & `PROJETO` == '{Projeto}'"
                    ).shape[0]
                    saindo = df_plamov_compilado.query(
                        f"`OM ATUAL` == '{om_k}' & {query_posto} & QUADRO == '{quadro}' & ESP == 'BMA' & `PROJETO` == '{Projeto}' & PLAMOV != ''"
                    ).shape[0]

                    try:
                        TP = int(
                            vagas_OM_selecionada.iloc[0]['TLP Ano Corrente'])
                        existentes_na_TP = int(
                            vagas_OM_selecionada.iloc[0]['Existentes'])
                        if 'Localidade' in vagas_OM_selecionada.columns:
                            lista_loc[k] = str(
                                vagas_OM_selecionada.iloc[0]['Localidade'])
                    except (KeyError, ValueError):
                        TP = 0
                        existentes_na_TP = 0

                    if TP == 0:
                        lista_taxa[k] = "Sem TP"
                    else:
                        existentes_futuro = existentes_na_TP + chegando - saindo
                        lista_vagas[k] = int(
                            TP - existentes_na_TP + saindo - chegando)
                        lista_taxa[k] = round(
                            (float(existentes_futuro) / float(TP)) * 100, 2)
            else:
                if posto == "CP":
                    vagas_OM_selecionada = df_TP.query(
                        f"Unidade == '{om_k}' & ((Posto == 'CP/TN') | (Posto == 'CP')) & Quadro == '{quadro}' & Especialidade == '{especialidade}'")
                elif posto == "TN":
                    vagas_OM_selecionada = df_TP.query(
                        f"Unidade == '{om_k}' & ((Posto == 'CP/TN') | (Posto == 'TN')) & Quadro == '{quadro}' & Especialidade == '{especialidade}'")
                else:
                    vagas_OM_selecionada = df_TP.query(
                        f"Unidade == '{om_k}' & Posto == '{posto}' & Quadro == '{quadro}' & Especialidade == '{especialidade}'")

                if not vagas_OM_selecionada.empty:
                    if "Localidade" in vagas_OM_selecionada.columns:
                        lista_loc[k] = str(
                            vagas_OM_selecionada.iloc[0]["Localidade"])

                    chegando = df_plamov_compilado.query(
                        f"PLAMOV == '{om_k}' & POSTO == '{posto}' & QUADRO == '{quadro}' & ESP == '{especialidade}'").shape[0]
                    saindo = df_plamov_compilado.query(
                        f"`OM ATUAL` == '{om_k}' & POSTO == '{posto}' & QUADRO == '{quadro}' & ESP == '{especialidade}' & PLAMOV != ''").shape[0]

                    try:
                        TP = float(
                            vagas_OM_selecionada.iloc[0]["TLP Ano Corrente"])
                        existentes_na_TP = float(
                            vagas_OM_selecionada.iloc[0]["Existentes"])
                    except (KeyError, ValueError):
                        TP = 0
                        existentes_na_TP = 0

                    if TP != 0:
                        existentes = existentes_na_TP + chegando - saindo
                        lista_vagas[k] = int(TP + saindo - chegando)
                        lista_taxa[k] = round(
                            (float(existentes) / float(TP)) * 100, 2)
                    else:
                        lista_taxa[k] = "Sem TP"
                else:
                    lista_taxa[k] = "Sem TP"

        # Atualiza df_OMs com resultados calculados
        df_OMs["Taxa de Ocup."] = np.array(lista_taxa, dtype=object)
        df_OMs["Vagas"] = np.array(lista_vagas, dtype=object)
        df_OMs["Localidade"] = np.array(lista_loc, dtype=object)

        # --- CONSTRUÇÃO DAS LISTAS: SÓ LOCALIDADES DIFÍCEIS (loc_bloco_a) ---
        # Apenas localidades difíceis recebem destaque no topo.
        # Localidades comuns (ex: Brasília) NÃO são destacadas mesmo que
        # o militar as tenha escolhido nas opções.
        loc_bloco_a = ["GUARANTÃ DO NORTE", "EIRUNEPÊ", "EIRUNEPE",
                       "SÃO GABRIEL DA CACHOEIRA", "SAO GABRIEL DA CACHOEIRA", "VILHENA"]
        loc_bloco_a_upper = [x.upper() for x in loc_bloco_a]

        # Identifica quais opções do militar são loc. difíceis
        # Mantém a ordem de escolha: 1ª → 2ª → 3ª
        opcoes_dificeis = []  # lista de tuplas (localidade, cor)
        cor_loc1 = QtGui.QColor(29, 181, 2)     # Verde (1ª Opção)
        cor_loc2 = QtGui.QColor(255, 243, 8)    # Amarelo (2ª Opção)
        cor_loc3 = QtGui.QColor(255, 0, 255)    # Magenta (3ª Opção)

        if loc1 in loc_bloco_a_upper and loc1 != "":
            opcoes_dificeis.append((loc1, cor_loc1))
        if loc2 in loc_bloco_a_upper and loc2 != "":
            opcoes_dificeis.append((loc2, cor_loc2))
        if loc3 in loc_bloco_a_upper and loc3 != "":
            opcoes_dificeis.append((loc3, cor_loc3))

        # Monta lista ordenada de OMs destacadas (na ordem de preferência)
        oms_destacadas = []  # lista de tuplas (indice_om, cor)
        indices_escolhidos = set()

        for loc_dificil, cor in opcoes_dificeis:
            for k in range(n_oms):
                om_loc = str(df_OMs.iloc[k, 3]).strip().upper(
                ) if "Localidade" in df_OMs.columns else ""
                if om_loc == loc_dificil and k not in indices_escolhidos:
                    oms_destacadas.append((k, cor))
                    indices_escolhidos.add(k)

        # OMs restantes: taxa crescente, vagas crescente
        demais = []
        for k in range(n_oms):
            if k not in indices_escolhidos:
                taxa_val = lista_taxa[k]
                vaga_val = lista_vagas[k]
                try:
                    taxa_num = float(
                        taxa_val) if taxa_val != "" and taxa_val != "Sem TP" else 99999
                except (ValueError, TypeError):
                    taxa_num = 99999
                try:
                    vaga_num = int(vaga_val) if vaga_val != "" else 99999
                except (ValueError, TypeError):
                    vaga_num = 99999
                demais.append((k, taxa_num, vaga_num))

        demais.sort(key=lambda x: (x[1], x[2]))

        # --- MONTAGEM DA TABELA ---
        n_escolhidas = len(oms_destacadas)
        tem_escolhidas = n_escolhidas > 0
        n_separador = 1 if tem_escolhidas else 0
        total_linhas = n_escolhidas + n_separador + len(demais)

        self.tableWidget_qld_direita.setColumnCount(3)
        self.tableWidget_qld_direita.setRowCount(total_linhas)
        self.tableWidget_qld_direita.setHorizontalHeaderLabels(
            ["OM", "Taxa de Ocup.", "Vagas"])

        cor_separador = QtGui.QColor(80, 80, 80)

        linha_tabela = 0

        def preencher_linha(linha_t, idx_om, cor_fundo):
            om_nome = str(df_OMs.iloc[idx_om, 0])
            taxa_val = lista_taxa[idx_om]
            vaga_val = lista_vagas[idx_om]
            if isinstance(taxa_val, (int, float)):
                taxa_str = "{:.2f}".format(taxa_val)
            else:
                taxa_str = str(taxa_val)
            valores = [om_nome, taxa_str, str(vaga_val)]
            for col, val in enumerate(valores):
                item = QtWidgets.QTableWidgetItem(val)
                item.setBackground(cor_fundo)
                if cor_fundo == cor_loc2:
                    item.setForeground(QtGui.QColor(0, 0, 0))
                else:
                    item.setForeground(QtGui.QColor(255, 255, 255))
                self.tableWidget_qld_direita.setItem(linha_t, col, item)

        # OMs de localidades difíceis (na ordem de preferência do militar)
        for idx_om, cor in oms_destacadas:
            preencher_linha(linha_tabela, idx_om, cor)
            linha_tabela += 1

        if tem_escolhidas:
            for col in range(3):
                item_sep = QtWidgets.QTableWidgetItem(
                    "━━━━━━━━━━" if col == 0 else "━━━━━")
                item_sep.setBackground(cor_separador)
                item_sep.setForeground(QtGui.QColor(200, 200, 200))
                item_sep.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
                self.tableWidget_qld_direita.setItem(
                    linha_tabela, col, item_sep)
            linha_tabela += 1

        for pos, (idx_om, _, _) in enumerate(demais):
            om_nome = str(df_OMs.iloc[idx_om, 0])
            taxa_val = lista_taxa[idx_om]
            vaga_val = lista_vagas[idx_om]
            if isinstance(taxa_val, (int, float)):
                taxa_str = "{:.2f}".format(taxa_val)
            else:
                taxa_str = str(taxa_val)
            valores = [om_nome, taxa_str, str(vaga_val)]
            for col, val in enumerate(valores):
                item = QtWidgets.QTableWidgetItem(val)
                if pos % 2:
                    item.setBackground(QtGui.QColor(100, 139, 245))
                self.tableWidget_qld_direita.setItem(linha_tabela, col, item)
            linha_tabela += 1

        self.tableWidget_qld_direita.resizeColumnsToContents()

    def atualizar_Painel_Direita(self):
        global df_OMs

        # Guard: só executa se os dados estiverem carregados
        if 'df_plamov_compilado' not in globals() or df_plamov_compilado.empty:
            return
        if 'df_TP' not in globals():
            return

        # Guard: linha=-1 significa que nenhuma linha está selecionada
        linha = self.linha_ativa_dados_militares()
        if linha < 0 or linha >= len(df_plamov_compilado):
            return

        # Guard: df_OMs deve existir e ter linhas
        if 'df_OMs' not in globals() or df_OMs.empty:
            return

        # 1. PEGA DADOS DO MILITAR (Sanitizados)
        posto = str(pegar_posto(linha)).strip()
        quadro = str(pegar_quadro(linha)).strip()
        especialidade = str(pegar_especialidade(linha)).strip()
        Projeto = str(pegar_Projeto(linha)).strip()

        loc1 = pegar_LOC1(linha)
        loc2 = pegar_LOC2(linha)
        loc3 = pegar_LOC3(linha)
        loc_atual = pegar_LOC_atual(linha)

        # 2. LISTAS DE RESULTADO — evita df_OMs.loc dentro do loop,
        #    que no pandas 2.x dispara TypeError por StringDtype.
        #    Os valores são acumulados em listas Python e atribuídos
        #    como arrays numpy de uma única vez ao final.
        n_oms = df_OMs.shape[0]
        lista_taxa = [""] * n_oms
        lista_vagas = [""] * n_oms
        lista_loc = list(df_OMs["Localidade"].astype(
            str)) if "Localidade" in df_OMs.columns else [""] * n_oms

        usar_bma = (
            especialidade == "BMA"
            and 'df_TP_BMA' in globals()
            and not df_TP_BMA.empty
            and 'Unidade' in df_TP_BMA.columns
        )

        # 3. CONFIGURAÇÃO BÁSICA DA TABELA
        self.ui.tableWidget_2.setColumnCount(3)
        self.ui.tableWidget_2.setRowCount(n_oms)
        self.ui.tableWidget_2.setHorizontalHeaderLabels(
            ["OM", "Taxa de Ocup.", "Vagas"])

        # 4. LOOP DE CÁLCULO — só lê df_OMs, nunca escreve nele
        for k in range(n_oms):
            chegando = 0
            saindo = 0
            om_k = str(df_OMs.iloc[k, 0]).strip()

            if usar_bma:
                filtro_bma = (
                    (df_TP_BMA['Unidade'].astype(str).str.strip() == om_k) &
                    (df_TP_BMA['Posto'].astype(str).str.strip() == posto) &
                    (df_TP_BMA['Quadro'].astype(str).str.strip() == quadro) &
                    (df_TP_BMA['Projeto'].astype(str).str.strip() == Projeto)
                )
                vagas_OM_selecionada = df_TP_BMA[filtro_bma]

                if not vagas_OM_selecionada.empty:
                    if posto == "SGT":
                        query_posto = "POSTO in ['1S', '2S', '3S', 'SO']"
                    elif posto == "TN":
                        query_posto = "POSTO in ['1T', '2T']"
                    else:
                        query_posto = f"POSTO == '{posto}'"

                    chegando = df_plamov_compilado.query(
                        f"PLAMOV == '{om_k}' & {query_posto} & QUADRO == '{quadro}' & ESP == 'BMA' & `PROJETO` == '{Projeto}'"
                    ).shape[0]
                    saindo = df_plamov_compilado.query(
                        f"`OM ATUAL` == '{om_k}' & {query_posto} & QUADRO == '{quadro}' & ESP == 'BMA' & `PROJETO` == '{Projeto}' & PLAMOV != ''"
                    ).shape[0]

                    try:
                        TP = int(
                            vagas_OM_selecionada.iloc[0]['TLP Ano Corrente'])
                        existentes_na_TP = int(
                            vagas_OM_selecionada.iloc[0]['Existentes'])
                        if 'Localidade' in vagas_OM_selecionada.columns:
                            lista_loc[k] = str(
                                vagas_OM_selecionada.iloc[0]['Localidade'])
                    except (KeyError, ValueError):
                        TP = 0
                        existentes_na_TP = 0

                    if TP == 0:
                        lista_taxa[k] = "Sem TP"
                    else:
                        existentes_futuro = existentes_na_TP + chegando - saindo
                        lista_vagas[k] = int(
                            TP - existentes_na_TP + saindo - chegando)
                        lista_taxa[k] = round(
                            (float(existentes_futuro) / float(TP)) * 100, 2)

            else:
                if posto == "CP":
                    vagas_OM_selecionada = df_TP.query(
                        f"Unidade == '{om_k}' & ((Posto == 'CP/TN') | (Posto == 'CP')) & Quadro == '{quadro}' & Especialidade == '{especialidade}'")
                elif posto == "TN":
                    vagas_OM_selecionada = df_TP.query(
                        f"Unidade == '{om_k}' & ((Posto == 'CP/TN') | (Posto == 'TN')) & Quadro == '{quadro}' & Especialidade == '{especialidade}'")
                else:
                    vagas_OM_selecionada = df_TP.query(
                        f"Unidade == '{om_k}' & Posto == '{posto}' & Quadro == '{quadro}' & Especialidade == '{especialidade}'")

                if not vagas_OM_selecionada.empty:
                    if "Localidade" in vagas_OM_selecionada.columns:
                        lista_loc[k] = str(
                            vagas_OM_selecionada.iloc[0]["Localidade"])

                    if posto == "CP" or posto == "TN":
                        chegando = df_plamov_compilado.query(
                            f"PLAMOV == '{om_k}' & POSTO == '{posto}' & QUADRO == '{quadro}' & ESP == '{especialidade}'").shape[0]
                        saindo = df_plamov_compilado.query(
                            f"`OM ATUAL` == '{om_k}' & POSTO == '{posto}' & QUADRO == '{quadro}' & ESP == '{especialidade}' & PLAMOV != ''").shape[0]
                    else:
                        chegando = df_plamov_compilado.query(
                            f"PLAMOV == '{om_k}' & POSTO == '{posto}' & QUADRO == '{quadro}' & ESP == '{especialidade}'").shape[0]
                        saindo = df_plamov_compilado.query(
                            f"`OM ATUAL` == '{om_k}' & POSTO == '{posto}' & QUADRO == '{quadro}' & ESP == '{especialidade}' & PLAMOV != ''").shape[0]

                    try:
                        TP = float(
                            vagas_OM_selecionada.iloc[0]["TLP Ano Corrente"])
                        existentes_na_TP = float(
                            vagas_OM_selecionada.iloc[0]["Existentes"])
                    except (KeyError, ValueError):
                        TP = 0
                        existentes_na_TP = 0

                    if TP != 0:
                        existentes = existentes_na_TP + chegando - saindo
                        lista_vagas[k] = int(TP + saindo - chegando)
                        lista_taxa[k] = round(
                            (float(existentes) / float(TP)) * 100, 2)
                    else:
                        lista_taxa[k] = "Sem TP"
                else:
                    lista_taxa[k] = "Sem TP"

        # 4b. ATRIBUIÇÃO EM BLOCO — numpy garante dtype=object, sem passar pelo pandas
        df_OMs["Taxa de Ocup."] = np.array(lista_taxa,  dtype=object)
        df_OMs["Vagas"] = np.array(lista_vagas, dtype=object)
        df_OMs["Localidade"] = np.array(lista_loc,   dtype=object)

        # 5. ORDENAÇÃO
        # Ordenamos usando o valor numérico completo antes de formatar
        df_OMs.sort_values(by=['Taxa de Ocup.', 'Vagas'],
                           ascending=[True, False], inplace=True)
        df_OMs.reset_index(drop=True, inplace=True)

        # 6. PREENCHIMENTO VISUAL (Com formatação de 2 casas)
        localidade_atual_do_militar = str(loc_atual).strip().upper()

        for i in range(df_OMs.shape[0]):
            for j in range(3):
                valor_original = df_OMs.iloc[i, j]

                # --- AQUI ESTÁ A MUDANÇA ---
                # Se for a coluna 1 (Taxa) e for número, formata com 2 casas
                if j == 1 and isinstance(valor_original, (int, float)):
                    texto_celula = "{:.2f}".format(valor_original)
                else:
                    texto_celula = str(valor_original)
                # ---------------------------

                item = QtWidgets.QTableWidgetItem(texto_celula)
                self.ui.tableWidget_2.setItem(i, j, item)

                # Coloração Alternada
                if i % 2:
                    self.ui.tableWidget_2.item(i, j).setBackground(
                        QtGui.QColor(100, 139, 245))

                # Cores de Localidade
                om_loc = str(df_OMs.iloc[i, 3]).strip().upper()
                l1 = str(loc1).strip().upper()
                l2 = str(loc2).strip().upper()
                l3 = str(loc3).strip().upper()

                if om_loc == l3 and l3 != "":
                    self.ui.tableWidget_2.item(i, j).setBackground(
                        QtGui.QColor(255, 0, 255))
                if om_loc == l2 and l2 != "":
                    self.ui.tableWidget_2.item(i, j).setBackground(
                        QtGui.QColor(255, 243, 8))
                if om_loc == l1 and l1 != "":
                    self.ui.tableWidget_2.item(i, j).setBackground(
                        QtGui.QColor(29, 181, 2))

            # Destaque Localidade Atual
            if om_loc == localidade_atual_do_militar and localidade_atual_do_militar != "":
                # Reinsere o item da primeira coluna para garantir que a cor pegue
                item = QtWidgets.QTableWidgetItem(str(df_OMs.iloc[i, 0]))
                self.ui.tableWidget_2.setItem(i, 0, item)
                self.ui.tableWidget_2.item(i, 0).setBackground(
                    QtGui.QColor(107, 107, 106))

        # Reativa Filtros
        if self.filtros_ativos_direita:
            self.executar_filtros_combinados(
                self.ui.tableWidget_2, self.filtros_ativos_direita)

        self.analisar_impacto_transferencia()

    def contar_militares_mesma_Projeto(self):
        global df_plamov_compilado

        # 1. Identifica a linha selecionada (Índice atual)
        linha_atual = self.linha_ativa_dados_militares()

        # 2. Pega a Projeto do militar selecionado
        Projeto_alvo = pegar_Projeto(linha_atual)

        # 3. Validação
        if not Projeto_alvo or Projeto_alvo == "nan":
            return 0

        # ==============================================================================
        # A MÁGICA ACONTECE AQUI: FATIAMENTO (SLICING)
        # ==============================================================================

        # Cria um novo DataFrame temporário contendo apenas as linhas
        # do índice seguinte (linha_atual + 1) até o final da lista (:).
        df_abaixo = df_plamov_compilado.iloc[linha_atual + 1:]

        # 4. Filtra apenas nesse DataFrame "recortado"
        filtro = df_abaixo["PROJETO"].astype(
            str).str.strip() == Projeto_alvo.strip()

        # Conta as linhas resultantes
        quantidade = df_abaixo[filtro].shape[0]

        # 5. Retorno/Exibição
        print(f"--- CONTAGEM ---")
        print(f"Militar atual (Linha): {linha_atual}")
        print(f"Projeto: {Projeto_alvo}")
        print(f"Militares abaixo (na fila): {quantidade}")

        return quantidade

    # --- FUNÇÕES PARA GRÁFICOS ---

    def Pag_Graficos(self):
        # Muda para a página nova que criamos (o índice é o último da lista)
        indice_graficos = self.ui.stackedWidget.count() - 1
        self.ui.stackedWidget.setCurrentIndex(indice_graficos)

        # Gera o gráfico atualizado
        self.gerar_dashboard()

    def gerar_dashboard(self):
        global df_plamov_compilado
        global df_OMs  # Precisamos disso para saber a localidade das OMs

        # 1. Limpa layout anterior
        for i in reversed(range(self.layout_graficos.count())):
            item = self.layout_graficos.itemAt(i)
            if item.widget():
                item.widget().setParent(None)

        if 'df_plamov_compilado' not in globals() or df_plamov_compilado.empty:
            aviso = QtWidgets.QLabel("Nenhum dado carregado.")
            aviso.setStyleSheet(
                "font-size: 18px; color: gray; qproperty-alignment: AlignCenter;")
            self.layout_graficos.addWidget(aviso)
            return

        # 2. Prepara o Canvas com tamanho maior para caber 3 gráficos
        # Usamos 'tight_layout' para ajustar automaticamente
        canvas = GraficoCanvas(self, width=10, height=12, dpi=100)

        # Cria uma grade de gráficos: 2 linhas, 2 colunas
        # ax1: Canto superior esquerdo (Pizza)
        # ax2: Canto superior direito (Barras Posto)
        # ax3: Parte inferior inteira (Barras OMs)
        gs = canvas.fig.add_gridspec(2, 2)
        ax1 = canvas.fig.add_subplot(gs[0, 0])
        ax2 = canvas.fig.add_subplot(gs[0, 1])
        # Ocupa as duas colunas de baixo
        ax3 = canvas.fig.add_subplot(gs[1, :])

        try:
            # Filtra apenas quem tem destino definido
            df_movimentados = df_plamov_compilado[df_plamov_compilado['PLAMOV'] != ""].copy(
            )

            if df_movimentados.empty:
                self.layout_graficos.addWidget(
                    QtWidgets.QLabel("Não há movimentações definidas."))
                return

            # =========================================================
            # GRÁFICO 1: TERMÔMETRO DE SATISFAÇÃO
            # =========================================================

            # Precisamos mapear a OM de Destino (PLAMOV) para sua Localidade
            # para comparar com LOC 1, LOC 2, LOC 3.

            # Cria um dicionário {OM: Localidade} baseado no df_OMs carregado
            if 'df_OMs' in globals() and not df_OMs.empty:
                dict_loc = dict(zip(df_OMs['OMs'].astype(
                    str).str.strip(), df_OMs['Localidade'].astype(str).str.strip()))
            else:
                dict_loc = {}

            # Função auxiliar para categorizar
            def classificar_atendimento(row):
                destino = str(row['PLAMOV']).strip()
                # Tenta pegar a localidade da OM de destino, se não achar, usa o próprio nome
                loc_destino = dict_loc.get(destino, destino)

                l1 = str(row['LOC 1']).strip()
                l2 = str(row['LOC 2']).strip()
                l3 = str(row['LOC 3']).strip()

                # Compara Localidade Destino com Localidades Escolhidas
                # (Ou compara direto o nome da OM se o usuário escreveu OM nas opções)
                if loc_destino == l1 or destino == l1:
                    return "1ª Opção"
                if loc_destino == l2 or destino == l2:
                    return "2ª Opção"
                if loc_destino == l3 or destino == l3:
                    return "3ª Opção"
                return "Interesse do Serviço"  # Não atendeu nenhuma das 3

            contagem_satisfacao = df_movimentados.apply(
                classificar_atendimento, axis=1).value_counts()

            # Cores para o gráfico de pizza
            cores_map = {
                "1ª Opção": "#2ca02c",  # Verde
                "2ª Opção": "#1f77b4",  # Azul
                "3ª Opção": "#ff7f0e",  # Laranja
                "Interesse do Serviço": "#d62728"  # Vermelho
            }
            cores = [cores_map.get(x, "#999999")
                     for x in contagem_satisfacao.index]

            ax1.pie(contagem_satisfacao, labels=contagem_satisfacao.index,
                    autopct='%1.1f%%', startangle=90, colors=cores)
            ax1.set_title('Índice de Atendimento (Satisfação)',
                          fontsize=10, fontweight='bold')

            # =========================================================
            # GRÁFICO 2: MOVIMENTAÇÃO POR POSTO
            # =========================================================
            contagem_posto = df_movimentados['POSTO'].value_counts()

            ax2.bar(contagem_posto.index,
                    contagem_posto.values, color='#4a90e2')
            ax2.set_title('Volume por Posto/Graduação',
                          fontsize=10, fontweight='bold')
            ax2.grid(axis='y', linestyle='--', alpha=0.5)

            # =========================================================
            # GRÁFICO 3: TOP 10 OMs DE DESTINO
            # =========================================================
            top_oms = df_movimentados['PLAMOV'].value_counts().head(10)

            barras = ax3.barh(top_oms.index, top_oms.values, color='#8856a7')
            ax3.invert_yaxis()  # Maior no topo
            ax3.set_title('Top 10 OMs de Destino (Para onde estão indo?)',
                          fontsize=10, fontweight='bold')
            ax3.bar_label(barras, padding=3)
            ax3.grid(axis='x', linestyle='--', alpha=0.5)

            # Ajuste final
            canvas.fig.tight_layout(pad=3.0)  # Mais espaço entre gráficos
            self.layout_graficos.addWidget(canvas)

        except Exception as e:
            erro = QtWidgets.QLabel(f"Erro ao gerar gráficos: {e}")
            self.layout_graficos.addWidget(erro)
            print(f"Erro detalhado Dashboard: {e}")

    def analisar_impacto_transferencia(self):
        """
        Verifica se a saída do militar vai quebrar a taxa de 70% da OM de origem
        e conta quantos reservas existem abaixo na lista.
        """
        self.ui.statusbar.clearMessage()
        self.ui.statusbar.setStyleSheet("")

        global df_plamov_compilado
        global df_TP_BMA

        # 1. Dados do Militar Selecionado
        linha_atual = self.linha_ativa_dados_militares()

        # Cuidado: Pegar a OM ATUAL (Origem), não o destino (PLAMOV)
        om_origem = str(
            df_plamov_compilado["OM ATUAL"].iloc[linha_atual]).strip()
        Projeto = pegar_Projeto(linha_atual)
        especialidade = pegar_especialidade(linha_atual)

        if not Projeto or Projeto == "nan":
            return  # Sem dados para analisar

        # 2. Diagnóstico da OM de Origem (TP BMA)
        # Filtra a TP BMA pela OM e Projeto (somando todos os postos)
        filtro_tp = (
            (df_TP_BMA['Unidade'].astype(str).str.strip() == om_origem) &
            (df_TP_BMA['Projeto'].astype(str).str.strip() == Projeto) &
            (df_TP_BMA['Especialidade'].astype(
                str).str.strip() == especialidade)
        )
        dados_tp = df_TP_BMA[filtro_tp]

        if dados_tp.empty:
            self.ui.statusbar.showMessage(
                f"A OM de origem ({om_origem}) desse militar não tem previsão na TP para {especialidade}.")
            self.ui.statusbar.setStyleSheet("color: red; font-weight: bold;")
            # Não faz sentido calcular a nova taxa de ocupação, pois, se não há TP, não existe taxa de ocupação.
            return

        # Soma TLP e Existentes (caso haja distinção de postos, somamos tudo daquela Projeto)
        # Ajuste os nomes das colunas 'TLP Ano Corrente' e 'Existentes' se necessário
        try:
            total_meta = dados_tp['TLP Ano Corrente'].sum()
            total_existentes = dados_tp['Existentes'].sum()
        except KeyError:
            # Fallback para índices se os nomes mudaram
            total_meta = dados_tp.iloc[:, 4].sum()
            total_existentes = dados_tp.iloc[:, 5].sum()

        if total_meta == 0:
            self.ui.statusbar.showMessage(
                f"Segundo o Retório TP, não há nenhum militar em {om_origem}. Verifique a OM de origem do militar está correta ou se o Retório TP está atualizado.")
            self.ui.statusbar.setStyleSheet("color: red; font-weight: bold;")
            return  # Evita divisão por zero

        # 3. Simulação da Saída
        taxa_atual = total_existentes / total_meta
        taxa_projetada = (total_existentes - 1) / total_meta

        # 4. Verificação do Gatilho (Abaixo de 70%)
        # Se a taxa JÁ ERA ruim, ou SE VAI FICAR ruim
        if taxa_projetada < 0.70:

            # 5. Busca de Reservas (Militares abaixo na lista)
            df_abaixo = df_plamov_compilado.iloc[linha_atual + 1:]

            # Filtra apenas pela mesma Projeto (conforme sua regra)
            reservas = df_abaixo[df_abaixo["PROJETO"].astype(
                str).str.strip() == Projeto].shape[0]

            # 6. GERAÇÃO DO ALERTA (Mensagem Prática)
            msg_alerta = (
                f"⚠️ ATENÇÃO: Se esse militar for transferido, a taxa de ocupação da {om_origem} diminuirá para {taxa_projetada:.1%} "
                f"(Meta: 70%).\n"
                f"RESERVAS DISPONÍVEIS ABAIXO: {reservas} militares de {Projeto}."
            )

            # SUGESTÃO PRÁTICA: Mostrar na Barra de Status do Programa (Rodapé)
            # Isso é discreto mas visível para o analista
            self.ui.statusbar.showMessage(msg_alerta)

            # Opcional: Mudar a cor da StatusBar para vermelho para chamar atenção
            self.ui.statusbar.setStyleSheet("color: red; font-weight: bold;")

        else:
            # Se estiver tudo seguro
            self.ui.statusbar.showMessage(
                f"✔ Saída segura. {om_origem} manterá taxa de {taxa_projetada:.1%} (Sub: {Projeto})")
            self.ui.statusbar.setStyleSheet("color: green;")

    def marcar_saram_com_bandeira(self, linha_alvo):
        """
        Insere o ícone ⚑ na coluna SARAM da linha especificada.
        """
        # 1. Descobre qual é o índice da coluna "SARAM"
        # Isso é importante caso você mude a ordem das colunas no futuro
        coluna_saram = -1
        for col in range(self.ui.tableWidget.columnCount()):
            item_header = self.ui.tableWidget.horizontalHeaderItem(col)
            if item_header and item_header.text() == "SARAM":
                coluna_saram = col
                break

        # Se não achou a coluna SARAM, para por aqui
        if coluna_saram == -1:
            return

        # 2. Pega o item (célula) específico naquela linha e coluna
        item = self.ui.tableWidget.item(linha_alvo, coluna_saram)

        if item:
            texto_atual = item.text()

            # 3. Verifica se já tem a bandeira para não colocar duas vezes
            if "⚑" not in texto_atual:
                novo_texto = f"⚑ {texto_atual}"
                item.setText(novo_texto)

                # Opcional: Mudar a cor do texto para Vermelho para destacar mais
                item.setForeground(QtGui.QColor("red"))

                print(
                    f"Bandeira adicionada na linha {linha_alvo}, SARAM {texto_atual}")
            else:
                print("Este militar já está marcado.")

    def Abrir_Dialogo_Carregar_Dados(self):
        resultado = QFileDialog.getOpenFileName(
            self, "Qual arquivo gostaria de carregar?", caminho_atual, 'Excel files (*.xlsx)')
        # obtém o endereço do arquivo do resultado
        endereco_do_arquivo = resultado[0]
        if endereco_do_arquivo:  # verifica se o endereço do arquivo não é vazio
            self.Carregar_Dados_dos_militares()  # chama a função para carregar os dados

    ##############################################################################
    ##############################################################################
    ##############################################################################
    #### FUNÇÃO PRINCIPAL DE CARREGAMENTO DE DADOS DOS MILITARES #################
    ##############################################################################

    # ----------------------------------------------------------------------
    # FUNÇÕES DE BANCO DE DADOS (SQLite) - Adicione na classe UI
    # ----------------------------------------------------------------------

    def salvar_tudo_no_banco(self):
        """Salva os dados atuais no arquivo SQLite."""
        global df_plamov_compilado
        global df_TP
        global df_TP_BMA

        try:
            conn = sqlite3.connect("tesserato_dados.db")

            if 'df_plamov_compilado' in globals() and not df_plamov_compilado.empty:
                # Converte para string para evitar erros de tipo no SQLite
                df_plamov_compilado.astype(str).to_sql(
                    "plamov", conn, if_exists="replace", index=False)

            if 'df_TP' in globals() and not df_TP.empty:
                df_TP.astype(str).to_sql("tp_geral", conn,
                                         if_exists="replace", index=False)

            if 'df_TP_BMA' in globals() and not df_TP_BMA.empty:
                df_TP_BMA.astype(str).to_sql(
                    "tp_bma", conn, if_exists="replace", index=False)

            conn.close()
            print("Dados salvos no Banco de Dados com sucesso!")
        except Exception as e:
            print(f"Erro ao salvar no banco: {e}")

    def carregar_tudo_do_banco(self):
        """Tenta carregar os dados do SQLite na inicialização."""
        global df_plamov_compilado
        global df_TP
        global df_TP_BMA
        global df_OMs
        global status_painel

        if not os.path.exists("tesserato_dados.db"):
            print("Nenhum banco de dados encontrado. Aguardando carga manual.")
            return False

        try:
            conn = sqlite3.connect("tesserato_dados.db")

            # Carrega PLAMOV
            try:
                df_plamov_compilado = pd.read_sql("SELECT * FROM plamov", conn)
                df_plamov_compilado = df_plamov_compilado.astype(
                    object).fillna("")
                # Garante que a coluna de ordem existe
                if 'ordem original' not in df_plamov_compilado.columns:
                    df_plamov_compilado['ordem original'] = df_plamov_compilado.index
            except:
                pass

            # Carrega TP Geral
            try:
                df_TP = pd.read_sql("SELECT * FROM tp_geral", conn)
            except:
                pass

            # Carrega TP BMA
            try:
                df_TP_BMA = pd.read_sql("SELECT * FROM tp_bma", conn)
                # Converte números de volta (banco traz como texto as vezes)
                for col in ['TLP Ano Corrente', 'Existentes']:
                    if col in df_TP_BMA.columns:
                        df_TP_BMA[col] = pd.to_numeric(
                            df_TP_BMA[col], errors='coerce').fillna(0)
            except:
                pass

            conn.close()

            # Se carregou algo, monta a tela
            if 'df_plamov_compilado' in globals() and not df_plamov_compilado.empty:
                print("Dados recuperados do Banco de Dados!")

                # 1. Gera df_OMs baseado no que carregou do banco
                df_OMs = pegar_OMs_do_COMPREP()

                # 2. Configura a Tabela Visual
                self.configurar_tabela_visual_pelo_banco()

                status_painel = "carregado"
                return True

        except Exception as e:
            print(f"Erro ao ler do banco: {e}")
            return False

    def configurar_tabela_visual_pelo_banco(self):
        """Reconstroi a visualização da tabela sem precisar do Excel."""
        global df_plamov_compilado

        # Definição das colunas
        COLUNAS_DESEJADAS = [
            "LOC ATUAL", "OM ATUAL", "SARAM", "POSTO", "QUADRO", "ESP", "PROJETO",
            "APRESENTAÇÃO NA LOC", "DATA DE PRAÇA", "NR PT",  # <--- ADICIONE AQUI TAMBÉM
            "LOC 1", "LOC 2", "LOC 3", "CÔNJUGE DA FAB?", "DADOS CÔNJUGE", "PLAMOV"
        ]

        colunas_existentes = [
            col for col in COLUNAS_DESEJADAS if col in df_plamov_compilado.columns]

        try:
            mapa_indices = {nome: df_plamov_compilado.columns.get_loc(
                nome) for nome in colunas_existentes}
            indices_a_exibir = [mapa_indices[nome]
                                for nome in colunas_existentes]
        except:
            return

        self.ui.tableWidget.setColumnCount(len(colunas_existentes))
        self.ui.tableWidget.setRowCount(df_plamov_compilado.shape[0])
        self.ui.tableWidget.setHorizontalHeaderLabels(colunas_existentes)

        self.aplicar_ordenacao_militares()

        self.formatar_datas_brasileiras()

        # Configurações visuais (Selection Behavior)
        self.ui.tableWidget.setSelectionBehavior(
            QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.ui.tableWidget.setSelectionMode(
            QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)

        coluna_tableWidget_esquerda = 0

        for i in range(df_plamov_compilado.shape[0]):
            for df_index in indices_a_exibir:
                valor_celula = str(df_plamov_compilado.iloc[i, df_index])
                item = QtWidgets.QTableWidgetItem(valor_celula)
                self.ui.tableWidget.setItem(
                    i, coluna_tableWidget_esquerda, item)

                if i % 2:
                    self.ui.tableWidget.item(i, coluna_tableWidget_esquerda).setBackground(
                        QtGui.QColor(100, 139, 245))

                coluna_tableWidget_esquerda += 1
            coluna_tableWidget_esquerda = 0

        self.destacar_saram_prioritarios()

    ##############################################################################
    ##### FIM DA FUNÇÃO PRINCIPAL DE CARREGAMENTO DE DADOS DOS MILITARES #########
    ##############################################################################
    ##############################################################################

    def Carregar_Dados_dos_militares(self):
        global endereco_do_arquivo
        global df_OMs
        global df_plamov_compilado
        global status_painel

        # 1. Tenta pegar o endereço do arquivo
        try:
            endereco_do_arquivo = QFileDialog.getOpenFileName(
                self, "Qual arquivo gostaria de carregar?", caminho_atual, 'Excel files (*.xlsx)')[0]
        except:
            endereco_do_arquivo = ""

        # 2. Só executa se tiver arquivo
        if endereco_do_arquivo:

            # --- CORREÇÃO DE PERFORMANCE AQUI ---
            # Bloqueia os sinais para que o 'celula_alterada' NÃO seja chamado durante o loop
            self.ui.tableWidget.blockSignals(True)
            # Define status temporário para garantir que não salve nada agora
            status_anterior = status_painel
            status_painel = "carregando"
            # ------------------------------------

            try:
                # Carrega a aba PLAMOV COMPILADO
                df_plamov_compilado = pd.read_excel(
                    endereco_do_arquivo, sheet_name="PLAMOV COMPILADO")
                df_plamov_compilado = df_plamov_compilado.astype(
                    object).fillna("")
                # Compatibilidade: planilha usa SUBDIVISAO, código usa PROJETO
                if 'SUBDIVISAO' in df_plamov_compilado.columns and 'PROJETO' not in df_plamov_compilado.columns:
                    df_plamov_compilado.rename(
                        columns={'SUBDIVISAO': 'PROJETO'}, inplace=True)
                df_plamov_compilado['ordem original'] = df_plamov_compilado.index

                # Configuração das Colunas
                COLUNAS_DESEJADAS = [
                    "LOC ATUAL", "OM ATUAL", "SARAM", "POSTO", "QUADRO", "ESP", "PROJETO",
                    "APRESENTAÇÃO NA LOC", "DATA DE PRAÇA", "NR PT",
                    "LOC 1", "LOC 2", "LOC 3", "CÔNJUGE DA FAB?", "DADOS CÔNJUGE", "PLAMOV"
                ]

                colunas_existentes = [
                    col for col in COLUNAS_DESEJADAS if col in df_plamov_compilado.columns]

                try:
                    mapa_indices = {nome: df_plamov_compilado.columns.get_loc(
                        nome) for nome in colunas_existentes}
                    indices_a_exibir = [mapa_indices[nome]
                                        for nome in colunas_existentes]
                except KeyError as e:
                    print(f"ERRO CRÍTICO: Coluna não encontrada: {e}")
                    return

                self.ui.tableWidget.setColumnCount(len(colunas_existentes))
                self.ui.tableWidget.setRowCount(df_plamov_compilado.shape[0])
                self.ui.tableWidget.setHorizontalHeaderLabels(
                    colunas_existentes)

                self.aplicar_ordenacao_militares()

                self.formatar_datas_brasileiras()

                # --- Preenchimento da Tabela Visual ---
                coluna_tableWidget_esquerda = 0
                for i in range(df_plamov_compilado.shape[0]):
                    for df_index in indices_a_exibir:
                        valor_celula = str(
                            df_plamov_compilado.iloc[i, df_index])
                        item = QtWidgets.QTableWidgetItem(valor_celula)
                        self.ui.tableWidget.setItem(
                            i, coluna_tableWidget_esquerda, item)

                        if i % 2:
                            self.ui.tableWidget.item(i, coluna_tableWidget_esquerda).setBackground(
                                QtGui.QColor(100, 139, 245))

                        coluna_tableWidget_esquerda += 1
                    coluna_tableWidget_esquerda = 0

                self.destacar_saram_prioritarios()

                # Bug E fix: carregar TP primeiro, depois montar df_OMs
                self.carregar_Relat_rio_TP()
                df_OMs = pegar_OMs_do_COMPREP()

                # Salva no banco UMA ÚNICA VEZ ao final de tudo
                self.salvar_tudo_no_banco()

            except Exception as e:
                print(f"Erro ao carregar planilha: {e}")

            finally:
                # --- RESTAURA O ESTADO NORMAL ---
                status_painel = "carregado"
                self.ui.tableWidget.blockSignals(False)
                # --------------------------------

        else:
            print("Nenhum arquivo selecionado.")

    def carregar_Relat_rio_TP(self):
        global df_TP
        global df_TP_BMA
        global endereco_do_arquivo

        # Guard: não executa se nenhum arquivo foi carregado nesta sessão
        if 'endereco_do_arquivo' not in globals() or not endereco_do_arquivo:
            print("Nenhum arquivo carregado. Carregue os dados dos militares primeiro.")
            return

        # Carrega a TP Padrão
        try:
            df_TP = pd.read_excel(endereco_do_arquivo,
                                  sheet_name="RELATÓRIO TP")
        except:
            pass

        # --- CARREGAMENTO DA TP BMA ---
        try:
            df_TP_BMA = pd.read_excel(
                endereco_do_arquivo, sheet_name="RELATÓRIO TP BMA")
            df_TP_BMA.fillna(0, inplace=True)

            # Remove espaços em branco antes e depois dos nomes das colunas
            df_TP_BMA.columns = df_TP_BMA.columns.str.strip()

            # Mapeia nomes incorretos para o nome correto "Projeto"
            mapa_correcao = {
                "projeto": "Projeto",
                "Projeto": "Projeto",
                "PROJETO": "Projeto",
                "Projeto ": "Projeto"
            }
            df_TP_BMA.rename(columns=mapa_correcao, inplace=True)

        except Exception as e:
            print(
                f"Sem aba 'RELATÓRIO TP BMA' ou erro ao carregar: {e}. Branch BMA usará RELATÓRIO TP.")
            df_TP_BMA = pd.DataFrame()

    def linha_ativa_dados_militares(self):
        global linha_selecionada_painel_esquerda
        linha_selecionada_painel_esquerda = self.ui.tableWidget.currentRow()

        return linha_selecionada_painel_esquerda

    def coluna_ativa_dados_militares(self):
        # nem sempre a coluna ativa no df_plamov_compilado vai ser a coluna ativa no tablewidget
        # depois que a célula da coluna "PLAMOV" checa se o militar foi movimentado e ajusta a quantidade de vagas na TP, dimunuindo a quantidade da "OM de destino" e aumentando da "OM ATUAL"
        # essa função vai precisar saber as dimensões do militar selecionado que foi obtida quando o usuário clicou na linha militare a linha ativa também.
        # parto do princípio que não existe mais de uma linha com a mesma combinação de OM,posto,quadro e especilidade
        # regras para ativar a função que atualiza as vagas na tabela TP
        # 1- checar se o militar está sendo transferido realmente, pq pode acontecer de colocar a unidade de destino igual à unidade atual
        # 2- checar se a coluna alterada é a coluna "PLAMOV"
        # 3- checar se a célula foi feita pelo usuário, caso contrário a função seria ativada quando o relatório fosse carregado.

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
            # nessa fase preciso saber qual a linha ativa que o usuário editou
            # nessa etapa preciso saber a OM_destino e OM_origem, isso vai ser buscado no df_plamov_compilado
            OM_atual = df_plamov_compilado.loc[linha_ativa, "OM ATUAL"]

            # Obtenha o novo valor da célula editada

            OM_Destino = self.ui.tableWidget.item(linha_alterada, 11).text()

            global posto
            posto = pegar_posto(linha_ativa)

            global quadro
            quadro = df_plamov_compilado["QUADRO"][linha_ativa]

            global especialidade
            especialidade = df_plamov_compilado["ESP"][linha_ativa]

            # nessa fase preciso achar duas linhas no df_TP
            # 1-linha da combinação entre a OM_destino e as três dimensões - dataframe.query("nome da coluna == 'valor da condição'").index[0])

            # Está funcionando mas tem que colocar um tratamento para quando não achar uma combinação.
            # a melhor opção é criar uma coluna com as pessoas "chegando" e "saindo" de cada OM
            # uma outra coluna com as "vagas dinâmicas" que refletem o existente, vagas na tp, chegando e saindo.
            # Se colocar o destino de alguém para alguma OM que não tenha TP prevista, vai ser criada linha com a combinação e uma unidade somada à coluna "chegando", dessa forma é possivel manter o controle de quantas pessoas estão chegando em cada unidade.
            # TODO idéia de gráfico, colocar um gráfico para cada OM uma quantidade de pessoas saíndo e chegando, talvez uma indicação de estão perdendo gente, ou seja, com uma quantidade maior de pessoas saindo do que chegando, ou o contrário.
            # O que fazer nesse caso, criar uma e deixar uma flag dizendo que não tem TP
            # Ver como está o tratamento no painel superior

            # se a OM inserida não estiver na relação, mostrar um popup com um warning
            # Se for do COMPREP  mas não tiver TP, mostrar um popup
            if posto == "TN":
                linha_OM_destino = df_TP.query(
                    f"Unidade == '{OM_Destino}' & (Posto == 'CP/TN' | Posto == 'TN') & Quadro == '{quadro}' & Especialidade == '{especialidade}'")
            elif posto == "CP":
                linha_OM_destino = df_TP.query(
                    f"Unidade == '{OM_Destino}' & (Posto == 'CP/TN' | Posto == 'CP') & Quadro == '{quadro}' & Especialidade == '{especialidade}'")
            else:
                linha_OM_destino = df_TP.query(
                    f"Unidade == '{OM_Destino}' & Posto == '{posto}' & Quadro == '{quadro}' & Especialidade == '{especialidade}'")

            if linha_OM_destino.empty:
                # DESCRIÇÃO: ESSE CASO CRIA UMA LINHA COM A COMBINAÇÃO DAS TRÊS DIMENSÕES DO MILITAR CASO ELE SEJA ALOCADO EM UMA OM QUE NÃO EXISTA A PREVISÃO PARA AS SUAS 3 DIMENSÕES NA TABELA DE TP
                # AQUI eu devo criar uma nova linha com a combinação da query acima, inserir no DF_TP e colocar os valores de vagas nas respectivas colunas.
                nova_linha = pd.DataFrame({'Unidade': [OM_Destino], 'Posto': [posto], 'Quadro': [quadro], 'Especialidade': [
                                          especialidade], 'TLP Ano Corrente': [0], 'Existentes': [1], 'Vagas': [-1]})
                df_TP = pd.concat([df_TP, nova_linha],
                                  axis=0, ignore_index=True)
                df_TP.fillna("", inplace=True)

            # UNIDADE QUE O MILITAR ESTÁ CHEGANDO
            if posto == "CP":
                # TIRA UMA VAGA NA UNIDADE QUE O MILITAR ESTÁ CHEGANDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_Destino}") & ((df_TP['Posto'] == "CP/TN") | (df_TP['Posto'] == "CP")) & (
                    df_TP['Quadro'] == f"{quadro}") & (df_TP['Especialidade'] == f"{especialidade}"), 'Vagas'] += -1
                # COLOCA UM EXISTENTE NA UNIDADE QUE O MILITAR ESTÁ CHEGANDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_Destino}") & ((df_TP['Posto'] == "CP/TN") | (df_TP['Posto'] == "CP")) & (
                    df_TP['Quadro'] == f"{quadro}") & (df_TP['Especialidade'] == f"{especialidade}"), 'Existentes'] += 1
            elif posto == "TN":
                # TIRA UMA VAGA NA UNIDADE QUE O MILITAR ESTÁ CHEGANDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_Destino}") & ((df_TP['Posto'] == "CP/TN") | (df_TP['Posto'] == "TN")) & (
                    df_TP['Quadro'] == f"{quadro}") & (df_TP['Especialidade'] == f"{especialidade}"), 'Vagas'] += -1
                # COLOCA UM EXISTENTE NA UNIDADE QUE O MILITAR ESTÁ CHEGANDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_Destino}") & ((df_TP['Posto'] == "CP/TN") | (df_TP['Posto'] == "TN")) & (
                    df_TP['Quadro'] == f"{quadro}") & (df_TP['Especialidade'] == f"{especialidade}"), 'Existentes'] += 1
            else:
                # TIRA UMA VAGA NA UNIDADE QUE O MILITAR ESTÁ CHEGANDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_Destino}") & (df_TP['Posto'] == f"{posto}") & (
                    df_TP['Quadro'] == f"{quadro}") & (df_TP['Especialidade'] == f"{especialidade}"), 'Vagas'] += -1
                # COLOCA UM EXISTENTE NA UNIDADE QUE O MILITAR ESTÁ CHEGANDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_Destino}") & (df_TP['Posto'] == f"{posto}") & (
                    df_TP['Quadro'] == f"{quadro}") & (df_TP['Especialidade'] == f"{especialidade}"), 'Existentes'] += 1

            # UNIDADE QUE O MILITAR ESTÁ SAINDO
            if posto == "CP":
                # COLOCA UMA VAGA NA UNIDADE QUE O MILITAR ESTÁ SAINDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_atual}") & ((df_TP['Posto'] == "CP/TN") | (df_TP['Posto'] == "CP")) & (
                    df_TP['Quadro'] == f"{quadro}") & (df_TP['Especialidade'] == f"{especialidade}"), 'Vagas'] += -1
                # TIRA UM EXISTENTE NA UNIDADE QUE O MILITAR ESTÁ SAINDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_atual}") & ((df_TP['Posto'] == "CP/TN") | (df_TP['Posto'] == "CP")) & (
                    df_TP['Quadro'] == f"{quadro}") & (df_TP['Especialidade'] == f"{especialidade}"), 'Existentes'] += 1
            elif posto == "TN":
                # COLOCA UMA VAGA NA UNIDADE QUE O MILITAR ESTÁ SAINDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_atual}") & ((df_TP['Posto'] == "CP/TN") | (df_TP['Posto'] == "TN")) & (
                    df_TP['Quadro'] == f"{quadro}") & (df_TP['Especialidade'] == f"{especialidade}"), 'Vagas'] += -1
                # TIRA UM EXISTENTE NA UNIDADE QUE O MILITAR ESTÁ SAINDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_atual}") & ((df_TP['Posto'] == "CP/TN") | (df_TP['Posto'] == "TN")) & (
                    df_TP['Quadro'] == f"{quadro}") & (df_TP['Especialidade'] == f"{especialidade}"), 'Existentes'] += 1
            else:
                # COLOCA UMA VAGA NA UNIDADE QUE O MILITAR ESTÁ SAINDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_atual}") & (df_TP['Posto'] == f"{posto}") & (
                    df_TP['Quadro'] == f"{quadro}") & (df_TP['Especialidade'] == f"{especialidade}"), 'Vagas'] += -1
                # TIRA UM EXISTENTE NA UNIDADE QUE O MILITAR ESTÁ SAINDO
                df_TP.loc[(df_TP['Unidade'] == f"{OM_atual}") & (df_TP['Posto'] == f"{posto}") & (
                    df_TP['Quadro'] == f"{quadro}") & (df_TP['Especialidade'] == f"{especialidade}"), 'Existentes'] += 1

    def escolher_OM_no_painel_direito(self):
        coluna_ativa_painel_direita = self.ui.tableWidget_2.currentColumn()
        linha_ativa_painel_direita = self.ui.tableWidget_2.currentRow()
        nome_coluna_ativa_painel_direita = df_OMs.columns[coluna_ativa_painel_direita]
        if (nome_coluna_ativa_painel_direita == "OMs"):

            OM_selecionada_painel_direita = QtWidgets.QTableWidgetItem(
                self.ui.tableWidget_2.item(linha_ativa_painel_direita, coluna_ativa_painel_direita))
            if (linha_selecionada_painel_esquerda % 2):
                # colorir de azul
                OM_selecionada_painel_direita.setBackground(
                    QtGui.QColor(100, 139, 245))
            else:
                # colorir de branco
                OM_selecionada_painel_direita.setBackground(
                    QtGui.QColor(255, 255, 255))

            self.ui.tableWidget.setItem(
                linha_selecionada_painel_esquerda, 15, OM_selecionada_painel_direita)
            df_plamov_compilado.loc[linha_selecionada_painel_esquerda, "PLAMOV"] = self.ui.tableWidget_2.item(
                linha_ativa_painel_direita, coluna_ativa_painel_direita).text()
            linha_ativa_painel_esquerda = self.linha_ativa_dados_militares()
            coluna_ativa_painel_esquerda = self.coluna_ativa_dados_militares()

            self.ui.tableWidget.setCurrentCell(
                linha_ativa_painel_esquerda + 1, coluna_ativa_painel_esquerda)
            self.salvar_tudo_no_banco()
            self.atualizar_Painel_Direita()
        #     self.ui.tableWidget_2.setItem(linha_selecionada_painel_esquerda, coluna_ativa_painel_esquerda)
        #     item = QtWidgets.QTableWidgetItem(str(df_OMs.iloc[k,i]))
        #     self.ui.tableWidget_2.setItem(k,i, item)


app = QApplication(sys.argv)
UIWindow = SplashScreen()
app.exec()
