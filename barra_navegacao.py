# -*- coding: utf-8 -*-
"""
barra_navegacao.py — Barra de Navegação com Botões em Formato de Chevron

Cria uma barra horizontal no topo da aplicação com botões em formato de seta
(chevron) para navegação entre as páginas do TESSERATO.
O botão da página atual é diferenciado visualmente.

RESPONSIVO: os botões expandem/encolhem para ocupar 100% da largura
disponível, independente do tamanho da janela.

Uso:
    from barra_navegacao import instalar_barra_navegacao
    # Chamar dentro do __init__ da classe UI, APÓS todas as páginas terem
    # sido adicionadas ao stackedWidget:
    instalar_barra_navegacao(self)
"""

from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtCore import Qt


# ============================================================
# Cores e Estilos
# ============================================================

COR_NORMAL = "#3A6FB0"          # Azul médio (botão inativo)
COR_NORMAL_HOVER = "#2D5A91"   # Azul mais escuro no hover
COR_ATIVO = "#1B3A5C"          # Azul escuro (botão da página atual)
COR_DESABILITADO = "#8FABC4"   # Azul acinzentado para páginas futuras
COR_TEXTO = "#FFFFFF"          # Branco
COR_FUNDO_BARRA = "#E8ECF0"   # Cinza claro de fundo da barra


# ============================================================
# Widget customizado: Botão Chevron (Responsivo)
# ============================================================

class BotaoChevron(QtWidgets.QWidget):
    """
    Widget customizado com formato de chevron (seta apontando para a direita).
    Usa QPainterPath para desenhar o polígono do chevron.
    
    RESPONSIVO: Usa QSizePolicy.Expanding para preencher toda a largura
    disponível proporcionalmente junto com os demais botões.
    """

    clicked = QtCore.pyqtSignal()

    def __init__(self, texto, indice, eh_primeiro=False, parent=None):
        super().__init__(parent)
        self.texto = texto
        self.indice = indice
        self.eh_primeiro = eh_primeiro
        self._ativo = False
        self._hover = False
        self._habilitado = True

        self.setCursor(Qt.CursorShape.PointingHandCursor)

        # CHAVE DA RESPONSIVIDADE: Expanding faz cada botão ocupar
        # uma fração proporcional da largura do layout pai
        self.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Fixed
        )
        self.setFixedHeight(34)
        self.setMinimumWidth(30)  # Mínimo absoluto para não sumir

    def set_ativo(self, ativo: bool):
        self._ativo = ativo
        self.update()

    def setEnabled(self, habilitado: bool):
        self._habilitado = habilitado
        if not habilitado:
            self.setCursor(Qt.CursorShape.ArrowCursor)
        else:
            self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.update()

    def setToolTip(self, texto):
        super().setToolTip(texto)

    def enterEvent(self, event):
        self._hover = True
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._hover = False
        self.update()
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        if self._habilitado and event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)

    def paintEvent(self, event):
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)

        w = self.width()
        h = self.height()
        # Tamanho da ponta proporcional à altura (mantém a forma em qualquer tamanho)
        ponta = min(h * 0.40, w * 0.12)

        # ── Cor de fundo ──
        if not self._habilitado:
            cor = QtGui.QColor(COR_DESABILITADO)
        elif self._ativo:
            cor = QtGui.QColor(COR_ATIVO)
        elif self._hover:
            cor = QtGui.QColor(COR_NORMAL_HOVER)
        else:
            cor = QtGui.QColor(COR_NORMAL)

        # ── Constrói o path do chevron ──
        path = QtGui.QPainterPath()

        if self.eh_primeiro:
            # Primeiro botão: lado esquerdo reto, lado direito com ponta
            path.moveTo(0, 0)
            path.lineTo(w - ponta, 0)
            path.lineTo(w, h / 2)
            path.lineTo(w - ponta, h)
            path.lineTo(0, h)
            path.closeSubpath()
        else:
            # Demais botões: lado esquerdo com recuo (V), lado direito com ponta
            path.moveTo(0, 0)
            path.lineTo(ponta, h / 2)
            path.lineTo(0, h)
            path.lineTo(w - ponta, h)
            path.lineTo(w, h / 2)
            path.lineTo(w - ponta, 0)
            path.closeSubpath()

        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(cor)
        painter.drawPath(path)

        # ── Texto ──
        cor_texto = QtGui.QColor(COR_TEXTO)
        if not self._habilitado:
            cor_texto.setAlpha(180)
        painter.setPen(cor_texto)

        # Fonte adaptativa: diminui em janelas estreitas
        font = painter.font()
        largura_util = w - ponta * 1.5 if not self.eh_primeiro else w - ponta * 0.5
        # Calcula tamanho de fonte que caiba no espaço disponível
        tamanho_fonte = self._calcular_tamanho_fonte(largura_util, h)
        font.setPointSizeF(tamanho_fonte)
        font.setBold(True)
        painter.setFont(font)

        # Área de texto deslocada para compensar as pontas do chevron
        if self.eh_primeiro:
            margem_esq = 4
            margem_dir = ponta * 0.6
        else:
            margem_esq = ponta * 0.6
            margem_dir = ponta * 0.6

        rect_texto = QtCore.QRectF(
            margem_esq, 1,
            w - margem_esq - margem_dir, h - 2
        )

        painter.drawText(
            rect_texto,
            Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap,
            self.texto
        )
        painter.end()

    def _calcular_tamanho_fonte(self, largura_disponivel, altura_disponivel):
        """Calcula o melhor tamanho de fonte para caber no espaço."""
        # Conta a linha mais longa do texto (para textos com \n)
        linhas = self.texto.split('\n')
        n_linhas = len(linhas)
        linha_mais_longa = max(linhas, key=len)

        # Tamanho máximo por altura (caber N linhas)
        max_por_altura = (altura_disponivel / (n_linhas + 0.5)) * 0.7

        # Tamanho máximo por largura (caber texto na linha)
        if len(linha_mais_longa) > 0:
            max_por_largura = (largura_disponivel / len(linha_mais_longa)) * 1.3
        else:
            max_por_largura = 10

        tamanho = min(max_por_altura, max_por_largura)

        # Clamp entre 5pt e 9pt
        return max(5.0, min(9.0, tamanho))

    def sizeHint(self):
        return QtCore.QSize(110, 34)

    def minimumSizeHint(self):
        return QtCore.QSize(30, 34)


# ============================================================
# Função principal de instalação
# ============================================================

def instalar_barra_navegacao(janela):
    """
    Instala a barra de navegação com chevrons no topo da janela principal,
    entre o frame do cabeçalho (COMPREP / TESSERATO / IAOp) e o stackedWidget.

    Parâmetros:
        janela: instância da classe UI (QMainWindow)

    A barra é atualizada automaticamente quando o stackedWidget muda de página.
    """

    # ──────────────────────────────────────────────────
    # Definição das páginas visíveis no menu de navegação
    # Cada tupla: (nome_exibido, funcao_de_navegacao)
    # ──────────────────────────────────────────────────
    paginas = [
        ("PAINEL\nPRINCIPAL",            lambda: janela.Pag_Militares()),
        ("MILITARES\nPRIORITÁRIOS",      lambda: janela.Pag_Prioritarios()),
        ("QUEREM LOC\nDIFÍCEIS",         lambda: janela.Pag_QuerLocDificeis()),
        ("VOLUNTÁRIOS\nC e B",           lambda: janela.Pag_VoluntariosCB()),
        ("MILITARES\nVOLUNTÁRIOS\nA-A",  None),  # Página futura
        ("ESCOLA DE\nFORMAÇÃO",          None),  # Página futura
        ("NÃO\nVOLUNTÁRIOS\nLOC TX > 1", None),  # Página futura
        ("NÃO\nVOLUNTÁRIOS\nT LOC >=8",  None),  # Página futura
        ("TRANSFERIDOS",                  None),  # Página futura
    ]

    # Mapeamento: índice do stackedWidget → índice do botão na barra
    mapa_pagina_botao = {}

    # Page index 0 = Painel Principal
    mapa_pagina_botao[0] = 0

    # Prioritários
    if hasattr(janela, 'page_prioritarios'):
        idx = janela.ui.stackedWidget.indexOf(janela.page_prioritarios)
        if idx >= 0:
            mapa_pagina_botao[idx] = 1

    # Querem Loc. Difíceis
    if hasattr(janela, 'page_quer_loc_dificeis'):
        idx = janela.ui.stackedWidget.indexOf(janela.page_quer_loc_dificeis)
        if idx >= 0:
            mapa_pagina_botao[idx] = 2

    # Voluntários C e B
    if hasattr(janela, 'page_vol_cb'):
        idx = janela.ui.stackedWidget.indexOf(janela.page_vol_cb)
        if idx >= 0:
            mapa_pagina_botao[idx] = 3

    # ──────────────────────────────────────────────────
    # Cria o widget da barra (SEM scroll area — layout direto)
    # ──────────────────────────────────────────────────
    widget_barra = QtWidgets.QWidget()
    widget_barra.setFixedHeight(40)
    widget_barra.setStyleSheet(f"background-color: {COR_FUNDO_BARRA};")

    # Layout horizontal: cada botão é Expanding, então todos dividem
    # a largura total igualmente e se ajustam ao redimensionar
    layout_barra = QtWidgets.QHBoxLayout(widget_barra)
    layout_barra.setContentsMargins(2, 3, 2, 3)
    layout_barra.setSpacing(-3)  # Overlap leve para continuidade visual

    botoes_nav = []

    for i, (nome, funcao) in enumerate(paginas):
        btn = BotaoChevron(nome, i, eh_primeiro=(i == 0))

        if funcao is not None:
            fn = funcao
            btn.clicked.connect(fn)
        else:
            btn.setEnabled(False)
            btn.setToolTip("Página em desenvolvimento")

        layout_barra.addWidget(btn)
        botoes_nav.append(btn)

    # SEM addStretch() — os botões Expanding preenchem tudo sozinhos

    # Armazena referências na janela
    janela._botoes_nav = botoes_nav
    janela._mapa_pagina_botao = mapa_pagina_botao
    janela._widget_barra_nav = widget_barra

    # ──────────────────────────────────────────────────
    # Insere a barra entre o frame do cabeçalho e o stackedWidget
    # ──────────────────────────────────────────────────
    # O layout principal é verticalLayout_2:
    #   index 0 = frame (cabeçalho COMPREP / TESSERATO)
    #   index 1 = stackedWidget
    # Inserimos na posição 1, empurrando o stacked para 2
    janela.ui.verticalLayout_2.insertWidget(1, widget_barra)

    # ──────────────────────────────────────────────────
    # Função para atualizar destaque do botão ativo
    # ──────────────────────────────────────────────────
    def _atualizar_destaque(indice_pagina):
        indice_botao = mapa_pagina_botao.get(indice_pagina, -1)
        for btn in botoes_nav:
            btn.set_ativo(btn.indice == indice_botao)

    # Conecta ao sinal de mudança de página
    janela.ui.stackedWidget.currentChanged.connect(_atualizar_destaque)

    # Atualiza o destaque para a página atual
    _atualizar_destaque(janela.ui.stackedWidget.currentIndex())
