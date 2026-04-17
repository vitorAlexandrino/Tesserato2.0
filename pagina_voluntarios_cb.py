# -*- coding: utf-8 -*-
"""
pagina_voluntarios_cb.py (v2 — com QTabWidget MCDA + ICA 30-4)

Instala a página "Voluntários C e B" no TESSERATO 2.0 sem modificar nem a UI
gerada pelo Qt Designer nem o corpo do 1-sideMenuMain.py além das 2 linhas
já existentes (import + install_vol_cb(self)).

Layout:
  Título + Resumo
  Descrição
  QSplitter horizontal
    ├─ Esquerda (~70%): QTabWidget
    │    ├─ Aba "MCDA"      → tabela ranqueada de alternativas (militar × OM)
    │    └─ Aba "ICA 30-4"  → lista de militares ordenada pela ICA
    └─ Direita (~30%): painel dinâmico
         - Aba MCDA: "Dados do Militar"
         - Aba ICA:  "OMs Disponíveis" (com destaque LOC 1/2/3)
         Botões [Transferir] [Manter na Origem] compartilhados

A aba ICA replica o fluxo do app principal (duplo clique na OM confirma
transferência), mas em widgets isolados e com paleta harmônica para não
herdar as cores fortes/hard-coded do app principal.
"""

import sys
import datetime
import types

import pandas as pd
import numpy as np

from PyQt6 import QtCore, QtGui, QtWidgets

from mcda_voluntarios_cb import calcular_mcda_voluntarios_cb


# ============================================================
# Helpers de acesso aos globals do módulo principal
# ============================================================

def _mod_principal():
    """Módulo principal (`__main__` quando 1-sideMenuMain.py é rodado direto)."""
    return sys.modules.get("__main__")


def _get_global(nome: str):
    mod = _mod_principal()
    if mod is None:
        return None
    return getattr(mod, nome, None)


def _set_global(nome: str, valor):
    mod = _mod_principal()
    if mod is not None:
        setattr(mod, nome, valor)


# ============================================================
# Paleta harmônica (funciona em modo claro e escuro)
# ============================================================

COR_LOC1 = QtGui.QColor(46, 160, 67)    # verde médio
COR_LOC2 = QtGui.QColor(201, 154, 17)   # âmbar médio
COR_LOC3 = QtGui.QColor(163, 97, 191)   # roxo médio


# ============================================================
# Formatação
# ============================================================

def _fmt_num(valor, casas=3):
    try:
        if pd.isna(valor):
            return ""
        return f"{float(valor):.{casas}f}"
    except (ValueError, TypeError):
        return str(valor) if valor is not None else ""


def _fmt_int(valor):
    try:
        if pd.isna(valor):
            return ""
        return str(int(float(valor)))
    except (ValueError, TypeError):
        return str(valor) if valor is not None else ""


def _fmt_str(valor):
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    return str(valor)


# ============================================================
# Auditoria
# ============================================================

COLUNAS_AUDITORIA = [
    "TIMESTAMP", "SARAM", "NOME", "OM_ORIGEM", "OM_DESTINO_SUGERIDO",
    "OM_DESTINO_ESCOLHIDO", "TIPO_INTERVENCAO", "VALOR_MCDA", "MOTIVO", "GESTOR",
]


def _get_df_auditoria():
    df_aud = _get_global("df_auditoria")
    if df_aud is None or not isinstance(df_aud, pd.DataFrame):
        df_aud = pd.DataFrame(columns=COLUNAS_AUDITORIA)
        _set_global("df_auditoria", df_aud)
    else:
        for col in COLUNAS_AUDITORIA:
            if col not in df_aud.columns:
                df_aud[col] = ""
    return df_aud


def _registrar_auditoria(evento: dict):
    df_aud = _get_df_auditoria()
    linha = {col: evento.get(col, "") for col in COLUNAS_AUDITORIA}
    linha.setdefault(
        "TIMESTAMP", datetime.datetime.now().isoformat(timespec="seconds"))
    novo = pd.concat([df_aud, pd.DataFrame([linha], columns=COLUNAS_AUDITORIA)],
                     ignore_index=True)
    _set_global("df_auditoria", novo)


# ============================================================
# Mutação do df_TP_BMA após confirmação de transferência
# ============================================================

def _normalizar(v):
    return "" if pd.isna(v) else str(v).strip().upper()


def _atualizar_tp_bma_apos_transferencia(unidade_origem, unidade_destino,
                                         projeto, posto=None, quadro=None):
    df_TP_BMA = _get_global("df_TP_BMA")
    if df_TP_BMA is None or df_TP_BMA.empty:
        return

    def _aplicar(unidade, delta_exist):
        mask = (
            df_TP_BMA["Unidade"].astype(str).map(
                _normalizar) == _normalizar(unidade)
        ) & (
            df_TP_BMA["Projeto"].astype(str).map(
                _normalizar) == _normalizar(projeto)
        )
        if posto is not None and "Posto" in df_TP_BMA.columns:
            mask = mask & (
                df_TP_BMA["Posto"].astype(str).map(
                    _normalizar) == _normalizar(posto)
            )
        if quadro is not None and "Quadro" in df_TP_BMA.columns:
            mask = mask & (
                df_TP_BMA["Quadro"].astype(str).map(
                    _normalizar) == _normalizar(quadro)
            )
        idxs = df_TP_BMA[mask].index.tolist()
        if not idxs:
            return
        idx = idxs[0]
        try:
            tlp = float(df_TP_BMA.at[idx, "TLP Ano Corrente"] or 0)
        except (ValueError, TypeError):
            tlp = 0.0
        try:
            exist = float(df_TP_BMA.at[idx, "Existentes"] or 0)
        except (ValueError, TypeError):
            exist = 0.0
        try:
            vagas = float(df_TP_BMA.at[idx, "Vagas"] or 0)
        except (ValueError, TypeError):
            vagas = 0.0
        novo_exist = exist + delta_exist
        novo_vagas = vagas - delta_exist
        df_TP_BMA.at[idx, "Existentes"] = novo_exist
        df_TP_BMA.at[idx, "Vagas"] = novo_vagas
        if tlp > 0:
            df_TP_BMA.at[idx, "Taxa atual"] = novo_exist / tlp

    _aplicar(unidade_destino, +1)
    _aplicar(unidade_origem, -1)


# ============================================================
# Helpers de posto
# ============================================================

def _normalizar_posto(posto_raw: str) -> str:
    p = _fmt_str(posto_raw).strip()
    if p in ("1S", "2S", "3S", "SO"):
        return "SGT"
    if p in ("1T", "2T"):
        return "TN"
    return p


# ============================================================
# Handlers MCDA
# ============================================================

def _Pag_VoluntariosCB(self):
    indice = self.ui.stackedWidget.indexOf(self.page_vol_cb)
    self.ui.stackedWidget.setCurrentIndex(indice)
    self.carregar_tabela_vol_cb_mcda()
    self.carregar_tabela_vol_cb_ica()
    _atualizar_painel_direito_por_aba(self)


def _carregar_tabela_vol_cb_mcda(self):
    df_plamov_compilado = _get_global("df_plamov_compilado")
    df_TP_BMA = _get_global("df_TP_BMA")

    if df_plamov_compilado is None or df_plamov_compilado.empty:
        self.tableWidget_vol_cb_mcda.setRowCount(0)
        self.lbl_resumo_vol_cb.setText("Nenhum dado carregado.")
        self.df_ranking_vol_cb = pd.DataFrame()
        return

    if not hasattr(self, "df_grupo_f"):
        self.tableWidget_vol_cb_mcda.setRowCount(0)
        self.lbl_resumo_vol_cb.setText(
            "Execute a classificação de blocos primeiro (carregue os dados dos militares).")
        self.df_ranking_vol_cb = pd.DataFrame()
        return

    if self.df_grupo_f is None or self.df_grupo_f.empty:
        self.tableWidget_vol_cb_mcda.setRowCount(0)
        self.lbl_resumo_vol_cb.setText(
            "Nenhum militar se enquadra no Grupo F (Voluntários C e B).")
        self.df_ranking_vol_cb = pd.DataFrame()
        return

    if df_TP_BMA is None or df_TP_BMA.empty:
        self.tableWidget_vol_cb_mcda.setRowCount(0)
        self.lbl_resumo_vol_cb.setText("Relatório TP BMA não carregado.")
        self.df_ranking_vol_cb = pd.DataFrame()
        return

    ranking = calcular_mcda_voluntarios_cb(self.df_grupo_f, df_TP_BMA)
    self.df_ranking_vol_cb = ranking.copy()

    n_militares = ranking["SARAM"].nunique() if not ranking.empty else 0
    n_alt = len(ranking)
    self.lbl_resumo_vol_cb.setText(
        f"{n_militares} militares  |  {n_alt} alternativas (MCDA)")

    _popular_tabela_mcda(self, ranking)


def _popular_tabela_mcda(self, ranking: pd.DataFrame):
    col_labels = [
        "SARAM", "OM Destino", "Localidade",
        "Tx Ocup Destino", "Capacitação", "Tx Ocup Origem",
        "Tloc", "Intenção", "Valor",
    ]
    self.tableWidget_vol_cb_mcda.setColumnCount(len(col_labels))
    self.tableWidget_vol_cb_mcda.setRowCount(len(ranking))
    self.tableWidget_vol_cb_mcda.setHorizontalHeaderLabels(col_labels)

    for i, (_, row) in enumerate(ranking.iterrows()):
        valores = [
            _fmt_str(row.get("SARAM")),
            _fmt_str(row.get("OM_DESTINO")),
            _fmt_str(row.get("LOCALIDADE_DESTINO")),
            _fmt_num(row.get("TX_VAI_BRUTO"), 3),
            _fmt_int(row.get("CAPACITACAO")),
            _fmt_num(row.get("TX_FICA_BRUTO"), 3),
            _fmt_num(row.get("TLOC_BRUTO"), 2),
            _fmt_num(row.get("INTENCAO"), 3),
            _fmt_num(row.get("VALOR"), 5),
        ]
        for j, v in enumerate(valores):
            item = QtWidgets.QTableWidgetItem(v)
            item.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            self.tableWidget_vol_cb_mcda.setItem(i, j, item)
    self.tableWidget_vol_cb_mcda.resizeColumnsToContents()


# ============================================================
# Handlers ICA
# ============================================================

def _carregar_tabela_vol_cb_ica(self):
    df_plamov_compilado = _get_global("df_plamov_compilado")

    if df_plamov_compilado is None or df_plamov_compilado.empty:
        self.tableWidget_vol_cb_ica.setRowCount(0)
        self.df_ica_vol_cb = pd.DataFrame()
        return

    if not hasattr(self, "df_grupo_f") or self.df_grupo_f is None or self.df_grupo_f.empty:
        self.tableWidget_vol_cb_ica.setRowCount(0)
        self.df_ica_vol_cb = pd.DataFrame()
        return

    df = self.df_grupo_f.copy()

    ordem_cols, ordem_asc = [], []
    if "SCORE_PRIORIDADE" in df.columns:
        ordem_cols.append("SCORE_PRIORIDADE")
        ordem_asc.append(False)
    if "MELHOR PRIO" in df.columns:
        ordem_cols.append("MELHOR PRIO")
        ordem_asc.append(True)
    if "TEMPO LOC" in df.columns:
        df["TEMPO LOC"] = pd.to_numeric(
            df["TEMPO LOC"], errors="coerce").fillna(0)
        ordem_cols.append("TEMPO LOC")
        ordem_asc.append(False)
    if "ANTIGUIDADE" in df.columns:
        df["ANTIGUIDADE"] = pd.to_numeric(df["ANTIGUIDADE"], errors="coerce")
        ordem_cols.append("ANTIGUIDADE")
        ordem_asc.append(True)

    if ordem_cols:
        df = df.sort_values(by=ordem_cols, ascending=ordem_asc,
                            kind="mergesort", na_position="last").reset_index(drop=True)

    self.df_ica_vol_cb = df.copy()

    col_labels_map = [
        ("LOC ATUAL", "Loc. Atual"),
        ("OM ATUAL", "OM Atual"),
        ("SARAM", "SARAM"),
        ("POSTO", "Posto"),
        ("QUADRO", "Quadro"),
        ("ESP", "Esp"),
        ("PROJETO", "Projeto"),
        ("TEMPO LOC", "T.Loc"),
        ("ANTIGUIDADE", "Antig."),
        ("LOC 1", "LOC 1"),
        ("LOC 2", "LOC 2"),
        ("LOC 3", "LOC 3"),
        ("PLAMOV", "PLAMOV"),
    ]
    col_keys = [k for (k, _) in col_labels_map if k in df.columns]
    col_labels = [lbl for (k, lbl) in col_labels_map if k in df.columns]

    self.tableWidget_vol_cb_ica.setColumnCount(len(col_keys))
    self.tableWidget_vol_cb_ica.setRowCount(len(df))
    self.tableWidget_vol_cb_ica.setHorizontalHeaderLabels(col_labels)

    for i, (_, row) in enumerate(df.iterrows()):
        for j, col_key in enumerate(col_keys):
            raw = row.get(col_key, "")
            if col_key == "TEMPO LOC":
                texto = _fmt_num(raw, 2)
            elif col_key == "ANTIGUIDADE":
                texto = _fmt_int(raw)
            else:
                texto = _fmt_str(raw)
            item = QtWidgets.QTableWidgetItem(texto)
            item.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            self.tableWidget_vol_cb_ica.setItem(i, j, item)

    self.tableWidget_vol_cb_ica.resizeColumnsToContents()


def _calcular_oms_compativeis(self):
    df_plamov_compilado = _get_global("df_plamov_compilado")
    df_TP_BMA = _get_global("df_TP_BMA")

    if (df_plamov_compilado is None or df_plamov_compilado.empty
            or df_TP_BMA is None or df_TP_BMA.empty):
        return []

    if not hasattr(self, "df_ica_vol_cb") or self.df_ica_vol_cb.empty:
        return []

    linha = self.tableWidget_vol_cb_ica.currentRow()
    if linha is None or linha < 0 or linha >= len(self.df_ica_vol_cb):
        return []

    militar = self.df_ica_vol_cb.iloc[linha]
    posto = _normalizar_posto(militar.get("POSTO", ""))
    quadro = _fmt_str(militar.get("QUADRO")).strip()
    especialidade = _fmt_str(militar.get("ESP")).strip()
    projeto = _fmt_str(militar.get("PROJETO")).strip()
    loc1 = _fmt_str(militar.get("LOC 1")).strip().upper()
    loc2 = _fmt_str(militar.get("LOC 2")).strip().upper()
    loc3 = _fmt_str(militar.get("LOC 3")).strip().upper()

    if especialidade != "BMA" or "Unidade" not in df_TP_BMA.columns:
        return []

    unidades_unicas = (
        df_TP_BMA["Unidade"].astype(str).str.strip().dropna().unique().tolist()
    )

    def _query_posto(p):
        if p == "SGT":
            return "POSTO in ['1S', '2S', '3S', 'SO']"
        if p == "TN":
            return "POSTO in ['1T', '2T']"
        return f"POSTO == '{p}'"

    qp = _query_posto(posto)

    resultados = []
    for om_k in unidades_unicas:
        if not om_k:
            continue
        mask = (
            (df_TP_BMA["Unidade"].astype(str).str.strip() == om_k)
            & (df_TP_BMA["Posto"].astype(str).str.strip() == posto)
            & (df_TP_BMA["Quadro"].astype(str).str.strip() == quadro)
            & (df_TP_BMA["Projeto"].astype(str).str.strip() == projeto)
        )
        vagas_OM = df_TP_BMA[mask]
        if vagas_OM.empty:
            continue

        localidade = ""
        if "Localidade" in vagas_OM.columns:
            localidade = _fmt_str(vagas_OM.iloc[0]["Localidade"]).strip()

        try:
            TP = int(float(vagas_OM.iloc[0]["TLP Ano Corrente"]))
            exist = int(float(vagas_OM.iloc[0]["Existentes"]))
        except (KeyError, ValueError, TypeError):
            TP, exist = 0, 0

        if TP == 0:
            taxa, vagas_num = "Sem TP", ""
        else:
            try:
                chegando = df_plamov_compilado.query(
                    f"PLAMOV == '{om_k}' & {qp} & QUADRO == '{quadro}' & "
                    f"ESP == 'BMA' & `PROJETO` == '{projeto}'"
                ).shape[0]
                saindo = df_plamov_compilado.query(
                    f"`OM ATUAL` == '{om_k}' & {qp} & QUADRO == '{quadro}' & "
                    f"ESP == 'BMA' & `PROJETO` == '{projeto}' & PLAMOV != ''"
                ).shape[0]
            except Exception:
                chegando, saindo = 0, 0
            existentes_fut = exist + chegando - saindo
            vagas_num = int(TP - exist + saindo - chegando)
            taxa = round((float(existentes_fut) / float(TP)) * 100, 2)

        loc_upper = localidade.upper()
        if loc_upper == loc1 and loc1 != "":
            categoria = "loc1"
        elif loc_upper == loc2 and loc2 != "":
            categoria = "loc2"
        elif loc_upper == loc3 and loc3 != "":
            categoria = "loc3"
        else:
            categoria = "demais"

        resultados.append({
            "om": om_k,
            "localidade": localidade,
            "taxa": taxa,
            "vagas": vagas_num,
            "categoria": categoria,
        })

    escolhidas = [r for r in resultados if r["categoria"]
                  in ("loc1", "loc2", "loc3")]
    ordem_cat = {"loc1": 0, "loc2": 1, "loc3": 2}
    escolhidas.sort(key=lambda r: ordem_cat[r["categoria"]])

    demais = [r for r in resultados if r["categoria"] == "demais"]

    def _chave(r):
        t = r["taxa"]
        v = r["vagas"]
        try:
            t_num = float(t) if t != "" and t != "Sem TP" else 99999.0
        except (ValueError, TypeError):
            t_num = 99999.0
        try:
            v_num = int(v) if v != "" else -99999
        except (ValueError, TypeError):
            v_num = -99999
        return (t_num, -v_num)

    demais.sort(key=_chave)

    linhas = list(escolhidas)
    if escolhidas and demais:
        linhas.append({"om": "", "localidade": "", "taxa": "", "vagas": "",
                       "categoria": "separador"})
    linhas.extend(demais)

    return linhas


def _atualizar_Painel_Direita_OMs_ica(self):
    linhas = _calcular_oms_compativeis(self)

    tbl = self.tableWidget_vol_cb_oms_ica
    tbl.setColumnCount(3)
    tbl.setRowCount(len(linhas))
    tbl.setHorizontalHeaderLabels(["OM", "Taxa de Ocup.", "Vagas"])

    for i, r in enumerate(linhas):
        cat = r["categoria"]
        if cat == "separador":
            for col in range(3):
                txt = "━━━━" if col != 0 else "━━━━━━━━━━"
                item = QtWidgets.QTableWidgetItem(txt)
                item.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
                item.setFlags(QtCore.Qt.ItemFlag.ItemIsEnabled)
                tbl.setItem(i, col, item)
            continue

        om = r["om"]
        loc = r["localidade"]
        taxa = r["taxa"]
        vagas = r["vagas"]

        if isinstance(taxa, (int, float)):
            taxa_str = f"{taxa:.2f}"
        else:
            taxa_str = str(taxa)

        texto_om = om if not loc else f"{om}  •  {loc}"

        itens = [
            QtWidgets.QTableWidgetItem(texto_om),
            QtWidgets.QTableWidgetItem(taxa_str),
            QtWidgets.QTableWidgetItem(str(vagas)),
        ]
        for col, item in enumerate(itens):
            if col > 0:
                item.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            if col == 0 and cat in ("loc1", "loc2", "loc3"):
                cor = {"loc1": COR_LOC1, "loc2": COR_LOC2,
                       "loc3": COR_LOC3}[cat]
                pix = QtGui.QPixmap(6, 16)
                pix.fill(cor)
                item.setIcon(QtGui.QIcon(pix))
                rot = {"loc1": "1ª opção", "loc2": "2ª opção",
                       "loc3": "3ª opção"}[cat]
                item.setToolTip(f"{rot} do militar")
            tbl.setItem(i, col, item)

    tbl.resizeColumnsToContents()


def _ica_om_duplo_clique(self, row, col):
    if row is None or row < 0:
        return
    item_om = self.tableWidget_vol_cb_oms_ica.item(row, 0)
    if item_om is None:
        return

    texto_om = item_om.text().strip()
    if not texto_om or texto_om.startswith("━"):
        return

    if "•" in texto_om:
        om_nome = texto_om.split("•")[0].strip()
    else:
        om_nome = texto_om

    if not hasattr(self, "df_ica_vol_cb") or self.df_ica_vol_cb.empty:
        return
    linha_mil = self.tableWidget_vol_cb_ica.currentRow()
    if linha_mil is None or linha_mil < 0 or linha_mil >= len(self.df_ica_vol_cb):
        QtWidgets.QMessageBox.information(
            self, "Voluntários C e B",
            "Selecione primeiro um militar na tabela à esquerda."
        )
        return

    militar = self.df_ica_vol_cb.iloc[linha_mil]
    _executar_transferencia(self, militar, om_nome, origem_aba="ICA",
                            valor_mcda=None)


# ============================================================
# Transferência (compartilhada entre MCDA e ICA)
# ============================================================

def _executar_transferencia(self, militar_row, om_destino, origem_aba, valor_mcda):
    df_plamov_compilado = _get_global("df_plamov_compilado")
    if df_plamov_compilado is None or df_plamov_compilado.empty:
        QtWidgets.QMessageBox.warning(
            self, "Voluntários C e B", "Dados dos militares não carregados."
        )
        return

    saram = militar_row.get("SARAM", "")
    nome = militar_row.get("NOME", "")
    om_origem = _fmt_str(militar_row.get("OM ATUAL",
                         militar_row.get("OM_ORIGEM", ""))).strip()
    projeto = _fmt_str(militar_row.get("PROJETO",
                       militar_row.get("PROJETO_MILITAR", ""))).strip()
    posto_raw = _fmt_str(militar_row.get("POSTO", "")).strip()
    quadro = _fmt_str(militar_row.get("QUADRO", "")).strip()

    try:
        mask = df_plamov_compilado["SARAM"].astype(str).str.strip() \
            == str(saram).strip()
    except Exception:
        mask = pd.Series([False] * len(df_plamov_compilado),
                         index=df_plamov_compilado.index)

    if mask.any():
        for idx in df_plamov_compilado[mask].index.tolist():
            df_plamov_compilado.at[idx, "PLAMOV"] = om_destino

    _atualizar_tp_bma_apos_transferencia(
        om_origem, om_destino, projeto,
        posto=posto_raw if posto_raw else None,
        quadro=quadro if quadro else None,
    )

    if hasattr(self, "df_grupo_f") and isinstance(self.df_grupo_f, pd.DataFrame):
        try:
            mask_f = self.df_grupo_f["SARAM"].astype(str).str.strip() \
                == str(saram).strip()
            self.df_grupo_f = self.df_grupo_f[~mask_f].copy()
        except Exception:
            pass

    _registrar_auditoria({
        "SARAM": _fmt_str(saram),
        "NOME": _fmt_str(nome),
        "OM_ORIGEM": om_origem,
        "OM_DESTINO_SUGERIDO": om_destino,
        "OM_DESTINO_ESCOLHIDO": om_destino,
        "TIPO_INTERVENCAO": f"Transferir ({origem_aba})",
        "VALOR_MCDA": _fmt_num(valor_mcda, 5) if valor_mcda is not None else "",
        "MOTIVO": "",
        "GESTOR": "",
    })

    self.carregar_tabela_vol_cb_mcda()
    self.carregar_tabela_vol_cb_ica()
    _atualizar_painel_direito_por_aba(self)

    try:
        self.ui.statusbar.showMessage(
            f"Transferência registrada: {saram} → {om_destino}", 4000
        )
    except Exception:
        pass


# ============================================================
# Botões
# ============================================================

def _btn_transferir_vol_cb(self):
    aba = self.tabWidget_vol_cb.currentIndex()

    if aba == 0:  # MCDA
        if not hasattr(self, "df_ranking_vol_cb") or self.df_ranking_vol_cb.empty:
            QtWidgets.QMessageBox.information(
                self, "Voluntários C e B", "Não há alternativas para transferir.")
            return
        linha = self.tableWidget_vol_cb_mcda.currentRow()
        if linha is None or linha < 0:
            QtWidgets.QMessageBox.information(
                self, "Voluntários C e B",
                "Selecione uma alternativa na tabela antes de transferir.")
            return
        alt = self.df_ranking_vol_cb.iloc[linha]
        mr = pd.Series({
            "SARAM": alt["SARAM"],
            "NOME": alt.get("NOME", ""),
            "OM ATUAL": alt["OM_ORIGEM"],
            "PROJETO": alt["PROJETO_MILITAR"],
            "POSTO": "",
            "QUADRO": "",
        })
        _executar_transferencia(self, mr, alt["OM_DESTINO"],
                                origem_aba="MCDA",
                                valor_mcda=alt.get("VALOR", 0))
    else:  # ICA
        if not hasattr(self, "df_ica_vol_cb") or self.df_ica_vol_cb.empty:
            QtWidgets.QMessageBox.information(
                self, "Voluntários C e B", "Não há militares na lista.")
            return
        linha = self.tableWidget_vol_cb_ica.currentRow()
        if linha is None or linha < 0:
            QtWidgets.QMessageBox.information(
                self, "Voluntários C e B",
                "Selecione um militar antes de transferir.")
            return
        linha_om = self.tableWidget_vol_cb_oms_ica.currentRow()
        if linha_om is None or linha_om < 0:
            QtWidgets.QMessageBox.information(
                self, "Voluntários C e B",
                "Selecione uma OM no painel direito antes de transferir.\n"
                "(Dica: você também pode dar duplo-clique na OM desejada.)")
            return
        item_om = self.tableWidget_vol_cb_oms_ica.item(linha_om, 0)
        if item_om is None:
            return
        texto_om = item_om.text().strip()
        if not texto_om or texto_om.startswith("━"):
            return
        om_nome = texto_om.split("•")[0].strip(
        ) if "•" in texto_om else texto_om

        militar = self.df_ica_vol_cb.iloc[linha]
        _executar_transferencia(self, militar, om_nome,
                                origem_aba="ICA", valor_mcda=None)


def _btn_manter_origem_vol_cb(self):
    aba = self.tabWidget_vol_cb.currentIndex()
    df_plamov_compilado = _get_global("df_plamov_compilado")
    if df_plamov_compilado is None:
        return

    if aba == 0:
        if not hasattr(self, "df_ranking_vol_cb") or self.df_ranking_vol_cb.empty:
            QtWidgets.QMessageBox.information(
                self, "Voluntários C e B", "Não há alternativas na tabela.")
            return
        linha = self.tableWidget_vol_cb_mcda.currentRow()
        if linha is None or linha < 0:
            QtWidgets.QMessageBox.information(
                self, "Voluntários C e B",
                "Selecione uma alternativa do militar antes de manter na origem.")
            return
        alt = self.df_ranking_vol_cb.iloc[linha]
        saram = alt["SARAM"]
        om_origem = _fmt_str(alt.get("OM_ORIGEM", "")).strip()
        nome = alt.get("NOME", "")
        try:
            m = self.df_ranking_vol_cb["SARAM"].astype(str).str.strip() \
                == str(saram).strip()
            self.df_ranking_vol_cb = self.df_ranking_vol_cb[~m].copy()
        except Exception:
            pass
    else:
        if not hasattr(self, "df_ica_vol_cb") or self.df_ica_vol_cb.empty:
            QtWidgets.QMessageBox.information(
                self, "Voluntários C e B", "Não há militares na lista.")
            return
        linha = self.tableWidget_vol_cb_ica.currentRow()
        if linha is None or linha < 0:
            QtWidgets.QMessageBox.information(
                self, "Voluntários C e B",
                "Selecione um militar antes de manter na origem.")
            return
        militar = self.df_ica_vol_cb.iloc[linha]
        saram = militar.get("SARAM", "")
        om_origem = _fmt_str(militar.get("OM ATUAL", "")).strip()
        nome = militar.get("NOME", "")

    if not df_plamov_compilado.empty:
        if "MANTIDO_ORIGEM_CB" not in df_plamov_compilado.columns:
            df_plamov_compilado["MANTIDO_ORIGEM_CB"] = False
        try:
            m = df_plamov_compilado["SARAM"].astype(str).str.strip() \
                == str(saram).strip()
            for idx in df_plamov_compilado[m].index.tolist():
                df_plamov_compilado.at[idx, "MANTIDO_ORIGEM_CB"] = True
        except Exception:
            pass

    if hasattr(self, "df_grupo_f") and isinstance(self.df_grupo_f, pd.DataFrame):
        try:
            mf = self.df_grupo_f["SARAM"].astype(str).str.strip() \
                == str(saram).strip()
            self.df_grupo_f = self.df_grupo_f[~mf].copy()
        except Exception:
            pass

    _registrar_auditoria({
        "SARAM": _fmt_str(saram),
        "NOME": _fmt_str(nome),
        "OM_ORIGEM": om_origem,
        "OM_DESTINO_SUGERIDO": "",
        "OM_DESTINO_ESCOLHIDO": "",
        "TIPO_INTERVENCAO": "Manter na Origem",
        "VALOR_MCDA": "",
        "MOTIVO": "",
        "GESTOR": "",
    })

    if aba == 0:
        _popular_tabela_mcda(self, self.df_ranking_vol_cb)
        n_mil = (self.df_ranking_vol_cb["SARAM"].nunique()
                 if not self.df_ranking_vol_cb.empty else 0)
        n_alt = len(self.df_ranking_vol_cb)
        self.lbl_resumo_vol_cb.setText(
            f"{n_mil} militares  |  {n_alt} alternativas (MCDA)")
    self.carregar_tabela_vol_cb_ica()
    _atualizar_painel_direito_por_aba(self)

    try:
        self.ui.statusbar.showMessage(
            f"Militar {saram} mantido na OM de origem ({om_origem}).", 4000
        )
    except Exception:
        pass


# ============================================================
# Painel direito dinâmico
# ============================================================

def _atualizar_painel_dados_militar_mcda(self):
    if not hasattr(self, "df_ranking_vol_cb") or self.df_ranking_vol_cb.empty:
        _limpar_painel_dados_militar(self)
        return
    linha = self.tableWidget_vol_cb_mcda.currentRow()
    if linha is None or linha < 0 or linha >= len(self.df_ranking_vol_cb):
        _limpar_painel_dados_militar(self)
        return
    saram_alvo = self.df_ranking_vol_cb.iloc[linha]["SARAM"]
    _popular_dados_militar(self, saram_alvo)


def _atualizar_painel_dados_militar_ica(self):
    _atualizar_Painel_Direita_OMs_ica(self)


def _atualizar_painel_direito_por_aba(self):
    aba = self.tabWidget_vol_cb.currentIndex()
    if aba == 0:
        self.stack_painel_dir_vol_cb.setCurrentWidget(
            self.widget_painel_dados_mcda)
        _atualizar_painel_dados_militar_mcda(self)
    else:
        self.stack_painel_dir_vol_cb.setCurrentWidget(
            self.widget_painel_oms_ica)
        _atualizar_painel_dados_militar_ica(self)


def _limpar_painel_dados_militar(self):
    if hasattr(self, "lbl_dados_militar_vol_cb"):
        self.lbl_dados_militar_vol_cb.setText(
            "Selecione uma linha para ver os dados do militar.")


def _popular_dados_militar(self, saram_alvo):
    df_plamov_compilado = _get_global("df_plamov_compilado")
    if df_plamov_compilado is None or df_plamov_compilado.empty:
        _limpar_painel_dados_militar(self)
        return
    try:
        mask = df_plamov_compilado["SARAM"].astype(str).str.strip() \
            == str(saram_alvo).strip()
    except Exception:
        _limpar_painel_dados_militar(self)
        return
    if not mask.any():
        _limpar_painel_dados_militar(self)
        return

    mil = df_plamov_compilado[mask].iloc[0]

    pares = [
        ("SARAM", _fmt_str(mil.get("SARAM"))),
        ("Nome", _fmt_str(mil.get("NOME"))),
        ("Posto", _fmt_str(mil.get("POSTO"))),
        ("Quadro", _fmt_str(mil.get("QUADRO"))),
        ("Especialidade", _fmt_str(mil.get("ESP"))),
        ("Projeto", _fmt_str(mil.get("PROJETO"))),
        ("OM Atual", _fmt_str(mil.get("OM ATUAL"))),
        ("Loc. Atual", _fmt_str(mil.get("LOC ATUAL"))),
        ("Tempo Loc.", _fmt_num(mil.get("TEMPO LOC"), 2)),
        ("Antiguidade", _fmt_int(mil.get("ANTIGUIDADE"))),
        ("LOC 1", _fmt_str(mil.get("LOC 1"))),
        ("LOC 2", _fmt_str(mil.get("LOC 2"))),
        ("LOC 3", _fmt_str(mil.get("LOC 3"))),
        ("Cônjuge FAB?", _fmt_str(mil.get("CÔNJUGE DA FAB?"))),
        ("Dados Cônjuge", _fmt_str(mil.get("DADOS CÔNJUGE"))),
    ]
    linhas_html = [
        f"<tr><td style='padding:2px 8px 2px 2px; font-weight:bold; "
        f"vertical-align:top;'>{k}:</td>"
        f"<td style='padding:2px;'>{v}</td></tr>"
        for k, v in pares
    ]
    html = (
        "<div style='font-family: Segoe UI, Arial; font-size: 11px;'>"
        "<table cellspacing='0'>" + "".join(linhas_html) + "</table></div>"
    )
    self.lbl_dados_militar_vol_cb.setText(html)


# ============================================================
# Instalação
# ============================================================

def install(janela):
    if getattr(janela, "_vol_cb_instalada", False):
        return
    janela._vol_cb_instalada = True

    janela.page_vol_cb = QtWidgets.QWidget()
    layout_root = QtWidgets.QVBoxLayout(janela.page_vol_cb)
    layout_root.setContentsMargins(2, 2, 2, 2)
    layout_root.setSpacing(2)

    # Título + resumo
    widget_topo = QtWidgets.QWidget()
    widget_topo.setFixedHeight(24)
    layout_topo = QtWidgets.QHBoxLayout(widget_topo)
    layout_topo.setContentsMargins(5, 0, 5, 0)
    layout_topo.setSpacing(15)

    lbl_titulo = QtWidgets.QLabel("Voluntários C e B")
    lbl_titulo.setStyleSheet("font-size: 13px; font-weight: bold;")
    layout_topo.addWidget(lbl_titulo)

    janela.lbl_resumo_vol_cb = QtWidgets.QLabel("")
    janela.lbl_resumo_vol_cb.setStyleSheet(
        "font-size: 11px; color: palette(mid);")
    layout_topo.addWidget(janela.lbl_resumo_vol_cb)
    layout_topo.addStretch()
    layout_root.addWidget(widget_topo)

    # Descrição
    lbl_desc = QtWidgets.QLabel(
        "Voluntários em Loc A que pediram Boa Vista, Porto Velho, "
        "Manaus ou Belém em alguma de suas opções (LOC 1, 2 ou 3).")
    lbl_desc.setStyleSheet(
        "font-size: 10px; color: palette(mid); padding: 2px 5px;")
    lbl_desc.setFixedHeight(18)
    layout_root.addWidget(lbl_desc)

    # Splitter
    splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Horizontal)

    # Esquerda: QTabWidget
    janela.tabWidget_vol_cb = QtWidgets.QTabWidget()
    janela.tabWidget_vol_cb.setDocumentMode(True)

    # Aba MCDA
    aba_mcda = QtWidgets.QWidget()
    layout_mcda = QtWidgets.QVBoxLayout(aba_mcda)
    layout_mcda.setContentsMargins(0, 4, 0, 0)
    layout_mcda.setSpacing(0)

    janela.tableWidget_vol_cb_mcda = QtWidgets.QTableWidget()
    janela.tableWidget_vol_cb_mcda.setSelectionBehavior(
        QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
    janela.tableWidget_vol_cb_mcda.setSelectionMode(
        QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
    janela.tableWidget_vol_cb_mcda.setEditTriggers(
        QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
    janela.tableWidget_vol_cb_mcda.setAlternatingRowColors(True)
    janela.tableWidget_vol_cb_mcda.setStyleSheet("""
        QHeaderView::section { padding-right: 5px; padding-left: 5px; }
    """)
    layout_mcda.addWidget(janela.tableWidget_vol_cb_mcda)
    janela.tabWidget_vol_cb.addTab(aba_mcda, "MCDA")

    # Aba ICA
    aba_ica = QtWidgets.QWidget()
    layout_ica = QtWidgets.QVBoxLayout(aba_ica)
    layout_ica.setContentsMargins(0, 4, 0, 0)
    layout_ica.setSpacing(0)

    janela.tableWidget_vol_cb_ica = QtWidgets.QTableWidget()
    janela.tableWidget_vol_cb_ica.setSelectionBehavior(
        QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
    janela.tableWidget_vol_cb_ica.setSelectionMode(
        QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
    janela.tableWidget_vol_cb_ica.setEditTriggers(
        QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
    janela.tableWidget_vol_cb_ica.setAlternatingRowColors(True)
    janela.tableWidget_vol_cb_ica.setStyleSheet("""
        QHeaderView::section { padding-right: 5px; padding-left: 5px; }
    """)
    layout_ica.addWidget(janela.tableWidget_vol_cb_ica)
    janela.tabWidget_vol_cb.addTab(aba_ica, "ICA 30-4")

    splitter.addWidget(janela.tabWidget_vol_cb)

    # Direita: stack de painéis + botões
    widget_dir = QtWidgets.QWidget()
    layout_dir = QtWidgets.QVBoxLayout(widget_dir)
    layout_dir.setContentsMargins(5, 0, 5, 0)
    layout_dir.setSpacing(5)

    janela.stack_painel_dir_vol_cb = QtWidgets.QStackedWidget()

    # Painel "Dados do Militar" (MCDA)
    janela.widget_painel_dados_mcda = QtWidgets.QWidget()
    lay_dados = QtWidgets.QVBoxLayout(janela.widget_painel_dados_mcda)
    lay_dados.setContentsMargins(0, 0, 0, 0)
    lay_dados.setSpacing(3)

    lbl_titulo_dados = QtWidgets.QLabel("Dados do Militar")
    lbl_titulo_dados.setFixedHeight(22)
    lbl_titulo_dados.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
    lbl_titulo_dados.setStyleSheet("font-size: 12px; font-weight: bold;")
    lay_dados.addWidget(lbl_titulo_dados)

    janela.lbl_dados_militar_vol_cb = QtWidgets.QLabel(
        "Selecione uma linha para ver os dados do militar.")
    janela.lbl_dados_militar_vol_cb.setWordWrap(True)
    janela.lbl_dados_militar_vol_cb.setAlignment(
        QtCore.Qt.AlignmentFlag.AlignTop | QtCore.Qt.AlignmentFlag.AlignLeft)
    janela.lbl_dados_militar_vol_cb.setStyleSheet(
        "border: 1px solid palette(mid); padding: 6px;")
    janela.lbl_dados_militar_vol_cb.setTextInteractionFlags(
        QtCore.Qt.TextInteractionFlag.TextSelectableByMouse)
    scroll_dados = QtWidgets.QScrollArea()
    scroll_dados.setWidgetResizable(True)
    scroll_dados.setWidget(janela.lbl_dados_militar_vol_cb)
    lay_dados.addWidget(scroll_dados, 1)

    janela.stack_painel_dir_vol_cb.addWidget(janela.widget_painel_dados_mcda)

    # Painel "OMs Disponíveis" (ICA)
    janela.widget_painel_oms_ica = QtWidgets.QWidget()
    lay_oms = QtWidgets.QVBoxLayout(janela.widget_painel_oms_ica)
    lay_oms.setContentsMargins(0, 0, 0, 0)
    lay_oms.setSpacing(3)

    lbl_titulo_oms = QtWidgets.QLabel("OMs Disponíveis")
    lbl_titulo_oms.setFixedHeight(22)
    lbl_titulo_oms.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
    lbl_titulo_oms.setStyleSheet("font-size: 12px; font-weight: bold;")
    lay_oms.addWidget(lbl_titulo_oms)

    # Legenda de cores
    widget_leg = QtWidgets.QWidget()
    lay_leg = QtWidgets.QHBoxLayout(widget_leg)
    lay_leg.setContentsMargins(2, 0, 2, 0)
    lay_leg.setSpacing(6)
    widget_leg.setFixedHeight(18)
    for texto, cor in [("1ª", COR_LOC1), ("2ª", COR_LOC2), ("3ª", COR_LOC3)]:
        swatch = QtWidgets.QLabel()
        swatch.setFixedSize(10, 10)
        swatch.setStyleSheet(
            f"background-color: rgb({cor.red()},{cor.green()},{cor.blue()}); "
            f"border-radius: 2px;")
        lay_leg.addWidget(swatch)
        lbl = QtWidgets.QLabel(texto)
        lbl.setStyleSheet("font-size: 9px;")
        lay_leg.addWidget(lbl)
    lay_leg.addStretch()
    lay_oms.addWidget(widget_leg)

    janela.tableWidget_vol_cb_oms_ica = QtWidgets.QTableWidget()
    janela.tableWidget_vol_cb_oms_ica.setSelectionBehavior(
        QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
    janela.tableWidget_vol_cb_oms_ica.setSelectionMode(
        QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
    janela.tableWidget_vol_cb_oms_ica.setEditTriggers(
        QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
    janela.tableWidget_vol_cb_oms_ica.setAlternatingRowColors(True)
    janela.tableWidget_vol_cb_oms_ica.setStyleSheet("""
        QHeaderView::section { padding-right: 5px; padding-left: 5px; }
    """)
    janela.tableWidget_vol_cb_oms_ica.setIconSize(QtCore.QSize(6, 16))
    lay_oms.addWidget(janela.tableWidget_vol_cb_oms_ica, 1)

    janela.stack_painel_dir_vol_cb.addWidget(janela.widget_painel_oms_ica)

    layout_dir.addWidget(janela.stack_painel_dir_vol_cb, 1)

    # Botões
    widget_botoes = QtWidgets.QWidget()
    lay_bt = QtWidgets.QVBoxLayout(widget_botoes)
    lay_bt.setContentsMargins(0, 5, 0, 0)
    lay_bt.setSpacing(4)

    janela.btn_transferir_vol_cb_widget = QtWidgets.QPushButton("Transferir")
    janela.btn_transferir_vol_cb_widget.setMinimumHeight(32)
    janela.btn_transferir_vol_cb_widget.setStyleSheet(
        "QPushButton { font-weight: bold; }")
    lay_bt.addWidget(janela.btn_transferir_vol_cb_widget)

    janela.btn_manter_origem_vol_cb_widget = QtWidgets.QPushButton(
        "Manter na Origem")
    janela.btn_manter_origem_vol_cb_widget.setMinimumHeight(32)
    lay_bt.addWidget(janela.btn_manter_origem_vol_cb_widget)

    layout_dir.addWidget(widget_botoes)

    splitter.addWidget(widget_dir)
    splitter.setStretchFactor(0, 7)
    splitter.setStretchFactor(1, 3)
    layout_root.addWidget(splitter, 1)

    # Estado inicial
    janela.df_ranking_vol_cb = pd.DataFrame()
    janela.df_ica_vol_cb = pd.DataFrame()

    # Anexa métodos
    janela.Pag_VoluntariosCB = types.MethodType(_Pag_VoluntariosCB, janela)
    janela.carregar_tabela_vol_cb_mcda = types.MethodType(
        _carregar_tabela_vol_cb_mcda, janela)
    janela.carregar_tabela_vol_cb_ica = types.MethodType(
        _carregar_tabela_vol_cb_ica, janela)
    janela.btn_transferir_vol_cb = types.MethodType(
        _btn_transferir_vol_cb, janela)
    janela.btn_manter_origem_vol_cb = types.MethodType(
        _btn_manter_origem_vol_cb, janela)
    # Retrocompat
    janela.carregar_tabela_vol_cb = janela.carregar_tabela_vol_cb_mcda

    # Stack + menu
    janela.ui.stackedWidget.addWidget(janela.page_vol_cb)
    janela.actionVoluntariosCB = QtGui.QAction("Voluntários C e B", janela)
    janela.ui.menuMenu.addAction(janela.actionVoluntariosCB)
    janela.actionVoluntariosCB.triggered.connect(
        lambda: janela.Pag_VoluntariosCB())

    # Conexões
    janela.tabWidget_vol_cb.currentChanged.connect(
        lambda _i: _atualizar_painel_direito_por_aba(janela))

    janela.tableWidget_vol_cb_mcda.itemSelectionChanged.connect(
        lambda: (_atualizar_painel_dados_militar_mcda(janela)
                 if janela.tabWidget_vol_cb.currentIndex() == 0 else None))

    janela.tableWidget_vol_cb_ica.itemSelectionChanged.connect(
        lambda: (_atualizar_painel_dados_militar_ica(janela)
                 if janela.tabWidget_vol_cb.currentIndex() == 1 else None))

    janela.tableWidget_vol_cb_oms_ica.cellDoubleClicked.connect(
        lambda r, c: _ica_om_duplo_clique(janela, r, c))

    janela.btn_transferir_vol_cb_widget.clicked.connect(
        lambda: janela.btn_transferir_vol_cb())
    janela.btn_manter_origem_vol_cb_widget.clicked.connect(
        lambda: janela.btn_manter_origem_vol_cb())
