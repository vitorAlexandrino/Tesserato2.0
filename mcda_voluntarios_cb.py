# -*- coding: utf-8 -*-
"""
mcda_voluntarios_cb.py

Módulo puro de cálculo do MCDA para o Grupo F do TESSERATO 2.0:
Voluntários em Loc A que pediram localidades C ou B (Boa Vista, Porto Velho,
Manaus ou Belém) em alguma de suas preferências LOC 1, LOC 2 ou LOC 3.

Modelo: função valor aditiva clássica
    V(Xij) = 0.469*TxVai + 0.276*Cap + 0.157*TxFica + 0.066*Tloc + 0.032*Intencao

IMPORTANTE — filtragem de alternativas:
Cada militar i gera alternativas APENAS contra as OMs cuja Localidade é uma
das localidades C/B que ELE especificamente pediu em LOC 1, LOC 2 ou LOC 3.
Ex.: militar com LOC1=Brasília, LOC2=Boa Vista, LOC3=Natal gera alternativas
somente para as OMs de Boa Vista.

Este módulo NÃO depende de PyQt6 e é importável/testável isoladamente.
"""

import pandas as pd
import numpy as np


# ============================================================
# Constantes do modelo
# ============================================================

PESOS = {
    "TX_VAI": 0.469,
    "CAPACITACAO": 0.276,
    "TX_FICA": 0.157,
    "TLOC": 0.066,
    "INTENCAO": 0.032,
}

# Ratings fixos do critério INTENÇÃO (Tabela do TCC)
RATING_LOC1 = 0.554
RATING_LOC2 = 0.244
RATING_LOC3 = 0.158
RATING_NAO_PEDIDA = 0.044  # não é usado quando filtramos por preferência do militar

# Localidades alvo do Grupo F (classes C e B)
LOCALIDADES_CB = {"BOA VISTA", "PORTO VELHO", "MANAUS", "BELÉM", "BELEM"}

# OMs que compõem a localidade "Santa Cruz" (não é alvo do Grupo F,
# mas mantemos o mapeamento caso seja útil no futuro)
OMS_SANTA_CRUZ = {"1/7 GAV", "3/8 GAV", "1 GAVCA"}


# ============================================================
# Funções auxiliares de normalização
# ============================================================

def normalizar_max(serie: pd.Series) -> pd.Series:
    """Normaliza para maximização: (x - xmin) / (xmax - xmin)."""
    s = pd.to_numeric(serie, errors="coerce").fillna(0.0)
    xmin, xmax = s.min(), s.max()
    if xmax == xmin:
        return pd.Series([1.0] * len(s), index=s.index)
    return (s - xmin) / (xmax - xmin)


def normalizar_min(serie: pd.Series) -> pd.Series:
    """Normaliza para minimização: (xmax - x) / (xmax - xmin)."""
    s = pd.to_numeric(serie, errors="coerce").fillna(0.0)
    xmin, xmax = s.min(), s.max()
    if xmax == xmin:
        return pd.Series([1.0] * len(s), index=s.index)
    return (xmax - s) / (xmax - xmin)


# ============================================================
# Helpers de acesso aos dados
# ============================================================

def _strip_upper(valor) -> str:
    """Normaliza texto para comparação: string, strip, upper."""
    if pd.isna(valor):
        return ""
    return str(valor).strip().upper()


def _delta_por_unidade_projeto(df_TP_BMA: pd.DataFrame, unidade: str, projeto: str) -> float:
    """
    Retorna o delta (Taxa ideal - Taxa atual) agregado para uma combinação
    de Unidade + Projeto em df_TP_BMA. Soma todos os postos dessa combinação.

    Se não encontrar linhas, ou se TLP total for 0, retorna 0.0 (a OM/projeto
    não fica excluída do ranking — apenas pontua baixo nesse critério).
    """
    if df_TP_BMA is None or df_TP_BMA.empty:
        return 0.0

    unidade_n = _strip_upper(unidade)
    projeto_n = _strip_upper(projeto)

    mask = (
        df_TP_BMA["Unidade"].astype(str).str.strip().str.upper() == unidade_n
    ) & (
        df_TP_BMA["Projeto"].astype(str).str.strip().str.upper() == projeto_n
    )
    filtradas = df_TP_BMA[mask]

    if filtradas.empty:
        return 0.0

    # Soma TLP e Existentes de todos os postos dessa Unidade+Projeto
    tlp_total = pd.to_numeric(
        filtradas["TLP Ano Corrente"], errors="coerce").fillna(0).sum()
    exist_total = pd.to_numeric(
        filtradas["Existentes"], errors="coerce").fillna(0).sum()

    if tlp_total == 0:
        return 0.0

    taxa_atual = exist_total / tlp_total
    # Taxa ideal agregada: média ponderada pelos TLPs
    taxa_ideal_col = pd.to_numeric(
        filtradas["Taxa ideal"], errors="coerce").fillna(0)
    # Ponderação simples pela linha (cada linha já vem com sua Taxa ideal):
    # se todas as linhas tiverem a mesma Taxa ideal (comum no mesmo Projeto),
    # o resultado é essa taxa; caso difiram, usamos a média ponderada por TLP.
    tlps = pd.to_numeric(
        filtradas["TLP Ano Corrente"], errors="coerce").fillna(0)
    if tlps.sum() > 0:
        taxa_ideal = (taxa_ideal_col * tlps).sum() / tlps.sum()
    else:
        taxa_ideal = taxa_ideal_col.mean() if len(taxa_ideal_col) > 0 else 0.0

    return float(taxa_ideal - taxa_atual)


def _oms_das_localidades_pedidas(
    df_TP_BMA: pd.DataFrame,
    localidades_pedidas_upper: list,
) -> pd.DataFrame:
    """
    Retorna subset de df_TP_BMA com OMs únicas (por Unidade+Localidade) cujas
    localidades estão entre as localidades pedidas pelo militar (que sejam C ou B).

    Deduplica por Unidade (uma linha por OM, pegando a primeira ocorrência).
    """
    if df_TP_BMA is None or df_TP_BMA.empty:
        return pd.DataFrame(columns=["Unidade", "Localidade"])

    loc_norm = df_TP_BMA["Localidade"].astype(str).str.strip().str.upper()
    mask = loc_norm.isin([x.upper() for x in localidades_pedidas_upper])
    sub = df_TP_BMA[mask].copy()

    if sub.empty:
        return pd.DataFrame(columns=["Unidade", "Localidade"])

    # Deduplica por Unidade (mantém primeira ocorrência)
    sub = sub.drop_duplicates(subset=["Unidade"], keep="first")
    return sub[["Unidade", "Localidade"]].reset_index(drop=True)


def _rating_intencao(localidade_om_j: str, loc1: str, loc2: str, loc3: str) -> float:
    """
    Retorna rating da INTENÇÃO conforme a posição da localidade da OM j nas
    preferências do militar. Dentro do Grupo F, como só geramos alternativas
    para localidades C/B que o militar pediu, esta função nunca deve retornar
    RATING_NAO_PEDIDA — mas mantemos como fallback defensivo.
    """
    loc_j = _strip_upper(localidade_om_j)
    l1 = _strip_upper(loc1)
    l2 = _strip_upper(loc2)
    l3 = _strip_upper(loc3)

    if loc_j == l1:
        return RATING_LOC1
    if loc_j == l2:
        return RATING_LOC2
    if loc_j == l3:
        return RATING_LOC3
    return RATING_NAO_PEDIDA


# ============================================================
# Função principal
# ============================================================

def calcular_mcda_voluntarios_cb(
    df_grupo_f: pd.DataFrame,
    df_TP_BMA: pd.DataFrame,
) -> pd.DataFrame:
    """
    Calcula o ranking MCDA do Grupo F (Voluntários C e B).

    Parameters
    ----------
    df_grupo_f : DataFrame
        Subconjunto de df_plamov_compilado com os militares do Grupo F.
        Deve conter as colunas: SARAM, NOME, OM ATUAL, PROJETO, TEMPO LOC,
        LOC 1, LOC 2, LOC 3, e as demais que forem exibidas.
    df_TP_BMA : DataFrame
        Relatório TP BMA carregado da planilha, com as colunas: Localidade,
        Unidade, Projeto, TLP Ano Corrente, Existentes, Taxa ideal, Taxa atual.

    Returns
    -------
    DataFrame com uma linha por par (SARAM, OM_DESTINO), ordenado por VALOR
    decrescente. Colunas:
        SARAM, NOME, OM_ORIGEM, PROJETO, OM_DESTINO, LOCALIDADE_DESTINO,
        TX_VAI_BRUTO, TX_VAI_NORM,
        CAPACITACAO,
        TX_FICA_BRUTO, TX_FICA_NORM,
        TLOC_BRUTO, TLOC_NORM,
        INTENCAO,
        VALOR
    """
    colunas_saida = [
        "SARAM", "NOME", "OM_ORIGEM", "PROJETO_MILITAR",
        "OM_DESTINO", "LOCALIDADE_DESTINO", "PROJETO_OM_DESTINO",
        "TX_VAI_BRUTO", "TX_VAI_NORM",
        "CAPACITACAO",
        "TX_FICA_BRUTO", "TX_FICA_NORM",
        "TLOC_BRUTO", "TLOC_NORM",
        "INTENCAO",
        "VALOR",
    ]

    if df_grupo_f is None or df_grupo_f.empty:
        return pd.DataFrame(columns=colunas_saida)
    if df_TP_BMA is None or df_TP_BMA.empty:
        return pd.DataFrame(columns=colunas_saida)

    # Colunas obrigatórias em df_grupo_f
    for col in ["SARAM", "OM ATUAL", "PROJETO", "TEMPO LOC",
                "LOC 1", "LOC 2", "LOC 3"]:
        if col not in df_grupo_f.columns:
            return pd.DataFrame(columns=colunas_saida)

    # Colunas obrigatórias em df_TP_BMA
    for col in ["Localidade", "Unidade", "Projeto",
                "TLP Ano Corrente", "Existentes", "Taxa ideal", "Taxa atual"]:
        if col not in df_TP_BMA.columns:
            return pd.DataFrame(columns=colunas_saida)

    # ---------- PASSO 1: gerar alternativas (brutas) em listas ----------
    saram_list = []
    nome_list = []
    om_origem_list = []
    projeto_mil_list = []
    om_destino_list = []
    loc_destino_list = []
    projeto_om_dest_list = []
    tx_vai_bruto_list = []
    capacitacao_list = []
    tx_fica_bruto_list = []
    tloc_bruto_list = []
    intencao_list = []

    for _, militar in df_grupo_f.iterrows():
        saram = militar.get("SARAM", "")
        nome = militar.get("NOME", "")
        om_origem = militar.get("OM ATUAL", "")
        projeto_mil = militar.get("PROJETO", "")
        try:
            tloc = float(militar.get("TEMPO LOC", 0) or 0)
        except (ValueError, TypeError):
            tloc = 0.0

        loc1 = militar.get("LOC 1", "")
        loc2 = militar.get("LOC 2", "")
        loc3 = militar.get("LOC 3", "")

        # Localidades C/B que ESTE militar pediu (normalizadas)
        pedidas = []
        for loc in (loc1, loc2, loc3):
            loc_u = _strip_upper(loc)
            if loc_u in LOCALIDADES_CB and loc_u not in pedidas:
                pedidas.append(loc_u)

        if not pedidas:
            # Militar está no Grupo F mas nenhuma de suas LOCs é C/B após
            # normalização: não gera alternativas.
            continue

        # OMs candidatas: as do df_TP_BMA cuja Localidade ∈ pedidas.
        # IMPORTANTE: para que CAPACITACAO funcione, precisamos considerar
        # linhas de df_TP_BMA tanto do Projeto do militar (capacitação=1)
        # quanto de outros Projetos (capacitação=0), mas apenas uma entrada
        # por Unidade destino. Se uma mesma Unidade aparece em múltiplos
        # Projetos, consideramos capacitacao=1 se o Projeto do militar
        # existe naquela Unidade.
        loc_norm_bma = df_TP_BMA["Localidade"].astype(
            str).str.strip().str.upper()
        mask_loc = loc_norm_bma.isin(pedidas)
        cand = df_TP_BMA[mask_loc].copy()

        if cand.empty:
            continue

        # Unidades únicas candidatas
        unidades_unicas = cand["Unidade"].astype(str).str.strip().unique()

        for unidade in unidades_unicas:
            # Localidade desta unidade (pega a primeira linha)
            linhas_un = cand[cand["Unidade"].astype(
                str).str.strip() == unidade]
            if linhas_un.empty:
                continue
            localidade_un = str(linhas_un.iloc[0]["Localidade"]).strip()

            # CAPACITACAO: 1 se existe linha para (Unidade, Projeto_do_militar)
            projetos_na_unidade = linhas_un["Projeto"].astype(
                str).str.strip().str.upper().tolist()
            projeto_mil_u = _strip_upper(projeto_mil)
            capacitacao = 1 if projeto_mil_u in projetos_na_unidade else 0

            # Projeto da OM que "representa" essa unidade no destino:
            # se o militar é capacitado, é o próprio projeto dele;
            # senão, pega o primeiro projeto listado (para exibição).
            if capacitacao == 1:
                projeto_om_dest = projeto_mil
            else:
                projeto_om_dest = str(linhas_un.iloc[0]["Projeto"]).strip()

            # TX_VAI_BRUTO: delta da OM destino no PROJETO do militar.
            # Se o militar não for capacitado para nenhum projeto da OM,
            # usamos o delta agregado da OM naquele Projeto do militar
            # (que provavelmente retornará 0, pois não há linha):
            # nesse caso, tentamos o delta agregado de qualquer linha da OM.
            tx_vai = _delta_por_unidade_projeto(
                df_TP_BMA, unidade, projeto_mil)
            if tx_vai == 0.0 and capacitacao == 0:
                # fallback: delta agregado sobre todos os projetos da OM
                tlp = pd.to_numeric(
                    linhas_un["TLP Ano Corrente"], errors="coerce").fillna(0).sum()
                exi = pd.to_numeric(
                    linhas_un["Existentes"], errors="coerce").fillna(0).sum()
                if tlp > 0:
                    ti = pd.to_numeric(
                        linhas_un["Taxa ideal"], errors="coerce").fillna(0)
                    tlps = pd.to_numeric(
                        linhas_un["TLP Ano Corrente"], errors="coerce").fillna(0)
                    if tlps.sum() > 0:
                        ti_ag = (ti * tlps).sum() / tlps.sum()
                    else:
                        ti_ag = ti.mean() if len(ti) > 0 else 0.0
                    tx_vai = float(ti_ag - exi / tlp)

            # TX_FICA_BRUTO: delta da OM origem do militar, no Projeto dele
            tx_fica = _delta_por_unidade_projeto(
                df_TP_BMA, om_origem, projeto_mil)

            # INTENCAO: rating pela posição da localidade
            intencao = _rating_intencao(localidade_un, loc1, loc2, loc3)

            # Registra alternativa
            saram_list.append(saram)
            nome_list.append(nome)
            om_origem_list.append(om_origem)
            projeto_mil_list.append(projeto_mil)
            om_destino_list.append(unidade)
            loc_destino_list.append(localidade_un)
            projeto_om_dest_list.append(projeto_om_dest)
            tx_vai_bruto_list.append(tx_vai)
            capacitacao_list.append(capacitacao)
            tx_fica_bruto_list.append(tx_fica)
            tloc_bruto_list.append(tloc)
            intencao_list.append(intencao)

    # Se nenhuma alternativa foi gerada, retorna vazio
    if not saram_list:
        return pd.DataFrame(columns=colunas_saida)

    # ---------- PASSO 2: montar DataFrame e normalizar ----------
    df = pd.DataFrame({
        "SARAM": saram_list,
        "NOME": nome_list,
        "OM_ORIGEM": om_origem_list,
        "PROJETO_MILITAR": projeto_mil_list,
        "OM_DESTINO": om_destino_list,
        "LOCALIDADE_DESTINO": loc_destino_list,
        "PROJETO_OM_DESTINO": projeto_om_dest_list,
        "TX_VAI_BRUTO": tx_vai_bruto_list,
        "CAPACITACAO": capacitacao_list,
        "TX_FICA_BRUTO": tx_fica_bruto_list,
        "TLOC_BRUTO": tloc_bruto_list,
        "INTENCAO": intencao_list,
    })

    # Normalização TX_VAI: sobre todo o vetor de deltas das OMs j (todas as linhas)
    df["TX_VAI_NORM"] = normalizar_max(df["TX_VAI_BRUTO"]).values

    # Normalização TX_FICA: sobre o vetor dos militares do subgrupo
    # (um valor por militar — minimização: quanto menor o delta de origem,
    # melhor a alternativa, pois a OM de origem precisa menos do militar).
    # Montamos um vetor único por SARAM e projetamos de volta.
    tx_fica_por_mil = df.drop_duplicates(subset=["SARAM"])[
        ["SARAM", "TX_FICA_BRUTO"]]
    tx_fica_norm_map = dict(zip(
        tx_fica_por_mil["SARAM"],
        normalizar_min(tx_fica_por_mil["TX_FICA_BRUTO"]).values,
    ))
    df["TX_FICA_NORM"] = df["SARAM"].map(
        tx_fica_norm_map).fillna(0.0).astype(float)

    # Normalização TLOC: sobre o vetor dos militares do subgrupo (maximização)
    tloc_por_mil = df.drop_duplicates(subset=["SARAM"])[
        ["SARAM", "TLOC_BRUTO"]]
    tloc_norm_map = dict(zip(
        tloc_por_mil["SARAM"],
        normalizar_max(tloc_por_mil["TLOC_BRUTO"]).values,
    ))
    df["TLOC_NORM"] = df["SARAM"].map(tloc_norm_map).fillna(0.0).astype(float)

    # ---------- PASSO 3: calcular VALOR ----------
    df["VALOR"] = (
        PESOS["TX_VAI"] * df["TX_VAI_NORM"].astype(float)
        + PESOS["CAPACITACAO"] * df["CAPACITACAO"].astype(float)
        + PESOS["TX_FICA"] * df["TX_FICA_NORM"].astype(float)
        + PESOS["TLOC"] * df["TLOC_NORM"].astype(float)
        + PESOS["INTENCAO"] * df["INTENCAO"].astype(float)
    )

    # Ordena por VALOR descendente
    df = df.sort_values(by="VALOR", ascending=False).reset_index(drop=True)

    return df[colunas_saida]


# ============================================================
# Smoke test isolado (só roda se chamado como script direto)
# ============================================================

if __name__ == "__main__":
    import os

    # Tenta localizar o BMA.xlsx no diretório do projeto para teste
    caminho_bma = None
    for possivel in [
        os.path.join(os.path.dirname(__file__), "BMA.xlsx"),
        "BMA.xlsx",
        "/mnt/project/BMA.xlsx",
    ]:
        if os.path.exists(possivel):
            caminho_bma = possivel
            break

    if caminho_bma is None:
        print("BMA.xlsx não encontrado. Pule o smoke test.")
    else:
        print(f"Carregando {caminho_bma}...")
        xl = pd.ExcelFile(caminho_bma)
        df_plamov = pd.read_excel(xl, sheet_name="PLAMOV COMPILADO")
        # Renomeia SUBDIVISAO -> PROJETO para alinhar ao código principal
        if "SUBDIVISAO" in df_plamov.columns and "PROJETO" not in df_plamov.columns:
            df_plamov = df_plamov.rename(columns={"SUBDIVISAO": "PROJETO"})
        df_TP_BMA = pd.read_excel(xl, sheet_name="RELATÓRIO TP BMA")

        # Filtro Grupo F heurístico apenas para smoke test: voluntários
        # em Loc A que pediram alguma localidade C/B
        if "LOC A" in df_plamov.columns:
            em_loc_a = df_plamov["LOC A"].astype(
                str).str.strip().str.upper() == "A"
            df_plamov = df_plamov[em_loc_a].copy()

        def pediu_cb(row):
            for col in ("LOC 1", "LOC 2", "LOC 3"):
                v = str(row.get(col, "")).strip().upper()
                if v in LOCALIDADES_CB:
                    return True
            return False

        mask_f = df_plamov.apply(pediu_cb, axis=1)
        df_grupo_f = df_plamov[mask_f].copy()

        print(f"Militares no Grupo F (smoke): {len(df_grupo_f)}")
        print(f"Linhas em df_TP_BMA: {len(df_TP_BMA)}")

        ranking = calcular_mcda_voluntarios_cb(df_grupo_f, df_TP_BMA)
        print(f"\nAlternativas geradas: {len(ranking)}")
        if not ranking.empty:
            print("\nTop 10:")
            print(ranking.head(10).to_string(index=False))
