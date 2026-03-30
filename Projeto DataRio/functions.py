import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


ARQUIVO = "IndServ_DataRio.xls"
ARQUIVO_SAIDA = "resultado_dadosrio.xlsx"
ANOS_COMPARACAO = [2000, 2006]
COLUNAS_NUMERICAS = ["salarios", "receita_total", "numero_empresas", "pessoal_ocupado"]


def padronizar_colunas(df):
    df = df.copy()

    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace(r"\s+", "_", regex=True)
        .str.lower()
    )

    novos_nomes = {}
    for col in df.columns:
        if "salários" in col or "salarios" in col:
            novos_nomes[col] = "salarios"
        elif "receita_total" in col:
            novos_nomes[col] = "receita_total"
        elif "número_de_empresas" in col or "numero_de_empresas" in col:
            novos_nomes[col] = "numero_empresas"
        elif "pessoal_ocupado" in col:
            novos_nomes[col] = "pessoal_ocupado"

    return df.rename(columns=novos_nomes)


def limpar_aba(df):
    df = padronizar_colunas(df)

    if "atividade" not in df.columns:
        return pd.DataFrame()

    df = df[df["atividade"].notna()].copy()

    if df["atividade"].astype(str).str.contains("Fonte", na=False).any():
        idx_fonte = df[df["atividade"].astype(str).str.contains("Fonte", na=False)].index[0]
        df = df.loc[:idx_fonte - 1].copy()

    df = df.replace(["x", "-", "'-"], pd.NA)

    for col in COLUNAS_NUMERICAS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def processar_planilhas(arquivo):
    planilhas = pd.read_excel(arquivo, sheet_name=None)

    dados_totais = []
    lista_atividades = []
    lista_limpos = []

    for ano, df in planilhas.items():
        try:
            ano_int = int(ano)
        except ValueError:
            continue

        df = limpar_aba(df)

        if df.empty:
            continue

        df["ano"] = ano_int
        lista_limpos.append(df.copy())

        linha_total = df[df["atividade"].astype(str).str.strip().str.lower() == "total"]

        if linha_total.empty:
            continue

        total = linha_total.iloc[0]

        dados_totais.append({
            "ano": ano_int,
            "salarios": total.get("salarios", np.nan),
            "receita_total": total.get("receita_total", np.nan),
            "numero_empresas": total.get("numero_empresas", np.nan),
            "pessoal_ocupado": total.get("pessoal_ocupado", np.nan),
        })

        atividades = df[df["atividade"].astype(str).str.strip().str.lower() != "total"].copy()
        colunas_disponiveis = ["atividade", "ano"] + [
            col for col in COLUNAS_NUMERICAS if col in atividades.columns
        ]
        lista_atividades.append(atividades[colunas_disponiveis])

    df_totais = pd.DataFrame(dados_totais).sort_values("ano").reset_index(drop=True)
    df_atividades = pd.concat(lista_atividades, ignore_index=True)
    df_limpo = pd.concat(lista_limpos, ignore_index=True)

    return df_totais, df_atividades, df_limpo


def criar_comparativo_total(df_totais):
    df = df_totais.copy()

    for col in COLUNAS_NUMERICAS:
        df[col] = df[col] / df[col].iloc[0] * 100

    return df


def criar_comparativo_atividades(df_atividades, anos=(2000, 2006)):
    df = df_atividades[df_atividades["ano"].isin(anos)].copy()

    tabela = df.pivot(
        index="atividade",
        columns="ano",
        values=COLUNAS_NUMERICAS
    )

    tabela = tabela.dropna()

    base = tabela.xs(anos[0], level=1, axis=1).replace(0, np.nan)
    comp = tabela.xs(anos[1], level=1, axis=1)

    crescimento = ((comp - base) / base) * 100
    crescimento = crescimento.apply(pd.to_numeric, errors="coerce")

    return crescimento.sort_index()


def top_3_por_indicador(df_crescimento):
    return {
        "receita_total": df_crescimento["receita_total"].nlargest(3),
        "salarios": df_crescimento["salarios"].nlargest(3),
        "numero_empresas": df_crescimento["numero_empresas"].nlargest(3),
        "pessoal_ocupado": df_crescimento["pessoal_ocupado"].nlargest(3),
    }


def organizar_tops_para_excel(tops):
    lista_tops = []

    for indicador, serie in tops.items():
        df_top = serie.reset_index()
        df_top.columns = ["atividade", "valor"]
        df_top["indicador"] = indicador
        df_top["posicao"] = range(1, len(df_top) + 1)
        lista_tops.append(df_top[["indicador", "posicao", "atividade", "valor"]])

    return pd.concat(lista_tops, ignore_index=True)


def exportar_resultados_excel(
    caminho_saida,
    df_limpo,
    df_totais,
    df_comparativo_total,
    df_comparativo_atividades,
    df_tops
):
    comparativo_atividades_excel = df_comparativo_atividades.reset_index()

    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
        df_limpo.to_excel(writer, sheet_name="limpeza", index=False)
        df_totais.to_excel(writer, sheet_name="totais", index=False)
        df_comparativo_total.to_excel(writer, sheet_name="comparativo_total", index=False)
        comparativo_atividades_excel.to_excel(writer, sheet_name="comparativo_atividades", index=False)
        df_tops.to_excel(writer, sheet_name="tops", index=False)


def plotar_totais(df_comparativo_total):
    ax = df_comparativo_total.plot(x="ano", figsize=(10, 6), marker="o")

    for col in df_comparativo_total.columns[1:]:
        for _, row in df_comparativo_total.iterrows():
            ax.text(row["ano"], row[col], f"{row[col]:.0f}%", fontsize=8)

    plt.title("Evolução percentual dos indicadores totais (base = 100 no primeiro ano)")
    plt.xlabel("Ano")
    plt.ylabel("Índice percentual")
    plt.tight_layout()
    plt.show()