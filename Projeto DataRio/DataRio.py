from functions import *


def main():
    df_totais, df_atividades, df_limpo = processar_planilhas(ARQUIVO)

    df_comparativo_total = criar_comparativo_total(df_totais)
    df_comparativo_atividades = criar_comparativo_atividades(
        df_atividades,
        anos=ANOS_COMPARACAO
    )

    tops = top_3_por_indicador(df_comparativo_atividades)
    df_tops = organizar_tops_para_excel(tops)

    exportar_resultados_excel(
        ARQUIVO_SAIDA,
        df_limpo,
        df_totais,
        df_comparativo_total,
        df_comparativo_atividades,
        df_tops
    )

    plotar_totais(df_comparativo_total)


main()