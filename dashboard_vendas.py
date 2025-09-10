# dashboard_vendas.py
import argparse
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path


def carregar_dados(caminho_excel: str) -> pd.DataFrame:
    """
    Lê as abas Loja_A, Loja_B, Loja_C de um Excel e unifica em um único DataFrame.
    Adiciona a coluna 'Loja' e converte 'Mês' -> 'Data' (datetime).
    """
    xls = pd.ExcelFile(caminho_excel)
    esperadas = ["Loja_A", "Loja_B", "Loja_C"]
    presentes = [s for s in esperadas if s in xls.sheet_names]
    if not presentes:
        raise ValueError(
            f"Nenhuma aba esperada encontrada em {xls.sheet_names}. "
            f"Esperado: {esperadas}"
        )

    frames = []
    for nome in presentes:
        df = pd.read_excel(caminho_excel, sheet_name=nome)
        df["Loja"] = nome.split("_")[-1]  # "A", "B" ou "C"
        frames.append(df)

    df_all = pd.concat(frames, ignore_index=True)

    # Normaliza nomes (evita erros por acentuação)
    colmap = {
        "Mês": "Mes",
        "Preço_Médio": "Preco_Medio",
        "Concorrencia_Promocoes": "Concorrencia_Promocoes",
        "Faltas_Func": "Faltas_Func",
        "Investimento_Marketing": "Investimento_Marketing",
        "Estoque_%": "Estoque_Perc",
        "Temperatura": "Temperatura",
        "Vendas": "Vendas",
    }
    df_all = df_all.rename(columns=colmap)

    # Garante tipos numéricos e trata faltas
    for col in ["Vendas", "Preco_Medio", "Estoque_Perc"]:
        if col not in df_all.columns:
            df_all[col] = 0
        df_all[col] = pd.to_numeric(df_all[col], errors="coerce").fillna(0)

    # Converte “Mes” (ex.: Jan/2020) para datetime
    def parse_mes(val):
        if pd.isna(val):
            return pd.NaT
        s = str(val).strip()
        # tenta PT direto
        try:
            return pd.to_datetime(s, format="%b/%Y", dayfirst=False)
        except Exception:
            pass
        # mapeia abreviações PT->EN e tenta de novo
        mapa = {"Fev": "Feb", "Abr": "Apr", "Mai": "May", "Ago": "Aug",
                "Set": "Sep", "Out": "Oct", "Dez": "Dec"}
        s_en = s
        for k, v in mapa.items():
            s_en = s_en.replace(k, v)
        for fmt in ("%b/%Y", "%b-%Y", "%b %Y"):
            try:
                return pd.to_datetime(s_en, format=fmt, dayfirst=False)
            except Exception:
                continue
        return pd.NaT

    if "Mes" in df_all.columns:
        df_all["Data"] = df_all["Mes"].map(parse_mes)
    elif "Mês" in df_all.columns:  # fallback, se não tiver sido renomeado
        df_all["Data"] = df_all["Mês"].map(parse_mes)
    else:
        raise KeyError("Coluna de mês não encontrada. Esperado 'Mês' ou 'Mes'.")

    return df_all



def sintetizar_metricas(df_all: pd.DataFrame) -> tuple[pd.DataFrame, pd.Series]:
    """
    Retorna:
      vendas_por_data: DataFrame com Vendas_Total e Preco_Medio por Data
      linha_max: linha (Series) com a data de maior Vendas_Total
    """
    vendas_por_data = (
        df_all.groupby("Data")
        .agg(
            Vendas_Total=("Vendas", "sum"),
            Preco_Medio=("Preco_Medio", "mean"),
        )
        .reset_index()
        .sort_values("Data")
    )
    linha_max = vendas_por_data.loc[vendas_por_data["Vendas_Total"].idxmax()]
    return vendas_por_data, linha_max


import numpy as np

def _prep_mensal(df_all: pd.DataFrame) -> pd.DataFrame:
    """Agrupa por mês (Data) somando Vendas e fazendo média dos demais fatores."""
    # garante colunas
    for col in ["Vendas","Preco_Medio","Concorrencia_Promocoes","Faltas_Func","Investimento_Marketing","Estoque_Perc"]:
        if col not in df_all.columns:
            df_all[col] = 0
        df_all[col] = pd.to_numeric(df_all[col], errors="coerce").fillna(0)

    meses = (
        df_all
        .groupby("Data")
        .agg(
            Vendas=("Vendas","sum"),
            Preco_Medio=("Preco_Medio","mean"),
            Concorrencia_Promocoes=("Concorrencia_Promocoes","mean"),
            Faltas_Func=("Faltas_Func","mean"),
            Investimento_Marketing=("Investimento_Marketing","mean"),
            Estoque_Perc=("Estoque_Perc","mean"),
        )
        .sort_index()
    )
    # Variação mês a mês (em %)
    meses["Vendas_MoM_%"] = meses["Vendas"].pct_change()*100

    # z-scores dos fatores (padronização)
    for col in ["Preco_Medio","Concorrencia_Promocoes","Faltas_Func","Investimento_Marketing","Estoque_Perc"]:
        std = meses[col].std(ddof=0)
        meses[f"{col}_z"] = (meses[col]-meses[col].mean())/std if std and std > 0 else 0
    return meses


def diagnostico_mes(df_all: pd.DataFrame, ano: int, mes: int):
    """
    Retorna:
      - meses: dataframe mensal com métricas, variação e z-scores
      - corr_vendas: correlações de 'Vendas' com cada fator
      - contrib_mes: contribuições estimadas dos fatores para ΔVendas do mês escolhido vs mês anterior
    """
    meses = _prep_mensal(df_all)

    # correlações simples com Vendas
    corr_vendas = (
        meses[["Vendas","Preco_Medio","Concorrencia_Promocoes","Faltas_Func","Investimento_Marketing","Estoque_Perc"]]
        .corr()
        .loc["Vendas"]
        .drop("Vendas")
        .sort_values(ascending=False)
    )

    # máscara do mês escolhido
    mask = (meses.index.month==mes) & (meses.index.year==ano)
    if not mask.any():
        return meses, corr_vendas, None  # não existe esse mês/ano nos dados

    idx_target = meses.index[mask][0]
    # mês anterior
    loc_idx = meses.index.get_loc(idx_target)
    if loc_idx == 0:
        return meses, corr_vendas, None
    idx_prev = meses.index[loc_idx-1]

    # regressão linear: Vendas ~ 1 + (5 fatores)
    fatores = ["Preco_Medio","Concorrencia_Promocoes","Faltas_Func","Investimento_Marketing","Estoque_Perc"]
    X = meses[fatores].values
    X = np.c_[np.ones(len(X)), X]  # intercepto
    y = meses["Vendas"].values
    beta, *_ = np.linalg.lstsq(X, y, rcond=None)

    # contribuição aproximada no Δ
    dx = (meses.loc[idx_target, fatores] - meses.loc[idx_prev, fatores])
    contrib = pd.Series(beta[1:], index=fatores) * dx
    contrib = contrib.sort_values(key=lambda s: s.abs(), ascending=False)

    contrib_mes = pd.DataFrame({
        "Δ fator (Mês - Anterior)": dx[contrib.index],
        "Coef_β": beta[1:][np.array([fatores.index(f) for f in contrib.index])],
        "Contrib_Estimada": contrib
    })
    delta_real = meses.loc[idx_target, "Vendas"] - meses.loc[idx_prev, "Vendas"]
    contrib_mes.loc["__TOTAL__","Contrib_Estimada"] = contrib.sum()
    contrib_mes.loc["__TOTAL__","Δ fator (Mês - Anterior)"] = delta_real
    return meses, corr_vendas, contrib_mes


def diagnostico_setembro(df_all: pd.DataFrame):
    # Setembro/2020
    return diagnostico_mes(df_all, ano=2020, mes=9)


def figuras_causas(df_all: pd.DataFrame) -> dict:
    """Cria figs extras: heatmap de correlação e barras de contribuição Setembro."""
    figs_extra = {}
    meses, corr_vendas, contrib_set = diagnostico_mes(df_all, ano=2020, mes=9)

    # Heatmap de correlação (Vendas x fatores)
    corr_df = corr_vendas.to_frame(name="Correlação").reset_index().rename(columns={"index":"Fator"})
    figs_extra["correlacoes_vendas"] = px.bar(
        corr_df, x="Fator", y="Correlação",
        title="Correlação de Vendas com Fatores",
        labels={"Fator":"Fator","Correlação":"Correlação de Pearson"},
    )
    figs_extra["correlacoes_vendas"].update_layout(yaxis=dict(tickformat=".2f"))

    # Contribuições para Setembro (se existir)
    if contrib_set is not None and "__TOTAL__" in contrib_set.index:
        plot_df = contrib_set.drop(index="__TOTAL__").reset_index().rename(columns={"index":"Fator"})
        plot_df = plot_df.sort_values("Contrib_Estimada", key=lambda s: s.abs(), ascending=False)
        figs_extra["contrib_setembro"] = px.bar(
            plot_df, x="Fator", y="Contrib_Estimada",
            title="Contribuição Estimada dos Fatores — Setembro vs Agosto/2020",
            labels={"Fator":"Fator","Contrib_Estimada":"Δ Vendas Estimada"},
        )
        figs_extra["contrib_setembro"].update_traces(
            hovertemplate="Fator=%{x}<br>Δ fator=%{customdata[0]:.2f}<br>β=%{customdata[1]:.3f}<br>Contribuição=%{y:.0f}<extra></extra>",
            customdata=plot_df[["Δ fator (Mês - Anterior)","Coef_β"]].values
        )

    return figs_extra


def figuras_causas_2022(df_all: pd.DataFrame) -> dict:
    figs_extra = {}
    meses, corr_vendas, contrib_nov22 = diagnostico_mes(df_all, ano=2022, mes=11)

    # Correlação geral (mesmo gráfico que antes)
    corr_df = corr_vendas.to_frame(name="Correlação").reset_index().rename(columns={"index":"Fator"})
    figs_extra["correlacoes_vendas_2022"] = px.bar(
        corr_df, x="Fator", y="Correlação",
        title="Correlação de Vendas com Fatores (Ano 2022)",
        labels={"Fator":"Fator","Correlação":"Correlação de Pearson"},
    )
    figs_extra["correlacoes_vendas_2022"].update_layout(yaxis=dict(tickformat=".2f"))

    # Contribuições para Novembro/2022 (se houver)
    if contrib_nov22 is not None and "__TOTAL__" in contrib_nov22.index:
        plot_df = contrib_nov22.drop(index="__TOTAL__").reset_index().rename(columns={"index":"Fator"})
        plot_df = plot_df.sort_values("Contrib_Estimada", key=lambda s: s.abs(), ascending=False)
        figs_extra["contrib_novembro_2022"] = px.bar(
            plot_df, x="Fator", y="Contrib_Estimada",
            title="Contribuição Estimada dos Fatores — Novembro vs Outubro/2022",
            labels={"Fator":"Fator","Contrib_Estimada":"Δ Vendas Estimada"},
        )
        figs_extra["contrib_novembro_2022"].update_traces(
            hovertemplate="Fator=%{x}<br>Δ fator=%{customdata[0]:.2f}<br>β=%{customdata[1]:.3f}<br>Contribuição=%{y:.0f}<extra></extra>",
            customdata=plot_df[["Δ fator (Mês - Anterior)","Coef_β"]].values
        )
    return figs_extra


def figuras_plotly(df_all: pd.DataFrame, vendas_por_data: pd.DataFrame) -> dict:
    figs = {}

    # 1) Vendas totais (todas as lojas)
    figs["vendas_total"] = px.bar(
        vendas_por_data,
        x="Data",
        y="Vendas_Total",
        title="Histórico de Vendas Totais (Todas as Lojas)",
        labels={"Data": "Período", "Vendas_Total": "Unidades Vendidas"},
        color="Vendas_Total",
        color_continuous_scale="Blues",
    )
    figs["vendas_total"].update_layout(
        xaxis=dict(tickformat="%b/%Y"),
        yaxis=dict(tickformat="~s")  # 12k, 150k...
    )

    # 2) Preço médio (média entre lojas)
    figs["preco_medio"] = px.line(
        vendas_por_data,
        x="Data",
        y="Preco_Medio",
        title="Preço Médio ao Longo do Tempo (Média entre Lojas)",
        labels={"Data": "Período", "Preco_Medio": "Preço Médio (R$)"},
        markers=True,
    )
    figs["preco_medio"].update_layout(
        xaxis=dict(tickformat="%b/%Y"),
        yaxis=dict(tickprefix="R$ ", tickformat=".2f")
    )

    # 3) Comparativo por loja + preço médio (eixo duplo)
    df_group = (
        df_all.groupby(["Data", "Loja"])
        .agg(Vendas=("Vendas", "sum"), Preco_Medio=("Preco_Medio", "mean"))
        .reset_index()
        .sort_values("Data")
    )

    fig_combined = px.bar(
        df_group,
        x="Data",
        y="Vendas",
        color="Loja",
        barmode="group",
        title="Histórico de Vendas por Loja",
        labels={"Data": "Período", "Vendas": "Unidades Vendidas"},
    )
    fig_combined.add_trace(
        go.Scatter(
            x=df_group["Data"],
            y=df_group["Preco_Medio"],
            mode="lines+markers",
            name="Preço Médio",
            yaxis="y2",
        )
    )
    fig_combined.update_layout(
        xaxis=dict(tickformat="%b/%Y"),
        yaxis=dict(title="Unidades Vendidas", tickformat="~s"),
        yaxis2=dict(title="Preço Médio (R$)", overlaying="y", side="right", tickprefix="R$ ", tickformat=".2f"),
        legend=dict(x=0.01, y=0.99, bordercolor="Black", borderwidth=1),
    )
    figs["comparativo_lojas"] = fig_combined

    # 4) Lucro estimado por loja
    # Fórmula: Lucro ≈ Preço_Médio × (Vendas − Estoque_unid_est)
    # onde Estoque_unid_est = (Estoque_Perc/100) × Vendas
    df_lucro = df_all.copy()
    df_lucro["Estoque_unid_est"] = (df_lucro["Estoque_Perc"] / 100.0) * df_lucro["Vendas"]
    df_lucro["Receita"] = df_lucro["Preco_Medio"] * df_lucro["Vendas"]
    df_lucro["Custo_Estoque"] = df_lucro["Preco_Medio"] * df_lucro["Estoque_unid_est"]
    df_lucro["Lucro_Estimado"] = df_lucro["Receita"] - df_lucro["Custo_Estoque"]

    lucro_por_loja = (
        df_lucro.groupby(["Data", "Loja"])
        .agg(Lucro_Estimado=("Lucro_Estimado", "sum"))
        .reset_index()
        .sort_values(["Loja", "Data"])
    )

    figs["lucro_estimado"] = px.line(
        lucro_por_loja,
        x="Data",
        y="Lucro_Estimado",
        color="Loja",
        markers=True,
        title="Lucro Estimado por Loja ao Longo do Tempo",
        labels={"Data": "Período", "Lucro_Estimado": "Lucro Estimado (R$)", "Loja": "Loja"},
    )
    figs["lucro_estimado"].update_layout(
        xaxis=dict(tickformat="%b/%Y"),
        yaxis=dict(tickprefix="R$ ", tickformat="~s")
    )
    figs["lucro_estimado"].update_traces(
        hovertemplate="Loja=%{legendgroup}<br>Período=%{x|%b/%Y}<br>Lucro Estimado=R$ %{y:,.2f}<extra></extra>"
    )


    # 5) Causas: correlações e (se disponível) explicação de Setembro
    figs_extra = figuras_causas(df_all)
    figs.update(figs_extra)

    # 6) Diagnóstico específico para Novembro/2022
    figs_extra_2022 = figuras_causas_2022(df_all)
    figs.update(figs_extra_2022)


    return figs



def salvar_html(figs: dict, saida_html: Path):
    from plotly.offline import plot

    ordem = [
        "vendas_total",
        "preco_medio",
        "comparativo_lojas",
        "lucro_estimado",
        "correlacoes_vendas",
        "contrib_setembro",
        "correlacoes_vendas_2022",
        "contrib_novembro_2022",
    ]



    titulos = {
        "vendas_total": "Vendas Totais",
        "preco_medio": "Preço Médio",
        "comparativo_lojas": "Comparativo entre Lojas",
        "lucro_estimado": "Lucro Estimado por Loja",
    }
    parts = []
    for key in ordem:
        if key in figs:
            fig = figs[key]
            parts.append(f"<h2 style='font-family:Arial'>{titulos.get(key, key)}</h2>")
            parts.append(plot(fig, include_plotlyjs="cdn", output_type="div"))

    html = f"""
        <html>
        <head>
        <meta charset='utf-8'>
        <title>Dashboard Vendas</title>
        </head>
        <body style="max-width:1100px;margin:24px auto;padding:0 16px;">
        <h1 style="font-family:Arial;margin-bottom:8px;">Dashboard de Vendas</h1>
        <p style="font-family:Arial;color:#444;margin-top:0">
            Análise por: Willian Batista Oliveira
        </p>

        <!-- Botão de Download -->
        <a href="imagens/Analise_Willian_Batista_Oliveira.pdf" download
            style="display:inline-block;padding:10px 20px;
                    background:#007BFF;color:#fff;
                    text-decoration:none;border-radius:6px;
                    font-family:Arial;">
            Baixar a Análise em PDF
        </a>

        {' '.join(parts)}
        </body>
        </html>
        """
    saida_html.write_text(html, encoding="utf-8")




def main():
    parser = argparse.ArgumentParser(
        description="Gera dashboard de vendas a partir de um Excel com abas Loja_A, Loja_B, Loja_C."
    )
    parser.add_argument(
        "--excel",
        type=str,
        required=True,
        help="Caminho para o arquivo Excel.",
    )
    parser.add_argument(
        "--saida_html",
        type=str,
        default="dashboard_vendas.html",
        help="Caminho do HTML de saída.",
    )
    args = parser.parse_args()

    df_all = carregar_dados(args.excel)
    vendas_por_data, linha_max = sintetizar_metricas(df_all)

    # Resumo no console
    data_max = pd.to_datetime(linha_max["Data"]).strftime("%b/%Y")
    print(
        f"Data com maior número de vendas: {data_max} | "
        f"Unidades: {int(linha_max['Vendas_Total'])} | "
        f"Preço médio: R$ {linha_max['Preco_Medio']:.2f}"
    )

    figs = figuras_plotly(df_all, vendas_por_data)
    salvar_html(figs, Path(args.saida_html))


if __name__ == "__main__":
    main()
