"""
Análise de estoque da lanchonete: demanda, estoque de segurança,
ponto de pedido, curva ABC e gráficos.

Fórmulas usadas
---------------
Demanda média diária (d)      = média do consumo diário
Desvio padrão diário (σ)      = desvio padrão amostral do consumo diário
Estoque de segurança (ES)     = Z × σ × √LT
Ponto de pedido (PP)          = d × LT + ES
Estoque máximo (EMAX)         = PP + d × dias de cobertura
Quantidade sugerida de compra = EMAX − estoque atual (se estoque atual ≤ PP)

Z vem do nível de serviço desejado (95% → Z ≈ 1,645).
"""
from math import ceil
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy.stats import norm

NIVEL_SERVICO = 0.95
DIAS_COBERTURA = 7  # quantos dias de consumo comprar além do ponto de pedido

BASE = Path(__file__).resolve().parents[1]
DATA_DIR = BASE / "data"
IMG_DIR = BASE / "images"

# Paleta (azul, laranja, aqua) e cores de status com rótulo de texto sempre junto
AZUL, LARANJA, AQUA = "#257de0", "#f47230", "#00b894"
STATUS_CORES = {"OK": "#0aac23", "REPOR": "#fbb315", "CRÍTICO": "#dc3c3d"}
TEXTO, TEXTO_2, GRADE = "#1f1f1f", "#6b6b6b", "#e6e6e6"

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10,
    "axes.edgecolor": GRADE, "axes.labelcolor": TEXTO_2,
    "xtick.color": TEXTO_2, "ytick.color": TEXTO_2,
    "axes.spines.top": False, "axes.spines.right": False,
    "axes.grid": True, "grid.color": GRADE, "grid.linewidth": 0.8,
    "axes.axisbelow": True, "figure.dpi": 150,
})


def calcular_indicadores(produtos: pd.DataFrame, vendas: pd.DataFrame) -> pd.DataFrame:
    z = norm.ppf(NIVEL_SERVICO)
    stats = vendas.groupby("codigo")["quantidade"].agg(demanda_media="mean", desvio_padrao="std")
    df = produtos.merge(stats, on="codigo")

    df["estoque_seguranca"] = np.ceil(z * df.desvio_padrao * np.sqrt(df.lead_time_dias))
    df["ponto_pedido"] = np.ceil(df.demanda_media * df.lead_time_dias + df.estoque_seguranca)
    df["estoque_maximo"] = np.ceil(df.ponto_pedido + df.demanda_media * DIAS_COBERTURA)
    df["dias_cobertura"] = df.estoque_atual / df.demanda_media

    df["status"] = np.select(
        [df.estoque_atual <= df.estoque_seguranca, df.estoque_atual <= df.ponto_pedido],
        ["CRÍTICO", "REPOR"], default="OK",
    )
    df["qtd_sugerida"] = np.where(df.status != "OK", df.estoque_maximo - df.estoque_atual, 0)
    df["custo_pedido"] = df.qtd_sugerida * df.custo_unitario

    # Curva ABC pelo valor de consumo mensal (30 dias)
    df["consumo_mensal_rs"] = df.demanda_media * 30 * df.custo_unitario
    df = df.sort_values("consumo_mensal_rs", ascending=False)
    df["pct_acumulado"] = df.consumo_mensal_rs.cumsum() / df.consumo_mensal_rs.sum()
    df["classe_abc"] = np.select([df.pct_acumulado <= 0.80, df.pct_acumulado <= 0.95], ["A", "B"], "C")
    return df


def grafico_custo_diario(produtos, vendas):
    v = vendas.merge(produtos[["codigo", "custo_unitario"]], on="codigo")
    v["custo"] = v.quantidade * v.custo_unitario
    diario = v.groupby("data")["custo"].sum()
    media7 = diario.rolling(7).mean()

    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(diario.index, diario.values, color=AZUL, linewidth=1.2, alpha=0.45)
    ax.plot(media7.index, media7.values, color=AZUL, linewidth=2)
    ax.text(media7.index[-1], media7.values[-1], "  média móvel 7 dias", color=TEXTO, va="center")
    ax.set_title("Custo diário dos insumos consumidos (R$)", loc="left", color=TEXTO, fontweight="bold")
    ax.yaxis.set_major_formatter(lambda x, _: f"R$ {x:,.0f}".replace(",", "."))
    ax.grid(axis="x", visible=False)
    fig.tight_layout()
    fig.savefig(IMG_DIR / "custo_diario.png", bbox_inches="tight")
    plt.close(fig)


def grafico_sazonalidade(vendas):
    v = vendas.copy()
    v["dia_semana"] = pd.to_datetime(v.data).dt.weekday
    # índice: consumo do dia ÷ consumo médio do produto (remove diferença de unidades)
    v["indice"] = v.quantidade / v.groupby("codigo").quantidade.transform("mean")
    idx = v.groupby("dia_semana").indice.mean()
    nomes = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]

    fig, ax = plt.subplots(figsize=(7, 3.8))
    barras = ax.bar(nomes, idx.values, color=AZUL, width=0.6)
    ax.axhline(1, color=TEXTO_2, linewidth=1, linestyle="--")
    for b, val in zip(barras, idx.values):
        ax.text(b.get_x() + b.get_width() / 2, val + 0.02, f"{val:.2f}", ha="center", color=TEXTO, fontsize=9)
    ax.set_title("Índice de consumo por dia da semana (1,00 = média)", loc="left", color=TEXTO, fontweight="bold")
    ax.grid(axis="x", visible=False)
    fig.tight_layout()
    fig.savefig(IMG_DIR / "sazonalidade_semanal.png", bbox_inches="tight")
    plt.close(fig)


def grafico_cobertura(df):
    d = df.sort_values("dias_cobertura")
    pp_dias = d.ponto_pedido / d.demanda_media

    fig, ax = plt.subplots(figsize=(9, 5.5))
    y = np.arange(len(d))
    ax.barh(y, d.dias_cobertura, color=[STATUS_CORES[s] for s in d.status], height=0.6)
    ax.scatter(pp_dias, y, marker="|", s=260, color=TEXTO, linewidths=2.2, zorder=3, label="Ponto de pedido (em dias)")
    for yi, (cob, pp, st) in enumerate(zip(d.dias_cobertura, pp_dias, d.status)):
        ax.text(max(cob, pp) + 0.2, yi, f"{cob:.1f} d · {st}", va="center", fontsize=9, color=TEXTO)
    ax.set_yticks(y, d.produto)
    ax.set_xlabel("Dias de cobertura do estoque atual")
    ax.set_title("Estoque atual × ponto de pedido", loc="left", color=TEXTO, fontweight="bold")
    ax.legend(loc="lower right", frameon=False)
    ax.grid(axis="y", visible=False)
    ax.set_xlim(0, d.dias_cobertura.max() * 1.3)
    fig.tight_layout()
    fig.savefig(IMG_DIR / "cobertura_estoque.png", bbox_inches="tight")
    plt.close(fig)


def grafico_abc(df):
    pct = df.consumo_mensal_rs / df.consumo_mensal_rs.sum() * 100
    cores = {"A": AZUL, "B": LARANJA, "C": AQUA}

    fig, ax = plt.subplots(figsize=(10, 4.8))
    x = np.arange(len(df))
    ax.bar(x, pct, color=[cores[c] for c in df.classe_abc], width=0.65)
    ax.plot(x, df.pct_acumulado * 100, color=TEXTO, linewidth=2, marker="o", markersize=5)
    for limite in (80, 95):
        ax.axhline(limite, color=TEXTO_2, linewidth=0.8, linestyle="--")
        ax.text(len(df) - 0.4, limite + 1, f"{limite}%", color=TEXTO_2, fontsize=8, ha="right")
    ax.set_xticks(x, df.produto, rotation=40, ha="right")
    ax.set_ylim(0, 105)
    ax.set_ylabel("% do valor consumido")
    ax.set_title("Curva ABC: participação no custo mensal (barras) e % acumulado (linha)",
                 loc="left", color=TEXTO, fontweight="bold")
    handles = [plt.Rectangle((0, 0), 1, 1, color=cores[c]) for c in "ABC"]
    ax.legend(handles, ["Classe A", "Classe B", "Classe C"], frameon=False, loc="center right")
    ax.grid(axis="x", visible=False)
    fig.tight_layout()
    fig.savefig(IMG_DIR / "curva_abc.png", bbox_inches="tight")
    plt.close(fig)


if __name__ == "__main__":
    IMG_DIR.mkdir(exist_ok=True)
    produtos = pd.read_csv(DATA_DIR / "produtos.csv")
    vendas = pd.read_csv(DATA_DIR / "vendas_diarias.csv", parse_dates=["data"])

    df = calcular_indicadores(produtos, vendas)
    df.to_csv(DATA_DIR / "indicadores_estoque.csv", index=False)

    grafico_custo_diario(produtos, vendas)
    grafico_sazonalidade(vendas)
    grafico_cobertura(df)
    grafico_abc(df)

    cols = ["produto", "demanda_media", "estoque_seguranca", "ponto_pedido",
            "estoque_atual", "status", "qtd_sugerida", "classe_abc"]
    print(df[cols].round(1).to_string(index=False))
    print(f"\nCusto total do pedido sugerido: R$ {df.custo_pedido.sum():,.2f}")
