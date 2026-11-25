"""
Gera dados SIMULADOS de vendas diárias de insumos de uma lanchonete.

Os dados são fictícios, criados para fins de estudo. A simulação inclui:
- sazonalidade semanal (sexta e sábado vendem mais, segunda vende menos);
- uma leve tendência de crescimento ao longo do período;
- ruído aleatório (distribuição de Poisson), para que haja variabilidade
  e o estoque de segurança faça sentido.

Saídas:
- data/produtos.csv        -> cadastro dos insumos
- data/vendas_diarias.csv  -> consumo diário por insumo (formato longo)
"""
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 42
DIAS = 90
DATA_FINAL = "2026-11-24"

BASE = Path(__file__).resolve().parents[1]
DATA_DIR = BASE / "data"

# código, produto, categoria, unidade, custo unitário (R$), lead time (dias),
# fornecedor, consumo médio diário (base), estoque atual
PRODUTOS = [
    ("P01", "Pão de hambúrguer", "Panificação", "un", 0.90, 1, "Padaria Central", 85, 300),
    ("P02", "Hambúrguer bovino 120g", "Carnes", "un", 3.20, 3, "Frigorífico Boi Bom", 80, 380),
    ("P03", "Queijo mussarela fatiado", "Laticínios", "kg", 42.00, 2, "Laticínios Serra", 3, 12),
    ("P04", "Presunto fatiado", "Frios", "kg", 32.00, 2, "Laticínios Serra", 2, 12),
    ("P05", "Bacon fatiado", "Carnes", "kg", 48.00, 3, "Frigorífico Boi Bom", 2, 3),
    ("P06", "Alface americana", "Hortifrúti", "pé", 3.50, 1, "Hortifrúti Verde", 10, 25),
    ("P07", "Tomate", "Hortifrúti", "kg", 7.00, 1, "Hortifrúti Verde", 4, 3),
    ("P08", "Batata pré-frita congelada", "Congelados", "kg", 14.00, 4, "Distribuidora Frios SP", 12, 60),
    ("P09", "Óleo de soja", "Mercearia", "L", 8.50, 5, "Atacadão Alimentos", 3, 25),
    ("P10", "Refrigerante lata 350ml", "Bebidas", "un", 2.80, 4, "Distribuidora Bebidas", 60, 280),
    ("P11", "Suco natural (polpa)", "Bebidas", "un", 4.50, 3, "Polpas Brasil", 18, 45),
    ("P12", "Embalagem para lanche", "Descartáveis", "un", 0.35, 7, "Embalagens Pack", 95, 900),
]

# multiplicador por dia da semana (0 = segunda ... 6 = domingo)
SAZONALIDADE = {0: 0.75, 1: 0.85, 2: 0.95, 3: 1.00, 4: 1.25, 5: 1.40, 6: 1.10}


def gerar() -> tuple[pd.DataFrame, pd.DataFrame]:
    rng = np.random.default_rng(SEED)
    datas = pd.date_range(end=DATA_FINAL, periods=DIAS, freq="D")

    produtos = pd.DataFrame(
        PRODUTOS,
        columns=["codigo", "produto", "categoria", "unidade", "custo_unitario",
                 "lead_time_dias", "fornecedor", "consumo_base", "estoque_atual"],
    )

    tendencia = np.linspace(0.95, 1.05, DIAS)  # crescimento leve de ~10% no período
    fator_dia = np.array([SAZONALIDADE[d.weekday()] for d in datas])

    linhas = []
    for _, p in produtos.iterrows():
        media = p.consumo_base * fator_dia * tendencia
        if p.unidade in ("kg", "L"):
            # itens a granel: consumo contínuo com 1 casa decimal
            qtd = np.round(rng.gamma(shape=12, scale=media / 12), 1)
        else:
            qtd = rng.poisson(media)
        linhas.append(pd.DataFrame({"data": datas, "codigo": p.codigo, "quantidade": qtd}))

    vendas = pd.concat(linhas, ignore_index=True)
    return produtos.drop(columns="consumo_base"), vendas


if __name__ == "__main__":
    DATA_DIR.mkdir(exist_ok=True)
    produtos, vendas = gerar()
    produtos.to_csv(DATA_DIR / "produtos.csv", index=False)
    vendas.to_csv(DATA_DIR / "vendas_diarias.csv", index=False)
    print(f"{len(produtos)} produtos e {len(vendas)} registros de vendas gerados em {DATA_DIR}")
