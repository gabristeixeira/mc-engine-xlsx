"""
graficos.py: histograma do VPL e gráfico de tornado a partir da saída do motor.

O motor (mc_engine.py) devolve os dados; este script é uma camada de
apresentação sobre eles. Uso:

    python graficos.py exemplos/CarLangA_Reconstrucao_NI43101.xlsx "FluxoCaixa!C46" 5000

Argumentos: planilha, saída de interesse (Aba!Célula) e número de iterações.
Gera histograma_vpl.png e tornado.png na pasta atual.
"""
import sys
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy.stats import spearmanr
from mc_engine import run_monte_carlo

xlsx, saida = sys.argv[1], sys.argv[2]
n = int(sys.argv[3]) if len(sys.argv) > 3 else 5000

r = run_monte_carlo(xlsx, n_trials=n, seed=42, compute_tir=False, verbose=False)
vpl = r["results"][saida]
p10, p50, p90 = np.percentile(vpl, [10, 50, 90])

# Histograma com P10, P50 e P90
fig, ax = plt.subplots(figsize=(8, 4.5))
ax.hist(vpl, bins=50, color="#8FAADC", edgecolor="white")
for p, nome in ((p10, "P10"), (p50, "P50"), (p90, "P90")):
    ax.axvline(p, color="#1F3864", linestyle="--", linewidth=1)
    ax.text(p, ax.get_ylim()[1] * 0.95, f"{nome}\n{p:,.0f}", ha="center", va="top", fontsize=8)
ax.set_xlabel("VPL")
ax.set_ylabel("Frequência")
ax.set_title(f"Distribuição simulada do VPL (N = {n:,})")
fig.tight_layout()
fig.savefig("histograma_vpl.png", dpi=200)

# Tornado: correlação de postos de Spearman entre cada entrada e o VPL
nomes, rhos = [], []
for chave, valores in r["samples"].items():
    rho = spearmanr(valores, vpl)[0]
    desc = next((i["desc"] for i in r["inputs"] if f"{i['aba']}!{i['celula']}" == chave), chave)
    nomes.append(desc)
    rhos.append(rho)
ordem = np.argsort(np.abs(rhos))
fig, ax = plt.subplots(figsize=(8, 3.5))
ax.barh([nomes[i] for i in ordem], [rhos[i] for i in ordem], color="#1F3864")
ax.axvline(0, color="black", linewidth=0.8)
ax.set_xlabel("Correlação de Spearman com o VPL")
ax.set_title("Sensibilidade dos drivers (tornado)")
fig.tight_layout()
fig.savefig("tornado.png", dpi=200)

print(f"P10/P50/P90: {p10:,.1f} / {p50:,.1f} / {p90:,.1f}")
for i in ordem[::-1]:
    print(f"  {nomes[i]}: {rhos[i]:+.2f}")
