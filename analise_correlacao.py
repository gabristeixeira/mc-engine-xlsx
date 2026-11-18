"""
analise_correlacao.py: análise complementar do efeito da correlação entre o preço do níquel
e o câmbio sobre o VPL do Projeto Serra Alta (seção 4.4.2 do texto principal).

1) Estima a correlação entre as variações mensais (log) do preço do níquel e do câmbio
   USD/BRL a partir de Historico_FX_Niquel_out2019_dez2022.csv, com intervalo bootstrap.
2) Roda o motor com a cópula gaussiana para coeficientes: 0 (independência), limites do
   intervalo e estimativa pontual, com a mesma semente e o mesmo número de iterações.

Uso: python3 analise_correlacao.py   (cerca de 6 minutos)
"""
import json
import sys
import numpy as np
import pandas as pd
from scipy import stats
from mc_engine import run_monte_carlo

XLSX = "Modelo_Proprio_Serra_Alta_v7.xlsx"
NI, FX = "Assumptions!C42", "Assumptions!C43"
N, SEED = 5000, 42

df = pd.read_csv("Historico_FX_Niquel_out2019_dez2022.csv")
r = np.log(df[["fx_usdbrl_media", "niquel_usd_tonne"]]).diff().dropna()
rho_hat = float(np.corrcoef(r.iloc[:, 0], r.iloc[:, 1])[0, 1])
rng = np.random.default_rng(1)
boot = []
for _ in range(5000):
    i = rng.integers(0, len(r), len(r))
    boot.append(np.corrcoef(r.iloc[i, 0], r.iloc[i, 1])[0, 1])
lo, hi = np.percentile(boot, [2.5, 97.5])
print(f"rho estimado (n={len(r)}): {rho_hat:.3f}  IC95% bootstrap: [{lo:.3f}, {hi:.3f}]")

saida = {"n_obs": len(r), "rho": rho_hat, "ic_inf": float(lo), "ic_sup": float(hi), "cenarios": {}}
CENARIOS = [("independente", 0.0), ("limite_superior", float(hi)), ("estimativa", rho_hat), ("limite_inferior", float(lo))]
# Uso opcional: python3 analise_correlacao.py <nome_do_cenario> executa um cenário por vez
# (e acumula em resultado_correlacao.json); sem argumento, executa todos em sequência.
escolhidos = [c for c in CENARIOS if len(sys.argv) < 2 or c[0] == sys.argv[1]]
try:
    saida = json.load(open("resultado_correlacao.json"))
    saida.update({"n_obs": len(r), "rho": rho_hat, "ic_inf": float(lo), "ic_sup": float(hi)})
except FileNotFoundError:
    pass
for nome, rho in escolhidos:
    out = run_monte_carlo(XLSX, n_trials=N, seed=SEED, compute_tir=False, verbose=False,
                          correlations={(NI, FX): rho})
    v = list(out["results"].values())[0]
    saida["cenarios"][nome] = dict(
        rho=rho, media=float(v.mean()), desvio=float(v.std()),
        p10=float(np.percentile(v, 10)), p50=float(np.percentile(v, 50)),
        p90=float(np.percentile(v, 90)), prob_pos=float((v > 0).mean()))
    print(nome, saida["cenarios"][nome])
    json.dump(saida, open("resultado_correlacao.json", "w"), indent=1)
