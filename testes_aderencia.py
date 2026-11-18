"""
testes_aderencia.py: testes formais de aderência do módulo de amostragem do motor
(Camada 2 da arquitetura de validação, seção 4.2.2 do texto principal).

Gera 200.000 amostras de cada distribuição declarada na aba MC_Config do Projeto
Serra Alta, usando a mesma função de amostragem do motor, e aplica o teste de
Kolmogorov-Smirnov de uma amostra (todas as distribuições) e o de Anderson-Darling
(distribuições normais) contra a distribuição teórica.

Uso:
    python3 testes_aderencia.py
Dependências: numpy, scipy, mc_engine.py no mesmo diretório.
"""
import numpy as np
from scipy import stats
from mc_engine import sample_distribution

N = 200_000
ESPECIFICACOES = [
    ("Choque preço do níquel", "Normal", 0.0, 0.08, None),
    ("Choque câmbio", "Normal", 0.0, 0.04, None),
    ("Choque opex", "Triangular", -0.05, 0.0, 0.10),
    ("Choque teor", "PERT", 0.90, 1.00, 1.08),
]


def cdf_teorica(dist, a, b, c):
    if dist == "Normal":
        return lambda v: stats.norm.cdf(v, a, b)
    if dist == "Triangular":
        return lambda v: stats.triang.cdf(v, (b - a) / (c - a), loc=a, scale=c - a)
    alpha = 1 + 4 * (b - a) / (c - a)
    beta = 1 + 4 * (c - b) / (c - a)
    return lambda v: stats.beta.cdf((v - a) / (c - a), alpha, beta)


if __name__ == "__main__":
    rng = np.random.default_rng(2026)
    print(f"Valor crítico KS (5%) para N={N}: {1.358 / np.sqrt(N):.5f}")
    for nome, dist, a, b, c in ESPECIFICACOES:
        x = sample_distribution(dist, a, b, c, N, rng)
        ks = stats.kstest(x, cdf_teorica(dist, a, b, c))
        linha = f"{nome:24s} {dist:11s} KS D={ks.statistic:.5f} p={ks.pvalue:.3f}"
        if dist == "Normal":
            ad = stats.anderson(x, "norm")
            linha += f"  AD={ad.statistic:.3f} (crítico 5%: {ad.critical_values[2]:.3f})"
        print(linha)
