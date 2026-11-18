"""
mc_engine.py: motor genérico de simulação de Monte Carlo para modelos financeiros em .xlsx

Ideia central (TCC de Gabriel Santos Teixeira, UFMG Controladoria e Finanças):
Ferramenta nativa em Python, alternativa gratuita ao @RISK, que roda sobre planilhas .xlsx
que adote a convenção "MC_Config" (ver aba MC_Config do Modelo_Proprio_Serra_Alta_v7.xlsx para um
exemplo de referência), sem precisar do Excel/LibreOffice instalado, sem precisar alterar as fórmulas
originais da planilha, e sem precisar de um modelo substituto/aproximação (o motor de fórmulas do
pycel é rápido o suficiente para reavaliar a árvore de dependências completa a cada iteração).

Uso:
    from mc_engine import run_monte_carlo
    resultado = run_monte_carlo("Modelo_Proprio_Serra_Alta_v7.xlsx", n_trials=10000, seed=42)

O motor lê a aba "MC_Config" da planilha (bloco INPUTS: quais células recebem qual distribuição de
probabilidade; bloco OUTPUTS: quais células de resultado monitorar) e não depende de nenhum
conhecimento hardcoded sobre o Projeto Serra Alta especificamente.
"""

import datetime
import time
import numpy as np
from openpyxl import load_workbook
from pycel import ExcelCompiler
from scipy.optimize import brentq

EXCEL_EPOCH = datetime.date(1899, 12, 30)


# ------------------------------------------------------------
# Leitura da convenção MC_Config
# ------------------------------------------------------------
def read_mc_config(xlsx_path):
    """Lê a aba MC_Config (convenção genérica) e devolve (inputs, outputs)."""
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    if "MC_Config" not in wb.sheetnames:
        raise ValueError(
            f"'{xlsx_path}' não tem uma aba 'MC_Config'. A ferramenta só processa planilhas que "
            "adotem essa convenção (ver Modelo_Proprio_Serra_Alta_v7.xlsx como referência)."
        )
    ws = wb["MC_Config"]

    # bloco INPUTS: comeca na linha do header "Aba" seguinte a secao "INPUTS"
    inputs, outputs = [], []
    mode = None
    for row in ws.iter_rows(min_row=1, values_only=False):
        vals = [c.value for c in row]
        if not any(vals):
            continue
        first = vals[0]
        if isinstance(first, str) and first.startswith("INPUTS"):
            mode = "inputs_header_next"
            continue
        if isinstance(first, str) and first.startswith("OUTPUTS"):
            mode = "outputs_header_next"
            continue
        if mode == "inputs_header_next":
            mode = "inputs"
            continue  # essa linha e o cabecalho (Aba, Celula, Descricao, ...)
        if mode == "outputs_header_next":
            mode = "outputs"
            continue
        if mode == "inputs" and vals[0] and vals[1]:
            aba, cel, desc, dist = vals[0], vals[1], vals[2], vals[3]
            p1, p2, p3 = vals[4], vals[5], vals[6]
            inputs.append(dict(aba=aba, celula=cel, desc=desc, dist=dist, p1=p1, p2=p2, p3=p3))
        elif mode == "outputs" and vals[0] and vals[1]:
            aba, cel, desc, tipo = vals[0], vals[1], vals[2], vals[3]
            outputs.append(dict(aba=aba, celula=cel, desc=desc, tipo=tipo))
    return inputs, outputs


# ------------------------------------------------------------
# Validação de consistência da convenção MC_Config
# ------------------------------------------------------------
def validate_mc_config(inputs, outputs):
    """
    Verifica a consistência interna da aba MC_Config antes de iniciar a simulação.
    Não verifica o sentido econômico dos parâmetros, apenas se são tecnicamente válidos:
    distribuição suportada, parâmetros numéricos presentes, ordem coerente entre mínimo,
    moda e máximo, desvio-padrão positivo e tipos de saída reconhecidos.
    Levanta ValueError listando todos os problemas encontrados de uma só vez.
    """
    problemas = []
    if not inputs:
        problemas.append("Nenhuma variável de entrada declarada no bloco INPUTS.")
    if not outputs:
        problemas.append("Nenhuma variável de saída declarada no bloco OUTPUTS.")
    tipos_validos = {"direto", "serie_fcl", "serie_datas"}
    for out in outputs:
        if out["tipo"] not in tipos_validos:
            problemas.append(f"Saída {out['aba']}!{out['celula']}: tipo '{out['tipo']}' desconhecido "
                             f"(válidos: {sorted(tipos_validos)}).")
    tipos = {o["tipo"] for o in outputs}
    if "serie_fcl" in tipos and "serie_datas" not in tipos:
        problemas.append("Saída 'serie_fcl' declarada sem a respectiva 'serie_datas' (necessária para a TIR).")
    for inp in inputs:
        nome = f"Entrada {inp['aba']}!{inp['celula']}"
        dist = (inp["dist"] or "").strip().lower()
        p1, p2, p3 = inp["p1"], inp["p2"], inp["p3"]
        if dist not in ("normal", "uniforme", "triangular", "pert"):
            problemas.append(f"{nome}: distribuição '{inp['dist']}' não suportada "
                             "(suportadas: Normal, Triangular, PERT, Uniforme).")
            continue
        n_params = 3 if dist in ("triangular", "pert") else 2
        params = [p1, p2, p3][:n_params]
        if any(not isinstance(p, (int, float)) for p in params):
            problemas.append(f"{nome}: parâmetros ausentes ou não numéricos para {inp['dist']}.")
            continue
        if dist == "normal" and p2 <= 0:
            problemas.append(f"{nome}: desvio-padrão da Normal deve ser positivo (recebido {p2}).")
        if dist == "uniforme" and p1 >= p2:
            problemas.append(f"{nome}: Uniforme exige mínimo < máximo (recebido {p1}, {p2}).")
        if dist in ("triangular", "pert") and not (p1 <= p2 <= p3 and p1 < p3):
            problemas.append(f"{nome}: {inp['dist']} exige mínimo <= moda <= máximo, com mínimo < máximo "
                             f"(recebido {p1}, {p2}, {p3}).")
    if problemas:
        raise ValueError("MC_Config inconsistente:\n  - " + "\n  - ".join(problemas))


# ------------------------------------------------------------
# Correlação entre variáveis de entrada (cópula gaussiana, opcional)
# ------------------------------------------------------------
def read_correlations(xlsx_path):
    """
    Lê o bloco opcional CORRELACOES da aba MC_Config. Cada linha após o cabeçalho tem:
    Aba1, Celula1, Aba2, Celula2, Coeficiente. Devolve {("Aba1!Cel1", "Aba2!Cel2"): rho}.
    Se o bloco não existir, devolve {} e a simulação segue com entradas independentes.
    """
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    if "MC_Config" not in wb.sheetnames:
        return {}
    corr, mode = {}, None
    for row in wb["MC_Config"].iter_rows(min_row=1, values_only=True):
        if not any(row):
            continue
        first = row[0]
        if isinstance(first, str) and first.startswith("CORRELACOES"):
            mode = "header_next"
            continue
        if isinstance(first, str) and (first.startswith("INPUTS") or first.startswith("OUTPUTS")):
            mode = None
            continue
        if mode == "header_next":
            mode = "corr"
            continue
        if mode == "corr" and row[0] and row[1] and row[2] and row[3] is not None:
            corr[(f"{row[0]}!{row[1]}", f"{row[2]}!{row[3]}")] = row[4]
    return corr


def _correlation_matrix(keys, correlations):
    """Monta e valida a matriz de correlação (simétrica, |rho| < 1, positiva definida)."""
    k = len(keys)
    mat = np.eye(k)
    idx = {key: i for i, key in enumerate(keys)}
    problemas = []
    for (a, b), rho in correlations.items():
        if a not in idx or b not in idx:
            problemas.append(f"Correlação {a} x {b}: variável não declarada em INPUTS.")
            continue
        if a == b:
            problemas.append(f"Correlação {a} x {b}: uma variável não pode se correlacionar consigo mesma.")
            continue
        if not isinstance(rho, (int, float)) or not (-1 < rho < 1):
            problemas.append(f"Correlação {a} x {b}: coeficiente deve estar entre -1 e 1, exclusive (recebido {rho}).")
            continue
        mat[idx[a], idx[b]] = mat[idx[b], idx[a]] = rho
    if not problemas and np.min(np.linalg.eigvalsh(mat)) <= 1e-10:
        problemas.append("A matriz de correlação não é positiva definida (coeficientes mutuamente incompatíveis).")
    if problemas:
        raise ValueError("Correlações inconsistentes:\n  - " + "\n  - ".join(problemas))
    return mat


def _ppf(dist, u, p1, p2, p3):
    """Função quantil (inversa da acumulada) das distribuições suportadas."""
    from scipy import stats
    dist = (dist or "").strip().lower()
    if dist == "normal":
        return stats.norm.ppf(u, loc=p1, scale=p2)
    if dist == "uniforme":
        return p1 + (p2 - p1) * u
    if dist == "triangular":
        return stats.triang.ppf(u, (p2 - p1) / (p3 - p1), loc=p1, scale=p3 - p1)
    if dist == "pert":
        if p3 <= p1:
            return np.full_like(u, p2)
        alpha = 1 + 4 * (p2 - p1) / (p3 - p1)
        beta = 1 + 4 * (p3 - p2) / (p3 - p1)
        return p1 + stats.beta.ppf(u, alpha, beta) * (p3 - p1)
    raise ValueError(f"Distribuição não suportada: '{dist}'")


def sample_correlated(inputs, correlations, n, rng):
    """
    Amostra as entradas com dependência pela cópula gaussiana: gera normais padrão
    correlacionadas (decomposição de Cholesky), converte em probabilidades uniformes e
    aplica a função quantil de cada distribuição marginal. As marginais declaradas no
    MC_Config são preservadas; o coeficiente informado é o da cópula (correlação de postos
    aproximada), e não necessariamente o de Pearson das variáveis amostradas.
    """
    from scipy import stats
    keys = [f"{i['aba']}!{i['celula']}" for i in inputs]
    mat = _correlation_matrix(keys, correlations)
    z = rng.standard_normal((n, len(keys))) @ np.linalg.cholesky(mat).T
    u = stats.norm.cdf(z)
    return {key: _ppf(i["dist"], u[:, j], i["p1"], i["p2"], i["p3"])
            for j, (key, i) in enumerate(zip(keys, inputs))}


# ------------------------------------------------------------
# Amostragem de distribuições (suporte: Normal, Triangular, PERT, Uniforme)
# ------------------------------------------------------------
def sample_distribution(dist, p1, p2, p3, size, rng):
    dist = (dist or "").strip().lower()
    if dist == "normal":
        return rng.normal(loc=p1, scale=p2, size=size)
    if dist == "uniforme":
        return rng.uniform(low=p1, high=p2, size=size)
    if dist == "triangular":
        return rng.triangular(left=p1, mode=p2, right=p3, size=size)
    if dist == "pert":
        # PERT modificada (lambda=4): beta reescalada em [min, max], com moda ponderada 4x
        mn, mo, mx = p1, p2, p3
        if mx <= mn:
            return np.full(size, mo)
        alpha = 1 + 4 * (mo - mn) / (mx - mn)
        beta = 1 + 4 * (mx - mo) / (mx - mn)
        return mn + rng.beta(alpha, beta, size=size) * (mx - mn)
    raise ValueError(f"Distribuição não suportada: '{dist}' (suportadas: Normal, Triangular, PERT, Uniforme)")


# ------------------------------------------------------------
# XIRR próprio (a pycel não implementa funções financeiras como XIRR, IRR, XNPV e PMT)
# Validado contra o XIRR do LibreOffice no modelo v6: bate em 14 casas decimais.
# ------------------------------------------------------------
def to_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    return EXCEL_EPOCH + datetime.timedelta(days=int(v))


def xirr(cashflows, dates):
    """
    Calcula a TIR (XIRR) por busca de raiz sobre VPL(taxa) = 0.

    Correção (Camada 2, validação estatística): a versão original usava um único bracket
    [-0.9999, 10] para o brentq. Isso falhava (~47% de NaN numa amostra de 5.000 cenários
    simulados) sempre que a curva VPL(taxa) tinha MAIS de uma mudança de sinal dentro desse
    intervalo, o que é comum e legítimo em fluxo de caixa de projeto de mineração com fase
    de construção (negativo), operação (positivo) e encerramento de mina (negativo de novo),
    um problema clássico de "TIR múltipla" (não-convencionalidade do fluxo de caixa). Um
    único bracket amplo não garante achar a raiz economicamente relevante (a mais próxima de
    uma taxa de desconto plausível), e em casos extremos (taxa perto de -100%) o cálculo da
    própria função VPL(taxa) fica numericamente instável (overflow), quebrando a premissa de
    sinais opostos nas pontas do bracket exigida pelo brentq.

    Correção aplicada: varre a curva VPL(taxa) em passos de 1%, primeiro no intervalo
    economicamente plausível [0%, 300%] e só depois, se nada for encontrado, em [-99%, 0%),
    procurando a PRIMEIRA mudança de sinal a partir de taxa zero, replicando o comportamento
    do XIRR do Excel, que converge para a raiz mais próxima de uma estimativa inicial (guess),
    por padrão 10%, em vez de aceitar qualquer raiz dentro de um bracket arbitrariamente amplo.
    Validado: no caso-base (sem choque), o resultado bate com o bracket original em 1e-16 (erro
    de ponto flutuante); nos ~47% de cenários que antes davam NaN, agora todos convergem para
    valores de TIR plausíveis (20%-35%), consistentes com o caso-base.
    """
    d0 = dates[0]
    def npv(rate):
        return sum(cf / (1 + rate) ** ((d - d0).days / 365.0) for cf, d in zip(cashflows, dates))
    for lo, hi in [(0.0, 3.0), (-0.99, 0.0)]:
        rates = np.arange(lo, hi, 0.01)
        prev_r, prev_v = rates[0], npv(rates[0])
        for r in rates[1:]:
            v = npv(r)
            if np.isfinite(prev_v) and np.isfinite(v) and prev_v * v < 0:
                try:
                    return brentq(npv, prev_r, r)
                except ValueError:
                    pass
            prev_r, prev_v = r, v
    return float("nan")  # nenhuma mudança de sinal em [-99%, 300%], TIR de fato indefinida


# ------------------------------------------------------------
# Motor principal
# ------------------------------------------------------------
def run_monte_carlo(xlsx_path, n_trials=2000, seed=None, compute_tir=True, tir_sample=None, verbose=True,
                    correlations=None):
    """
    Roda a simulação de Monte Carlo sobre xlsx_path usando a convenção MC_Config.

    n_trials: numero de iteracoes para as saidas "diretas" (ex.: VPL), ~55-60 iter/s no
              Projeto Serra Alta (varia com a complexidade de formulas de cada planilha; no
              CarLang A, estrutura mais enxuta, chega a ~295 iter/s).
    compute_tir: se True, tambem calcula TIR (XIRR) via series de fluxo de caixa, mais lento
                 por precisar ler ~160 celulas por iteracao (uma por coluna de fluxo de caixa
                 e de data), entao roda numa subamostra (tir_sample) em vez das n_trials
                 completas, por padrao.
    tir_sample: numero de iteracoes para a TIR (default: min(n_trials, 500)).
    correlations: dict opcional {("Aba!Cel1", "Aba!Cel2"): rho} com a correlação entre entradas
                  (cópula gaussiana). Se None, usa o bloco CORRELACOES da MC_Config, se existir;
                  sem nenhum dos dois, as entradas são amostradas de forma independente.
    """
    inputs, outputs = read_mc_config(xlsx_path)
    validate_mc_config(inputs, outputs)
    if verbose:
        print(f"MC_Config: {len(inputs)} inputs, {len(outputs)} outputs")
        for inp in inputs:
            print(f"  INPUT  {inp['aba']}!{inp['celula']:6s} {inp['dist']:10s} "
                  f"({inp['p1']}, {inp['p2']}, {inp['p3']}) : {inp['desc']}")
        for out in outputs:
            print(f"  OUTPUT {out['aba']}!{out['celula']:12s} [{out['tipo']}]: {out['desc']}")

    direct_outputs = [o for o in outputs if o["tipo"] == "direto"]
    fcl_output = next((o for o in outputs if o["tipo"] == "serie_fcl"), None)
    dates_output = next((o for o in outputs if o["tipo"] == "serie_datas"), None)

    rng = np.random.default_rng(seed)
    if correlations is None:
        correlations = read_correlations(xlsx_path)
    if correlations:
        samples = sample_correlated(inputs, correlations, n_trials, rng)
        if verbose:
            for (a, b), rho in correlations.items():
                print(f"  CORRELACAO {a} x {b}: {rho}")
    else:
        samples = {
            f"{inp['aba']}!{inp['celula']}": sample_distribution(inp["dist"], inp["p1"], inp["p2"], inp["p3"], n_trials, rng)
            for inp in inputs
        }

    excel = ExcelCompiler(filename=xlsx_path)
    results = {f"{o['aba']}!{o['celula']}": np.empty(n_trials) for o in direct_outputs}

    # "aquece" o mapa de celulas do pycel: so depois de avaliar uma celula (e toda a arvore de
    # dependencias dela) e que o pycel sabe que as celulas de input existem e aceita set_value nelas.
    for o in direct_outputs:
        excel.evaluate(f"{o['aba']}!{o['celula']}")
    if fcl_output:
        for c in _expand_range(fcl_output["celula"])[:1]:
            excel.evaluate(f"{fcl_output['aba']}!{c}")
    if dates_output:
        for c in _expand_range(dates_output["celula"])[:1]:
            excel.evaluate(f"{dates_output['aba']}!{c}")

    t0 = time.time()
    for i in range(n_trials):
        for key, arr in samples.items():
            excel.set_value(key, float(arr[i]))
        for o in direct_outputs:
            key = f"{o['aba']}!{o['celula']}"
            results[key][i] = excel.evaluate(key)
    dt = time.time() - t0
    if verbose:
        print(f"{n_trials} iterações (saídas diretas) em {dt:.1f}s ({n_trials/dt:.1f} iter/s)")

    tir_values = None
    if compute_tir and fcl_output and dates_output:
        tir_sample = tir_sample or min(n_trials, 500)
        aba = fcl_output["aba"]
        fcl_range = _expand_range(fcl_output["celula"])
        dates_range = _expand_range(dates_output["celula"])
        tir_values = np.empty(tir_sample)
        t0 = time.time()
        for i in range(tir_sample):
            for key, arr in samples.items():
                excel.set_value(key, float(arr[i]))
            fcl = [excel.evaluate(f"{aba}!{c}") for c in fcl_range]
            dts = [to_date(excel.evaluate(f"{aba}!{c}")) for c in dates_range]
            tir_values[i] = xirr(fcl, dts)
        dt = time.time() - t0
        if verbose:
            print(f"{tir_sample} iterações (TIR/XIRR) em {dt:.1f}s ({tir_sample/dt:.2f} iter/s)")

    return dict(inputs=inputs, outputs=outputs, samples=samples, results=results, tir_values=tir_values)


def _expand_range(range_str):
    """Expande 'C86:CF86' (ou similar) numa lista de referências de coluna+linha."""
    from openpyxl.utils.cell import range_boundaries
    from openpyxl.utils import get_column_letter
    min_col, min_row, max_col, max_row = range_boundaries(range_str)
    if min_row == max_row:
        return [f"{get_column_letter(c)}{min_row}" for c in range(min_col, max_col + 1)]
    return [f"{get_column_letter(min_col)}{r}" for r in range(min_row, max_row + 1)]


if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else "Modelo_Proprio_Serra_Alta_v7.xlsx"
    n = int(sys.argv[2]) if len(sys.argv) > 2 else 2000
    out = run_monte_carlo(path, n_trials=n, seed=42)
    vpl = list(out["results"].values())[0]
    print()
    print("=== RESULTADO ===")
    print(f"VPL médio: {vpl.mean():,.1f}  desvio-padrão: {vpl.std():,.1f}")
    print(f"VPL P10/P50/P90: {np.percentile(vpl,10):,.1f} / {np.percentile(vpl,50):,.1f} / {np.percentile(vpl,90):,.1f}")
    print(f"Probabilidade VPL > 0: {(vpl > 0).mean()*100:.2f}%")
    if out["tir_values"] is not None:
        tir = out["tir_values"]
        tir = tir[~np.isnan(tir)]
        print(f"TIR média: {tir.mean()*100:.2f}%  P10/P50/P90: {np.percentile(tir,10)*100:.2f}% / "
              f"{np.percentile(tir,50)*100:.2f}% / {np.percentile(tir,90)*100:.2f}%")
