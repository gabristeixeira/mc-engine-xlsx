# mc-engine-xlsx

Motor de simulação de Monte Carlo, em Python, para modelos financeiros construídos em planilha (.xlsx). Ele lê as fórmulas do arquivo, sorteia as variáveis incertas declaradas em uma aba de configuração (`MC_Config`) e recalcula o modelo milhares de vezes, sem alterar nenhuma fórmula e sem precisar do Excel instalado.

*Python engine that runs Monte Carlo simulation directly on spreadsheet financial models (.xlsx), driven by an in-sheet `MC_Config` configuration tab. No Excel needed.*

Material do Trabalho de Conclusão de Curso "Proposta de framework metodológico para integração nativa da Simulação de Monte Carlo em modelos financeiros de avaliação de projetos de mineração" (Controladoria e Finanças, UFMG, 2026). Autor: Gabriel Santos Teixeira. Orientador: Prof. Dr. Marcos Antônio de Camargos. Coorientador: MSc. César Augusto Camargos Rocha.

## O que há neste repositório

| Arquivo | O que é |
|---|---|
| `mc_engine.py` | O motor de simulação (código completo, comentado). |
| `graficos.py` | Gera o histograma do VPL e o gráfico de tornado a partir da saída do motor. |
| `testes_aderencia.py` | Testes de aderência da amostragem às distribuições declaradas (Kolmogorov-Smirnov e Anderson-Darling), Camada 2 da validação. |
| `Historico_FX_Niquel_out2019_dez2022.xlsx` | Série mensal de câmbio USD/BRL e preço do níquel (out/2019 a dez/2022), formatada para abrir no Excel. O mesmo conteúdo está no `.csv`, que é o arquivo lido pelo script. |
| `analise_correlacao.py` | Estimativa da correlação entre preço do níquel e câmbio (dados em `Historico_FX_Niquel_out2019_dez2022.csv`) e seu efeito no VPL. Requer o modelo do Projeto Serra Alta, que não é público. |
| `exemplos/CarLangA_Reconstrucao_NI43101.xlsx` | Reconstrução do projeto CarLang A a partir do relatório técnico NI 43-101 (SRK; Caracle Creek, 2025), usada como teste de robustez do motor. |

O modelo financeiro do Projeto Serra Alta não é distribuído, pois deriva de dados de uma empresa privada.

## Instalação

```
pip install -r requirements.txt
```

## Uso rápido

Pela linha de comando (imprime o resumo no terminal):

```
python mc_engine.py exemplos/CarLangA_Reconstrucao_NI43101.xlsx 2000
```

Resultado esperado (CarLang A, 2.000 iterações, semente 42): VPL médio C$ 1.203,8 milhões; desvio-padrão 903,3; P10/P50/P90 de 31,9 / 1.219,0 / 2.368,5; probabilidade de VPL negativo de 9,6%.

Dentro do seu código (devolve os dados completos):

```python
from mc_engine import run_monte_carlo
saida = run_monte_carlo("exemplos/CarLangA_Reconstrucao_NI43101.xlsx", n_trials=2000, seed=42)
vpl = saida["results"]["FluxoCaixa!C46"]   # vetor com o VPL de cada iteração
```

## Como preparar a sua planilha

O modelo precisa ter uma aba `MC_Config`, organizada em blocos identificados pelo texto do marcador na primeira coluna (a linha seguinte ao marcador é o cabeçalho; linhas em branco são ignoradas; cada bloco termina no marcador seguinte).

| Bloco | Colunas, na ordem |
|---|---|
| `INPUTS` | Aba, Célula, Descrição, Distribuição, Param1, Param2, Param3 |
| `OUTPUTS` | Aba, Célula, Descrição, Tipo |
| `CORRELACOES` (opcional) | Aba1, Célula1, Aba2, Célula2, Coeficiente |

**Distribuições aceitas**

| Distribuição | Param1 | Param2 | Param3 | Condição verificada |
|---|---|---|---|---|
| Normal | média | desvio-padrão | não usado | desvio-padrão positivo |
| Uniforme | mínimo | máximo | não usado | mínimo menor que máximo |
| Triangular | mínimo | moda | máximo | mínimo ≤ moda ≤ máximo e mínimo < máximo |
| PERT | mínimo | moda | máximo | mesma condição da Triangular (λ = 4) |

**Tipos de saída**

| Tipo | Significado |
|---|---|
| `direto` | célula única cujo valor é coletado a cada iteração (por exemplo, o VPL) |
| `serie_fcl` | intervalo com a série de fluxos de caixa, usada para calcular a TIR a cada iteração |
| `serie_datas` | intervalo com as datas correspondentes à série de fluxos de caixa |

Para a TIR, declare `serie_fcl` e `serie_datas` juntas, com intervalos de mesmo tamanho.

**Regras do modelo**

1. Cada célula de entrada declarada deve conter um valor constante (não uma fórmula) e ser usada, direta ou indiretamente, pelas fórmulas que levam às saídas. Recomenda-se uma alavanca de choque relativa ao valor-base (0 para nenhum choque percentual, ou 1 para um fator multiplicativo neutro).
2. Os parâmetros das distribuições devem ser números digitados na aba, com ao menos uma entrada e uma saída.
3. As fórmulas só podem usar funções da `pycel` (por exemplo SUM, SUMPRODUCT, IF, INDEX, MATCH, COUNTIF, NPV). IRR, XIRR, XNPV e PMT não são implementadas, e macros VBA não são executadas; a TIR é calculada pelo próprio motor.
4. No bloco `CORRELACOES`, cada coeficiente deve estar entre -1 e 1 (exclusive), referir-se a variáveis declaradas em `INPUTS`, e o conjunto deve formar matriz positiva definida. A dependência é introduzida por cópula gaussiana, preservando as distribuições declaradas. Sem o bloco, as entradas são independentes.

## A função `run_monte_carlo`

```python
run_monte_carlo(xlsx_path, n_trials=2000, seed=None, compute_tir=True,
                tir_sample=None, verbose=True, correlations=None)
```

| Parâmetro | Significado |
|---|---|
| `n_trials` | número de iterações das saídas diretas (padrão: 2.000) |
| `seed` | semente do gerador aleatório; fixá-la torna o resultado reprodutível |
| `compute_tir` | se verdadeiro, calcula também a TIR quando há `serie_fcl` e `serie_datas` |
| `tir_sample` | número de iterações usadas para a TIR (padrão: o menor entre `n_trials` e 500) |
| `correlations` | dicionário opcional com as correlações; se omitido, usa o bloco `CORRELACOES` |
| `verbose` | se verdadeiro, imprime a configuração lida e o tempo de execução |

Devolve um dicionário com `inputs` e `outputs` (a configuração lida), `samples` (valores sorteados de cada entrada), `results` (um vetor por saída direta, indexado por `"Aba!Célula"`) e `tir_values`. Média, desvio-padrão, percentis e probabilidade de VPL positivo são calculados sobre esses vetores. Pela linha de comando, a semente é fixa em 42.

Tempo de execução: cerca de 55 a 60 iterações por segundo no Projeto Serra Alta e cerca de 295 no CarLang A.

## MC_Config da reconstrução do CarLang A

Planilha `exemplos/CarLangA_Reconstrucao_NI43101.xlsx`.

| Aba | Célula | Descrição | Distribuição | Param1 | Param2 | Param3 |
|---|---|---|---|---|---|---|
| Assumptions | C26 | Choque preço do níquel / NSR (%) | Triangular | -0,30 | 0 | 0,30 |
| Assumptions | C28 | Choque opex (%) | Triangular | -0,20 | 0 | 0,20 |
| Assumptions | C29 | Choque capex (%) | Triangular | -0,20 | 0 | 0,30 |

Saída: `FluxoCaixa!C46`, VPL pós-impostos a 8% (C$ milhões), tipo `direto`.

## Mensagens de erro mais comuns

| Mensagem | Causa | Solução |
|---|---|---|
| não tem uma aba 'MC_Config' | a aba não existe ou tem outro nome | criar ou renomear a aba para `MC_Config` |
| MC_Config inconsistente | distribuição não suportada, parâmetros ausentes ou fora de ordem, desvio-padrão não positivo, tipo de saída desconhecido | corrigir os itens listados na própria mensagem |
| AssertionError: Address "Aba!Cel" not found in the cell map | a célula de entrada não é usada por nenhuma fórmula que leve às saídas | ligar a alavanca ao cálculo ou removê-la do bloco `INPUTS` |
| Function X is not implemented (UnknownFunction) | a planilha usa uma função que a pycel não implementa (IRR, XIRR, XNPV, PMT) | substituir por fórmula equivalente; para a TIR, declarar `serie_fcl` e `serie_datas` |
| Correlações inconsistentes | coeficiente fora de (-1, 1), variável não declarada ou matriz não positiva definida | revisar o bloco `CORRELACOES` |

## Licença

MIT. Veja o arquivo `LICENSE`.
