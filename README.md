# Dashboard de Vendas — Guia Explicativo

Este repositório contém um **pipeline completo** de análise e visualização de vendas de três lojas (A, B e C), incluindo um **diagnóstico de causas** para quedas/picos mensais. O material foi preparado para uso em **sala de aula de Analistas de Dados**.

---

## 1) O que o projeto faz

1. **Carrega** um Excel com três abas (`Loja_A`, `Loja_B`, `Loja_C`).
2. **Padroniza** nomes de colunas, trata tipos e converte o campo de mês (`Mes` / `Mês`) para `datetime`.
3. **Agrega** vendas e preço médio por mês para todo o portfólio.
4. **Gera visualizações** interativas Plotly:
   - Vendas totais por mês (todas as lojas).
   - Preço médio mensal (média entre lojas).
   - Comparativo por loja com **eixo duplo** (vendas e preço médio).
   - **Lucro estimado** por loja (modelo simples).
5. **Diagnóstico de causas** (análise explicativa):
   - **Correlação** de vendas com fatores (preço, estoque, concorrência, faltas e marketing).
   - **Atribuição de contribuição** mês versus mês anterior usando **regressão linear** (por exemplo **Set/2020** e **Nov/2022**).

---

## 2) Como executar

> Pré‑requisitos: Python 3.10+ e as bibliotecas `pandas`, `plotly` e `numpy`.

No Windows (PowerShell):

```powershell
python .\dashboard_vendas.py --excel ".\dataset_causa_raiz_multiplas_lojas_DISCREPANTE.xlsx" --saida_html ".\dashboard_vendas.html"
start .\dashboard_vendas.html
```

---

## 3) Estrutura dos dados esperada (Excel)

- Abas: `Loja_A`, `Loja_B`, `Loja_C`
- Colunas principais (exemplos; nomes são **normalizados** pelo script):
  - `Mes` (ou `Mês`) — mês no formato `Jan/2020` etc.
  - `Vendas` (unidades)
  - `Preco_Medio` (R$)
  - `Concorrencia_Promocoes`, `Faltas_Func`, `Investimento_Marketing`, `Estoque_%` (percentual)

> O script cria a coluna **`Loja`** automaticamente a partir do nome da aba.

---

## 4) Pipeline de preparação

### 4.1 Carregamento e padronização
- Lemos apenas as abas esperadas e empilhamos tudo em um único `DataFrame`.
- Padronizamos nomes com acentuação e garantimos que `Vendas`, `Preco_Medio` e `Estoque_Perc` sejam **numéricos**.
- Convertemos `Mes/Mês` em `Data` com uma função robusta que entende abreviações em PT‑BR (ex.: `Set` → `Sep`).

### 4.2 Agregação mensal
- `vendas_por_data`: soma de `Vendas` e média de `Preco_Medio` por mês (todos os canais).
- Identificação do **mês de maior venda** para exibir um resumo no terminal.

---

## 5) Visualizações

### 5.1 Vendas totais (todas as lojas)
Gráfico de barras por mês com escala de cores pela intensidade de vendas.

**Leitura sugerida:** observe _picos sazonais_ (por exemplo, fim de ano) e eventuais **quedas atípicas**.

### 5.2 Preço médio
Série temporal de `Preco_Medio` mensal (média entre lojas). Útil para inspecionar movimentos de preço e avaliar se há **relação inversa** com vendas.

### 5.3 Comparativo por loja + preço médio (eixo duplo)
Barras agrupadas por loja (unidades vendidas) com uma linha de **preço médio** no **eixo secundário**. Ajuda a comparar desempenho relativo e o nível de preço no mesmo período.

### 5.4 Lucro estimado por loja
Modelo simples:
- `Estoque_unid_est = (Estoque_Perc / 100) * Vendas`
- `Lucro_Estimado ≈ Preco_Medio * (Vendas - Estoque_unid_est)`

> **Atenção:** é uma aproximação didática — não considera impostos, descontos, custos fixos, etc.

---

## 6) Diagnóstico de causas (explicativo)

### 6.1 Correlação (Pearson)
Calculamos a correlação de `Vendas` com cada fator (`Preco_Medio`, `Concorrencia_Promocoes`, `Faltas_Func`, `Investimento_Marketing`, `Estoque_Perc`). Isso aponta **associações lineares** (não causalidade).

### 6.2 Atribuição de contribuição (regressão OLS)
Para um mês **alvo**, estimamos _quanto cada fator contribuiu_ para a variação de vendas **vs. mês anterior**:

1. Ajustamos um modelo linear:  
   `Vendas_t ≈ β0 + β1*Preco_Medio_t + β2*Concorrencia_Promocoes_t + β3*Faltas_Func_t + β4*Investimento_Marketing_t + β5*Estoque_Perc_t`
2. Calculamos `Δfator = fator_alvo − fator_mês_anterior`.
3. A **contribuição estimada** do fator = `β * Δfator`.
4. Exibimos as contribuições ordenadas por **impacto absoluto**.

> Em sala de aula, discuta limitações: multicolinearidade, variável omitida e o fato de que **correlação ≠ causalidade**.

---

## 7) Principais achados (dados fornecidos)

- O pico de vendas do período ocorreu em **Dez/2020** com **174.590** unidades.
- Dezembro também foi forte em **2021** (≈ **150.691**) e **2022** (≈ **128.370**), sugerindo **sazonalidade de fim de ano**.
- O módulo de diagnóstico facilita investigar quedas, como em **Set/2020**, decompondo a variação de vendas por contribuição de **preço**, **estoque**, **concorrência**, **faltas** e **marketing**.

> Use o gráfico de contribuições para discutir hipóteses: por exemplo, “alta de preço” (β negativo * Δpreço positivo) tende a puxar as vendas para baixo; já **marketing** e **estoque saudável** costumam contribuir positivamente (dependendo dos dados).

---

## 8) Como expandir (exercícios)

1. **Adicionar custo unitário** e margem para melhorar o modelo de lucro.
2. Testar **defasagens** (ex.: marketing do mês anterior) e/ou **variáveis dummies** sazonais.
3. Aplicar **regularização** (Ridge/Lasso) para mitigar multicolinearidade.
4. Construir um **painel** no Streamlit a partir das mesmas funções.
5. Validar o modelo com **treino/validação** (métrica: MAPE/RMSE).

---

## 9) Arquivos importantes

- `dashboard_vendas.py` — pipeline, métricas, gráficos e diagnóstico.
- `dataset_causa_raiz_multiplas_lojas_DISCREPANTE.xlsx` — dados de exemplo (três abas, uma por loja).
- `dashboard_vendas.html` — dashboard gerado (interativo).

---

## 10) Créditos e licença Willian Batista Oliveira

Material educacional para prática de **Análise Exploratória** e **Explicabilidade** em séries temporais curtas. Use, modifique e compartilhe livremente com crédito ao autor do notebook/código base.
