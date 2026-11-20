# Extração e análise de dados do SIH/SUS – Implante de marcapasso

Repositório destinado à disponibilização dos scripts utilizados na extração, organização, processamento e análise dos dados do Sistema de Informações Hospitalares do Sistema Único de Saúde (SIH/SUS), para o projeto de iniciação científica sobre as internações para implante de marcapasso cardíaco transvenoso de dupla câmara no Brasil, no período de 2016 a 2025.

O repositório permite a reprodução das etapas de extração dos microdados, obtenção dos denominadores populacionais e processamento das bases para geração dos indicadores, tabelas, gráficos e análises de tendência.

## 1. Estrutura do repositório

| Arquivo | Descrição |
|---|---|
| `extrator-ftp-SIH-SUS.py` | Script responsável pela extração dos microdados das Autorizações de Internação Hospitalar (AIH-RD) do SIH/SUS, diretamente do FTP do DATASUS, filtrando os registros pelo código do procedimento estudado. |
| `extrator_ibge.py` | Script responsável pela obtenção e organização dos dados populacionais do IBGE, por ano, Grande Região e grupo etário, utilizados no cálculo das taxas de internação. |
| `modelador.py` | Script responsável pelo processamento, agregação, correção monetária, cálculo dos indicadores, análise de tendência pelo método Prais-Winsten, geração de gráficos e exportação dos resultados. |
| `marcapasso_bruto.csv` | Base de dados extraída pelo script `extrator-ftp-SIH-SUS.py`. |
| `ibge_populacao.csv` | Base populacional produzida pelo script `extrator_ibge.py`. |
| `resultados.xlsx` | Planilha de resultados gerada pelo script `modelador.py`. |

## 2. Requisitos

Para executar os scripts, é necessário ter instalado:

- Python 3.12;
- conexão com a internet, para acesso aos arquivos do DATASUS, IBGE e, quando disponível, Banco Central.

As bibliotecas Python utilizadas incluem:

- pandas;
- numpy;
- scipy;
- matplotlib;
- openpyxl;
- datasus-dbc.

## 3. Instalação das dependências

Após instalar o Python 3.12, abra o terminal ou Prompt de Comando e execute:

```bash
py -3.12 -m pip install pandas numpy scipy matplotlib openpyxl datasus-dbc
```

## 4. Ordem de execução

**Os scripts devem ser executados na seguinte ordem:**

### Etapa 1 – Extração dos microdados do SIH/SUS

Execute:

```bash
py -3.12 extrator-ftp-SIH-SUS.py
```

O script acessa o FTP oficial do DATASUS, identifica os arquivos de dados de internação disponíveis nas competências de 2015 a 2026 e extrai os registros correspondentes ao código de procedimento `0406010650`.

A inclusão das competências de 2015 e 2026 permite contemplar internações ocorridas entre 2016 e 2025, considerando o processamento das AIHs em competências anteriores ou posteriores ao ano da internação.

Os registros extraídos são reunidos e exportados para o arquivo `marcapasso_bruto.csv`, que será utilizado na etapa de processamento.

O script também utiliza arquivos em cache para reaproveitar extrações já processadas e permitir a retomada do procedimento em caso de interrupção.

### Etapa 2 – Extração dos dados populacionais

Execute:

```bash
py -3.12 extrator_ibge.py
```

O script obtém os dados da publicação *Projeções da População do Brasil e Unidades da Federação – Revisão 2024*, do IBGE, com base no Censo Demográfico de 2022.

São organizados os denominadores populacionais referentes à população geral e à população de 60 anos ou mais, por ano e Grande Região, entre 2016 e 2025.

As populações estaduais são agregadas em Grandes Regiões, e os dados são exportados para `ibge_populacao.csv`.

### Etapa 3 – Processamento e análise dos dados

Execute:

```bash
py -3.12 modelador.py
```

Este script utiliza os arquivos `marcapasso_bruto.csv` e `ibge_populacao.csv`, gerados nas etapas anteriores, para processar e agregar os dados, calcular os indicadores e executar as análises previstas no projeto.

A execução gera uma pasta denominada `resultados`, contendo a planilha de resultados, os gráficos e os arquivos destinados à reprodução das análises estatísticas no Gretl.

**A execução do `modelador.py` depende da conclusão das duas etapas anteriores.**

## 5. Produtos gerados

Ao final da execução do `modelador.py`, será criada a pasta `resultados`, contendo:

- `resultados.xlsx`: planilha com as tabelas de internações, taxas, valores, análises de tendência e gráficos;
- `graficos/`: diretório contendo os gráficos gerados pelo processamento;
- `gretl/`: diretório contendo os arquivos de dados e scripts necessários para reprodução das análises no Gretl;
- `prais_winsten.csv`: tabela com os resultados das análises de tendência pelo método Prais-Winsten.

## 6. Fontes dos dados

- **DATASUS – SIH/SUS:** microdados das Autorizações de Internação Hospitalar (AIH-RD), disponibilizados no FTP oficial do DATASUS.
- **IBGE:** Projeções da População do Brasil e Unidades da Federação – Revisão 2024, com base no Censo Demográfico de 2022.
- **Banco Central do Brasil:** série do Índice Nacional de Preços ao Consumidor Amplo (IPCA), utilizada para correção monetária dos valores, quando disponível para obtenção automatizada.

## 7. Reprodutibilidade

Os scripts disponibilizados neste repositório documentam os procedimentos computacionais utilizados na extração e no processamento dos dados, possibilitando a reprodução das etapas de obtenção das bases e geração dos resultados analíticos do projeto.
