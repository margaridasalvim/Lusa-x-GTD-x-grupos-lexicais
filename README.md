# Lusa × GTD — Análise de títulos e terrorismo

Aplicação desenvolvida no âmbito de uma dissertação de Mestrado em Jornalismo para análise de títulos da agência Lusa e comparação com dados do Global Terrorism Database (GTD).

## Sobre o projeto

A aplicação permite explorar um corpus de títulos da Lusa relacionado com terrorismo, radicalismo, extremismo e rebelião, bem como analisar a sua distribuição temporal, geográfica e temática.

A aplicação utiliza uma base de dados DuckDB preparada durante o processo de investigação.

## Funcionalidades

- análise temporal dos títulos;
- distribuição geográfica;
- identificação e análise de referências a terrorismo;
- classificação relacionada com radicalismo e extremismo;
- análise de conteúdo político;
- análise de tipos de ação;
- comparação entre dados da Lusa e acontecimentos registados no Global Terrorism Database;
- filtros interativos para exploração do corpus.

## Dados

A aplicação online utiliza uma base de dados DuckDB (analysis.duckdb) que contém os dados estruturados necessários ao funcionamento da aplicação.

Os ficheiros CSV utilizados durante o processo de recolha e preparação dos dados não são necessários para executar a aplicação online e, por isso, não fazem parte da versão de publicação da aplicação.

## Tecnologias

- Python
- Streamlit
- DuckDB
- Pandas
- Altair
- OpenPyXL

## Execução local

Para executar a aplicação localmente:

streamlit run streamlit_app.py

## Contexto académico

Projeto desenvolvido no âmbito de uma dissertação de Mestrado em Jornalismo.

A aplicação constitui uma ferramenta de apoio à análise de narrativas mediáticas relacionadas com terrorismo e fenómenos associados, permitindo cruzar informação proveniente do arquivo da Lusa com dados do Global Terrorism Database.

## Nota metodológica

O corpus da Lusa foi construído a partir de ficheiros CSV previamente recolhidos e tratados durante a investigação. A aplicação publicada utiliza a base de dados DuckDB resultante desse processo.
