# Leitor de Faturas de Energia

Este aplicativo Streamlit processa faturas de energia elétrica em PDF, extraindo dados estruturados com a ajuda da API da OpenAI.

## Funcionalidades

- Interface interativa para seleção de áreas nos PDFs
- Desbloqueio automático de PDFs protegidos
- Identificação automática da distribuidora
- Extração seletiva de texto para reduzir custos com a API
- Classificação dos itens usando dicionário de mapeamento
- Visualização e exportação de resultados em Excel

## Instalação

1. Clone este repositório:
```
git clone https://github.com/seu-usuario/leitor-faturas.git
cd leitor-faturas
```

2. Instale as dependências:
```
pip install -r requirements.txt
```

3. Execute o aplicativo:
```
streamlit run app.py
```

## Como usar

1. Insira sua chave da API OpenAI
2. Faça upload das faturas em PDF
3. Selecione as áreas de referência e tabela
4. Processe as faturas
5. Exporte os resultados para Excel

## Requisitos

- Python 3.8+
- Chave de API da OpenAI

## Deploy no Streamlit Cloud

Para fazer o deploy no Streamlit Cloud:

1. Crie uma conta no [Streamlit Cloud](https://streamlit.io/cloud)
2. Conecte sua conta GitHub
3. Crie um novo repositório com este código
4. Faça o deploy do aplicativo no Streamlit Cloud

## Limitações

- O processamento depende da qualidade do PDF
- Faturas com layouts muito diferentes podem exigir ajustes manuais
- O custo do processamento depende do tamanho das tabelas extraídas
