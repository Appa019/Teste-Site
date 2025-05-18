# Leitor de Faturas de Energia Elétrica

Este projeto é uma aplicação Streamlit para extração automática de dados de faturas de energia elétrica de diferentes distribuidoras (Cemig, Light, Enel), com categorização de itens e geração de relatórios.

## Funcionalidades

- Upload de múltiplos arquivos PDF de faturas de energia elétrica
- Desbloqueio automático de PDFs protegidos por senha
- Extração de dados relevantes (instalação, referência, valores)
- Categorização automática dos itens da fatura
- Geração de planilha Excel com os dados extraídos
- Relatório de erros para arquivos não processados
- Interface amigável e responsiva

## Requisitos

- Python 3.8 ou superior
- Bibliotecas listadas em `requirements.txt`

## Instalação

1. Clone este repositório:
```bash
git clone https://github.com/seu-usuario/leitor-faturas-energia.git
cd leitor-faturas-energia
```

2. Instale as dependências:
```bash
pip install -r requirements.txt
```

## Uso

1. Execute a aplicação Streamlit:
```bash
streamlit run streamlit_app.py
```

2. Acesse a aplicação no navegador (geralmente em http://localhost:8501)

3. Faça upload das faturas em PDF

4. Escolha entre criar um novo arquivo Excel ou atualizar um existente

5. Clique em "Processar Faturas"

6. Verifique os resultados nas abas "Resultados" e "Relatório de Erros"

## Estrutura do Projeto

- `streamlit_app.py`: Aplicação principal com interface Streamlit
- `utils.py`: Módulo com funções utilitárias para processamento de PDFs
- `requirements.txt`: Lista de dependências do projeto

## Distribuidoras Suportadas

- Cemig
- Light
- Enel/Ampla

## Categorias de Itens

- Encargo
- Desc. Encargo
- Demanda
- Desc. Demanda
- Reativa
- Contribuição Pública
- Outros (itens não categorizados)

## Notas

- Faturas protegidas por senha são desbloqueadas automaticamente usando senhas conhecidas
- O sistema gera um relatório de erros para arquivos que falharam na extração
- Itens não categorizados são classificados como "Outros"
- Os valores são formatados no padrão brasileiro (vírgula como separador decimal)

## Licença

Este projeto está licenciado sob a licença MIT - veja o arquivo LICENSE para detalhes.
