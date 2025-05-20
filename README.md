# Leitor de Faturas de Energia

Este projeto implementa uma aplicação web interativa usando Streamlit para processar faturas de energia elétrica em PDF com estrutura variável. A aplicação utiliza a API da OpenAI para extrair dados das tabelas de itens de faturamento, otimizando o processo para reduzir o custo de tokens.

## Funcionalidades

- **Interface interativa para seleção de áreas nos PDFs:**
  - Visualização do PDF como imagem para compatibilidade com todos os navegadores
  - Seleção por coordenadas com destaque visual (azul para referência/instalação, verde para tabela)
  - Salvamento automático de padrões de seleção por distribuidora

- **Processamento otimizado:**
  - Desbloqueio automático de PDFs protegidos usando lista de senhas conhecidas
  - Identificação automática da distribuidora pelo cabeçalho
  - Extração seletiva de texto para reduzir custos com a API da OpenAI
  - Classificação dos itens usando dicionário de mapeamento

- **Visualização e exportação de resultados:**
  - Tabela interativa com todos os dados processados
  - Cálculo e exibição do custo em tokens e USD/BRL
  - Exportação para Excel com um clique

## Requisitos

```
streamlit==1.32.0
pdfplumber==0.11.6
PyPDF2==3.0.1
pikepdf==8.10.1
pandas==2.2.0
openai==1.12.0
openpyxl==3.1.2
Pillow==10.2.0
requests==2.31.0
pdf2image==1.17.0
```

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

3. Execute a aplicação:
```
streamlit run app.py
```

## Como usar

1. Insira sua chave da API OpenAI
2. Faça upload das faturas em PDF
3. Selecione as áreas diretamente no PDF:
   - Área azul: Referência/Instalação
   - Área verde: Tabela de itens
4. Processe as faturas
5. Exporte os resultados para Excel

## Estrutura do código

- `app.py`: Aplicação principal do Streamlit
- `requirements.txt`: Dependências do projeto

## Notas importantes

- A aplicação requer uma chave válida da API OpenAI
- Para PDFs protegidos, a aplicação tentará desbloquear usando uma lista predefinida de senhas
- O dicionário de mapeamento está incorporado no código para classificação dos itens
- A aplicação salva padrões de seleção por distribuidora para processamento mais rápido de faturas similares

## Limitações conhecidas

- A visualização do PDF como imagem pode ter resolução reduzida em telas muito pequenas
- A extração de dados depende da qualidade do PDF e da precisão da seleção das áreas
- O processamento com a API da OpenAI pode variar em precisão dependendo da complexidade da tabela

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests com melhorias.

## Licença

Este projeto está licenciado sob a licença MIT - veja o arquivo LICENSE para detalhes.
