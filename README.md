# Automação de Consulta de Faturas LIGHT

Este projeto implementa uma aplicação web para automatizar a consulta de faturas em aberto no portal da distribuidora de energia elétrica LIGHT, com base em uma planilha de códigos de instalação fornecida pelo usuário.

## Funcionalidades

- Upload de planilha com códigos de instalação
- Automação de login no portal da LIGHT
- Tentativa automática de contornar o captcha (com opção para intervenção manual)
- Consulta de faturas em aberto para cada código de instalação
- Download automático das faturas disponíveis
- Geração de relatório detalhado com status de cada instalação
- Interface web responsiva e amigável

## Requisitos

- Python 3.8+
- Navegador Chrome instalado
- Bibliotecas Python listadas em `requirements.txt`

## Instalação

1. Clone o repositório:
```
git clone https://github.com/seu-usuario/light-scraper.git
cd light-scraper
```

2. Instale as dependências:
```
pip install -r requirements.txt
```

3. Execute a aplicação:
```
streamlit run app.py
```

## Estrutura da Planilha

A planilha de entrada deve conter as seguintes colunas:
- FILIAL: valores possíveis incluem CSN, CMIN e CIBR
- RAZAO SOCIAL: padrão "LIGHT SERVIÇOS DE ELETRICIDADE SA"
- SITE: URL do portal da LIGHT
- LOGIN: credenciais de acesso
- SENHA: senha de acesso
- NUM INSTALACAO 1: código numérico da instalação
- NUM INSTALACAO 2: (opcional) código adicional

## Fluxo de Execução

1. Faça upload da planilha com os códigos de instalação
2. Clique em "Iniciar Consulta"
3. Se solicitado, resolva o captcha manualmente
4. Aguarde o processamento das consultas
5. Visualize o relatório final e faça o download em formato CSV

## Tratamento do Captcha

O sistema tenta contornar o captcha automaticamente usando técnicas de automação. Caso não seja possível, solicita a intervenção manual do usuário.

## Notas Importantes

- As faturas baixadas são salvas na pasta de downloads padrão do navegador
- O relatório final inclui hyperlinks para abrir as faturas
- A aplicação está preparada para ser hospedada no Streamlit

## Licença

Este projeto é licenciado sob a licença MIT.
