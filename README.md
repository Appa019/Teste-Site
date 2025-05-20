# Extrator Automático de Tabelas de Faturas de Energia

Este script processa faturas de energia em PDF, extraindo automaticamente as tabelas de itens de faturamento e classificando-os nas categorias apropriadas.

## Características

- Processamento em lote de múltiplos PDFs
- Desbloqueio automático de PDFs protegidos com senhas conhecidas
- Detecção automática de tabelas usando palavras-chave
- Extração otimizada com uso eficiente da API OpenAI
- Classificação dos itens em categorias predefinidas
- Geração de tabela resumo consolidada
- Exportação para Excel com múltiplas abas

## Requisitos

- Python 3.6 ou superior
- Bibliotecas Python listadas em `requirements.txt`
- Chave de API da OpenAI (opcional, mas recomendada para melhor precisão)

## Instalação

1. Clone ou baixe este repositório
2. Instale as dependências:

```bash
pip install -r requirements.txt
```

## Uso

1. Coloque os arquivos PDF das faturas na pasta `pdfs/`
2. Execute o script:

```bash
python extrator_faturas_final.py
```

3. Quando solicitado, insira sua chave da API OpenAI
4. O script processará todos os PDFs e gerará um arquivo Excel com os resultados na pasta `resultados/`

## Estrutura do Arquivo Excel Gerado

O arquivo Excel gerado contém:

- Uma aba "Resumo" com a tabela resumo consolidada, contendo:
  - Número da instalação
  - Data de referência
  - Valores agrupados por categoria (Encargo, Desc. Encargo, Reativa, Demanda, Desc. Demanda, Contb. Publica)

- Uma aba para cada fatura processada, contendo a tabela de itens de faturamento com classificação

## Categorias de Classificação

Os itens das faturas são classificados nas seguintes categorias:

- **Encargo**: Consumo ativo, energia ativa, etc.
- **Desc. Encargo**: Descontos e devoluções relacionados a encargos
- **Reativa**: Consumo reativo, energia reativa, etc.
- **Demanda**: Demanda ativa, demanda ponta, etc.
- **Desc. Demanda**: Descontos e devoluções relacionados a demanda
- **Contb. Publica**: Contribuição para iluminação pública
- **Outros**: Itens que não se encaixam nas categorias acima

## Tratamento de PDFs Protegidos

O script tenta desbloquear PDFs protegidos usando uma lista de senhas conhecidas:
```
["33042", "56993", "60869", "08902", "3304", "5699", "6086", "0890"]
```

## Logs

O script gera logs detalhados no arquivo `extrator_faturas.log`, úteis para diagnóstico em caso de problemas.

## Autor

Desenvolvido por Manus AI - Maio 2025
