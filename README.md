
## Visão Geral
 Ferramenta especializada na extração e análise de informações estruturadas de faturas de energia elétrica em formato PDF. Utilizando processamento de linguagem natural via API da OpenAI, o sistema categoriza automaticamente os diferentes tipos de cobrança e gera relatórios analíticos detalhados.

---

## Arquitetura Técnica

### Processamento de Documentos

- **Extração de Texto**  
  Utiliza a biblioteca `PyMuPDF (fitz)` para converter PDFs em texto plano.

- **Reconhecimento de Estrutura**  
  Identifica a seção *"Itens da Fatura"* dentro do documento, separando-a para processamento específico.

- **Pré-processamento Textual**  
  Aplica normalização Unicode, remoção de caracteres especiais e padronização de texto para melhorar a precisão da classificação.

### Análise Semântica

- **API de Machine Learning**  
  Integração com a API GPT da OpenAI para interpretar conteúdo textual não estruturado.

- **Prompt Engineering**  
  Utiliza prompts específicos para extração e classificação de informações relevantes.

- **Extração de Metadados**  
  Identifica automaticamente o número da instalação e o período de referência do documento.

---

## Categorização Automática

O sistema utiliza um dicionário extensivo de mapeamento para associar termos comuns nas faturas às seguintes categorias:

- `Encargo`: Cobranças relacionadas ao consumo de energia.  
- `Desc. Encargo`: Descontos aplicados aos encargos.  
- `Demanda`: Cobranças fixas relacionadas à demanda contratada.  
- `Desc. Demanda`: Descontos aplicados à demanda.  
- `Reativa`: Cobranças associadas à energia reativa excedente.  
- `Contb. Pública`: Contribuições para iluminação pública.  
- `Outros`: Cobranças diversas não classificadas nas categorias anteriores.

---

## Processamento Numérico

- **Correção de Valores Negativos**  
  Detecta e padroniza valores negativos, inclusive quando o símbolo está posicionado após o número (ex: `32123,23-` → `-32123,23`).

- **Formatação Regional**  
  Converte valores para o formato brasileiro: vírgula como separador decimal e ponto como separador de milhar.

- **Cálculo Agregado**  
  Soma automaticamente os valores por categoria para fornecer visões consolidadas e resumos financeiros.

---

## Geração de Relatórios

- **Processamento Tabular**  
  Estrutura os dados extraídos em `DataFrames` para manipulação eficiente.

- **Formatação Excel**  
  Aplica estilos, cabeçalhos e ajuste automático de colunas para melhor apresentação.

- **Multi-sheet Reports**  
  Gera relatórios com múltiplas abas em Excel: uma para resumo e outra para cada PDF processado.

---

## Métricas e Análises

- **Token Accounting**  
  Contabiliza os tokens de entrada e saída utilizados nas chamadas à API da OpenAI.

- **Estimativa de Custo**  
  Calcula automaticamente o custo aproximado com base nos preços atuais da API.

- **Conversão Monetária**  
  Utiliza uma API externa para conversão de valores de USD para BRL em tempo real.

---

## Segurança e Performance

- **Processamento Local**  
  A extração de texto é feita localmente; apenas o conteúdo relevante é enviado às APIs externas.

- **Armazenamento Temporário**  
  Os arquivos são tratados temporariamente e removidos automaticamente após o uso.

- **Tratamento de Erros**  
  Implementa um sistema robusto de captura e exibição de erros, facilitando a identificação e resolução de problemas.

- **Limitações de Escala**  
  Define limites de processamento (ex: até 80 PDFs por execução) para garantir desempenho estável.

---

## Considerações Técnicas

- O sistema é otimizado para faturas brasileiras de energia elétrica.
- A qualidade da classificação depende da fidelidade da extração textual.
- A arquitetura modular permite adaptação para outros tipos de documentos e categorias.
- A precisão da extração pode variar conforme o formato e complexidade do PDF original.
