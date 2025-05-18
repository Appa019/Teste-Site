import os
import re
import PyPDF2
import pdfplumber
import pandas as pd
from datetime import datetime
import unicodedata
import logging
import json

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger('leitor_faturas')

# Senhas conhecidas para desbloqueio de PDFs
SENHAS_CONHECIDAS = ["33042", "56993", "60869", "08902", "3304", "5699", "6086", "0890"]

# Dicionário de mapeamento de palavras-chave
MAPEAMENTO_PALAVRAS_CHAVE = {
    # Encargo
    "Energia elet. adquirida 3": "Encargo",
    "energia elet adquirida 3": "Encargo",
    "energiaelet.adquirida3": "Encargo",
    "energia elet. adquirida3": "Encargo",
    "energia elet. adquirida 3* c/imposto": "Encargo",
    "energia elet. adquirida 3* c/lmposto": "Encargo",
    "energia elet. adquirida 3* c imp": "Encargo",
    "energia elet adquirida 3 c imposto": "Encargo",
    "Liminar Red Icms Fatura": "Encargo",
    "liminar red icms fatura": "Encargo",
    "liminarredicmsfatura": "Encargo",
    "liminar red. icms fatura": "Encargo",
    "Subtotal Faturamento": "Encargo",
    "subtotal faturamento": "Encargo",
    "sub total faturamento": "Encargo",
    "subtot. faturamento": "Encargo",
    "Componente Encargo kWh HFP": "Encargo",
    "componente encargo kwh hfp": "Encargo",
    "componente encargokwh hfp": "Encargo",
    "componente encargo kw hfp": "Encargo",
    "Componente Encargo kWh HP": "Encargo",
    "componente encargo kwh hp": "Encargo",
    "componente encargokwh hp": "Encargo",
    "componente encargo kw hp": "Encargo",
    "Componente Encargo HFP": "Encargo",
    "componente encargo hfp": "Encargo",
    "Componente Encargo HP": "Encargo",
    "componente encargo hp": "Encargo",
    "ENERGIA ACL": "Encargo",
    "energia acl": "Encargo",
    "energiaacl": "Encargo",
    "energia acl ": "Encargo",
    "CONSUMOATIVO PONTA TUSD": "Encargo",
    "consumoativo ponta tusd": "Encargo",
    "consumo ativo ponta tusd": "Encargo",
    "consumo ativoponta tusd": "Encargo",
    "CONSUMO ATIVO F.PONTA TUSD": "Encargo",
    "consumo ativo f.ponta tusd": "Encargo",
    "consumo ativo f. ponta tusd": "Encargo",
    "consumo ativo f ponta tusd": "Encargo",
    "BENEFICIO TARIFARIO BRUTO": "Encargo",
    "beneficio tarifario bruto": "Encargo",
    "benefício tarifario bruto": "Encargo",
    "beneficio tar. bruto": "Encargo",
    "benefício tar. bruto": "Encargo",
    "PIS/COFINS da subvencao/descon": "Encargo",
    "PIS/COFINS da subvencäo/descon": "Encargo",
    "pis cofins da subvençao/descon": "Encargo",
    "piscofins da subvencao descon": "Encargo",
    "Correcao IPCA/IGPM s/ conta 01/25 pg 19/02/25": "Encargo",
    "correção ipca igpm s/ conta": "Encargo",
    "correcao ipca igpm s/conta": "Encargo",
    "correcao ipca/igpm s/conta": "Encargo",
    "correcao ipca igpm s conta": "Encargo",
    "Encargo Cta Covid REN 885/2020": "Encargo",
    "encargo cta covid ren885/2020": "Encargo",
    "encargo covid ren 885/2020": "Encargo",
    "encargo covid ren885/2020": "Encargo",
    "Energia Ativa kWh HFP/nico": "Encargo",
    "energia ativa kwh hfp/unico": "Encargo",
    "energia ativa kwh hfp unico": "Encargo",
    "energia ativa kwh hfp": "Encargo",
    "Energia Ativa kWh HP": "Encargo",
    "energia ativa kwh hp": "Encargo",
    "DEBITO VAR IPCA": "Encargo",
    "débito var ipca": "Encargo",
    "debito var ipca": "Encargo",
    "deb1to var ipca": "Encargo",
    "TUSD ponta": "Encargo",
    "TUSD fora ponta": "Encargo",
    "tusd ponta": "Encargo",
    "tusd fora ponta": "Encargo",
    "Consumo Energia Elétrica HP": "Encargo",
    "consumo energia eletrica hp": "Encargo",
    "consumo energia elétrica hp": "Encargo",
    "DEMANDALEIESTADUAL16.886/18": "Encargo",
    "Componente Fio HP": "Encargo",
    "Componente Fio HFP": "Encargo",
    
    # Desc. Encargo
    "DEDUCAO ENERGIA ACL": "Desc. Encargo",
    "DEDUCÃO ENERGIA ACL": "Desc. Encargo",
    "DED. ENERGIA ACL": "Desc. Encargo",
    "D. ENERGIA ACL": "Desc. Encargo",
    "Valor de Autoprodução HFP": "Desc. Encargo",
    "BENEFICIO TARIFARIO LIQUIDO": "Desc. Encargo",
    "BENEFÍCIO TARIFÁRIO LIQUIDO": "Desc. Encargo",
    "BENEFiCIO TARIFARIO LIQUIDO": "Desc. Encargo",
    "BENEFÍCIO TARIFÁRIO LÍQUIDO": "Desc. Encargo",
    "BENEFICIO TARIFÁRIO LÍQUIDO": "Desc. Encargo",
    "liminar icm5 tusd-consumo": "Desc. Encargo",
    "liminar icms tusd-consu mo": "Desc. Encargo",
    "liminar icms tusd-con5umo": "Desc. Encargo",
    "liminaricms tusd consumo": "Desc. Encargo",
    "liminar icms t usd-consumo": "Desc. Encargo",
    "liminar icms tu sd consumo": "Desc. Encargo",
    "liminar icms tusd-con sumo": "Desc. Encargo",
    "liminar icms t u s d-consumo": "Desc. Encargo",
    "liminar icms tusd-demanda": "Desc. Encargo",
    "liminar icms tusd demanda": "Desc. Encargo",
    "liminar icms tusddemanda": "Desc. Encargo",
    "liminar icms tus ddemanda": "Desc. Encargo",
    "liminar 1cms tusd-demanda": "Desc. Encargo",
    "liminar icms tuso-demanda": "Desc. Encargo",
    "liminar icm5 tusd-demanda": "Desc. Encargo",
    "liminar icms tusd demand a": "Desc. Encargo",
    "liminaricms tusd demanda": "Desc. Encargo",
    "liminar icms tu sd-demanda": "Desc. Encargo",
    "liminar icms t usd-demanda": "Desc. Encargo",
    "liminar icms tus-d-demanda": "Desc. Encargo",
    "liminar icms tusd d emanda": "Desc. Encargo",
    "liminar icms t u s d-demanda": "Desc. Encargo",
    "Liminar ICMS TUSD-Demanda": "Desc. Encargo",
    "Liminar ICMS TUSD Demanda": "Desc. Encargo",
    "Liminar ICMS TUSDDemanda": "Desc. Encargo",
    "Liminar ICMS TUS D Demanda": "Desc. Encargo",
    "Liminar ICMS TUS-D-Demanda": "Desc. Encargo",
    "Liminar ICMS TUSD Deman da": "Desc. Encargo",
    "Liminar 1CMS TUSD-Demanda": "Desc. Encargo",
    "Llminar ICMS TUSD-Demanda": "Desc. Encargo",
    "Liminar ICMS TUSO-Demanda": "Desc. Encargo",
    "Liminar ICMS TUS-Demanda": "Desc. Encargo",
    "Liminar ICMS TUSDDemand a": "Desc. Encargo",
    "Liminar ICMS TUSD-Consumo": "Desc. Encargo",
    "Liminar ICMS TUSD Consumo": "Desc. Encargo",
    "DEDUCAO ENERGIAACL": "Desc. Encargo",
    "DEDUÇÃO ENERGIAACL": "Desc. Encargo",
    "DED. ENERGIAACL": "Desc. Encargo",
    "deducao energiaacl": "Desc. Encargo",
    "deducaoenergiaacl": "Desc. Encargo",
    "deducao energia acl": "Desc. Encargo",
    "DEDUÇÃO ENERGIA ACL": "Desc. Encargo",
    "Liminar ICMS TUSDConsumo": "Desc. Encargo",
    "Liminar ICMS TUS D Consumo": "Desc. Encargo",
    "Liminar ICMS TUSD-Consu mo": "Desc. Encargo",
    "Liminar ICMS TUSO-Consumo": "Desc. Encargo",
    "Liminar 1CMS TUSD-Consumo": "Desc. Encargo",
    "Llminar ICMS TUSD-Consumo": "Desc. Encargo",
    "Liminar ICMS TUSD-Consuno": "Desc. Encargo",
    "Liminar ICMS TUS DConsumo": "Desc. Encargo",
    "Valor de Autoproducao FP": "Desc. Encargo",
    "Valor de Autoprodução HP": "Desc. Encargo",
    "Desconto Comp. Encargo HP": "Desc. Encargo",
    "Ajuste de Desconto C. Enc HP": "Desc. Encargo",
    "Desconto Comp. Encargo HFP": "Desc. Encargo",
    "Ajuste de Desconto C. Enc HFP": "Desc. Encargo",
    "desconto compencargo hp": "Desc. Encargo",
    "ajuste de desconto cenc hp": "Desc. Encargo",
    "Valor Autoprodução HFP": "Desc. Encargo",
    "Valor de Autoproducao HFP": "Desc. Encargo",
    "Valor Autoprodução HP": "Desc. Encargo",
    "Valor de Autoproducao HP": "Desc. Encargo",
    "Desconto Comp Encargo HP": "Desc. Encargo",
    "Desconto Compencargo HP": "Desc. Encargo",
    "Ajuste de Desconto C Enc HP": "Desc. Encargo",
    "Ajuste de Desconto CEnc HP": "Desc. Encargo",
    "Ajuste de Desconto C Enc HFP": "Desc. Encargo",
    "Ajuste de Desconto CEnc HFP": "Desc. Encargo",
    "Desconto Comp Encargo HFP": "Desc. Encargo",
    "Desconto Compencargo HFP": "Desc. Encargo",
    "Restituicao de Pagamento": "Desc. Encargo",
    "Restituição de Pagamento": "Desc. Encargo",
    "Restituic:ao de Pagamento": "Desc. Encargo",
    "Restitüição de Pagamento": "Desc. Encargo",
    "Restituicaoo de Pagamento": "Desc. Encargo",
    "Restituicão Pagamento": "Desc. Encargo",
    "TAR.REDUZ.RES.ANEEL166/05P": "Desc. Encargo",
    "TAR.REDUZ.RES.ANEEL166/05FP": "Desc. Encargo",
    "Diferenca Energia Retroativa": "Desc. Encargo",
    "Desconto Comp.Encargo HP": "Desc. Encargo",
    "Ajuste de Desconto C.Enc HP": "Desc. Encargo",
    "Ajuste de Desconto C.Enc HFP": "Desc. Encargo",
    "Energia Terc Comercializad HP": "Desc. Encargo",
    "Energia Terc Comercializad HFP": "Desc. Encargo",
    "Desconto Comp.Encargo HFP": "Desc. Encargo",
    "Ajuste de Desconto C.Enc HFP": "Desc. Encargo",
    "Desconto Comp Fio HP": "Desc. Encargo",
    "Desconto Comp Fio HFP": "Desc. Encargo",
    
    # Demanda
    "demanda ponta c/desconto": "Demanda",
    "demanda ponta cidesconto": "Demanda",
    "demanda ponta cidescon": "Demanda",
    "demanda ponta c/desc0nt0": "Demanda",
    "DEMANDA FORA PONTA C/DESC": "Demanda",
    "DEMANDA FORA PONTA C/DESC0": "Demanda",
    "DEMANDA FORA PONTA C/DESCONTO": "Demanda",
    "DEMANDA FORA PONTA C/DESC0NT0": "Demanda",
    "DEMANDA FORA PONTA CIDESCONTO": "Demanda",
    "DEMANDA FORA PONTA CIDESC0NTO": "Demanda",
    "demanda p0nta c/desconto": "Demanda",
    "demanda ponta c / desconto": "Demanda",
    "ULTRAPASSAGEM PONTA/TUSD": "Demanda",
    "ULTRAPASSAGEM PONTA TUSD": "Demanda",
    "ULTRAPASSAGEM PONTA TUSD-": "Demanda",
    "ULTRAPASSAGEM PONTA/TUSD ": "Demanda",
    "ULTRAPASSAGEM PONTA/T USD": "Demanda",
    "ULTRAPASSAGEM PONTA/TUSD-": "Demanda",
    "ULTRAPASSAGEM PONTA/TUSD ": "Demanda",
    "ULTRAPASSAGEM PONTA-TUSD": "Demanda",
    "ULTRA PASSAGEM PONTA/TUSD": "Demanda",
    "ULTRA PASSAGEM PONTA TUSD": "Demanda",
    "ULTRAPASSAGEM PONTA TUSD ": "Demanda",
    "ULTRAPASSAGEM FORAPONTA/TUSD": "Demanda",
    "ULTRAPASSAGEM FORA PONTA TUSD": "Demanda",
    "ULTRAPASSAGEM FORAPONTA TUSD": "Demanda",
    "ULTRAPASSAGEM FORA PONTA/TUSD": "Demanda",
    "ULTRAPASSAGEM FORA PONTA TUSD ": "Demanda",
    "UFDR PONTA/TUSD": "Demanda",
    "UFDRPONTA/TUSD": "Demanda",
    "UFDR PONTA TUSD": "Demanda",
    "UFDRPONTATUSD": "Demanda",
    "UFDR P0NTA/TUSD": "Demanda",
    "UFDR PONTA/T USD": "Demanda",
    "UFDR P0NTA TUSD": "Demanda",
    "UFDR PONTA-TUSD": "Demanda",
    "UFDR PONTA/TUS D": "Demanda",
    "UFDRFORA PONTA/TUSD": "Demanda",
    "UFDR FORA PONTA/TUSD": "Demanda",
    "UFDR FORAPONTA/TUSD": "Demanda",
    "UFDR FORA PONTA TUSD": "Demanda",
    "UFDRFORAPONTA TUSD": "Demanda",
    "UFDRFORA PONTA TUSD": "Demanda",
    "UFDR FORA P0NTA/TUSD": "Demanda",
    "Demanda Ponta": "Demanda",
    "Demanda Fora Ponta": "Demanda",
    "Componente Fio kW HFP": "Demanda",
    "Componente Fio kW HP": "Demanda",
    
    # Desc. Demanda
    "Liminar ICMS TUSD-Demanda": "Desc. Demanda",
    "Liminar ICMS TUSD Demanda": "Desc. Demanda",
    "Liminar ICMS TUSDDemanda": "Desc. Demanda",
    "Liminar ICMS TUS D Demanda": "Desc. Demanda",
    "Liminar ICMS TUS-D-Demanda": "Desc. Demanda",
    "Liminar ICMS TUSD Deman da": "Desc. Demanda",
    "Liminar 1CMS TUSD-Demanda": "Desc. Demanda",
    "Llminar ICMS TUSD-Demanda": "Desc. Demanda",
    "Liminar ICMS TUSO-Demanda": "Desc. Demanda",
    "Liminar ICMS TUS-Demanda": "Desc. Demanda",
    "Liminar ICMS TUSDDemand a": "Desc. Demanda",
    "liminar icms tusd-demanda": "Desc. Demanda",
    "liminar icms tusd demanda": "Desc. Demanda",
    "liminar icms tusddemanda": "Desc. Demanda",
    "liminar icms tus ddemanda": "Desc. Demanda",
    "liminar 1cms tusd-demanda": "Desc. Demanda",
    "liminar icms tuso-demanda": "Desc. Demanda",
    "liminar icm5 tusd-demanda": "Desc. Demanda",
    "liminar icms tusd demand a": "Desc. Demanda",
    "liminaricms tusd demanda": "Desc. Demanda",
    "liminar icms tu sd-demanda": "Desc. Demanda",
    "liminar icms t usd-demanda": "Desc. Demanda",
    "liminar icms tus-d-demanda": "Desc. Demanda",
    "liminar icms tusd d emanda": "Desc. Demanda",
    "liminar icms t u s d-demanda": "Desc. Demanda",
    
    # Reativa
    "Energia Reativa HFP": "Reativa",
    "Energia Reativa HP": "Reativa",
    "energia reativa hfp": "Reativa",
    "energia reativa hp": "Reativa",
    "ENERGIA REATIVA HFP": "Reativa",
    "ENERGIA REATIVA HP": "Reativa",
    "ENERGIA REATIVA": "Reativa",
    "energia reativa": "Reativa",
    "Energia Reativa kWh HFP": "Reativa",
    "Energia Reativa kWh HP": "Reativa",
    "energia reativa kwh hfp": "Reativa",
    "energia reativa kwh hp": "Reativa",
    "Energia Reativa kWh HFP/Único": "Reativa",
    
    # Contribuição Pública
    "Contrib Ilum Publica Municipal": "Contribuição Pública",
    "Contrib Ilum Pública Municipal": "Contribuição Pública",
    "contrib ilum publica municipal": "Contribuição Pública",
    "contrib ilum pública municipal": "Contribuição Pública",
    "CONTRIB ILUM PUBLICA MUNICIPAL": "Contribuição Pública",
    "CONTRIB ILUM PÚBLICA MUNICIPAL": "Contribuição Pública",
    "Contribuição Iluminação Pública": "Contribuição Pública",
    "Contribuicao Iluminacao Publica": "Contribuição Pública",
    "contribuição iluminação pública": "Contribuição Pública",
    "contribuicao iluminacao publica": "Contribuição Pública",
    "CONTRIBUIÇÃO ILUMINAÇÃO PÚBLICA": "Contribuição Pública",
    "CONTRIBUICAO ILUMINACAO PUBLICA": "Contribuição Pública",
    "Contib. Ilum. Publica Municipal": "Contribuição Pública",
    "Contib. Ilum. Pública Municipal": "Contribuição Pública",
    "contib. ilum. publica municipal": "Contribuição Pública",
    "contib. ilum. pública municipal": "Contribuição Pública",
    "CONTIB. ILUM. PUBLICA MUNICIPAL": "Contribuição Pública",
    "CONTIB. ILUM. PÚBLICA MUNICIPAL": "Contribuição Pública",
}

# Palavras-chave para identificar tabelas relevantes por distribuidora
PALAVRAS_CHAVE_TABELAS = {
    "cemig": ["Itens da Fatura", "Valores Faturados"],
    "light": ["Itens de fatura", "ltens de fatura", "Items de fatura", "Itens da fatura"],
    "enel": ["DISCRIMINAÇÃO DO FATURAMENTO", "Itens de fatura"]
}

def normalizar_texto(texto):
    """Normaliza o texto removendo acentos e convertendo para minúsculas"""
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return texto.lower()

def desbloquear_pdf(pdf_path):
    """Tenta desbloquear um PDF com as senhas conhecidas"""
    for senha in SENHAS_CONHECIDAS:
        try:
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                if reader.is_encrypted:
                    if reader.decrypt(senha) > 0:
                        # Senha correta, salvar versão desbloqueada
                        writer = PyPDF2.PdfWriter()
                        for page in reader.pages:
                            writer.add_page(page)
                        
                        temp_path = os.path.join(os.path.dirname(pdf_path), f"desbloqueado_{os.path.basename(pdf_path)}")
                        with open(temp_path, 'wb') as output_file:
                            writer.write(output_file)
                        
                        return temp_path, senha
                else:
                    # PDF não está criptografado
                    return pdf_path, None
        except Exception as e:
            logger.warning(f"Erro ao tentar senha {senha}: {str(e)}")
            continue
    
    # Nenhuma senha funcionou
    return None, None

def identificar_distribuidora(texto, nome_arquivo=None):
    """Identifica a distribuidora com base no texto e nome do arquivo"""
    texto_lower = texto.lower()
    
    # Verificar pelo nome do arquivo primeiro (mais confiável)
    if nome_arquivo:
        nome_lower = nome_arquivo.lower()
        if 'light' in nome_lower:
            return "light"
        if 'cemig' in nome_lower:
            return "cemig"
        if 'enel' in nome_lower or 'ampla' in nome_lower:
            return "enel"
    
    # Verificar pelo conteúdo do texto
    if 'light' in texto_lower or 'light serviços de eletricidade' in texto_lower:
        return "light"
    if 'cemig' in texto_lower or 'companhia energética de minas gerais' in texto_lower:
        return "cemig"
    if 'enel' in texto_lower or 'ampla' in texto_lower:
        return "enel"
    
    # Verificar por outros padrões específicos
    if 'discriminação do faturamento' in texto_lower:
        return "enel"
    if 'valores faturados' in texto_lower:
        return "cemig"
    if 'itens de fatura' in texto_lower:
        return "light"
    
    # Não foi possível identificar
    return "desconhecida"

def extrair_instalacao(texto, distribuidora, nome_arquivo=None):
    """Extrai o número de instalação com base na distribuidora"""
    texto_lower = texto.lower()
    
    # Padrões específicos por distribuidora
    if distribuidora == 'cemig':
        # Padrão Cemig: "Instalação: 3009999999"
        match = re.search(r'instalação:?\s*(\d+)', texto_lower)
        if match:
            return match.group(1).strip()
        
        # Tentar outro padrão: "Nº DA INSTALAÇÃO: 3009999999"
        match = re.search(r'n[º°]?\s*da\s*instalação:?\s*(\d+)', texto_lower)
        if match:
            return match.group(1).strip()
            
    elif distribuidora == 'light':
        # Padrão Light: "Nº da instalação 99999999"
        match = re.search(r'n[º°]?\s*da\s*instalação:?\s*(\d+)', texto_lower)
        if match:
            return match.group(1).strip()
        
        # Outro padrão Light: "Instalação: 99999999"
        match = re.search(r'instalação:?\s*(\d+)', texto_lower)
        if match:
            return match.group(1).strip()
            
    elif distribuidora == 'enel':
        # Padrão Enel: "Nº da instalação: 9999999"
        match = re.search(r'n[º°]?\s*da\s*instalação:?\s*(\d+)', texto_lower)
        if match:
            return match.group(1).strip()
        
        # Outro padrão Enel: "Instalação: 9999999"
        match = re.search(r'instalação:?\s*(\d+)', texto_lower)
        if match:
            return match.group(1).strip()
        
        # Outro padrão Enel: "CÓDIGO DA INSTALAÇÃO: 9999999"
        match = re.search(r'código\s*da\s*instalação:?\s*(\d+)', texto_lower)
        if match:
            return match.group(1).strip()
    
    # Tentar extrair do nome do arquivo como último recurso
    if nome_arquivo:
        # Procurar por padrões numéricos no nome do arquivo
        match = re.search(r'(\d{7,10})', nome_arquivo)
        if match:
            return match.group(1).strip()
    
    # Busca genérica por qualquer número que pareça uma instalação
    match = re.search(r'(?:instalação|instalacao|inst|código|codigo)[\s:]*(\d{5,10})', texto_lower)
    if match:
        return match.group(1).strip()
    
    return "Não identificado"

def extrair_referencia(texto):
    """Extrai a referência (mês/ano) da fatura"""
    texto_lower = texto.lower()
    
    # Padrão: "Referência: MM/AAAA"
    match = re.search(r'referência:?\s*(\d{2}/\d{4})', texto_lower)
    if match:
        return match.group(1).strip()
    
    # Padrão: "Referente a: MM/AAAA"
    match = re.search(r'referente\s*a:?\s*(\d{2}/\d{4})', texto_lower)
    if match:
        return match.group(1).strip()
    
    # Padrão: "Ref.: MM/AAAA"
    match = re.search(r'ref\.?:?\s*(\d{2}/\d{4})', texto_lower)
    if match:
        return match.group(1).strip()
    
    # Padrão: "Mês de referência: MM/AAAA"
    match = re.search(r'mês\s*de\s*referência:?\s*(\d{2}/\d{4})', texto_lower)
    if match:
        return match.group(1).strip()
    
    # Padrão: "Data de emissão: DD/MM/AAAA"
    match = re.search(r'data\s*de\s*emissão:?\s*(\d{2}/\d{2}/\d{4})', texto_lower)
    if match:
        data = match.group(1).strip()
        partes = data.split('/')
        if len(partes) == 3:
            return f"{partes[1]}/{partes[2]}"
    
    # Padrão: "FEV/2025" ou "FEVEREIRO/2025"
    meses = {
        'jan': '01', 'fev': '02', 'mar': '03', 'abr': '04', 'mai': '05', 'jun': '06',
        'jul': '07', 'ago': '08', 'set': '09', 'out': '10', 'nov': '11', 'dez': '12'
    }
    
    for mes_abrev, mes_num in meses.items():
        # Formato abreviado: "FEV/2025"
        match = re.search(fr'{mes_abrev}[a-z]*[/\s-](\d{{4}})', texto_lower)
        if match:
            return f"{mes_num}/{match.group(1)}"
    
    # Busca por qualquer formato de data que pareça uma referência
    match = re.search(r'(\d{2})[/\s-](\d{4})', texto_lower)
    if match and 1 <= int(match.group(1)) <= 12:
        return f"{match.group(1)}/{match.group(2)}"
    
    return "Não identificado"

def extrair_tabela_relevante(texto_completo, distribuidora):
    """Extrai a tabela relevante com base na distribuidora"""
    linhas = texto_completo.split('\n')
    tabela_encontrada = False
    tabela_linhas = []
    
    # Palavras-chave específicas para cada distribuidora
    palavras_chave = PALAVRAS_CHAVE_TABELAS.get(distribuidora, [])
    
    # Se não temos palavras-chave para esta distribuidora, usar abordagem genérica
    if not palavras_chave:
        palavras_chave = ["Itens", "Discriminação", "Valores", "Faturamento", "Fatura"]
    
    # Buscar pela tabela usando as palavras-chave
    for i, linha in enumerate(linhas):
        # Verificar se a linha contém alguma das palavras-chave
        if any(palavra.lower() in linha.lower() for palavra in palavras_chave):
            tabela_encontrada = True
            tabela_linhas.append(linha)
            continue
        
        # Se já encontrou a tabela, continuar coletando linhas até encontrar um marcador de fim
        if tabela_encontrada:
            # Verificar se chegou ao fim da tabela
            if "TOTAL" in linha and len(tabela_linhas) > 5:
                tabela_linhas.append(linha)
                break
            
            # Verificar se a linha está vazia ou contém apenas caracteres de formatação
            if not linha.strip() or all(c in ' -=:.|' for c in linha.strip()):
                continue
            
            # Adicionar a linha à tabela
            if linha.strip():
                tabela_linhas.append(linha)
    
    # Se não encontrou a tabela pelo método padrão, tentar abordagem alternativa
    if not tabela_linhas and distribuidora == 'light':
        # Buscar por padrões específicos da Light
        for i, linha in enumerate(linhas):
            if 'Unid.' in linha and 'Quant.' in linha and 'Tarifa' in linha:
                tabela_encontrada = True
                continue
            
            if tabela_encontrada:
                if 'TOTAL' in linha and len(tabela_linhas) > 5:
                    tabela_linhas.append(linha)
                    break
                
                if linha.strip() and ('kWh' in linha or 'Componente' in linha or 'Energia' in linha or 
                                     'Demanda' in linha or 'Ajuste' in linha or 'Desconto' in linha or 
                                     'Contrib' in linha or 'TUSD' in linha):
                    tabela_linhas.append(linha)
    
    # Se ainda não encontrou, tentar uma abordagem mais genérica para a Light
    if not tabela_linhas and distribuidora == 'light':
        componente_encontrado = False
        for i, linha in enumerate(linhas):
            if 'Componente' in linha and 'Fio' in linha:
                componente_encontrado = True
                tabela_linhas.append(linha)
                continue
            
            if componente_encontrado:
                if 'TOTAL' in linha and len(tabela_linhas) > 5:
                    tabela_linhas.append(linha)
                    break
                
                if linha.strip():
                    tabela_linhas.append(linha)
    
    # Abordagem específica para Cemig
    if not tabela_linhas and distribuidora == 'cemig':
        for i, linha in enumerate(linhas):
            if 'Valores Faturados' in linha:
                tabela_encontrada = True
                tabela_linhas.append(linha)
                continue
            
            if tabela_encontrada:
                if 'TOTAL' in linha and len(tabela_linhas) > 5:
                    tabela_linhas.append(linha)
                    break
                
                if linha.strip() and not linha.startswith('Página'):
                    tabela_linhas.append(linha)
    
    # Abordagem específica para Enel
    if not tabela_linhas and distribuidora == 'enel':
        for i, linha in enumerate(linhas):
            if 'DISCRIMINAÇÃO DO FATURAMENTO' in linha:
                tabela_encontrada = True
                tabela_linhas.append(linha)
                continue
            
            if tabela_encontrada:
                if 'TOTAL A PAGAR' in linha and len(tabela_linhas) > 5:
                    tabela_linhas.append(linha)
                    break
                
                if linha.strip() and not linha.startswith('Página'):
                    tabela_linhas.append(linha)
    
    return '\n'.join(tabela_linhas) if tabela_linhas else None

def categorizar_item(descricao):
    """Categoriza um item com base no dicionário de mapeamento"""
    descricao_lower = descricao.lower()
    
    # Tentar correspondência exata primeiro
    for chave, categoria in MAPEAMENTO_PALAVRAS_CHAVE.items():
        if chave.lower() == descricao_lower:
            return categoria
    
    # Tentar correspondência parcial com limitações
    for chave, categoria in MAPEAMENTO_PALAVRAS_CHAVE.items():
        if chave.lower() in descricao_lower or descricao_lower in chave.lower():
            # Verificar se a correspondência é significativa (pelo menos 70% de similaridade)
            if len(chave) > 3 and (len(chave) >= 0.7 * len(descricao) or len(descricao) >= 0.7 * len(chave)):
                return categoria
    
    return "Outros"

def extrair_itens_tabela(tabela_texto):
    """Extrai e categoriza itens da tabela"""
    linhas = tabela_texto.split('\n')
    itens = []
    
    for linha in linhas:
        # Tentar extrair descrição e valor
        # Padrão: descrição seguida de valor numérico
        match = re.search(r'([a-zA-Z\s.]+)[\s]*([0-9,.]+)', linha)
        if match:
            descricao = match.group(1).strip()
            valor_str = match.group(2).strip()
            
            # Verificar se o valor é negativo (pode estar no formato 123- ou -123)
            valor_negativo = False
            if valor_str.endswith('-'):
                valor_str = valor_str[:-1]
                valor_negativo = True
            
            try:
                # Converter para float, substituindo vírgula por ponto
                valor_str = valor_str.replace('.', '').replace(',', '.')
                valor = float(valor_str)
                if valor_negativo or '-' in linha:
                    valor = -valor
                
                categoria = categorizar_item(descricao)
                itens.append({
                    'descricao': descricao,
                    'valor': valor,
                    'categoria': categoria
                })
            except ValueError:
                continue
    
    return itens

def processar_pdf(pdf_path):
    """Processa um PDF e extrai os dados"""
    try:
        nome_arquivo = os.path.basename(pdf_path)
        
        # Extrair texto do PDF
        texto_completo = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                texto_completo += page.extract_text() + "\n"
        
        # Identificar distribuidora
        distribuidora = identificar_distribuidora(texto_completo, nome_arquivo)
        
        # Extrair instalação e referência
        instalacao = extrair_instalacao(texto_completo, distribuidora, nome_arquivo)
        referencia = extrair_referencia(texto_completo)
        
        # Extrair tabela relevante
        tabela_texto = extrair_tabela_relevante(texto_completo, distribuidora)
        
        if not tabela_texto:
            return {
                'sucesso': False,
                'erro': 'Tabela relevante não encontrada',
                'distribuidora': distribuidora,
                'instalacao': instalacao,
                'referencia': referencia
            }
        
        # Extrair e categorizar itens
        itens = extrair_itens_tabela(tabela_texto)
        
        # Calcular valores por categoria
        encargo = sum(item['valor'] for item in itens if item['categoria'] == 'Encargo')
        reativa = sum(item['valor'] for item in itens if item['categoria'] == 'Reativa')
        desc_encargo = sum(item['valor'] for item in itens if item['categoria'] == 'Desc. Encargo')
        demanda = sum(item['valor'] for item in itens if item['categoria'] == 'Demanda')
        desc_demanda = sum(item['valor'] for item in itens if item['categoria'] == 'Desc. Demanda')
        contribuicao_publica = sum(item['valor'] for item in itens if item['categoria'] == 'Contribuição Pública')
        outros = sum(item['valor'] for item in itens if item['categoria'] == 'Outros')
        
        # Calcular subtotais e total
        sap1 = encargo + reativa + desc_encargo
        sap2 = demanda + desc_demanda
        total = sap1 + sap2 + contribuicao_publica + outros
        
        return {
            'sucesso': True,
            'distribuidora': distribuidora,
            'instalacao': instalacao,
            'referencia': referencia,
            'itens': itens,
            'encargo': encargo,
            'reativa': reativa,
            'desc_encargo': desc_encargo,
            'sap1': sap1,
            'demanda': demanda,
            'desc_demanda': desc_demanda,
            'sap2': sap2,
            'contribuicao_publica': contribuicao_publica,
            'outros': outros,
            'total': total
        }
    
    except Exception as e:
        logger.error(f"Erro ao processar PDF {pdf_path}: {str(e)}")
        return {
            'sucesso': False,
            'erro': str(e)
        }

def atualizar_excel(excel_path, resultados):
    """Atualiza o arquivo Excel com os resultados"""
    try:
        # Verificar se o arquivo existe
        if os.path.exists(excel_path):
            df = pd.read_excel(excel_path)
        else:
            # Criar um novo DataFrame com as colunas necessárias
            df = pd.DataFrame(columns=[
                'Ref', 'Instalação', 'Encargo', 'Reativa', 'Desc. Encargo', 'SAP1',
                'Demanda', 'Desc. Demanda', 'SAP2', 'Contribuição Pública', 'Total', 'Outros'
            ])
        
        # Adicionar novas linhas
        for resultado in resultados:
            if resultado['sucesso']:
                nova_linha = {
                    'Ref': resultado['referencia'],
                    'Instalação': resultado['instalacao'],
                    'Encargo': resultado['encargo'],
                    'Reativa': resultado['reativa'],
                    'Desc. Encargo': resultado['desc_encargo'],
                    'SAP1': resultado['sap1'],
                    'Demanda': resultado['demanda'],
                    'Desc. Demanda': resultado['desc_demanda'],
                    'SAP2': resultado['sap2'],
                    'Contribuição Pública': resultado['contribuicao_publica'],
                    'Total': resultado['total'],
                    'Outros': resultado['outros']
                }
                df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
        
        # Salvar o DataFrame atualizado
        df.to_excel(excel_path, index=False)
        return True
    
    except Exception as e:
        logger.error(f"Erro ao atualizar Excel: {str(e)}")
        return False

def gerar_relatorio_erros(erros, output_path):
    """Gera um relatório de erros em formato JSON"""
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(erros, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        logger.error(f"Erro ao gerar relatório: {str(e)}")
        return False
