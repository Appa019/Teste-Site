#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Extrator Automático de Tabelas de Faturas de Energia
-----------------------------------------------------

Este script processa faturas de energia em PDF, extraindo automaticamente as tabelas
de itens de faturamento e classificando-os nas categorias apropriadas.

Características:
- Processamento em lote de múltiplos PDFs
- Desbloqueio automático de PDFs protegidos
- Detecção automática de tabelas usando palavras-chave
- Extração otimizada com uso eficiente da API OpenAI
- Classificação dos itens em categorias predefinidas
- Geração de tabela resumo consolidada
- Exportação para Excel com múltiplas abas

Autor: Manus AI
Data: Maio 2025
"""

import os
import re
import io
import json
import pandas as pd
import numpy as np
import PyPDF2
import pdfplumber
import pikepdf
import tempfile
import warnings
import logging
from datetime import datetime
from openai import OpenAI
from tqdm import tqdm

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("extrator_faturas.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Ignorar avisos específicos
warnings.filterwarnings("ignore", category=UserWarning)

# Lista de senhas conhecidas para tentar desbloquear PDFs protegidos
SENHAS_CONHECIDAS = ["33042", "56993", "60869", "08902", "3304", "5699", "6086", "0890"]

# Dicionário de mapeamento para classificação dos itens
CATEGORIAS = {
    "Encargo": [
        "CONSUMO ATIVO", "CONSUMO PONTA", "CONSUMO FORA PONTA", "ENERGIA ATIVA", 
        "ENERGIA PONTA", "ENERGIA FORA PONTA", "CONSUMO", "ENERGIA", "ATIVO", 
        "CONSUMO ATIVO PONTA", "CONSUMO ATIVO FORA PONTA", "ENERGIA ATIVA PONTA", 
        "ENERGIA ATIVA FORA PONTA", "CONSUMO ATIVO ULTRAPASSAGEM", "CONSUMO ATIVO UNICA",
        "CONSUMO ATIVO RESERVADO", "CONSUMO ATIVO CONTRATADO", "CONSUMO ATIVO MEDIDO",
        "ENERGIA ATIVA UNICA", "ENERGIA ATIVA RESERVADA", "ENERGIA ATIVA CONTRATADA",
        "ENERGIA ATIVA MEDIDA", "CONSUMO ATIVO VERDE", "CONSUMO ATIVO AZUL",
        "ENERGIA ATIVA VERDE", "ENERGIA ATIVA AZUL", "CONSUMO ATIVO HORARIO CAPACITIVO",
        "CONSUMO ATIVO HORARIO INDUTIVO", "ENERGIA ATIVA HORARIO CAPACITIVO",
        "ENERGIA ATIVA HORARIO INDUTIVO"
    ],
    "Desc. Encargo": [
        "DESCONTO CONSUMO ATIVO", "DESCONTO CONSUMO PONTA", "DESCONTO CONSUMO FORA PONTA",
        "DESCONTO ENERGIA ATIVA", "DESCONTO ENERGIA PONTA", "DESCONTO ENERGIA FORA PONTA",
        "DESCONTO CONSUMO", "DESCONTO ENERGIA", "DESCONTO ATIVO", "DEVOLUCAO CONSUMO ATIVO",
        "DEVOLUCAO CONSUMO PONTA", "DEVOLUCAO CONSUMO FORA PONTA", "DEVOLUCAO ENERGIA ATIVA",
        "DEVOLUCAO ENERGIA PONTA", "DEVOLUCAO ENERGIA FORA PONTA", "DEVOLUCAO CONSUMO",
        "DEVOLUCAO ENERGIA", "DEVOLUCAO ATIVO", "AJUSTE CONSUMO ATIVO", "AJUSTE CONSUMO PONTA",
        "AJUSTE CONSUMO FORA PONTA", "AJUSTE ENERGIA ATIVA", "AJUSTE ENERGIA PONTA",
        "AJUSTE ENERGIA FORA PONTA", "AJUSTE CONSUMO", "AJUSTE ENERGIA", "AJUSTE ATIVO",
        "CREDITO CONSUMO ATIVO", "CREDITO CONSUMO PONTA", "CREDITO CONSUMO FORA PONTA",
        "CREDITO ENERGIA ATIVA", "CREDITO ENERGIA PONTA", "CREDITO ENERGIA FORA PONTA",
        "CREDITO CONSUMO", "CREDITO ENERGIA", "CREDITO ATIVO"
    ],
    "Reativa": [
        "CONSUMO REATIVO", "CONSUMO REATIVO PONTA", "CONSUMO REATIVO FORA PONTA",
        "ENERGIA REATIVA", "ENERGIA REATIVA PONTA", "ENERGIA REATIVA FORA PONTA",
        "REATIVO", "CONSUMO REATIVO EXCEDENTE", "ENERGIA REATIVA EXCEDENTE",
        "CONSUMO REATIVO ULTRAPASSAGEM", "ENERGIA REATIVA ULTRAPASSAGEM",
        "CONSUMO REATIVO UNICA", "ENERGIA REATIVA UNICA", "CONSUMO REATIVO RESERVADO",
        "ENERGIA REATIVA RESERVADA", "CONSUMO REATIVO CONTRATADO", "ENERGIA REATIVA CONTRATADA",
        "CONSUMO REATIVO MEDIDO", "ENERGIA REATIVA MEDIDA", "CONSUMO REATIVO VERDE",
        "ENERGIA REATIVA VERDE", "CONSUMO REATIVO AZUL", "ENERGIA REATIVA AZUL",
        "CONSUMO REATIVO HORARIO CAPACITIVO", "ENERGIA REATIVA HORARIO CAPACITIVO",
        "CONSUMO REATIVO HORARIO INDUTIVO", "ENERGIA REATIVA HORARIO INDUTIVO",
        "DEMANDA REATIVA", "DEMANDA REATIVA PONTA", "DEMANDA REATIVA FORA PONTA",
        "DEMANDA REATIVA EXCEDENTE", "DEMANDA REATIVA ULTRAPASSAGEM", "DEMANDA REATIVA UNICA",
        "DEMANDA REATIVA RESERVADA", "DEMANDA REATIVA CONTRATADA", "DEMANDA REATIVA MEDIDA",
        "DEMANDA REATIVA VERDE", "DEMANDA REATIVA AZUL", "DEMANDA REATIVA HORARIO CAPACITIVO",
        "DEMANDA REATIVA HORARIO INDUTIVO"
    ],
    "Demanda": [
        "DEMANDA ATIVA", "DEMANDA PONTA", "DEMANDA FORA PONTA", "DEMANDA",
        "DEMANDA ATIVA PONTA", "DEMANDA ATIVA FORA PONTA", "DEMANDA ATIVA UNICA",
        "DEMANDA ATIVA RESERVADA", "DEMANDA ATIVA CONTRATADA", "DEMANDA ATIVA MEDIDA",
        "DEMANDA ATIVA VERDE", "DEMANDA ATIVA AZUL", "DEMANDA ATIVA HORARIO CAPACITIVO",
        "DEMANDA ATIVA HORARIO INDUTIVO", "DEMANDA ULTRAPASSAGEM", "DEMANDA ATIVA ULTRAPASSAGEM",
        "DEMANDA PONTA ULTRAPASSAGEM", "DEMANDA FORA PONTA ULTRAPASSAGEM",
        "ULTRAPASSAGEM DEMANDA", "ULTRAPASSAGEM DEMANDA ATIVA", "ULTRAPASSAGEM DEMANDA PONTA",
        "ULTRAPASSAGEM DEMANDA FORA PONTA"
    ],
    "Desc. Demanda": [
        "DESCONTO DEMANDA ATIVA", "DESCONTO DEMANDA PONTA", "DESCONTO DEMANDA FORA PONTA",
        "DESCONTO DEMANDA", "DEVOLUCAO DEMANDA ATIVA", "DEVOLUCAO DEMANDA PONTA",
        "DEVOLUCAO DEMANDA FORA PONTA", "DEVOLUCAO DEMANDA", "AJUSTE DEMANDA ATIVA",
        "AJUSTE DEMANDA PONTA", "AJUSTE DEMANDA FORA PONTA", "AJUSTE DEMANDA",
        "CREDITO DEMANDA ATIVA", "CREDITO DEMANDA PONTA", "CREDITO DEMANDA FORA PONTA",
        "CREDITO DEMANDA"
    ],
    "Contb. Publica": [
        "CONTRIBUICAO ILUMINACAO PUBLICA", "CIP", "COSIP", "TAXA ILUMINACAO PUBLICA",
        "TAXA DE ILUMINACAO PUBLICA", "CONTRIBUICAO PARA CUSTEIO DO SERVICO DE ILUMINACAO PUBLICA",
        "CONTRIBUICAO PARA ILUMINACAO PUBLICA", "ILUMINACAO PUBLICA"
    ]
}

class ExtratorFaturas:
    """
    Classe principal para extração de tabelas de faturas de energia.
    """
    
    def __init__(self, api_key=None, diretorio_entrada=None, diretorio_saida=None):
        """
        Inicializa o extrator de faturas.
        
        Args:
            api_key (str): Chave da API OpenAI
            diretorio_entrada (str): Diretório contendo os PDFs a serem processados
            diretorio_saida (str): Diretório para salvar os resultados
        """
        self.api_key = api_key
        self.diretorio_entrada = diretorio_entrada or os.path.join(os.getcwd(), 'pdfs')
        self.diretorio_saida = diretorio_saida or os.path.join(os.getcwd(), 'resultados')
        self.cliente_openai = None
        self.resultados = []
        self.tabela_resumo = None
        
        # Criar diretórios se não existirem
        os.makedirs(self.diretorio_entrada, exist_ok=True)
        os.makedirs(self.diretorio_saida, exist_ok=True)
        
        if api_key:
            self.configurar_openai(api_key)
    
    def configurar_openai(self, api_key):
        """
        Configura o cliente da API OpenAI.
        
        Args:
            api_key (str): Chave da API OpenAI
        """
        self.api_key = api_key
        self.cliente_openai = OpenAI(api_key=api_key)
        logger.info("Cliente OpenAI configurado com sucesso")
    
    def desbloquear_pdf(self, caminho_pdf):
        """
        Tenta desbloquear um PDF protegido usando a lista de senhas conhecidas.
        
        Args:
            caminho_pdf (str): Caminho para o arquivo PDF
            
        Returns:
            str: Caminho para o PDF desbloqueado ou o original se não estiver protegido
        """
        try:
            # Verificar se o PDF está protegido
            with open(caminho_pdf, 'rb') as f:
                leitor = PyPDF2.PdfReader(f)
                if not leitor.is_encrypted:
                    return caminho_pdf
            
            # Tentar desbloquear com as senhas conhecidas
            for senha in SENHAS_CONHECIDAS:
                try:
                    pdf = pikepdf.open(caminho_pdf, password=senha)
                    caminho_desbloqueado = os.path.join(
                        os.path.dirname(caminho_pdf),
                        f"desbloqueado_{os.path.basename(caminho_pdf)}"
                    )
                    pdf.save(caminho_desbloqueado)
                    pdf.close()
                    logger.info(f"PDF desbloqueado com sucesso usando a senha: {senha}")
                    return caminho_desbloqueado
                except:
                    continue
            
            logger.warning(f"Não foi possível desbloquear o PDF: {caminho_pdf}")
            return caminho_pdf
        except Exception as e:
            logger.error(f"Erro ao tentar desbloquear o PDF: {str(e)}")
            return caminho_pdf
    
    def extrair_texto_pdf(self, caminho_pdf):
        """
        Extrai todo o texto de um arquivo PDF usando PyPDF2.
        
        Args:
            caminho_pdf (str): Caminho para o arquivo PDF
            
        Returns:
            str: Texto extraído do PDF
        """
        try:
            texto_completo = ""
            with open(caminho_pdf, 'rb') as f:
                leitor = PyPDF2.PdfReader(f)
                for pagina in leitor.pages:
                    texto = pagina.extract_text() or ""
                    texto_completo += texto + "\n\n"
            return texto_completo
        except Exception as e:
            logger.error(f"Erro ao extrair texto do PDF com PyPDF2: {str(e)}")
            # Tentar com pdfplumber como fallback
            return self.extrair_texto_com_pdfplumber(caminho_pdf)
    
    def extrair_texto_com_pdfplumber(self, caminho_pdf):
        """
        Extrai todo o texto de um arquivo PDF usando pdfplumber como fallback.
        
        Args:
            caminho_pdf (str): Caminho para o arquivo PDF
            
        Returns:
            str: Texto extraído do PDF
        """
        try:
            texto_completo = ""
            with pdfplumber.open(caminho_pdf) as pdf:
                for pagina in pdf.pages:
                    texto = pagina.extract_text() or ""
                    texto_completo += texto + "\n\n"
            return texto_completo
        except Exception as e:
            logger.error(f"Erro ao extrair texto do PDF com pdfplumber: {str(e)}")
            return ""
    
    def extrair_tabelas_com_pdfplumber(self, caminho_pdf):
        """
        Extrai tabelas de um arquivo PDF usando pdfplumber.
        
        Args:
            caminho_pdf (str): Caminho para o arquivo PDF
            
        Returns:
            list: Lista de DataFrames contendo as tabelas extraídas
        """
        try:
            tabelas = []
            with pdfplumber.open(caminho_pdf) as pdf:
                for pagina in pdf.pages:
                    try:
                        # Extrair tabelas da página
                        tabelas_pagina = pagina.extract_tables()
                        
                        # Converter para DataFrames
                        for tabela in tabelas_pagina:
                            if tabela and len(tabela) > 1:  # Pelo menos cabeçalho e uma linha
                                df = pd.DataFrame(tabela[1:], columns=tabela[0])
                                # Remover linhas vazias
                                df = df.replace('', np.nan)
                                df = df.dropna(how='all')
                                # Remover colunas vazias
                                df = df.replace('', np.nan)
                                df = df.dropna(axis=1, how='all')
                                # Resetar índices
                                df = df.reset_index(drop=True)
                                tabelas.append(df)
                    except Exception as e:
                        logger.warning(f"Erro ao extrair tabelas da página: {str(e)}")
                        continue
            
            return tabelas
        except Exception as e:
            logger.error(f"Erro ao extrair tabelas do PDF com pdfplumber: {str(e)}")
            return []
    
    def extrair_tabelas_com_openai(self, texto_pdf):
        """
        Usa a API OpenAI para extrair tabelas do texto do PDF.
        
        Args:
            texto_pdf (str): Texto completo do PDF
            
        Returns:
            list: Lista de DataFrames contendo as tabelas extraídas
        """
        try:
            # Verificar se o cliente OpenAI está configurado
            if not self.cliente_openai:
                logger.warning("Cliente OpenAI não configurado. Não é possível extrair tabelas com a API.")
                return []
            
            # Limitar o texto para não exceder o limite de tokens
            texto_limitado = texto_pdf
            if len(texto_limitado) > 12000:  # Aproximadamente 3000 tokens
                # Procurar por palavras-chave para encontrar a seção relevante
                padrao_inicio = re.compile(r'itens\s+d[ae]\s+fatura', re.IGNORECASE)
                padrao_fim = re.compile(r'tarifa\s+unit', re.IGNORECASE)
                
                match_inicio = padrao_inicio.search(texto_pdf)
                match_fim = padrao_fim.search(texto_pdf)
                
                if match_inicio and match_fim and match_inicio.start() < match_fim.start():
                    # Extrair a seção relevante com contexto
                    inicio = max(0, match_inicio.start() - 500)
                    fim = min(len(texto_pdf), match_fim.end() + 2000)
                    texto_limitado = texto_pdf[inicio:fim]
                else:
                    # Se não encontrar os padrões, usar os primeiros 12000 caracteres
                    texto_limitado = texto_pdf[:12000]
            
            # Criar o prompt para a API
            prompt = f"""
            Extraia a tabela de itens de faturamento da fatura de energia abaixo.
            A tabela geralmente começa após "Itens da Fatura" ou "Itens de Fatura" e termina antes de "Tarifa Unit".
            Ela contém descrições de serviços, quantidades, valores unitários e valores totais em R$.
            
            Texto da fatura:
            {texto_limitado}
            
            Responda com um JSON no formato:
            {{
                "tabela": [
                    {{"descricao": "CONSUMO ATIVO PONTA", "quantidade": "100", "unidade": "kWh", "tarifa": "0,50", "valor": "50,00"}},
                    {{"descricao": "CONSUMO ATIVO FORA PONTA", "quantidade": "200", "unidade": "kWh", "tarifa": "0,30", "valor": "60,00"}},
                    ...
                ]
            }}
            
            Inclua apenas os itens da tabela de faturamento, não inclua totais ou outras informações.
            """
            
            # Chamar a API
            resposta = self.cliente_openai.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "Você é um assistente especializado em análise de faturas de energia."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=1000
            )
            
            # Extrair a tabela da resposta
            texto_resposta = resposta.choices[0].message.content.strip()
            
            # Extrair o JSON da resposta
            match = re.search(r'\{.*\}', texto_resposta, re.DOTALL)
            if match:
                try:
                    dados = json.loads(match.group(0))
                    itens = dados.get('tabela', [])
                    
                    if itens:
                        # Criar DataFrame
                        df = pd.DataFrame(itens)
                        return [df]
                except json.JSONDecodeError:
                    logger.warning("Erro ao decodificar JSON da resposta da OpenAI")
            
            return []
        except Exception as e:
            logger.error(f"Erro ao extrair tabelas com OpenAI: {str(e)}")
            return []
    
    def identificar_tabela_itens_fatura(self, tabelas, texto_pdf):
        """
        Identifica a tabela de itens de fatura entre as tabelas extraídas.
        
        Args:
            tabelas (list): Lista de DataFrames com as tabelas extraídas
            texto_pdf (str): Texto completo do PDF
            
        Returns:
            pd.DataFrame: DataFrame contendo a tabela de itens de fatura, ou None se não encontrada
        """
        try:
            # Verificar se há tabelas extraídas
            if not tabelas:
                # Se não há tabelas extraídas, tentar extrair com OpenAI
                if self.cliente_openai:
                    tabelas = self.extrair_tabelas_com_openai(texto_pdf)
                
                # Se ainda não há tabelas, retornar None
                if not tabelas:
                    return None
            
            # Procurar por palavras-chave no texto do PDF
            padrao_inicio = re.compile(r'itens\s+d[ae]\s+fatura', re.IGNORECASE)
            padrao_fim = re.compile(r'tarifa\s+unit', re.IGNORECASE)
            
            # Se encontrar os padrões, procurar a tabela correspondente
            if padrao_inicio.search(texto_pdf) and padrao_fim.search(texto_pdf):
                # Verificar cada tabela
                for tabela in tabelas:
                    # Converter a tabela para string para facilitar a busca
                    tabela_str = tabela.to_string().lower()
                    
                    # Verificar se a tabela contém valores monetários (R$)
                    if 'r$' in tabela_str:
                        # Verificar se a tabela tem pelo menos 3 colunas (descrição, quantidade, valor)
                        if tabela.shape[1] >= 3:
                            return tabela
            
            # Se não encontrar usando os padrões, usar a API OpenAI para identificar
            if self.cliente_openai:
                return self.identificar_tabela_com_openai(tabelas, texto_pdf)
            
            # Se não encontrar, retornar a maior tabela como fallback
            if tabelas:
                return max(tabelas, key=lambda df: df.size)
            
            return None
        except Exception as e:
            logger.error(f"Erro ao identificar tabela de itens de fatura: {str(e)}")
            return None
    
    def identificar_tabela_com_openai(self, tabelas, texto_pdf):
        """
        Usa a API OpenAI para identificar a tabela de itens de fatura.
        
        Args:
            tabelas (list): Lista de DataFrames com as tabelas extraídas
            texto_pdf (str): Texto completo do PDF
            
        Returns:
            pd.DataFrame: DataFrame contendo a tabela de itens de fatura, ou None se não encontrada
        """
        try:
            # Preparar as tabelas para envio à API
            tabelas_texto = []
            for i, df in enumerate(tabelas):
                tabelas_texto.append(f"Tabela {i+1}:\n{df.to_string()}")
            
            # Limitar o texto para não exceder o limite de tokens
            texto_combinado = "\n\n".join(tabelas_texto)
            if len(texto_combinado) > 12000:  # Aproximadamente 3000 tokens
                texto_combinado = texto_combinado[:12000] + "..."
            
            # Criar o prompt para a API
            prompt = f"""
            Analise as tabelas extraídas de uma fatura de energia e identifique qual delas é a tabela de itens de faturamento.
            A tabela de itens de faturamento geralmente contém descrições de serviços, valores em R$, e pode ter palavras como "Consumo", "Demanda", "Energia", etc.
            
            {texto_combinado}
            
            Responda apenas com o número da tabela que contém os itens de faturamento (ex: "Tabela 3").
            """
            
            # Chamar a API
            resposta = self.cliente_openai.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "Você é um assistente especializado em análise de faturas de energia."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=50
            )
            
            # Extrair o número da tabela da resposta
            texto_resposta = resposta.choices[0].message.content.strip()
            match = re.search(r'tabela\s+(\d+)', texto_resposta, re.IGNORECASE)
            if match:
                indice_tabela = int(match.group(1)) - 1
                if 0 <= indice_tabela < len(tabelas):
                    return tabelas[indice_tabela]
            
            # Se não conseguir identificar, retornar a maior tabela
            return max(tabelas, key=lambda df: df.size)
        except Exception as e:
            logger.error(f"Erro ao identificar tabela com OpenAI: {str(e)}")
            # Fallback para a maior tabela
            if tabelas:
                return max(tabelas, key=lambda df: df.size)
            return None
    
    def extrair_informacoes_cabecalho(self, texto_pdf):
        """
        Extrai informações do cabeçalho da fatura, como número de instalação e data de referência.
        
        Args:
            texto_pdf (str): Texto completo do PDF
            
        Returns:
            dict: Dicionário com as informações extraídas
        """
        try:
            info = {
                "num_instalacao": None,
                "data_referencia": None,
                "distribuidora": None
            }
            
            # Extrair número de instalação
            padrao_instalacao = re.compile(r'(instalação|instalacao|n[°º\.]?\s*(?:da)?\s*instala[çc][ãa]o)\s*[:\.]?\s*(\d[\d\.\-\/]*\d)', re.IGNORECASE)
            match_instalacao = padrao_instalacao.search(texto_pdf)
            if match_instalacao:
                info["num_instalacao"] = match_instalacao.group(2).strip()
            
            # Extrair data de referência
            padrao_referencia = re.compile(r'(referente|referência|referencia|ref\.)\s*[:\.]?\s*((?:jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[\/\.\-\s]+\d{2,4})', re.IGNORECASE)
            match_referencia = padrao_referencia.search(texto_pdf)
            if match_referencia:
                info["data_referencia"] = match_referencia.group(2).strip()
            
            # Extrair distribuidora
            distribuidoras = [
                "AMPLA", "LIGHT", "ENEL", "CEMIG", "CPFL", "ELEKTRO", "EDP", "ENERGISA",
                "EQUATORIAL", "NEOENERGIA", "COPEL", "CELESC", "CEEE", "COELBA", "CELPE",
                "COSERN", "COELCE", "CEMAR", "CEPISA", "CERON", "ELETROACRE", "CEAL",
                "CELG", "CEB", "CEMAT", "ENERGISA SUL-SUDESTE", "ENERGISA MINAS GERAIS",
                "ENERGISA NOVA FRIBURGO", "ENERGISA BORBOREMA", "ENERGISA PARAÍBA",
                "ENERGISA SERGIPE", "ENERGISA MATO GROSSO", "ENERGISA MATO GROSSO DO SUL",
                "ENERGISA TOCANTINS", "ENERGISA RONDÔNIA", "ENERGISA ACRE"
            ]
            
            for distribuidora in distribuidoras:
                if re.search(distribuidora, texto_pdf, re.IGNORECASE):
                    info["distribuidora"] = distribuidora
                    break
            
            return info
        except Exception as e:
            logger.error(f"Erro ao extrair informações do cabeçalho: {str(e)}")
            return {"num_instalacao": None, "data_referencia": None, "distribuidora": None}
    
    def classificar_itens_fatura(self, tabela):
        """
        Classifica os itens da fatura nas categorias predefinidas.
        
        Args:
            tabela (pd.DataFrame): DataFrame contendo a tabela de itens de fatura
            
        Returns:
            pd.DataFrame: DataFrame com os itens classificados
        """
        try:
            # Verificar se a tabela existe
            if tabela is None or tabela.empty:
                return None
            
            # Criar cópia da tabela para não modificar a original
            df = tabela.copy()
            
            # Adicionar coluna de categoria
            df['Categoria'] = 'Outros'
            
            # Identificar a coluna de descrição (geralmente a primeira coluna)
            coluna_descricao = None
            for i, coluna in enumerate(df.columns):
                # Verificar se a coluna contém descrições de itens
                if df[coluna].astype(str).str.contains('CONSUMO|DEMANDA|ENERGIA', case=False).any():
                    coluna_descricao = i
                    break
            
            # Se não encontrou, usar a primeira coluna
            if coluna_descricao is None:
                coluna_descricao = 0
            
            # Identificar a coluna de valor (geralmente contém "R$" ou valores numéricos)
            coluna_valor = None
            for i, coluna in enumerate(df.columns):
                # Verificar se a coluna contém "R$" ou valores monetários
                if df[coluna].astype(str).str.contains('R\$|r\$').any():
                    coluna_valor = i
                    break
            
            # Se não encontrou a coluna de valor, assumir a última coluna
            if coluna_valor is None:
                coluna_valor = df.shape[1] - 1
            
            # Classificar os itens
            for i, row in df.iterrows():
                descricao = str(row[coluna_descricao]).upper()
                
                # Verificar cada categoria
                for categoria, termos in CATEGORIAS.items():
                    for termo in termos:
                        if termo in descricao:
                            df.at[i, 'Categoria'] = categoria
                            break
                    
                    # Se já encontrou uma categoria, não precisa verificar as demais
                    if df.at[i, 'Categoria'] != 'Outros':
                        break
            
            return df
        except Exception as e:
            logger.error(f"Erro ao classificar itens da fatura: {str(e)}")
            return tabela
    
    def classificar_itens_com_openai(self, tabela):
        """
        Usa a API OpenAI para classificar os itens da fatura.
        
        Args:
            tabela (pd.DataFrame): DataFrame contendo a tabela de itens de fatura
            
        Returns:
            pd.DataFrame: DataFrame com os itens classificados
        """
        try:
            # Verificar se a tabela existe e se o cliente OpenAI está configurado
            if tabela is None or tabela.empty or self.cliente_openai is None:
                return self.classificar_itens_fatura(tabela)
            
            # Criar cópia da tabela para não modificar a original
            df = tabela.copy()
            
            # Converter a tabela para texto
            tabela_texto = df.to_string()
            
            # Criar o prompt para a API
            categorias_texto = json.dumps(CATEGORIAS, indent=2)
            prompt = f"""
            Analise a tabela de itens de faturamento abaixo e classifique cada item em uma das seguintes categorias:
            {categorias_texto}
            
            Se um item não se encaixar em nenhuma categoria, classifique como "Outros".
            
            Tabela:
            {tabela_texto}
            
            Responda com um JSON no formato:
            {{
                "classificacoes": [
                    {{"linha": 0, "categoria": "Encargo"}},
                    {{"linha": 1, "categoria": "Demanda"}},
                    ...
                ]
            }}
            """
            
            # Chamar a API
            resposta = self.cliente_openai.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "Você é um assistente especializado em análise de faturas de energia."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=1000
            )
            
            # Extrair as classificações da resposta
            texto_resposta = resposta.choices[0].message.content.strip()
            
            # Extrair o JSON da resposta
            match = re.search(r'\{.*\}', texto_resposta, re.DOTALL)
            if match:
                try:
                    dados = json.loads(match.group(0))
                    classificacoes = dados.get('classificacoes', [])
                    
                    # Adicionar coluna de categoria
                    df['Categoria'] = 'Outros'
                    
                    # Aplicar as classificações
                    for item in classificacoes:
                        linha = item.get('linha')
                        categoria = item.get('categoria')
                        if linha is not None and categoria:
                            if 0 <= linha < len(df):
                                df.at[linha, 'Categoria'] = categoria
                    
                    return df
                except json.JSONDecodeError:
                    logger.warning("Erro ao decodificar JSON da resposta da OpenAI")
            
            # Fallback para classificação local
            return self.classificar_itens_fatura(tabela)
        except Exception as e:
            logger.error(f"Erro ao classificar itens com OpenAI: {str(e)}")
            return self.classificar_itens_fatura(tabela)
    
    def criar_tabela_resumo(self):
        """
        Cria uma tabela resumo com os valores agrupados por categoria.
        
        Returns:
            pd.DataFrame: DataFrame contendo a tabela resumo
        """
        try:
            # Verificar se há resultados
            if not self.resultados:
                return None
            
            # Criar lista para armazenar os dados do resumo
            dados_resumo = []
            
            # Processar cada resultado
            for resultado in self.resultados:
                # Extrair informações
                nome_arquivo = resultado.get('nome_arquivo', '')
                num_instalacao = resultado.get('info_cabecalho', {}).get('num_instalacao', '')
                data_referencia = resultado.get('info_cabecalho', {}).get('data_referencia', '')
                tabela_classificada = resultado.get('tabela_classificada')
                
                # Verificar se a tabela existe
                if tabela_classificada is None or tabela_classificada.empty:
                    continue
                
                # Identificar a coluna de valor
                coluna_valor = None
                for coluna in tabela_classificada.columns:
                    # Verificar se a coluna contém "R$" ou valores monetários
                    if tabela_classificada[coluna].astype(str).str.contains('R\$|r\$').any():
                        coluna_valor = coluna
                        break
                
                # Se não encontrou a coluna de valor, tentar a última coluna
                if coluna_valor is None and tabela_classificada.shape[1] > 0:
                    coluna_valor = tabela_classificada.columns[-1]
                
                # Se ainda não encontrou, pular este resultado
                if coluna_valor is None:
                    continue
                
                # Inicializar valores por categoria
                valores_categoria = {
                    'Encargo': 0.0,
                    'Desc. Encargo': 0.0,
                    'Reativa': 0.0,
                    'Demanda': 0.0,
                    'Desc. Demanda': 0.0,
                    'Contb. Publica': 0.0,
                    'Outros': 0.0
                }
                
                # Somar valores por categoria
                for _, row in tabela_classificada.iterrows():
                    categoria = row.get('Categoria', 'Outros')
                    valor_str = str(row.get(coluna_valor, '0')).replace('R$', '').replace(' ', '').strip()
                    
                    # Converter para float
                    try:
                        # Substituir vírgula por ponto para conversão
                        valor_str = valor_str.replace('.', '').replace(',', '.')
                        valor = float(valor_str)
                        valores_categoria[categoria] += valor
                    except ValueError:
                        continue
                
                # Adicionar linha ao resumo
                dados_resumo.append({
                    'Nome Arquivo': nome_arquivo,
                    'Num. Instalacao': num_instalacao,
                    'Data Referencia': data_referencia,
                    'Encargo': valores_categoria['Encargo'],
                    'Desc. Encargo': valores_categoria['Desc. Encargo'],
                    'Reativa': valores_categoria['Reativa'],
                    'Demanda': valores_categoria['Demanda'],
                    'Desc. Demanda': valores_categoria['Desc. Demanda'],
                    'Contb. Publica': valores_categoria['Contb. Publica'],
                    'Outros': valores_categoria['Outros']
                })
            
            # Criar DataFrame com o resumo
            if dados_resumo:
                df_resumo = pd.DataFrame(dados_resumo)
                self.tabela_resumo = df_resumo
                return df_resumo
            
            return None
        except Exception as e:
            logger.error(f"Erro ao criar tabela resumo: {str(e)}")
            return None
    
    def processar_pdf(self, caminho_pdf):
        """
        Processa um arquivo PDF, extraindo e classificando a tabela de itens de fatura.
        
        Args:
            caminho_pdf (str): Caminho para o arquivo PDF
            
        Returns:
            dict: Dicionário com os resultados do processamento
        """
        try:
            logger.info(f"Processando arquivo: {caminho_pdf}")
            nome_arquivo = os.path.basename(caminho_pdf)
            
            # Desbloquear o PDF se necessário
            caminho_pdf_desbloqueado = self.desbloquear_pdf(caminho_pdf)
            
            # Extrair texto do PDF
            texto_pdf = self.extrair_texto_pdf(caminho_pdf_desbloqueado)
            
            # Extrair informações do cabeçalho
            info_cabecalho = self.extrair_informacoes_cabecalho(texto_pdf)
            logger.info(f"Informações do cabeçalho: {info_cabecalho}")
            
            # Extrair tabelas do PDF
            tabelas = self.extrair_tabelas_com_pdfplumber(caminho_pdf_desbloqueado)
            
            # Se não encontrou tabelas com pdfplumber, tentar com OpenAI
            if not tabelas and self.cliente_openai:
                tabelas = self.extrair_tabelas_com_openai(texto_pdf)
            
            # Identificar a tabela de itens de fatura
            tabela_itens = self.identificar_tabela_itens_fatura(tabelas, texto_pdf)
            
            # Se não encontrou a tabela, retornar resultado vazio
            if tabela_itens is None:
                logger.warning(f"Não foi possível identificar a tabela de itens de fatura em: {caminho_pdf}")
                return {
                    'nome_arquivo': nome_arquivo,
                    'info_cabecalho': info_cabecalho,
                    'tabela_original': None,
                    'tabela_classificada': None,
                    'erro': "Tabela de itens de fatura não encontrada"
                }
            
            # Classificar os itens da tabela
            if self.cliente_openai:
                tabela_classificada = self.classificar_itens_com_openai(tabela_itens)
            else:
                tabela_classificada = self.classificar_itens_fatura(tabela_itens)
            
            # Retornar os resultados
            resultado = {
                'nome_arquivo': nome_arquivo,
                'info_cabecalho': info_cabecalho,
                'tabela_original': tabela_itens,
                'tabela_classificada': tabela_classificada,
                'erro': None
            }
            
            # Adicionar aos resultados
            self.resultados.append(resultado)
            
            return resultado
        except Exception as e:
            logger.error(f"Erro ao processar o PDF {caminho_pdf}: {str(e)}")
            return {
                'nome_arquivo': os.path.basename(caminho_pdf),
                'info_cabecalho': None,
                'tabela_original': None,
                'tabela_classificada': None,
                'erro': str(e)
            }
    
    def processar_pdfs(self):
        """
        Processa todos os PDFs no diretório de entrada.
        
        Returns:
            list: Lista de dicionários com os resultados do processamento
        """
        try:
            # Listar arquivos PDF no diretório de entrada
            arquivos_pdf = [os.path.join(self.diretorio_entrada, arquivo) for arquivo in os.listdir(self.diretorio_entrada) if arquivo.lower().endswith('.pdf')]
            
            if not arquivos_pdf:
                logger.warning(f"Nenhum arquivo PDF encontrado no diretório: {self.diretorio_entrada}")
                return []
            
            logger.info(f"Processando {len(arquivos_pdf)} arquivos PDF")
            
            # Processar cada arquivo
            self.resultados = []
            for arquivo in tqdm(arquivos_pdf, desc="Processando PDFs"):
                self.processar_pdf(arquivo)
            
            # Criar tabela resumo
            self.criar_tabela_resumo()
            
            return self.resultados
        except Exception as e:
            logger.error(f"Erro ao processar PDFs: {str(e)}")
            return []
    
    def exportar_para_excel(self, caminho_saida=None):
        """
        Exporta os resultados para um arquivo Excel.
        
        Args:
            caminho_saida (str): Caminho para o arquivo Excel de saída
            
        Returns:
            str: Caminho para o arquivo Excel gerado
        """
        try:
            # Verificar se há resultados
            if not self.resultados:
                logger.warning("Não há resultados para exportar")
                return None
            
            # Definir caminho de saída
            if caminho_saida is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                caminho_saida = os.path.join(self.diretorio_saida, f"resultados_{timestamp}.xlsx")
            
            # Criar o arquivo Excel
            with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
                # Adicionar tabela resumo
                if self.tabela_resumo is not None and not self.tabela_resumo.empty:
                    self.tabela_resumo.to_excel(writer, sheet_name='Resumo', index=False)
                
                # Adicionar cada tabela classificada em uma aba separada
                for resultado in self.resultados:
                    nome_arquivo = resultado.get('nome_arquivo', '')
                    tabela_classificada = resultado.get('tabela_classificada')
                    
                    # Limitar o nome da aba a 31 caracteres (limite do Excel)
                    nome_aba = nome_arquivo[:31].replace('.pdf', '')
                    
                    # Verificar se a tabela existe
                    if tabela_classificada is not None and not tabela_classificada.empty:
                        tabela_classificada.to_excel(writer, sheet_name=nome_aba, index=False)
            
            logger.info(f"Resultados exportados para: {caminho_saida}")
            return caminho_saida
        except Exception as e:
            logger.error(f"Erro ao exportar para Excel: {str(e)}")
            return None

def main():
    """
    Função principal para execução do script.
    """
    print("=" * 80)
    print("Extrator Automático de Tabelas de Faturas de Energia")
    print("=" * 80)
    print("\nEste script processa faturas de energia em PDF, extraindo automaticamente")
    print("as tabelas de itens de faturamento e classificando-os nas categorias apropriadas.")
    print("\nPara começar, você precisa fornecer sua chave da API OpenAI.")
    
    # Solicitar chave da API OpenAI
    api_key = input("\nDigite sua chave da API OpenAI: ").strip()
    
    # Verificar se a chave foi fornecida
    if not api_key:
        print("\nChave da API OpenAI não fornecida. O script será executado sem usar a API.")
        print("Isso pode reduzir a precisão da extração e classificação.")
    
    # Criar o extrator
    extrator = ExtratorFaturas(api_key=api_key)
    
    print("\nProcessando PDFs...")
    
    # Processar os PDFs
    resultados = extrator.processar_pdfs()
    
    # Verificar se há resultados
    if not resultados:
        print("\nNenhum resultado encontrado. Verifique os logs para mais detalhes.")
        return
    
    # Exportar para Excel
    caminho_excel = extrator.exportar_para_excel()
    
    if caminho_excel:
        print(f"\nResultados exportados com sucesso para: {caminho_excel}")
    else:
        print("\nErro ao exportar resultados. Verifique os logs para mais detalhes.")
    
    print("\nProcessamento concluído!")

if __name__ == "__main__":
    main()
