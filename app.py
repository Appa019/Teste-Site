import streamlit as st
import pandas as pd
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import base64
from datetime import datetime
import cv2
import numpy as np
import tensorflow as tf
import io

# Configuração da página Streamlit
st.set_page_config(
    page_title="Consulta de Faturas LIGHT",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo CSS personalizado
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #0066cc;
        text-align: center;
        margin-bottom: 2rem;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #444;
        margin-bottom: 1rem;
    }
    .info-box {
        background-color: #f0f7ff;
        padding: 1rem;
        border-radius: 5px;
        border-left: 5px solid #0066cc;
        margin-bottom: 1rem;
    }
    .success-box {
        background-color: #e6ffe6;
        padding: 1rem;
        border-radius: 5px;
        border-left: 5px solid #00cc66;
        margin-bottom: 1rem;
    }
    .warning-box {
        background-color: #fff8e6;
        padding: 1rem;
        border-radius: 5px;
        border-left: 5px solid #ffcc00;
        margin-bottom: 1rem;
    }
    .error-box {
        background-color: #ffe6e6;
        padding: 1rem;
        border-radius: 5px;
        border-left: 5px solid #cc0000;
        margin-bottom: 1rem;
    }
    .stButton>button {
        background-color: #0066cc;
        color: white;
        font-weight: bold;
        border: none;
        border-radius: 5px;
        padding: 0.5rem 1rem;
        width: 100%;
    }
    .stButton>button:hover {
        background-color: #004c99;
    }
</style>
""", unsafe_allow_html=True)

# Título e descrição
st.markdown("<h1 class='main-header'>Consulta de Faturas LIGHT</h1>", unsafe_allow_html=True)
st.markdown("<div class='info-box'>Este sistema automatiza a consulta de faturas em aberto no portal da LIGHT para os códigos de instalação fornecidos.</div>", unsafe_allow_html=True)

# Função para tentar resolver o captcha
def solve_captcha(driver):
    try:
        # Tentativa de encontrar e clicar no checkbox do reCAPTCHA
        WebDriverWait(driver, 5).until(
            EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe[title='reCAPTCHA']"))
        )
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "div.recaptcha-checkbox-border"))
        ).click()
        driver.switch_to.default_content()
        
        # Aguardar para ver se o captcha foi resolvido
        time.sleep(2)
        
        # Verificar se o captcha foi resolvido
        try:
            WebDriverWait(driver, 3).until(
                EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe[title='reCAPTCHA']"))
            )
            checkbox = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.recaptcha-checkbox-checkmark"))
            )
            driver.switch_to.default_content()
            return True
        except:
            driver.switch_to.default_content()
            return False
    except Exception as e:
        print(f"Erro ao tentar resolver captcha: {e}")
        return False

# Função para fazer login no portal da LIGHT
def login_light_portal(driver, login, senha):
    try:
        # Acessar o site da LIGHT
        driver.get("https://agenciavirtual.light.com.br/Portal/MainFlow.Login_Agencia.aspx")
        
        # Preencher login e senha
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_ctl00_ConteudoPagina_ConteudoPagina_txtLogin"))
        ).send_keys(login)
        
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_ctl00_ConteudoPagina_ConteudoPagina_txtSenha"))
        ).send_keys(senha)
        
        # Tentar resolver o captcha automaticamente
        captcha_solved = solve_captcha(driver)
        
        if not captcha_solved:
            # Se não conseguir resolver automaticamente, solicitar intervenção do usuário
            st.warning("Não foi possível resolver o captcha automaticamente. Por favor, resolva o captcha manualmente na janela do navegador.")
            # Aguardar intervenção do usuário
            captcha_solved_manually = st.checkbox("Marquei o captcha manualmente")
            if not captcha_solved_manually:
                return False
        
        # Clicar no botão de entrar
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ctl00_ConteudoPagina_ConteudoPagina_btnEntrar"))
        ).click()
        
        # Verificar se o login foi bem-sucedido
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ctl00_ctl00_ConteudoPagina_ConteudoPagina_ddlEmpresa"))
            )
            return True
        except:
            st.error("Falha no login. Verifique as credenciais ou tente novamente.")
            return False
            
    except Exception as e:
        st.error(f"Erro durante o login: {e}")
        return False

# Função para selecionar a filial
def selecionar_filial(driver, filial):
    try:
        # Mapear os valores da coluna Filial para as opções do dropdown
        filial_mapping = {
            "CSN": "Companhia Siderúrgica Nacional",
            "CMIN": "Congonhas Minérios S.A.",
            "CIBR": "CSN Cimentos Brasil"
        }
        
        # Selecionar a empresa correspondente
        dropdown = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_ctl00_ConteudoPagina_ConteudoPagina_ddlEmpresa"))
        )
        
        from selenium.webdriver.support.ui import Select
        select = Select(dropdown)
        
        # Encontrar a opção que contém o texto mapeado
        for option in select.options:
            if filial_mapping[filial] in option.text:
                select.select_by_visible_text(option.text)
                break
        
        # Clicar no botão de confirmar
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ctl00_ConteudoPagina_ConteudoPagina_btnConfirmar"))
        ).click()
        
        # Aguardar carregamento da página
        time.sleep(2)
        return True
        
    except Exception as e:
        st.error(f"Erro ao selecionar filial: {e}")
        return False

# Função para acessar a segunda via da conta
def acessar_segunda_via(driver):
    try:
        # Clicar no menu de segunda via
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Segunda Via')]"))
        ).click()
        
        # Aguardar carregamento da página
        time.sleep(2)
        return True
        
    except Exception as e:
        st.error(f"Erro ao acessar segunda via: {e}")
        return False

# Função para consultar faturas por código de instalação
def consultar_faturas(driver, codigo_instalacao):
    try:
        # Limpar campo de código de instalação
        campo_codigo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_ctl00_ConteudoPagina_ConteudoPagina_txtInstalacao"))
        )
        campo_codigo.clear()
        campo_codigo.send_keys(str(codigo_instalacao))
        
        # Clicar no botão de consultar
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ctl00_ConteudoPagina_ConteudoPagina_btnConsultar"))
        ).click()
        
        # Aguardar carregamento dos resultados
        time.sleep(3)
        
        # Verificar se existem faturas em aberto
        try:
            tabela_faturas = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "ctl00_ctl00_ConteudoPagina_ConteudoPagina_gvFaturas"))
            )
            
            # Extrair dados das faturas
            linhas = tabela_faturas.find_elements(By.TAG_NAME, "tr")
            faturas = []
            
            # Pular a primeira linha (cabeçalho)
            for i in range(1, len(linhas)):
                colunas = linhas[i].find_elements(By.TAG_NAME, "td")
                if len(colunas) >= 4:
                    fatura = {
                        "Mês Referência": colunas[0].text,
                        "Data Vencimento": colunas[1].text,
                        "Valor": colunas[2].text,
                        "Status": "Em aberto"
                    }
                    faturas.append(fatura)
                    
                    # Tentar fazer download da fatura
                    try:
                        # Clicar no botão de download
                        colunas[3].find_element(By.TAG_NAME, "a").click()
                        time.sleep(2)
                        
                        # Selecionar motivo (perda ou esquecimento)
                        try:
                            WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.ID, "ctl00_ctl00_ConteudoPagina_ConteudoPagina_rblMotivo_0"))
                            ).click()
                            
                            # Confirmar download
                            WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.ID, "ctl00_ctl00_ConteudoPagina_ConteudoPagina_btnConfirmar"))
                            ).click()
                            
                            time.sleep(2)
                            
                            # Se aparecer a opção "Não, obrigado", clicar nela
                            try:
                                WebDriverWait(driver, 5).until(
                                    EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Não, obrigado')]"))
                                ).click()
                            except:
                                pass
                                
                        except:
                            pass
                    except:
                        pass
            
            return {"status": "success", "faturas": faturas}
            
        except:
            # Verificar se há mensagem de erro específica
            try:
                mensagem_erro = driver.find_element(By.ID, "ctl00_ctl00_ConteudoPagina_ConteudoPagina_lblMensagem").text
                if "não possui" in mensagem_erro.lower():
                    return {"status": "no_faturas", "mensagem": mensagem_erro}
                else:
                    return {"status": "error", "mensagem": mensagem_erro}
            except:
                return {"status": "no_faturas", "mensagem": "Não foram encontradas faturas para este código de instalação."}
                
    except Exception as e:
        return {"status": "error", "mensagem": f"Erro ao consultar faturas: {e}"}

# Função para processar a planilha
def processar_planilha(uploaded_file):
    try:
        # Ler a planilha
        df = pd.read_excel(uploaded_file)
        
        # Verificar se as colunas obrigatórias existem
        colunas_obrigatorias = ["FILIAL", "RAZAO SOCIAL", "NUM INSTALACAO 1"]
        for coluna in colunas_obrigatorias:
            if coluna not in df.columns:
                st.error(f"Coluna obrigatória '{coluna}' não encontrada na planilha.")
                return None
        
        # Configurar o driver do Chrome
        chrome_options = Options()
        chrome_options.add_argument("--headless")  # Modo headless
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        
        # Iniciar o driver
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        
        # Fazer login no portal da LIGHT
        login_success = login_light_portal(driver, df["LOGIN"].iloc[0], df["SENHA"].iloc[0])
        if not login_success:
            driver.quit()
            return None
        
        # Resultados
        resultados = []
        
        # Processar cada linha da planilha
        for index, row in df.iterrows():
            filial = row["FILIAL"]
            codigo_instalacao = row["NUM INSTALACAO 1"]
            
            # Selecionar a filial
            filial_success = selecionar_filial(driver, filial)
            if not filial_success:
                resultados.append({
                    "Filial": filial,
                    "Código Instalação": codigo_instalacao,
                    "Status": "Erro",
                    "Mensagem": "Falha ao selecionar filial"
                })
                continue
            
            # Acessar a segunda via
            segunda_via_success = acessar_segunda_via(driver)
            if not segunda_via_success:
                resultados.append({
                    "Filial": filial,
                    "Código Instalação": codigo_instalacao,
                    "Status": "Erro",
                    "Mensagem": "Falha ao acessar segunda via"
                })
                continue
            
            # Consultar faturas
            resultado_consulta = consultar_faturas(driver, codigo_instalacao)
            
            if resultado_consulta["status"] == "success":
                for fatura in resultado_consulta["faturas"]:
                    resultados.append({
                        "Filial": filial,
                        "Código Instalação": codigo_instalacao,
                        "Status": "Fatura em aberto",
                        "Mês Referência": fatura["Mês Referência"],
                        "Data Vencimento": fatura["Data Vencimento"],
                        "Valor": fatura["Valor"]
                    })
            elif resultado_consulta["status"] == "no_faturas":
                resultados.append({
                    "Filial": filial,
                    "Código Instalação": codigo_instalacao,
                    "Status": "Sem pendências",
                    "Mensagem": resultado_consulta.get("mensagem", "Não há faturas em aberto")
                })
            else:
                resultados.append({
                    "Filial": filial,
                    "Código Instalação": codigo_instalacao,
                    "Status": "Erro",
                    "Mensagem": resultado_consulta.get("mensagem", "Erro desconhecido")
                })
        
        # Fechar o driver
        driver.quit()
        
        # Converter resultados para DataFrame
        df_resultados = pd.DataFrame(resultados)
        
        return df_resultados
        
    except Exception as e:
        st.error(f"Erro ao processar planilha: {e}")
        return None

# Função para criar um link de download para o DataFrame
def get_table_download_link(df, filename, text):
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}" style="text-decoration:none;color:white;background-color:#0066cc;padding:0.5rem 1rem;border-radius:5px;font-weight:bold;">{text}</a>'
    return href

# Interface principal
def main():
    # Sidebar
    st.sidebar.markdown("<h2 class='sub-header'>Instruções</h2>", unsafe_allow_html=True)
    st.sidebar.markdown("""
    1. Faça upload da planilha com os códigos de instalação
    2. A planilha deve conter as colunas:
       - FILIAL (CSN, CMIN ou CIBR)
       - RAZAO SOCIAL
       - NUM INSTALACAO 1
    3. Clique em "Iniciar Consulta"
    4. Aguarde o processamento
    5. Baixe o relatório final
    """)
    
    # Upload da planilha
    st.markdown("<h2 class='sub-header'>Upload da Planilha</h2>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Selecione a planilha com os códigos de instalação", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        # Exibir preview da planilha
        df = pd.read_excel(uploaded_file)
        st.markdown("<h3 class='sub-header'>Preview da Planilha</h3>", unsafe_allow_html=True)
        st.dataframe(df.head())
        
        # Botão para iniciar consulta
        if st.button("Iniciar Consulta"):
            with st.spinner("Processando consultas... Isso pode levar alguns minutos."):
                # Resetar o ponteiro do arquivo para permitir nova leitura
                uploaded_file.seek(0)
                
                # Processar a planilha
                resultados = processar_planilha(uploaded_file)
                
                if resultados is not None:
                    # Exibir resultados
                    st.markdown("<h2 class='sub-header'>Resultados da Consulta</h2>", unsafe_allow_html=True)
                    st.dataframe(resultados)
                    
                    # Estatísticas
                    total_instalacoes = len(resultados["Código Instalação"].unique())
                    faturas_em_aberto = len(resultados[resultados["Status"] == "Fatura em aberto"])
                    sem_pendencias = len(resultados[resultados["Status"] == "Sem pendências"])
                    erros = len(resultados[resultados["Status"] == "Erro"])
                    
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("Total de Instalações", total_instalacoes)
                    col2.metric("Faturas em Aberto", faturas_em_aberto)
                    col3.metric("Sem Pendências", sem_pendencias)
                    col4.metric("Erros", erros)
                    
                    # Link para download do relatório
                    st.markdown("<h3 class='sub-header'>Download do Relatório</h3>", unsafe_allow_html=True)
                    data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"relatorio_light_{data_atual}.csv"
                    st.markdown(get_table_download_link(resultados, filename, "Baixar Relatório CSV"), unsafe_allow_html=True)
                    
                    # Exibir mensagem de conclusão
                    st.markdown("<div class='success-box'>Consulta concluída com sucesso!</div>", unsafe_allow_html=True)

# Executar a aplicação
if __name__ == "__main__":
    main()
