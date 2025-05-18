import streamlit as st
import os
import pandas as pd
import tempfile
import shutil
from datetime import datetime
import json
from utils import (
    desbloquear_pdf, 
    processar_pdf, 
    atualizar_excel, 
    gerar_relatorio_erros
)

# Configuração da página
st.set_page_config(
    page_title="Leitor de Faturas de Energia Elétrica",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo CSS personalizado
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1E88E5;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #0D47A1;
        margin-bottom: 1rem;
    }
    .success-message {
        background-color: #E8F5E9;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 5px solid #4CAF50;
    }
    .error-message {
        background-color: #FFEBEE;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 5px solid #F44336;
    }
    .info-box {
        background-color: #E3F2FD;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 5px solid #2196F3;
        margin-bottom: 1rem;
    }
    .warning-box {
        background-color: #FFF8E1;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 5px solid #FFC107;
        margin-bottom: 1rem;
    }
    .stButton button {
        background-color: #1E88E5;
        color: white;
        font-weight: bold;
    }
    .stButton button:hover {
        background-color: #0D47A1;
    }
</style>
""", unsafe_allow_html=True)

# Título principal
st.markdown("<h1 class='main-header'>Leitor de Faturas de Energia Elétrica</h1>", unsafe_allow_html=True)

# Inicializar variáveis de sessão
if 'resultados' not in st.session_state:
    st.session_state.resultados = []
if 'erros' not in st.session_state:
    st.session_state.erros = {
        'arquivos_falha': [],
        'tabelas_nao_reconhecidas': [],
        'itens_nao_mapeados': []
    }
if 'excel_path' not in st.session_state:
    st.session_state.excel_path = ""
if 'processamento_concluido' not in st.session_state:
    st.session_state.processamento_concluido = False
if 'temp_dir' not in st.session_state:
    st.session_state.temp_dir = tempfile.mkdtemp()

# Função para resetar o estado da aplicação
def resetar_aplicacao():
    st.session_state.resultados = []
    st.session_state.erros = {
        'arquivos_falha': [],
        'tabelas_nao_reconhecidas': [],
        'itens_nao_mapeados': []
    }
    st.session_state.excel_path = ""
    st.session_state.processamento_concluido = False
    
    # Limpar diretório temporário
    if os.path.exists(st.session_state.temp_dir):
        shutil.rmtree(st.session_state.temp_dir)
    st.session_state.temp_dir = tempfile.mkdtemp()

# Sidebar com informações e opções
with st.sidebar:
    st.markdown("<h2 class='sub-header'>Informações</h2>", unsafe_allow_html=True)
    
    st.markdown("""
    <div class='info-box'>
        <p><strong>Leitor de Faturas</strong> é uma ferramenta para extração automática de dados de faturas de energia elétrica.</p>
        <p>Distribuidoras suportadas:</p>
        <ul>
            <li>Cemig</li>
            <li>Light</li>
            <li>Enel</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<h2 class='sub-header'>Opções</h2>", unsafe_allow_html=True)
    
    if st.button("Reiniciar Aplicação"):
        resetar_aplicacao()
        st.experimental_rerun()
    
    st.markdown("""
    <div class='warning-box'>
        <p><strong>Nota:</strong> Faturas protegidas por senha serão desbloqueadas automaticamente usando senhas conhecidas.</p>
    </div>
    """, unsafe_allow_html=True)

# Área principal
tab1, tab2, tab3 = st.tabs(["Upload de Arquivos", "Resultados", "Relatório de Erros"])

with tab1:
    st.markdown("<h2 class='sub-header'>Upload de Faturas</h2>", unsafe_allow_html=True)
    
    # Upload de arquivos PDF
    uploaded_files = st.file_uploader("Selecione as faturas em PDF", type="pdf", accept_multiple_files=True)
    
    # Campo para caminho do Excel
    excel_option = st.radio("Arquivo Excel para salvar os resultados:", ["Criar novo arquivo", "Atualizar arquivo existente"])
    
    if excel_option == "Criar novo arquivo":
        excel_filename = st.text_input("Nome do arquivo Excel a ser criado:", "resultados_faturas.xlsx")
        st.session_state.excel_path = os.path.join(os.getcwd(), excel_filename)
    else:
        excel_file = st.file_uploader("Selecione o arquivo Excel existente:", type=["xlsx", "xls"])
        if excel_file:
            # Salvar o arquivo Excel carregado no diretório temporário
            excel_temp_path = os.path.join(st.session_state.temp_dir, excel_file.name)
            with open(excel_temp_path, "wb") as f:
                f.write(excel_file.getbuffer())
            st.session_state.excel_path = excel_temp_path
    
    # Botão para processar
    if st.button("Processar Faturas"):
        if not uploaded_files:
            st.markdown("<div class='error-message'>Por favor, selecione pelo menos um arquivo PDF.</div>", unsafe_allow_html=True)
        elif not st.session_state.excel_path:
            st.markdown("<div class='error-message'>Por favor, defina o arquivo Excel para salvar os resultados.</div>", unsafe_allow_html=True)
        else:
            # Resetar resultados anteriores
            st.session_state.resultados = []
            st.session_state.erros = {
                'arquivos_falha': [],
                'tabelas_nao_reconhecidas': [],
                'itens_nao_mapeados': []
            }
            
            # Barra de progresso
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Processar cada arquivo
            for i, file in enumerate(uploaded_files):
                status_text.text(f"Processando {file.name}... ({i+1}/{len(uploaded_files)})")
                
                # Salvar o arquivo no diretório temporário
                pdf_path = os.path.join(st.session_state.temp_dir, file.name)
                with open(pdf_path, "wb") as f:
                    f.write(file.getbuffer())
                
                # Tentar desbloquear o PDF
                pdf_path, senha = desbloquear_pdf(pdf_path)
                
                if not pdf_path:
                    st.session_state.erros['arquivos_falha'].append({
                        'arquivo': file.name,
                        'erro': 'Não foi possível desbloquear o PDF'
                    })
                    continue
                
                # Processar o PDF
                resultado = processar_pdf(pdf_path)
                resultado['arquivo'] = file.name
                resultado['senha_usada'] = senha
                
                if not resultado['sucesso']:
                    st.session_state.erros['arquivos_falha'].append({
                        'arquivo': file.name,
                        'erro': resultado.get('erro', 'Erro desconhecido')
                    })
                    continue
                
                # Verificar se a tabela foi encontrada
                if 'itens' not in resultado or not resultado['itens']:
                    st.session_state.erros['tabelas_nao_reconhecidas'].append({
                        'arquivo': file.name,
                        'distribuidora': resultado.get('distribuidora', 'desconhecida')
                    })
                    continue
                
                # Registrar itens não mapeados
                for item in resultado.get('itens', []):
                    if item['categoria'] == 'Outros':
                        st.session_state.erros['itens_nao_mapeados'].append({
                            'arquivo': file.name,
                            'item': item['descricao']
                        })
                
                st.session_state.resultados.append(resultado)
                
                # Atualizar barra de progresso
                progress_bar.progress((i + 1) / len(uploaded_files))
            
            # Atualizar o Excel
            if st.session_state.resultados:
                sucesso = atualizar_excel(st.session_state.excel_path, st.session_state.resultados)
                
                if not sucesso:
                    st.markdown("<div class='error-message'>Erro ao atualizar o arquivo Excel.</div>", unsafe_allow_html=True)
                else:
                    st.session_state.processamento_concluido = True
                    
                    # Gerar relatório de erros
                    relatorio_path = os.path.join(st.session_state.temp_dir, 'relatorio_erros.json')
                    gerar_relatorio_erros(st.session_state.erros, relatorio_path)
                    
                    status_text.text("Processamento concluído!")
                    st.markdown("<div class='success-message'>Processamento concluído com sucesso! Verifique as abas 'Resultados' e 'Relatório de Erros'.</div>", unsafe_allow_html=True)
            else:
                st.markdown("<div class='error-message'>Nenhum arquivo foi processado com sucesso.</div>", unsafe_allow_html=True)

with tab2:
    st.markdown("<h2 class='sub-header'>Resultados da Extração</h2>", unsafe_allow_html=True)
    
    if st.session_state.resultados:
        # Criar DataFrame para visualização
        dados = []
        for resultado in st.session_state.resultados:
            if resultado['sucesso']:
                dados.append({
                    'Arquivo': resultado['arquivo'],
                    'Distribuidora': resultado['distribuidora'],
                    'Instalação': resultado['instalacao'],
                    'Referência': resultado['referencia'],
                    'Encargo': f"R$ {resultado['encargo']:.2f}".replace('.', ','),
                    'Reativa': f"R$ {resultado['reativa']:.2f}".replace('.', ','),
                    'Desc. Encargo': f"R$ {resultado['desc_encargo']:.2f}".replace('.', ','),
                    'SAP1': f"R$ {resultado['sap1']:.2f}".replace('.', ','),
                    'Demanda': f"R$ {resultado['demanda']:.2f}".replace('.', ','),
                    'Desc. Demanda': f"R$ {resultado['desc_demanda']:.2f}".replace('.', ','),
                    'SAP2': f"R$ {resultado['sap2']:.2f}".replace('.', ','),
                    'Contribuição Pública': f"R$ {resultado['contribuicao_publica']:.2f}".replace('.', ','),
                    'Outros': f"R$ {resultado['outros']:.2f}".replace('.', ','),
                    'Total': f"R$ {resultado['total']:.2f}".replace('.', ',')
                })
        
        if dados:
            df = pd.DataFrame(dados)
            st.dataframe(df, use_container_width=True)
            
            # Opção para download do Excel
            if st.session_state.processamento_concluido and os.path.exists(st.session_state.excel_path):
                with open(st.session_state.excel_path, "rb") as file:
                    excel_data = file.read()
                
                st.download_button(
                    label="Baixar Arquivo Excel",
                    data=excel_data,
                    file_name=os.path.basename(st.session_state.excel_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.markdown("<div class='warning-box'>Nenhum resultado disponível para exibição.</div>", unsafe_allow_html=True)
    else:
        st.markdown("<div class='info-box'>Nenhum resultado disponível. Faça o upload e processamento das faturas primeiro.</div>", unsafe_allow_html=True)

with tab3:
    st.markdown("<h2 class='sub-header'>Relatório de Erros</h2>", unsafe_allow_html=True)
    
    if st.session_state.erros:
        # Arquivos que falharam na extração
        if st.session_state.erros['arquivos_falha']:
            st.markdown("<h3>Arquivos que falharam na extração:</h3>", unsafe_allow_html=True)
            for erro in st.session_state.erros['arquivos_falha']:
                st.markdown(f"- **{erro['arquivo']}**: {erro['erro']}")
        
        # Tabelas não reconhecidas
        if st.session_state.erros['tabelas_nao_reconhecidas']:
            st.markdown("<h3>Tabelas não reconhecidas:</h3>", unsafe_allow_html=True)
            for erro in st.session_state.erros['tabelas_nao_reconhecidas']:
                st.markdown(f"- **{erro['arquivo']}**: Distribuidora {erro['distribuidora']}")
        
        # Itens não mapeados corretamente
        if st.session_state.erros['itens_nao_mapeados']:
            st.markdown("<h3>Itens não mapeados corretamente:</h3>", unsafe_allow_html=True)
            for erro in st.session_state.erros['itens_nao_mapeados']:
                st.markdown(f"- **{erro['arquivo']}**: Item '{erro['item']}' (classificado como 'Outros')")
        
        # Opção para download do relatório de erros
        if st.session_state.processamento_concluido:
            relatorio_path = os.path.join(st.session_state.temp_dir, 'relatorio_erros.json')
            if os.path.exists(relatorio_path):
                with open(relatorio_path, "rb") as file:
                    relatorio_data = file.read()
                
                st.download_button(
                    label="Baixar Relatório de Erros (JSON)",
                    data=relatorio_data,
                    file_name="relatorio_erros.json",
                    mime="application/json"
                )
    else:
        st.markdown("<div class='info-box'>Nenhum erro reportado. Faça o upload e processamento das faturas primeiro.</div>", unsafe_allow_html=True)

# Rodapé
st.markdown("---")
st.markdown(f"<p style='text-align: center; color: gray;'>Leitor de Faturas de Energia Elétrica • {datetime.now().strftime('%Y')}</p>", unsafe_allow_html=True)
