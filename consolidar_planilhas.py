"""
Consolidador de Planilhas Excel - Streamlit
Para rodar no VSCode
"""

import streamlit as st
import pandas as pd
import io
from collections import Counter

# Configuração da página
st.set_page_config(
    page_title="Consolidador de Planilhas Excel",
    page_icon="📊",
    layout="wide"
)

# CSS customizado
st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stButton>button {
        width: 100%;
        background-color: #4F46E5;
        color: white;
        border-radius: 8px;
        padding: 0.5rem 1rem;
        font-weight: 600;
    }
    .stButton>button:hover {
        background-color: #4338CA;
    }
    .success-box {
        padding: 1rem;
        background-color: #D1FAE5;
        border-left: 4px solid #10B981;
        border-radius: 4px;
        margin: 1rem 0;
    }
    .info-box {
        padding: 1rem;
        background-color: #DBEAFE;
        border-left: 4px solid #3B82F6;
        border-radius: 4px;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

# Título
st.title("📊 Consolidador de Planilhas Excel")
st.markdown("**Unifique múltiplas planilhas em um único arquivo de forma simples e rápida**")
st.markdown("---")

# Inicializar session state
if 'estrutura' not in st.session_state:
    st.session_state.estrutura = None
if 'arquivos_dados' not in st.session_state:
    st.session_state.arquivos_dados = {}
if 'modo' not in st.session_state:
    st.session_state.modo = None
if 'abas_selecionadas' not in st.session_state:
    st.session_state.abas_selecionadas = []
if 'multiselect_key' not in st.session_state:
    st.session_state.multiselect_key = 0

# Etapa 1: Upload de arquivos
st.header("1️⃣ Upload dos Arquivos")
uploaded_files = st.file_uploader(
    "Selecione os arquivos Excel: (.xlsx, .xls)",
    type=['xlsx', 'xls'],
    accept_multiple_files=True,
    help="Você pode selecionar múltiplos arquivos de uma vez"
)

if uploaded_files:
    st.markdown(f"<div class='success-box'>✅ <b>{len(uploaded_files)} arquivo(s) carregado(s)</b></div>", unsafe_allow_html=True)
    
    with st.expander("📁 Ver arquivos carregados"):
        for file in uploaded_files:
            st.write(f"• {file.name}")
    
    # Analisar estrutura dos arquivos
    if st.button("🔍 Analisar Estrutura das Planilhas"):
        with st.spinner("Analisando arquivos..."):
            estrutura_completa = []
            arquivos_processados = {}
            todas_abas_nomes = []
            
            for uploaded_file in uploaded_files:
                try:
                    # Ler arquivo Excel
                    excel_file = pd.ExcelFile(uploaded_file)
                    num_abas = len(excel_file.sheet_names)
                    nomes_abas = excel_file.sheet_names
                    
                    # Armazenar informações
                    arquivos_processados[uploaded_file.name] = {
                        'num_abas': num_abas,
                        'nomes': nomes_abas,
                        'arquivo': uploaded_file
                    }
                    
                    estrutura_completa.append({
                        'arquivo': uploaded_file.name,
                        'num_abas': num_abas,
                        'nomes': nomes_abas
                    })
                    
                    todas_abas_nomes.extend(nomes_abas)
                    
                except Exception as e:
                    st.error(f"❌ Erro ao processar {uploaded_file.name}: {e}")
            
            # Salvar no session state
            st.session_state.estrutura = {
                'completa': estrutura_completa,
                'total_arquivos': len(estrutura_completa),
                'todas_abas_nomes': todas_abas_nomes,
                'contador_nomes': Counter(todas_abas_nomes)
            }
            st.session_state.arquivos_dados = arquivos_processados
            
            st.success("✅ Análise concluída!")

# Etapa 2: Mostrar estrutura e escolher modo
if st.session_state.estrutura:
    st.markdown("---")
    st.header("2️⃣ Estrutura Detectada")
    
    estrutura = st.session_state.estrutura
    
    # Resumo
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("📄 Total de Arquivos", estrutura['total_arquivos'])
    with col2:
        nums_abas = [info['num_abas'] for info in estrutura['completa']]
        media_abas = sum(nums_abas)/len(nums_abas) if len(nums_abas) > 0 else 0
        st.metric("📊 Abas por Arquivo (média)", f"{media_abas:.1f}")
    with col3:
        st.metric("📋 Total de Abas", len(estrutura['todas_abas_nomes']))
    
    # Detalhes de cada arquivo
    with st.expander("🔎 Ver detalhes de cada arquivo."):
        for info in estrutura['completa']:
            st.write(f"**{info['arquivo']}** - {info['num_abas']} aba(s)")
            for idx, nome in enumerate(info['nomes'], 1):
                st.write(f"  {idx}. {nome}")
            st.write("")
    
    # Escolher modo
    st.markdown("---")
    st.header("3️⃣ Modo de Consolidação")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📍 Consolidar por POSIÇÃO", use_container_width=True, key="btn_modo_posicao"):
            st.session_state.modo = 'posicao'
            st.session_state.abas_selecionadas = []
            st.session_state.multiselect_key = 0
        st.caption("Junta todas as 1ª abas, todas as 2ª abas, etc.")
    
    with col2:
        if st.button("🏷️ Consolidar por NOME", use_container_width=True, key="btn_modo_nome"):
            st.session_state.modo = 'nome'
            st.session_state.abas_selecionadas = []
            st.session_state.multiselect_key = 0
        st.caption("Junta todas as abas 'Vendas', todas 'Estoque', etc.")

# Etapa 3: Selecionar abas
if st.session_state.modo:
    st.markdown("---")
    st.header("4️⃣ Selecione as Abas para Consolidar")
    
    st.markdown(f"<div class='info-box'>📌 <b>Modo selecionado:</b> {st.session_state.modo.upper()}</div>", unsafe_allow_html=True)
    
    if st.session_state.modo == 'posicao':
        # Consolidar por posição
        max_abas = max([info['num_abas'] for info in st.session_state.estrutura['completa']])
        
        opcoes_posicao = {}
        for pos in range(1, max_abas + 1):
            # Pegar nomes das abas nesta posição
            nomes_nesta_posicao = []
            for info in st.session_state.estrutura['completa']:
                if pos <= len(info['nomes']):
                    nomes_nesta_posicao.append(info['nomes'][pos-1])
            
            nomes_unicos = list(set(nomes_nesta_posicao))
            opcoes_posicao[pos] = {
                'label': f"Posição {pos} ({', '.join(nomes_unicos[:3])}{'...' if len(nomes_unicos) > 3 else ''})",
                'count': len(nomes_nesta_posicao)
            }
        
        st.write("**Selecione as posições:**")
        
        # Armazenar todas as opções
        todas_posicoes = list(opcoes_posicao.keys())
        
        col_btn1, col_btn2 = st.columns([1, 3])
        with col_btn1:
            if st.button("✅ Selecionar Todas", key="select_all_pos"):
                st.session_state.abas_selecionadas = todas_posicoes
                st.session_state.multiselect_key += 1
        with col_btn2:
            if st.button("❌ Limpar Seleção", key="clear_all_pos"):
                st.session_state.abas_selecionadas = []
                st.session_state.multiselect_key += 1
        
        # Criar lista de opções formatadas
        opcoes_formatadas = {pos: info['label'] for pos, info in opcoes_posicao.items()}
        
        # Usar multiselect com key dinâmica para forçar atualização
        selecionadas = st.multiselect(
            "Escolha as posições para consolidar:",
            options=todas_posicoes,
            default=st.session_state.abas_selecionadas,
            format_func=lambda x: opcoes_formatadas[x],
            key=f"multiselect_pos_{st.session_state.multiselect_key}"
        )
        
        # Atualizar apenas se houve mudança manual
        if selecionadas != st.session_state.abas_selecionadas:
            st.session_state.abas_selecionadas = selecionadas
        
    else:  # modo == 'nome'
        # Consolidar por nome
        nomes_unicos = sorted(set(st.session_state.estrutura['todas_abas_nomes']))
        contador = st.session_state.estrutura['contador_nomes']
        
        st.write("**Selecione os nomes das abas:**")
        
        col_btn1, col_btn2 = st.columns([1, 3])
        with col_btn1:
            if st.button("✅ Selecionar Todas", key="select_all_name"):
                st.session_state.abas_selecionadas = nomes_unicos.copy()
                st.session_state.multiselect_key += 1
        with col_btn2:
            if st.button("❌ Limpar Seleção", key="clear_all_name"):
                st.session_state.abas_selecionadas = []
                st.session_state.multiselect_key += 1
        
        # Criar função de formatação para mostrar contagem
        def formatar_nome(nome):
            return f"{nome} ({contador[nome]}x)"
        
        # Usar multiselect com key dinâmica para forçar atualização
        selecionadas = st.multiselect(
            "Escolha os nomes das abas para consolidar:",
            options=nomes_unicos,
            default=st.session_state.abas_selecionadas,
            format_func=formatar_nome,
            key=f"multiselect_name_{st.session_state.multiselect_key}"
        )
        
        # Atualizar apenas se houve mudança manual
        if selecionadas != st.session_state.abas_selecionadas:
            st.session_state.abas_selecionadas = selecionadas
    
    # Botão de consolidar
    if st.session_state.abas_selecionadas:
        st.markdown(f"<div class='success-box'>✅ <b>{len(st.session_state.abas_selecionadas)} aba(s) selecionada(s)</b></div>", unsafe_allow_html=True)
        
        st.markdown("---")
        st.header("5️⃣ Consolidar e Baixar")
        
        if st.button("🚀 CONSOLIDAR E BAIXAR PLANILHA", use_container_width=True):
            with st.spinner("🔄 Processando... Isso pode levar alguns instantes..."):
                try:
                    # Preparar dados para consolidação
                    dados_consolidados = {aba: [] for aba in st.session_state.abas_selecionadas}
                    
                    # Processar cada arquivo
                    progress_bar = st.progress(0)
                    total_arquivos = len(st.session_state.arquivos_dados)
                    
                    for idx, (nome_arquivo, info) in enumerate(st.session_state.arquivos_dados.items()):
                        info['arquivo'].seek(0)
                        excel_file = pd.ExcelFile(info['arquivo'])
                        
                        if st.session_state.modo == 'posicao':
                            for pos in st.session_state.abas_selecionadas:
                                if pos <= len(info['nomes']):
                                    nome_aba = info['nomes'][pos - 1]
                                    df = pd.read_excel(excel_file, sheet_name=nome_aba)
                                    dados_consolidados[pos].append(df)
                        else:
                            for nome_aba in st.session_state.abas_selecionadas:
                                if nome_aba in info['nomes']:
                                    df = pd.read_excel(excel_file, sheet_name=nome_aba)
                                    dados_consolidados[nome_aba].append(df)
                        
                        progress_bar.progress((idx + 1) / total_arquivos)
                    
                    # Criar arquivo consolidado
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        for aba_id, lista_dfs in dados_consolidados.items():
                            if lista_dfs:
                                df_consolidado = pd.concat(lista_dfs, ignore_index=True)
                                
                                if st.session_state.modo == 'posicao':
                                    # Obter o nome original da primeira aba nesta posição
                                    nome_aba_destino = None
                                    for info in st.session_state.estrutura['completa']:
                                        if aba_id <= len(info['nomes']):
                                            nome_aba_destino = info['nomes'][aba_id - 1]
                                            break
                                    if nome_aba_destino is None:
                                        nome_aba_destino = f'Aba{aba_id}'
                                else:
                                    # Usar o nome original da aba
                                    nome_aba_destino = str(aba_id)
                                
                                # Limitar a 31 caracteres (limite do Excel)
                                nome_aba_destino = nome_aba_destino[:31]
                                
                                df_consolidado.to_excel(writer, sheet_name=nome_aba_destino, index=False)
                    
                    output.seek(0)
                    
                    # Botão de download
                    st.success("✅ Consolidação concluída com sucesso!")
                    
                    st.download_button(
                        label="📥 BAIXAR PLANILHA CONSOLIDADA",
                        data=output,
                        file_name="planilha_consolidada.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                    # Estatísticas
                    st.markdown("### 📊 Resumo da Consolidação")
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Arquivos processados", total_arquivos)
                    with col2:
                        st.metric("Abas consolidadas", len([d for d in dados_consolidados.values() if d]))
                    
                except Exception as e:
                    st.error(f"❌ Erro durante a consolidação: {e}")
    else:
        st.info("👆 Selecione pelo menos uma aba para continuar.")

# Botão de reset
if st.session_state.estrutura or st.session_state.modo:
    st.markdown("---")
    if st.button("🔄 Recomeçar com Novos Arquivos"):
        st.session_state.estrutura = None
        st.session_state.arquivos_dados = {}
        st.session_state.modo = None
        st.session_state.abas_selecionadas = []
        st.session_state.multiselect_key = 0

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #6B7280; padding: 1rem;'>
    <small>💡 <b>Dica:</b> Certifique-se de que suas planilhas tenham estruturas consistentes para melhores resultados</small>
</div>
""", unsafe_allow_html=True)