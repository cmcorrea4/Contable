import streamlit as st
import pandas as pd
import os
from langchain_experimental.agents.agent_toolkits import create_pandas_dataframe_agent
from langchain_openai import ChatOpenAI
import warnings
warnings.filterwarnings('ignore')

def main():
    st.set_page_config(
        page_title="Agente de Datos",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # ── Estilos ──────────────────────────────────────────────────────────────────
    st.markdown("""
    <style>
        .stApp { background-color: #F5F7FA; }
        .header-block {
            background: linear-gradient(135deg, #1E3A5F 0%, #2E6DA4 100%);
            border-radius: 12px;
            padding: 28px 36px;
            margin-bottom: 28px;
            color: white;
        }
        .header-block h1 { font-size: 1.9rem; margin: 0; font-weight: 700; color: white !important; }
        .header-block p  { margin: 6px 0 0; opacity: .85; font-size: .95rem; }
        .metric-card {
            background: white;
            border-radius: 10px;
            padding: 20px 24px;
            box-shadow: 0 2px 8px rgba(0,0,0,.08);
            text-align: center;
            height: 100%;
        }
        .metric-card .number {
            font-size: 2.2rem;
            font-weight: 700;
            line-height: 1.1;
            color: #2E6DA4;
        }
        .metric-card .label {
            font-size: .82rem;
            color: #666;
            margin-top: 4px;
            text-transform: uppercase;
            letter-spacing: .05em;
        }
        .upload-section {
            background: white;
            border-radius: 10px;
            padding: 24px;
            box-shadow: 0 2px 8px rgba(0,0,0,.08);
            margin-bottom: 20px;
        }
        .section-title {
            font-size: 1rem;
            font-weight: 600;
            color: #1E3A5F;
            margin-bottom: 12px;
            padding-bottom: 8px;
            border-bottom: 2px solid #E8EDF3;
        }
        .result-block {
            background: white;
            border-radius: 10px;
            padding: 24px;
            box-shadow: 0 2px 8px rgba(0,0,0,.08);
            margin-top: 20px;
        }
        #MainMenu, footer, header { visibility: hidden; }
        .block-container { padding-top: 1.5rem; }
    </style>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div class="header-block">
        <h1>📊 Agente de Datos</h1>
        <p>Carga tu archivo CSV o Excel y realiza preguntas sobre tus datos usando Inteligencia Artificial.</p>
    </div>
    """, unsafe_allow_html=True)
    
    # ── Configuración del modelo ─────────────────────────────────────────────────
    openai_api_key = st.secrets["key"]
    model_name = "gpt-4.1-2025-04-14"
    temperature = 0.1
    
    # Carga de archivo
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    st.markdown('<p class="section-title">📄 Archivo de Datos (CSV o Excel)</p>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Sube tu archivo para explorar:",
        type=['csv', 'xlsx', 'xls'],
        label_visibility="collapsed"
    )
    st.markdown('</div>', unsafe_allow_html=True)
    
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.success(f"✅ Archivo cargado exitosamente: {uploaded_file.name}")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="number">{df.shape[0]}</div>
                    <div class="label">Filas</div>
                </div>""", unsafe_allow_html=True)
            with col2:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="number">{df.shape[1]}</div>
                    <div class="label">Columnas</div>
                </div>""", unsafe_allow_html=True)
            with col3:
                tamano_kb = df.memory_usage(deep=True).sum() / 1024
                st.markdown(f"""
                <div class="metric-card">
                    <div class="number">{tamano_kb:.1f}</div>
                    <div class="label">Tamaño (KB)</div>
                </div>""", unsafe_allow_html=True)
            
            st.markdown('<div class="result-block">', unsafe_allow_html=True)
            st.markdown('<p class="section-title">👀 Análisis Exploratorio</p>', unsafe_allow_html=True)
            
            tab1, tab2, tab3 = st.tabs(["📋 Datos", "📈 Información", "🔍 Estadísticas"])
            
            with tab1:
                st.dataframe(df.head(100), use_container_width=True)
                
            with tab2:
                st.write("**Estructura del Dataset:**")
                info_df = pd.DataFrame({
                    'Columna': df.columns,
                    'Tipo': df.dtypes.astype(str),
                    'No Nulos': df.count(),
                    'Nulos': df.isnull().sum(),
                    '% Nulos': (df.isnull().sum() / len(df) * 100).round(2)
                })
                st.dataframe(info_df, use_container_width=True)
                
            with tab3:
                st.write("**Estadísticas Descriptivas:**")
                numeric_df = df.select_dtypes(include=['number'])
                if not numeric_df.empty:
                    st.dataframe(numeric_df.describe(), use_container_width=True)
                else:
                    st.info("No hay columnas numéricas para mostrar estadísticas.")
            
            st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown('<div class="result-block">', unsafe_allow_html=True)
            st.markdown('<p class="section-title">🤖 Asistente de IA</p>', unsafe_allow_html=True)
            
            try:
                llm = ChatOpenAI(
                    model=model_name,
                    temperature=temperature,
                    openai_api_key=openai_api_key
                )
                
                agent = create_pandas_dataframe_agent(
                    llm,
                    df,
                    verbose=True,
                    agent_type="openai-tools",
                    allow_dangerous_code=True
                )
                
                if 'chat_history' not in st.session_state:
                    st.session_state.chat_history = []
                
                user_question = st.text_input(
                    "Ingresa tu consulta sobre los datos:",
                    placeholder="Ej: ¿Hay valores nulos en el dataset o cuál es su estructura general?",
                    key="user_input"
                )
                
                col_btn1, col_btn2 = st.columns([1, 4])
                with col_btn1:
                    ask_button = st.button("🚀 Preguntar", type="primary")
                with col_btn2:
                    clear_button = st.button("🗑️ Limpiar historial")
                
                if clear_button:
                    st.session_state.chat_history = []
                    st.rerun()
                
                if ask_button and user_question:
                    with st.spinner("🔄 Procesando consulta mediante IA..."):
                        try:
                            response = agent.invoke({"input": user_question})
                            st.session_state.chat_history.append({
                                "question": user_question,
                                "answer": response["output"]
                            })
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Error al procesar la pregunta: {str(e)}")
                            st.info("💡 Intenta reformular tu pregunta.")
                
                if st.session_state.chat_history:
                    st.markdown("<br><b>💬 Conversación reciente:</b>", unsafe_allow_html=True)
                    for i, chat in enumerate(reversed(st.session_state.chat_history)):
                        with st.expander(f"❓ {chat['question'][:80]}..." if len(chat['question']) > 80 else f"❓ {chat['question']}", expanded=(i==0)):
                            st.write("**Respuesta:**")
                            st.write(chat['answer'])
                
            except Exception as e:
                st.error(f"❌ Error al inicializar el agente: {str(e)}")
            st.markdown('</div>', unsafe_allow_html=True)
                
        except Exception as e:
            st.error(f"❌ Error al cargar el archivo. ¿Verificaste el formato? ({str(e)})")
    
    else:
        st.markdown("""
        <div style="text-align:center; padding: 60px 20px; color: #999;">
            <div style="font-size: 3rem; margin-bottom: 16px;">📂</div>
            <p style="font-size: 1.1rem; font-weight: 600;">Carga un archivo de datos para explorar mediante IA</p>
            <p style="font-size: .9rem;">Formatos soportados: CSV, Excel (.xlsx, .xls)</p>
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
