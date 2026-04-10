"""
Automação SAP - EDOC_COCKPIT
Interface Streamlit para automação da transação EDOC_COCKPIT do SAP
"""

import streamlit as st
import win32com.client as win32
import pythoncom
import time
import pandas as pd
import os
from datetime import datetime, timedelta


class SAPConnection:
    """Gerencia conexão e interação com SAP GUI"""
    
    def __init__(self):
        self.session = None
        self.connection = None
        self.application = None
        
    def connect_and_create_session(self):
        """Conecta ao SAP GUI e cria uma nova sessão"""
        try:
            pythoncom.CoInitialize()
            
            # Usar GetObject (mais estável que Dispatch)
            sap_gui_auto = win32.GetObject("SAPGUI")
            application = sap_gui_auto.GetScriptingEngine
            
            # Verificar se tem conexão
            if application.Children.Count == 0:
                return False, "❌ SAP não tem conexões ativas. Faça login no SAP primeiro."
            
            connection = application.Children(0)
            
            # Verificar se tem sessão
            if connection.Children.Count == 0:
                return False, "❌ SAP não tem sessões ativas. Faça login no SAP primeiro."
            
            # Criar nova sessão
            session = connection.Children(0).createSession()
            
            # Salva nas propriedades da classe
            self.application = application
            self.connection = connection
            self.session = session
            
            return True, "✅ Nova sessão criada com sucesso!"
            
        except Exception as e:
            return False, f"❌ Erro ao conectar ao SAP: {str(e)}"
    
    def execute_edoc_cockpit_automation(self, df_base):
        """
        Executa a automação completa do EDOC_COCKPIT com filtros
        
        Args:
            df_base: DataFrame pandas com os dados para filtrar
        
        Returns:
            tuple: (success, message)
        """
        try:
            if not self.session:
                return False, "❌ Não há sessão SAP ativa"
            
            # Maximizar janela
            self.session.FindById("wnd[0]").Maximize()
            time.sleep(1)
            
            # Abrir transação EDOC_COCKPIT
            self.session.FindById("wnd[0]/tbar[0]/okcd").Text = "EDOC_COCKPIT"
            self.session.FindById("wnd[0]").SendVKey(0)
            time.sleep(2)
            
            # Pressionar botão de mudança de seleção
            self.session.FindById("wnd[0]/usr/cntlC_CONTAINER/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont[1]/shell/shellcont[1]/shell").PressButton("CHANGE_SELECTION")
            time.sleep(1)
            
            # Pressionar botão na toolbar
            self.session.FindById("wnd[1]/tbar[0]/btn[25]").Press()
            time.sleep(1)
            
            # Navegar pela árvore de seleção
            self.session.FindById("wnd[2]/shellcont/shell").CollapseNode("          1")
            time.sleep(0.5)
            
            self.session.FindById("wnd[2]/shellcont/shell").ExpandNode("         84")
            time.sleep(0.5)
            
            self.session.FindById("wnd[2]/shellcont/shell").SelectNode("         92")
            self.session.FindById("wnd[2]/shellcont/shell").TopNode = "         84"
            time.sleep(0.5)
            
            self.session.FindById("wnd[2]/shellcont/shell").DoubleClickNode("         92")
            time.sleep(1)
            
            self.session.FindById("wnd[2]/shellcont/shell").TopNode = "          1"
            time.sleep(0.5)
            
            # Pressionar botão de valores
            self.session.FindById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN002_%_APP_%-VALU_PUSH").Press()
            time.sleep(1)
            
            # Colar dados do DataFrame
            self._paste_dataframe_to_sap(df_base)
            
            # Confirmar dados colados
            self.session.FindById("wnd[3]/tbar[0]/btn[8]").Press()
            time.sleep(1)
            
            self.session.FindById("wnd[2]/tbar[0]/btn[8]").Press()
            time.sleep(1)
            
            # Calcular datas (D-7 e D-1)
            data_inicial = (datetime.now() - timedelta(days=7)).strftime("%d%m%Y")
            data_final = (datetime.now() - timedelta(days=1)).strftime("%d%m%Y")
            
            # Preencher campo de data inicial (D-7)
            self.session.FindById("wnd[1]/usr/subSUB_EDOCUMENT:SAPLEDOC_COCKPIT:0300/ctxtS00CREA-LOW").Text = data_inicial
            time.sleep(0.5)
            
            # Preencher campo de data final (D-1)
            self.session.FindById("wnd[1]/usr/subSUB_EDOCUMENT:SAPLEDOC_COCKPIT:0300/ctxtS00CREA-HIGH").Text = data_final
            self.session.FindById("wnd[1]/usr/subSUB_EDOCUMENT:SAPLEDOC_COCKPIT:0300/ctxtS00CREA-HIGH").SetFocus()
            self.session.FindById("wnd[1]/usr/subSUB_EDOCUMENT:SAPLEDOC_COCKPIT:0300/ctxtS00CREA-HIGH").CaretPosition = 8
            time.sleep(0.5)
            
            # Executar busca
            self.session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
            time.sleep(2)
            
            return True, f"✅ Automação executada! Período: {data_inicial} a {data_final}"
            
        except Exception as e:
            return False, f"❌ Erro na automação: {str(e)}"
    
    def _paste_dataframe_to_sap(self, df):
        """
        Cola dados do DataFrame no SAP usando clipboard
        Usa apenas a coluna 'CNPJ' do DataFrame
        
        Args:
            df: DataFrame pandas com coluna 'CNPJ'
        """
        try:
            # Extrair apenas a coluna CNPJ
            if 'CNPJ' not in df.columns:
                raise Exception("Coluna 'CNPJ' não encontrada no arquivo base.csv")
            
            cnpj_data = df[['CNPJ']].copy()
            
            # Converter DataFrame para formato de clipboard (tab-separated)
            cnpj_data.to_clipboard(sep='\t', index=False, header=False)
            
            # Aguardar um momento para o clipboard ser atualizado
            time.sleep(0.5)
            
            # Simular Ctrl+V no SAP (o SAP deve estar focado no campo correto)
            self.session.FindById("wnd[3]").SendVKey(24)  # Ctrl+Shift ou método de paste
            
        except Exception as e:
            raise Exception(f"Erro ao colar dados: {str(e)}")
    
    def close_session(self):
        """Fecha a sessão SAP criada"""
        try:
            if self.session:
                self.session.findById("wnd[0]").close()
                self.session = None
        except:
            pass


def main():
    """Interface principal do Streamlit"""
    
    # Configuração da página
    st.set_page_config(
        page_title="SAP EDOC_COCKPIT Automation",
        page_icon="🔧",
        layout="wide"
    )
    
    # Título
    st.title("🔧 Automação SAP - EDOC_COCKPIT")
    st.markdown("---")
    
    # Inicializa o estado da sessão
    if 'df_base' not in st.session_state:
        st.session_state.df_base = None
    
    # Sidebar com informações
    with st.sidebar:
        st.header("ℹ️ Informações")
        
        st.markdown("""
        ### Pré-requisitos:
        1. ✅ SAP Logon instalado
        2. ✅ SAP GUI aberto e logado
        3. ✅ Scripting habilitado
        
        ### Transação:
        - **EDOC_COCKPIT**: Cockpit de Documentos Eletrônicos
        
        ### Automação:
        - Cria nova sessão SAP automaticamente
        - Período de busca: D-7 até D-1
        - Filtros baseados no arquivo base.csv
        """)
    
    # Área principal
    st.subheader("⏰ Status")
    st.info(f"**Data/Hora:** {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    
    st.markdown("---")
    
    # Seção de automação
    st.subheader("🚀 Automação EDOC_COCKPIT")
    
    # Carregar base.csv automaticamente
    st.markdown("### 📂 Base de Dados")
    
    base_csv_path = os.path.join(os.path.dirname(__file__), "base.csv")
    
    col_base1, col_base2, col_base3 = st.columns([2, 1, 1])
    
    with col_base1:
        if os.path.exists(base_csv_path):
            if st.session_state.df_base is None:
                try:
                    st.session_state.df_base = pd.read_csv(base_csv_path)
                    st.success(f"✅ Arquivo base.csv carregado")
                except Exception as e:
                    st.error(f"❌ Erro ao ler base.csv: {str(e)}")
            else:
                st.info(f"📄 base.csv carregado")
        else:
            st.error(f"❌ Arquivo não encontrado: base.csv")
    
    with col_base2:
        if st.session_state.df_base is not None:
            st.metric("Registros", len(st.session_state.df_base))
    
    with col_base3:
        if st.session_state.df_base is not None:
            if 'CNPJ' in st.session_state.df_base.columns:
                cnpj_count = st.session_state.df_base['CNPJ'].notna().sum()
                st.metric("CNPJs", cnpj_count)
            else:
                st.warning("⚠️ Coluna CNPJ não encontrada")
    
    # Visualizar dados carregados
    if st.session_state.df_base is not None:
        with st.expander("👁️ Visualizar CNPJs da Base"):
            if 'CNPJ' in st.session_state.df_base.columns:
                st.dataframe(st.session_state.df_base[['CNPJ']], width="stretch")
            else:
                st.dataframe(st.session_state.df_base, width="stretch")
    
    st.markdown("---")
    
    # Configuração de datas
    st.markdown("### 📅 Período de Busca")
    
    col_date1, col_date2, col_date3 = st.columns(3)
    
    with col_date1:
        data_inicial = datetime.now() - timedelta(days=7)
        st.info(f"**Data Inicial (D-7):**\n\n{data_inicial.strftime('%d/%m/%Y')}")
    
    with col_date2:
        data_final = datetime.now() - timedelta(days=1)
        st.info(f"**Data Final (D-1):**\n\n{data_final.strftime('%d/%m/%Y')}")
    
    with col_date3:
        dias_periodo = 6
        st.info(f"**Período:**\n\n{dias_periodo} dias")
    
    st.markdown("---")
    
    # Botão de execução
    can_execute = (
        st.session_state.df_base is not None and 
        'CNPJ' in st.session_state.df_base.columns
    )
    
    if can_execute:
        if st.button("🤖 Executar Automação EDOC_COCKPIT", type="primary", width="stretch"):
            with st.spinner("Executando automação... Não interaja com o SAP!"):
                progress_bar = st.progress(0, text="Iniciando...")
                
                # Criar nova conexão e sessão
                sap = SAPConnection()
                progress_bar.progress(10, text="Conectando ao SAP...")
                success, message = sap.connect_and_create_session()
                
                if not success:
                    st.error(message)
                else:
                    progress_bar.progress(20, text="Sessão criada com sucesso...")
                    time.sleep(0.5)
                    
                    progress_bar.progress(40, text="Abrindo transação...")
                    time.sleep(0.5)
                    
                    progress_bar.progress(60, text="Aplicando filtros...")
                    
                    success, message = sap.execute_edoc_cockpit_automation(
                        st.session_state.df_base
                    )
                    
                    progress_bar.progress(100, text="Concluído!")
                    
                    if success:
                        st.success(message)
                        st.balloons()
                    else:
                        st.error(message)
                    
                    # Fechar sessão
                    sap.close_session()
    else:
        st.button(
            "🤖 Executar Automação EDOC_COCKPIT",
            type="primary",
            width="stretch",
            disabled=True,
            help="Arquivo base.csv não encontrado ou não possui coluna CNPJ"
        )
    
    # Área de informações
    with st.expander("ℹ️ Sobre a Automação"):
        st.markdown("""
        ### 🔄 Fluxo da Automação:
        
        1. **Cria nova sessão** no SAP automaticamente
        2. **Maximiza** a janela do SAP
        3. **Abre** a transação EDOC_COCKPIT
        4. **Navega** pelos menus de seleção
        5. **Aplica filtros** com CNPJs da base.csv
        6. **Define período** de busca (D-7 até D-1)
        7. **Executa** a consulta
        8. **Fecha** a sessão criada
        
        ### 📋 Requisitos:
        - ✅ Arquivo base.csv com coluna 'CNPJ'
        - ✅ SAP GUI aberto e logado
        - ✅ Não interagir com o SAP durante execução
        
        ### ⏱️ Tempo estimado:
        - 15-30 segundos dependendo da conexão
        """)
    
    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: gray;'>"
        "🤖 Desenvolvido com Streamlit + SAP GUI Scripting"
        "</div>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
