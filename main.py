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
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path
import zipfile

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
            
            # Pegar o número de sessões antes de criar uma nova
            sessions_antes = connection.Children.Count
            
            # Criar nova sessão
            connection.Children(0).createSession()
            
            # Aguardar a criação da sessão
            time.sleep(2)
            
            # Pegar a última sessão criada (a nova)
            new_session = connection.Children(connection.Children.Count - 1)
            
            # Salva nas propriedades da classe
            self.application = application
            self.connection = connection
            self.session = new_session
            
            return True, "✅ Nova sessão criada com sucesso!"
            
        except Exception as e:
            return False, f"❌ Erro ao conectar ao SAP: {str(e)}"
    
    def execute_edoc_cockpit_automation(self, df_base, data_inicial, data_final):
        """
        Executa a automação completa do EDOC_COCKPIT com filtros
        
        Args:
            df_base: DataFrame pandas com os dados para filtrar
            data_inicial: Data inicial no formato datetime.date
            data_final: Data final no formato datetime.date
        
        Returns:
            tuple: (success, message)
        """
        try:
            if not self.session:
                return False, "❌ Não há sessão SAP ativa"
            
            # Limpar pastas de relatórios antes de iniciar
            print("[INFO] 🧹 Limpando pastas de relatórios...")
            clean_success, clean_message = self.limpar_pastas_relatorios()
            if not clean_success:
                print(f"[WARN] {clean_message}")
            else:
                print(clean_message)
            
            # Maximizar janela
            self.session.FindById("wnd[0]").Maximize()
            
            # Abrir transação EDOC_COCKPIT
            print("[INFO] Abrindo transação EDOC_COCKPIT...")
            self.session.FindById("wnd[0]/tbar[0]/okcd").Text = "EDOC_COCKPIT"
            self.session.FindById("wnd[0]").SendVKey(0)

            self.session.findById("wnd[0]/usr/cntlC_CONTAINER/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont[1]/shell/shellcont[1]/shell").pressButton("CHANGE_SELECTION")

            # Pressionar botão na toolbar
            print("[INFO] Pressionando botão 25 da toolbar...")
            self.session.FindById("wnd[1]/tbar[0]/btn[25]").Press()
            
            # Navegar pela árvore de seleção
            self.session.FindById("wnd[2]/shellcont/shell").CollapseNode("          1")
            self.session.FindById("wnd[2]/shellcont/shell").ExpandNode("         84")
            self.session.FindById("wnd[2]/shellcont/shell").SelectNode("         92")
            self.session.FindById("wnd[2]/shellcont/shell").TopNode = "         84"
            self.session.FindById("wnd[2]/shellcont/shell").DoubleClickNode("         92")
            self.session.FindById("wnd[2]/shellcont/shell").TopNode = "          1"
            
            # Pressionar botão de valores
            self.session.FindById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN002_%_APP_%-VALU_PUSH").Press()
            
            # Colar dados do DataFrame
            self._paste_dataframe_to_sap(df_base)
            
            # Confirmar dados colados
            self.session.FindById("wnd[3]/tbar[0]/btn[8]").Press()
            self.session.FindById("wnd[2]/tbar[0]/btn[8]").Press()
            
            # Converter datas para formato SAP (DDMMYYYY)
            data_inicial_sap = data_inicial.strftime("%d%m%Y")
            data_final_sap = data_final.strftime("%d%m%Y")
            
            # Preencher campo de data inicial
            self.session.FindById("wnd[1]/usr/subSUB_EDOCUMENT:SAPLEDOC_COCKPIT:0300/ctxtS00CREA-LOW").Text = data_inicial_sap
            
            # Preencher campo de data final
            self.session.FindById("wnd[1]/usr/subSUB_EDOCUMENT:SAPLEDOC_COCKPIT:0300/ctxtS00CREA-HIGH").Text = data_final_sap
            self.session.FindById("wnd[1]/usr/subSUB_EDOCUMENT:SAPLEDOC_COCKPIT:0300/ctxtS00CREA-HIGH").SetFocus()
            self.session.FindById("wnd[1]/usr/subSUB_EDOCUMENT:SAPLEDOC_COCKPIT:0300/ctxtS00CREA-HIGH").CaretPosition = 8
            
            # Executar busca
            self.session.FindById("wnd[1]/tbar[0]/btn[8]").Press()

            project_dir = os.path.abspath(os.path.dirname(__file__))
            reports_dir = os.path.join(project_dir, "Relatórios", "EDOC")
            os.makedirs(reports_dir, exist_ok=True)
            print(f"[INFO] Caminho dos relatórios: {reports_dir}")
            
            exported_count = 0
            tree_id = "wnd[0]/usr/cntlC_CONTAINER/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell/shellcont[1]/shell"
            tree = self.session.FindById(tree_id)

            # Coluna da árvore (alguns ambientes aceitam "&Hierarchy", outros "HIERARCHY")
            COL_HIER = "&Hierarchy"

            # 1) Pega todos os NodeKeys
            all_keys = list(tree.GetAllNodeKeys())

            
            target_keys = []
            for k in all_keys:
                try:
                    txt = tree.GetNodeTextByKey(k)
                except Exception:
                    continue

                if txt and txt.strip().startswith("B"):
                    target_keys.append(k)


            for opcao_num in range(2, 11):
                try:
                    # opcao_num=2 vira índice 1 (0-based)
                    idx = opcao_num - 1

                    if idx >= len(target_keys):
                        print(f"[WARN] Não existe opcao {opcao_num} na árvore (só {len(target_keys)} itens filtrados).")
                        continue

                    node_key = target_keys[idx]

                    # --- Selecionar item (use NodeKey) ---
                    tree.SelectNode(node_key)

                    # --- Garantir visibilidade horizontal (use NodeKey e coluna real) ---
                    try:
                        tree.EnsureVisibleHorizontalItem(node_key, COL_HIER)
                    except Exception:
                        # Nem todo Tree suporta esse método; não é crítico
                        pass

                    # --- Clicar no link (use NodeKey) ---
                    try:
                        tree.ClickLink(node_key, COL_HIER)
                    except Exception:
                        # fallback: duplo clique no nó (muitas vezes é o que atualiza a grade da direita)
                        tree.DoubleClickNode(node_key)

                    # >>> daqui para baixo segue seu fluxo de exportação exatamente como já estava <<<
                    export_shell = "wnd[0]/usr/cntlC_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell"

                    self.session.FindById(export_shell).PressToolbarContextButton("&MB_EXPORT")

                    self.session.FindById(export_shell).SelectContextMenuItem("&XXL")

                    filename = f"edoc_opcao{opcao_num}.xlsx"
                    self.session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = reports_dir

                    self.session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = filename
                    self.session.FindById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = len(filename)
                    
                    self.session.FindById("wnd[1]/tbar[0]/btn[11]").Press()
                    self._close_excel_files()

                    exported_count += 1

                    
                except Exception as e:
                    # Opção não existe ou erro, continua
                    print(f"Opção {opcao_num}: {str(e)}")
                    continue
            
            # ==================== FIM DA EXPORTAÇÃO ====================
            
            # Fechar arquivos Excel que podem ter sido abertos
            if exported_count > 0:
                self._close_excel_files()
            
            if exported_count > 0:
                return True, f"✅ Automação concluída! Período: {data_inicial.strftime('%d/%m/%Y')} a {data_final.strftime('%d/%m/%Y')}\n✅ {exported_count} relatório(s) exportado(s)"
            else:
                return True, f"✅ Busca executada! Período: {data_inicial.strftime('%d/%m/%Y')} a {data_final.strftime('%d/%m/%Y')}\n⚠️ Nenhum relatório foi exportado"
            
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
    
    def _close_excel_files(self):
        """
        Fecha o programa Microsoft Excel usando taskkill do Windows
        """
        try:
            import subprocess
            
            print("[INFO] Encerrando Microsoft Excel...")
            
            # Usar taskkill para forçar o fechamento do Excel
            # /F = força o fechamento
            # /IM = nome da imagem (executável)
            result = subprocess.run(
                ['taskkill', '/F', '/IM', 'EXCEL.EXE'],
                capture_output=True,
                text=True
            )
            
            if result.returncode == 0:
                print("[INFO] ✅ Microsoft Excel fechado com sucesso")
            else:
                # Código 128 significa que o processo não foi encontrado (Excel já estava fechado)
                if result.returncode == 128:
                    print("[INFO] Excel não estava aberto")
                else:
                    print(f"[WARN] Código de retorno do taskkill: {result.returncode}")
                    
        except Exception as e:
            print(f"[WARN] Erro ao tentar fechar Excel: {str(e)}")
    
    def limpar_pastas_relatorios(self):
        """
        Limpa todas as pastas de relatórios antes de iniciar nova rodada
        Remove arquivos das pastas: EDOC, ZBRMMT416 e VTIN
        """
        try:
            import shutil
            
            project_dir = os.path.abspath(os.path.dirname(__file__))
            
            # Lista de pastas a serem limpas
            pastas_para_limpar = [
                os.path.join(project_dir, "Relatórios", "EDOC"),
                os.path.join(project_dir, "Relatórios", "ZBRMMT416"),
                os.path.join(project_dir, "Relatórios", "VTIN")
            ]
            
            arquivos_removidos = 0
            
            for pasta in pastas_para_limpar:
                if os.path.exists(pasta):
                    # Listar todos os arquivos na pasta
                    arquivos = [f for f in os.listdir(pasta) if os.path.isfile(os.path.join(pasta, f))]
                    
                    for arquivo in arquivos:
                        try:
                            arquivo_path = os.path.join(pasta, arquivo)
                            os.remove(arquivo_path)
                            arquivos_removidos += 1
                            print(f"[INFO] Removido: {arquivo}")
                        except Exception as e:
                            print(f"[WARN] Não foi possível remover {arquivo}: {str(e)}")
                else:
                    print(f"[INFO] Pasta não existe (será criada): {pasta}")
            
            if arquivos_removidos > 0:
                print(f"[INFO] ✅ {arquivos_removidos} arquivo(s) removido(s) das pastas de relatórios")
            else:
                print("[INFO] Nenhum arquivo para limpar")
            
            return True, f"✅ Pastas de relatórios limpas ({arquivos_removidos} arquivo(s) removido(s))"
            
        except Exception as e:
            print(f"[ERROR] Erro ao limpar pastas: {str(e)}")
            return False, f"⚠️ Erro ao limpar pastas: {str(e)}"
    
    def concatenate_edoc_reports(self):
        """
        Concatena todos os arquivos Excel da pasta Relatórios/EDOC em um único DataFrame
        
        Returns:
            tuple: (success, result_or_message, chaves)
                - Se success=True: result é o DataFrame consolidado, chaves é a lista de chaves únicas
                - Se success=False: result é a mensagem de erro, chaves é None
        """
        try:
            project_dir = os.path.abspath(os.path.dirname(__file__))
            reports_dir = os.path.join(project_dir, "Relatórios", "EDOC")
            
            # Verificar se a pasta existe
            if not os.path.exists(reports_dir):
                return False, f"❌ Pasta não encontrada: {reports_dir}", None
            
            # Listar todos os arquivos Excel na pasta
            excel_files = [f for f in os.listdir(reports_dir) if f.endswith('.xlsx') or f.endswith('.xls')]
            
            if not excel_files:
                return False, "⚠️ Nenhum arquivo Excel encontrado na pasta Relatórios/EDOC", None
            
            # Lista para armazenar os DataFrames
            dfs = []
            
            # Ler cada arquivo Excel e adicionar à lista
            for file in excel_files:
                file_path = os.path.join(reports_dir, file)
                try:
                    df = pd.read_excel(file_path)
                    
                    if 'Descr.st.processo' in df.columns:
                                    df = df[df['Descr.st.processo'] != 'Rejeitado']
                    else:
                        print(f"[WARN] Coluna 'Descr.st.processo' não encontrada em {file}")

                    dfs.append(df)
                    print(f"[INFO] Arquivo lido: {file} ({len(df)} linhas)")
                except Exception as e:
                    print(f"[WARN] Erro ao ler {file}: {str(e)}")
                    continue
            
            # Verificar se conseguiu ler algum arquivo
            if not dfs:
                return False, "❌ Não foi possível ler nenhum arquivo Excel", None
            
            # Concatenar todos os DataFrames
            df_consolidado = pd.concat(dfs, ignore_index=True)
            
            print(f"[INFO] Concatenação concluída: {len(df_consolidado)} linhas no total")
            
            # Extrair chaves de acesso únicas - busca flexível
            chaves = None
            coluna_chave = None
            
            # Tentar encontrar a coluna de chave de acesso (busca case-insensitive e variações)
            possibilidades = ['Chave de Acesso', 'Chave de acesso', 'chave de acesso', 
                            'CHAVE DE ACESSO', 'Chave Acesso', 'ChaveAcesso']
            
            for col in df_consolidado.columns:
                # Remover espaços extras e comparar
                col_limpo = str(col).strip()
                if col_limpo in possibilidades or 'chave' in col_limpo.lower() and 'acesso' in col_limpo.lower():
                    coluna_chave = col
                    break
            
            if coluna_chave:
                chaves = df_consolidado[coluna_chave].dropna().unique()
                print(f"[INFO] Coluna encontrada: '{coluna_chave}'")
                print(f"[INFO] {len(chaves)} chaves únicas encontradas")
            else:
                print(f"[WARN] Coluna de chave de acesso não encontrada no DataFrame consolidado")
                print(f"[WARN] Colunas disponíveis: {list(df_consolidado.columns)}")
            
            return True, df_consolidado, chaves

        except Exception as e:
            return False, f"❌ Erro ao concatenar relatórios: {str(e)}", None
    
    def close_session(self):
        """Fecha a sessão SAP criada"""
        try:
            if self.session:
                self.session.findById("wnd[0]").close()
                self.session = None
        except:
            pass
    
    def buscar_chaves_zbr416(self, chaves):
        """
        Busca informações das chaves de acesso na transação ZBRMMT416
        
        Args:
            chaves: Lista ou array com as chaves de acesso
        
        Returns:
            tuple: (success, message)
        """
        try:
            if not self.session:
                return False, "❌ Não há sessão SAP ativa"
            
            # Converter chaves para DataFrame (uma coluna)
            df_chaves = pd.DataFrame({'Chaves': chaves})
            
            # Copiar para clipboard
            df_chaves.to_clipboard(sep='\t', index=False, header=False)
            
            # Criar pasta de relatórios
            project_dir = os.path.abspath(os.path.dirname(__file__))
            reports_dir = os.path.join(project_dir, "Relatórios", "ZBRMMT416")
            os.makedirs(reports_dir, exist_ok=True)
            
            # Definir caminho do arquivo de saída
            output_path = os.path.join(reports_dir)
            
            print(f"[INFO] Caminho de saída: {output_path}")
            
            # Maximizar janela
            self.session.FindById("wnd[0]").Maximize()
            time.sleep(1)
            
            # Abrir transação ZBRMMT416
            self.session.FindById("wnd[0]/tbar[0]/okcd").Text = "ZBRMMT416"
            self.session.FindById("wnd[0]").SendVKey(0)
            time.sleep(2)
            
            # Pressionar botão de valores múltiplos
            self.session.FindById("wnd[0]/usr/btn%_S_NFEID_%_APP_%-VALU_PUSH").Press()
            time.sleep(1)
            
            # Pressionar botão 23 (importar da área de transferência)
            self.session.FindById("wnd[1]/tbar[0]/btn[23]").Press()
            time.sleep(1)
            
            # Confirmar a janela popup (SendVKey 12 = Enter em popup)
            self.session.FindById("wnd[2]").SendVKey(12)
            time.sleep(1)
            
            # Pressionar botão 24 (colar)
            self.session.FindById("wnd[1]/tbar[0]/btn[24]").Press()
            time.sleep(1)
            
            # Confirmar (botão 8)
            self.session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
            time.sleep(1)
            
            # Preencher campo de arquivo de saída
            self.session.FindById("wnd[0]/usr/ctxtP_FILE").Text = reports_dir
            time.sleep(0.5)
            
            # Preencher outros campos
            self.session.FindById("wnd[0]/usr/txtP_BACK").Text = "99999"
            time.sleep(0.3)
            
            self.session.FindById("wnd[0]/usr/txtP_DIV").Text = "99999"
            time.sleep(0.3)
            
            # Focar no campo de arquivo e posicionar cursor
            self.session.FindById("wnd[0]/usr/ctxtP_FILE").SetFocus()
            self.session.FindById("wnd[0]/usr/ctxtP_FILE").CaretPosition = len(output_path)
            time.sleep(0.5)
            
            # Executar (botão 8 da toolbar)
            self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
            time.sleep(5)  # Aguardar processamento
            
            # Verificar se o arquivo foi criado
            if os.path.exists(output_path):
                file_size = os.path.getsize(output_path)
                arquivos_zip = list(Path(reports_dir).glob('*.zip'))
            
                if not arquivos_zip:
                    return False
                
                # Pegar o primeiro ZIP encontrado
                arquivo_zip = arquivos_zip[0]
                
                # Extrair ZIP na mesma pasta
                with zipfile.ZipFile(arquivo_zip, 'r') as zip_ref:
                    zip_ref.extractall(reports_dir)
                return True, f"✅ Relatório ZBRMMT416 gerado: ({file_size} bytes)"
                
            else:
                return True, f"⚠️ Comando executado, mas arquivo não encontrado:"
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"[ERROR] Detalhes do erro:\n{error_details}")
            return False, f"❌ Erro ao buscar informações na ZBRMMT416: {str(e)}"

    def extrair_zip(pasta_destino):
        """Extrai arquivo ZIP encontrado na pasta"""
        try:
            # Procurar arquivo ZIP
            arquivos_zip = list(Path(pasta_destino).glob('*.zip'))
            
            if not arquivos_zip:
                return False
            
            # Pegar o primeiro ZIP encontrado
            arquivo_zip = arquivos_zip[0]
            
            # Extrair ZIP na mesma pasta
            with zipfile.ZipFile(arquivo_zip, 'r') as zip_ref:
                zip_ref.extractall(pasta_destino)
            
            return True
            
        except Exception as e:
            st.error(f"❌ Erro ao extrair ZIP: {str(e)}")
            return False
    
    def ler_xmls_zbr416(self):
        """
        Lê todos os arquivos XML da pasta ZBRMMT416 e extrai informações
        
        Returns:
            tuple: (success, result_or_message)
                - Se success=True: result é um DataFrame com as informações extraídas
                - Se success=False: result é a mensagem de erro
        """
        try:
            project_dir = os.path.abspath(os.path.dirname(__file__))
            reports_dir = os.path.join(project_dir, "Relatórios", "ZBRMMT416")
            
            # Verificar se a pasta existe
            if not os.path.exists(reports_dir):
                return False, f"❌ Pasta não encontrada: {reports_dir}"
            
            # Listar todos os arquivos XML na pasta
            xml_files = list(Path(reports_dir).glob('*.xml'))
            
            if not xml_files:
                return False, "⚠️ Nenhum arquivo XML encontrado na pasta Relatórios/ZBRMMT416"
            
            print(f"[INFO] {len(xml_files)} arquivo(s) XML encontrado(s)")
            
            # Lista para armazenar os dados extraídos
            dados_extraidos = []
            
            # Namespace do XML da NFe
            ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
            
            # Processar cada arquivo XML
            for xml_file in xml_files:
                try:
                    tree = ET.parse(xml_file)
                    root = tree.getroot()
                    
                    # Extrair Chave NF (do atributo Id da tag infNFe)
                    info_nfe = root.find('.//nfe:infNFe', ns)
                    chave_nf = ''
                    if info_nfe is not None and 'Id' in info_nfe.attrib:
                        # Remove o prefixo 'NFe' da chave
                        chave_nf = info_nfe.attrib['Id'].replace('NFe', '')
                    
                    # Extrair Nome do Fornecedor (xNome do emitente)
                    nome_fornecedor_element = root.find('.//nfe:emit/nfe:xNome', ns)
                    nome_fornecedor = nome_fornecedor_element.text if nome_fornecedor_element is not None else ''
                    
                    # Extrair CNPJ Emissor
                    cnpj_emissor_element = root.find('.//nfe:emit/nfe:CNPJ', ns)
                    cnpj_emissor = cnpj_emissor_element.text if cnpj_emissor_element is not None else ''
                    
                    # Extrair CNPJ Destinatário
                    cnpj_dest_element = root.find('.//nfe:dest/nfe:CNPJ', ns)
                    cnpj_destinatario = cnpj_dest_element.text if cnpj_dest_element is not None else ''
                    
                    # Extrair CFOP
                    cfop_element = root.find('.//nfe:det/nfe:prod/nfe:CFOP', ns)
                    cfop = cfop_element.text if cfop_element is not None else ''
                    
                    # Extrair Quantidade
                    quantidade_element = root.find('.//nfe:det/nfe:prod/nfe:qTrib', ns)
                    quantidade = quantidade_element.text if quantidade_element is not None else ''
                    
                    # Extrair Valor Total da Nota
                    valor_total_element = root.find('.//nfe:total/nfe:ICMSTot/nfe:vNF', ns)
                    valor_total = valor_total_element.text if valor_total_element is not None else ''
                    
                    # Extrair Peso Líquido (transp/vol/pesoL)
                    peso_liquido_element = root.find('.//nfe:transp/nfe:vol/nfe:pesoL', ns)
                    peso_liquido = peso_liquido_element.text if peso_liquido_element is not None else ''
                    
                    # Adicionar dados à lista
                    dados_extraidos.append({
                        'Chave NF': chave_nf,
                        'Nome Fornecedor': nome_fornecedor,
                        'CNPJ Emissor': cnpj_emissor,
                        'CNPJ Destinatário': cnpj_destinatario,
                        'CFOP': cfop,
                        'Quantidade': quantidade,
                        'Valor Total NF': valor_total,
                        'Peso Líquido': peso_liquido,
                        'Arquivo': xml_file.name
                    })
                    
                    print(f"[INFO] Processado: {xml_file.name}")
                    
                except Exception as e:
                    print(f"[WARN] Erro ao processar {xml_file.name}: {str(e)}")
                    continue
            
            # Verificar se conseguiu processar algum arquivo
            if not dados_extraidos:
                return False, "❌ Não foi possível processar nenhum arquivo XML"
            
            # Criar DataFrame com os dados extraídos
            df_xmls = pd.DataFrame(dados_extraidos)
            
            print(f"[INFO] {len(df_xmls)} registro(s) extraído(s) dos XMLs")
            
            return True, df_xmls
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"[ERROR] Detalhes do erro:\n{error_details}")
            return False, f"❌ Erro ao processar XMLs: {str(e)}"
    
    def extrair_vtin(self, chave_nf):
        """
        Extrai dados da transação /VTIN/MDE
        
        Args:
            chave_nf: Series/Array do pandas com as chaves de NF
            
        Returns:
            tuple: (success, message, df_vtin)
                - success: bool indicando sucesso
                - message: mensagem de status
                - df_vtin: DataFrame com os dados extraídos (None se falhar)
        """
        try:
            if not self.session:
                return False, "❌ Não há sessão SAP ativa", None
            
            print(f"[INFO] Iniciando extração VTIN com {len(chave_nf)} chaves...")
            
            # Converter chaves para DataFrame (uma coluna)
            df_chaves = pd.DataFrame({'Chaves': chave_nf})
            
            # Copiar para clipboard
            df_chaves.to_clipboard(sep='\t', index=False, header=False)
            time.sleep(0.5)
            
            # Criar pasta de relatórios
            project_dir = os.path.abspath(os.path.dirname(__file__))
            reports_dir = os.path.join(project_dir, "Relatórios", "VTIN")
            os.makedirs(reports_dir, exist_ok=True)
            
            # Maximizar e abrir transação
            self.session.FindById("wnd[0]").Maximize()
            time.sleep(0.5)
            
            self.session.FindById("wnd[0]/tbar[0]/okcd").Text = "/N/VTIN/MDE"
            self.session.FindById("wnd[0]").SendVKey(0)
            time.sleep(2)
            
            # Pressionar botão de valores múltiplos
            self.session.FindById("wnd[0]/usr/btn%_S_ID_%_APP_%-VALU_PUSH").Press()
            time.sleep(1)
            
            # Colar valores (botão 24 = colar da área de transferência)
            try:
                self.session.FindById("wnd[1]/tbar[0]/btn[24]").Press()
                time.sleep(1)
                print("[INFO] Chaves coladas com sucesso")
            except Exception as e:
                print(f"[WARN] Erro ao colar, tentando método alternativo: {str(e)}")
                # Método alternativo: colar diretamente
                self.session.FindById("wnd[1]").SendVKey(24)
                time.sleep(1)
            
            # Confirmar valores
            self.session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
            time.sleep(1)
            
            # Executar busca
            self.session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
            time.sleep(5)  # Aguardar processamento
            
            # Exportar para Excel
            self.session.FindById("wnd[0]/usr/shell/shellcont[0]/shell").PressToolbarContextButton("&MB_EXPORT")
            time.sleep(0.5)
            
            self.session.FindById("wnd[0]/usr/shell/shellcont[0]/shell").SelectContextMenuItem("&XXL")
            time.sleep(1.5)
            
            # Definir caminho e nome do arquivo
            filename = "relatorioVtin.xlsx"
            self.session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = reports_dir
            time.sleep(0.3)
            
            self.session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = filename
            self.session.FindById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = len(filename)
            time.sleep(0.3)
            
            # Confirmar exportação
            self.session.FindById("wnd[1]/tbar[0]/btn[11]").Press()
            time.sleep(2)
            
            # Fechar arquivos Excel que possam ter sido abertos automaticamente
            print("[INFO] Fechando Excel para liberar arquivo...")
            self._close_excel_files()
            time.sleep(2)  # Aguardar Excel liberar o arquivo completamente
            
            print(f"[INFO] Relatório VTIN exportado para: {reports_dir}")
            
            # Ler o arquivo Excel exportado
            file_path = os.path.join(reports_dir, filename)
            
            # Aguardar o arquivo ser criado completamente
            max_wait = 10  # segundos
            wait_count = 0
            while not os.path.exists(file_path) and wait_count < max_wait:
                time.sleep(1)
                wait_count += 1
            
            if not os.path.exists(file_path):
                return False, f"❌ Arquivo não encontrado após exportação: {file_path}", None
            
            # Ler o arquivo Excel
            try:
                df_vtin = pd.read_excel(file_path)
                print(f"[INFO] Arquivo VTIN lido com sucesso: {len(df_vtin)} linhas, {len(df_vtin.columns)} colunas")
                
                return True, f"✅ Relatório VTIN exportado e lido com sucesso! ({len(df_vtin)} registros)", df_vtin
                
            except Exception as e:
                print(f"[ERROR] Erro ao ler arquivo Excel: {str(e)}")
                return False, f"❌ Erro ao ler arquivo Excel: {str(e)}", None
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"[ERROR] Erro na extração VTIN:\n{error_details}")
            return False, f"❌ Erro ao extrair VTIN: {str(e)}", None
    
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
        - Período de busca: Personalizável (padrão D-2 até D-1)
        - Filtros baseados no arquivo base.csv
        - Exporta relatórios para Excel automaticamente
        - Relatórios salvos em: Relatórios/EDOC/
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
                    
                    # Tratamento da coluna CNPJ: garantir 14 dígitos com zeros à esquerda
                    if 'CNPJ' in st.session_state.df_base.columns:
                        st.session_state.df_base['CNPJ'] = st.session_state.df_base['CNPJ'].astype(str).str.zfill(14)
                    
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
    
    st.markdown("---")
    
    # Configuração de datas
    st.markdown("### 📅 Período de Busca")
    
    col_date1, col_date2 = st.columns(2)
    
    with col_date1:
        data_inicial = st.date_input(
            "Data Inicial",
            value=datetime.now() - timedelta(days=2),
            format="DD/MM/YYYY"
        )
    
    with col_date2:
        data_final = st.date_input(
            "Data Final",
            value=datetime.now() - timedelta(days=1),
            format="DD/MM/YYYY"
        )
    
    st.markdown("---")
    
    # Botão de execução
    can_execute = (
        st.session_state.df_base is not None and 
        'CNPJ' in st.session_state.df_base.columns
    )
    if can_execute:
        if st.button("Executar Automação", type="primary", width="stretch"):
            with st.spinner("Executando automação... Não interaja com o SAP!"):
                progress_bar = st.progress(0, text="Iniciando...")
                
                # Criar nova conexão e sessão
                sap = SAPConnection()
                progress_bar.progress(5, text="Limpando pastas de relatórios...")
                progress_bar.progress(10, text="Conectando ao SAP...")
                success, message = sap.connect_and_create_session()
                
                if not success:
                    st.error(message)
                else:
                    progress_bar.progress(20, text="Sessão criada com sucesso...")
                    progress_bar.progress(30, text="Abrindo transação...")
                    progress_bar.progress(50, text="Aplicando filtros...")
                    progress_bar.progress(70, text="Executando busca e exportando relatórios...")
                    
                    success, message = sap.execute_edoc_cockpit_automation(
                        st.session_state.df_base,
                        data_inicial,
                        data_final
                    )
                    
                    progress_bar.progress(100, text="Concluído!")
                    
                    if success:
                        # Concatenar relatórios exportados
                        progress_bar.progress(75, text="Consolidando relatórios EDOC...")
                        concat_success, result, chaves = sap.concatenate_edoc_reports()
                        sap.close_session()
                        
                        if concat_success:
                            df_consolidado = result
                            
                            # Buscar informações na ZBRMMT416 (somente se houver chaves)
                            if chaves is not None and len(chaves) > 0:
                                progress_bar.progress(80, text="Buscando XMLs na ZBRMMT416...")
                                success, message = sap.connect_and_create_session()
                                zbr_success, zbr_message = sap.buscar_chaves_zbr416(chaves)
                                sap.close_session()
                                
                                if zbr_success:
                                    # Processar arquivos XML gerados
                                    progress_bar.progress(85, text="Processando arquivos XML...")
                                    xml_success, xml_result = sap.ler_xmls_zbr416()
                                    
                                    if xml_success:
                                        df_xmls = xml_result
                                        
                                        # Salvar DataFrame em Excel para referência
                                        project_dir = os.path.abspath(os.path.dirname(__file__))
                                        reports_dir = os.path.join(project_dir, "Relatórios", "ZBRMMT416")
                                        excel_path = os.path.join(reports_dir, "dados_xmls_consolidados.xlsx")
                                        df_xmls.to_excel(excel_path, index=False)
                                        
                                        # Fazer merge dos DataFrames
                                        progress_bar.progress(90, text="Consolidando dados EDOC+XML...")
                                        
                                        # Tentar merge por nome de coluna primeiro, depois por índice
                                        try:
                                            # Tentar encontrar a coluna de chave no df_consolidado
                                            coluna_chave_consolidado = None
                                            for col in df_consolidado.columns:
                                                if 'chave' in str(col).lower() and 'acesso' in str(col).lower():
                                                    coluna_chave_consolidado = col
                                                    break
                                            
                                            if coluna_chave_consolidado and 'Chave NF' in df_xmls.columns:
                                                # Padronizar chaves antes do merge
                                                def padronizar_chave_merge1(chave):
                                                    """Padroniza chave de NF para formato limpo"""
                                                    if pd.isna(chave):
                                                        return ''
                                                    chave_str = str(chave).strip()
                                                    chave_str = chave_str.upper().replace('NFE', '').replace('NF-E', '').strip()
                                                    chave_str = ''.join(filter(str.isdigit, chave_str))
                                                    if len(chave_str) == 44:
                                                        return chave_str
                                                    elif len(chave_str) < 44:
                                                        return chave_str.zfill(44)
                                                    else:
                                                        return chave_str[:44]
                                                
                                                df_consolidado['_chave_merge_temp'] = df_consolidado[coluna_chave_consolidado].apply(padronizar_chave_merge1)
                                                df_xmls['_chave_merge_temp'] = df_xmls['Chave NF'].apply(padronizar_chave_merge1)
                                                
                                                # Merge usando chaves padronizadas
                                                df_int = pd.merge(df_consolidado, df_xmls, 
                                                                left_on='_chave_merge_temp', 
                                                                right_on='_chave_merge_temp', 
                                                                how='left',
                                                                suffixes=('', '_xml'))
                                                
                                                # Remover coluna temporária
                                                df_int = df_int.drop(columns=['_chave_merge_temp'])
                                                
                                                # Contar matches
                                                matches = df_int['Chave NF'].notna().sum()
                                            else:
                                                # Verificar se as colunas existem
                                                if len(df_consolidado.columns) > 4 and len(df_xmls.columns) > 0:
                                                    # Pegar nomes das colunas pelos índices
                                                    col_consolidado = df_consolidado.columns[4]  # Coluna E (índice 4)
                                                    col_xmls = df_xmls.columns[0]  # Coluna A (índice 0)
                                                    
                                                    # Padronizar chaves antes do merge
                                                    def padronizar_chave_merge2(chave):
                                                        """Padroniza chave de NF para formato limpo"""
                                                        if pd.isna(chave):
                                                            return ''
                                                        chave_str = str(chave).strip()
                                                        chave_str = chave_str.upper().replace('NFE', '').replace('NF-E', '').strip()
                                                        chave_str = ''.join(filter(str.isdigit, chave_str))
                                                        if len(chave_str) == 44:
                                                            return chave_str
                                                        elif len(chave_str) < 44:
                                                            return chave_str.zfill(44)
                                                        else:
                                                            return chave_str[:44]
                                                    
                                                    df_consolidado['_chave_merge_temp'] = df_consolidado[col_consolidado].apply(padronizar_chave_merge2)
                                                    df_xmls['_chave_merge_temp'] = df_xmls[col_xmls].apply(padronizar_chave_merge2)
                                                    
                                                    # Merge usando chaves padronizadas
                                                    df_int = pd.merge(df_consolidado, df_xmls,
                                                                    left_on='_chave_merge_temp',
                                                                    right_on='_chave_merge_temp',
                                                                    how='left',
                                                                    suffixes=('', '_xml'))
                                                    
                                                    # Remover coluna temporária
                                                    df_int = df_int.drop(columns=['_chave_merge_temp'])
                                                    
                                                    # Contar matches
                                                    col_check = col_xmls if col_xmls in df_int.columns else 'Chave NF'
                                                    if col_check in df_int.columns:
                                                        matches = df_int[col_check].notna().sum()
                                                else:
                                                    df_int = df_consolidado.copy()
                                            
                                            # Salvar DataFrame integrado
                                            integrated_excel_path = os.path.join(reports_dir, "dados_integrados_final.xlsx")
                                            df_int.to_excel(integrated_excel_path, index=False)
                                            
                                            # Executar consulta VTIN
                                            progress_bar.progress(93, text="Extraindo dados VTIN...")
                                            success, message = sap.connect_and_create_session()
                                            
                                            if success:
                                                vtin_success, vtin_message, df_vtin = sap.extrair_vtin(chaves)
                                                
                                                if vtin_success:
                                                    # Salvar DataFrame VTIN
                                                    if df_vtin is not None:
                                                        # Salvar em Excel consolidado
                                                        project_dir = os.path.abspath(os.path.dirname(__file__))
                                                        reports_dir_zbr = os.path.join(project_dir, "Relatórios", "ZBRMMT416")
                                                        vtin_consolidated_path = os.path.join(reports_dir_zbr, "dados_vtin_consolidados.xlsx")
                                                        df_vtin.to_excel(vtin_consolidated_path, index=False)
                                                        
                                                        # Identificar coluna de chave em df_int
                                                        colunas_possiveis_int = []
                                                        for col in df_int.columns:
                                                            col_lower = str(col).lower().strip()
                                                            if 'chave' in col_lower:
                                                                colunas_possiveis_int.append(col)
                                                        
                                                        coluna_chave_int = None
                                                        for col in df_int.columns:
                                                            col_lower = str(col).lower().strip()
                                                            if 'chave' in col_lower and ('nf' in col_lower or 'acesso' in col_lower):
                                                                coluna_chave_int = col
                                                                break
                                                        
                                                        # Se não encontrou, tentar apenas com "chave"
                                                        if not coluna_chave_int and len(colunas_possiveis_int) > 0:
                                                            coluna_chave_int = colunas_possiveis_int[0]
                                                        
                                                        # Identificar coluna de chave em df_vtin
                                                        coluna_chave_vtin = None
                                                        colunas_candidatas_vtin = []
                                                        
                                                        for col in df_vtin.columns:
                                                            col_lower = str(col).lower().strip()
                                                            if 'id' in col_lower and 'chave' in col_lower:
                                                                colunas_candidatas_vtin.append((col, col_lower))
                                                        
                                                        if len(colunas_candidatas_vtin) > 0:
                                                            # Priorizar coluna que contenha "id" (mais específica)
                                                            for col, col_lower in colunas_candidatas_vtin:
                                                                if 'id' in col_lower and 'chave' in col_lower and 'acesso' in col_lower:
                                                                    coluna_chave_vtin = col
                                                                    break
                                                            
                                                            # Se não encontrou com "id", usar a primeira
                                                            if not coluna_chave_vtin:
                                                                coluna_chave_vtin = colunas_candidatas_vtin[0][0]
                                                        
                                                        # Identificar coluna de documento em df_vtin
                                                        coluna_doc_vtin = None
                                                        for col in df_vtin.columns:
                                                            col_lower = str(col).lower().strip()
                                                            if 'documento' in col_lower or 'nº' in col_lower:
                                                                coluna_doc_vtin = col
                                                                break
                                                        
                                                        if coluna_chave_int and coluna_chave_vtin:
                                                            # Padronizar chaves - manter apenas números com 44 dígitos
                                                            def padronizar_chave(chave):
                                                                """Padroniza chave de NF para formato limpo (apenas números)"""
                                                                if pd.isna(chave):
                                                                    return ''
                                                                # Converter para string e remover espaços
                                                                chave_str = str(chave).strip()
                                                                # Manter apenas dígitos numéricos
                                                                chave_str = ''.join(filter(str.isdigit, chave_str))
                                                                # Garantir que tenha exatamente 44 dígitos
                                                                if len(chave_str) == 44:
                                                                    return chave_str
                                                                elif len(chave_str) < 44:
                                                                    return chave_str.zfill(44)
                                                                else:
                                                                    return chave_str[:44]
                                                            
                                                            # Aplicar padronização
                                                            df_int['_chave_temp'] = df_int[coluna_chave_int].apply(padronizar_chave)
                                                            df_vtin['_chave_temp'] = df_vtin[coluna_chave_vtin].apply(padronizar_chave)
                                                            
                                                            # Selecionar colunas para merge
                                                            colunas_vtin = ['_chave_temp', 'Número da nota', 'Data escrituração']
                                                            if coluna_doc_vtin:
                                                                colunas_vtin.append(coluna_doc_vtin)
                                                            
                                                            # Realizar merge
                                                            df_final = pd.merge(
                                                                df_int, 
                                                                df_vtin[colunas_vtin], 
                                                                on='_chave_temp', 
                                                                how='left'
                                                            )

                                                            # Remover coluna temporária
                                                            df_final = df_final.drop(columns=['_chave_temp'])
                                                            
                                                            # Salvar DataFrame final
                                                            progress_bar.progress(100, text="Concluído!")
                                                            final_excel_path = os.path.join(reports_dir_zbr, "dados_final_completo.xlsx")
                                                            df_final = df_final[['Data de emissão da NF-e', 'Data escrituração', 'Número da nota', 'Série de dados', 'Chave NF', 'Nome Fornecedor', 'CNPJ Emissor', 'CNPJ Destinatário', 'CFOP', 'Valor Total NF', 'Peso Líquido', 'Nº documento']]
                                                            for col in ['Data de emissão da NF-e', 'Data escrituração']:
                                                                if col in df_final.columns:
                                                                    df_final[col] = pd.to_datetime(df_final[col], errors='coerce').dt.strftime('%d/%m/%Y')
                                                            df_final.to_excel(final_excel_path, index=False)
                                                            
                                                            # Exibir apenas a tabela final
                                                            st.success(f"✅ Processo concluído! {len(df_final)} registros processados")
                                                            st.markdown("### 📊 Dados Consolidados")
                                                            st.dataframe(df_final, use_container_width=True)
                                                            st.info(f"💾 Arquivo salvo: dados_final_completo.xlsx")
                                                            st.balloons()
                                                        else:
                                                            st.error(f"❌ Erro ao identificar colunas de chave para merge")
                                                else:
                                                    st.error(vtin_message)
                                                
                                                sap.close_session()
                                            else:
                                                st.error(f"❌ Erro ao criar sessão para VTIN: {message}")
                                            
                                        except Exception as e:
                                            st.error(f"❌ Erro ao fazer merge: {str(e)}")
                                            
                                    else:
                                        st.error(xml_result)
                                else:
                                    st.error(zbr_message)
                            else:
                                st.error("⚠️ Nenhuma chave de acesso encontrada")
                        else:
                            st.error(result)
                    else:
                        st.error(message)
                    
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
        
        1. **Limpa** as pastas de relatórios anteriores
        2. **Cria nova sessão** no SAP automaticamente
        3. **Maximiza** a janela do SAP
        4. **Abre** a transação EDOC_COCKPIT
        5. **Navega** pelos menus de seleção
        6. **Aplica filtros** com CNPJs da base.csv
        7. **Define período** de busca conforme selecionado
        8. **Executa** a consulta
        9. **Exporta relatórios** para Excel
        10. **Salva arquivos** em Relatórios/EDOC/
        11. **Fecha** a sessão criada
        
        ### 📋 Requisitos:
        - ✅ Arquivo base.csv com coluna 'CNPJ'
        - ✅ SAP GUI aberto e logado
        - ✅ Não interagir com o SAP durante execução
        
        ### 📁 Arquivos Gerados:
        - edoc_opcao0.xlsx até edoc_opcao10.xlsx
        - Salvos em: Relatórios/EDOC/
        
        ### ⏱️ Tempo estimado:
        - 1-2 minutos (depende da quantidade de relatórios)
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
