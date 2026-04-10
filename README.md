# 🔧 Automação SAP - EDOC_COCKPIT

Interface web desenvolvida com Streamlit para automação da transação EDOC_COCKPIT do SAP.

## 📋 Descrição

Este projeto fornece uma interface amigável para interagir com o SAP GUI, automatizando o acesso à transação EDOC_COCKPIT (Cockpit de Documentos Eletrônicos).

## 🚀 Funcionalidades

- ✅ Conexão automática com SAP GUI
- ✅ Abertura da transação EDOC_COCKPIT
- ✅ Interface web moderna e intuitiva
- ✅ Monitoramento de status da conexão
- ✅ Informações da sessão SAP em tempo real

## 📦 Pré-requisitos

1. **SAP GUI** instalado (versão 7.60+)
2. **Python 3.7+** instalado
3. **SAP GUI Scripting habilitado**:
   - No SAP Logon: `Opções → Accessibility & Scripting → Scripting`
   - Marcar: "Enable scripting"

## 🔧 Instalação

1. Clone o repositório ou navegue até a pasta do projeto:
```powershell
cd "c:\Users\HACKBG\OneDrive - Louis Dreyfus Company\Documents\Git\painel-dinamico\projects\fertz-escrituracao"
```

2. Instale as dependências:
```powershell
pip install -r requirements.txt
```

## 🎮 Como Usar

1. **Abra o SAP Logon** e faça login no sistema

2. **Execute o aplicativo Streamlit**:
```powershell
streamlit run main.py
```

3. **No navegador** (abre automaticamente):
   - Clique em "🔗 Conectar ao SAP"
   - Clique em "📂 Abrir EDOC_COCKPIT"

## 📂 Estrutura do Projeto

```
fertz-escrituracao/
├── main.py              # Aplicação principal Streamlit
├── requirements.txt     # Dependências Python
└── README.md           # Este arquivo
```

## 🔐 Segurança

- A aplicação se conecta apenas a sessões SAP já autenticadas
- Não armazena credenciais
- Utiliza a API oficial do SAP GUI Scripting

## 🛠️ Tecnologias

- **Streamlit**: Interface web
- **pywin32**: Integração com SAP GUI
- **Python**: Lógica de automação

## 📝 Próximas Funcionalidades

- [ ] Extração de dados do EDOC_COCKPIT
- [ ] Filtros e pesquisas automáticas
- [ ] Export para Excel/CSV
- [ ] Dashboard com métricas
- [ ] Processamento em lote

## 🐛 Troubleshooting

### Erro "SAPGUI not found"
- Certifique-se de que o SAP GUI está instalado e aberto
- Verifique se o scripting está habilitado no SAP Logon

### Erro "No active sessions"
- Faça login no SAP antes de executar a automação
- Mantenha a janela do SAP aberta

## 📞 Suporte

Para dúvidas ou sugestões, consulte a documentação em:
- `skills/sap-gui-scripting/SKILL.md`

## 📄 Licença

Uso interno - Louis Dreyfus Company
