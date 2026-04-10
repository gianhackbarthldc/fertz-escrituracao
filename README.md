# 🔧 Automação SAP - Escrituração Fertz

Interface web desenvolvida com Streamlit para automação completa do processo de escrituração de NF-e, integrando as transações SAP **EDOC_COCKPIT**, **ZBRMMT416** e **/VTIN/MDE**.

## 📋 Descrição

A aplicação executa um pipeline de 3 etapas para identificar notas fiscais eletrônicas pendentes de escrituração, filtrando por um conjunto de CNPJs configurados em `base.csv` e por um período de datas selecionado pelo usuário. O resultado final é uma planilha Excel com as NF-e que ainda **não foram emitidas** no sistema.

---

## 🔄 Fluxo Completo da Automação

```
base.csv (CNPJs)
      │
      ▼
┌─────────────────────────────────────────────────────┐
│  ETAPA 1 — EDOC_COCKPIT                             │
│  • Cria nova sessão SAP                             │
│  • Limpa pastas de relatórios anteriores            │
│  • Filtra por CNPJs (base.csv) + período de datas   │
│  • Exporta até 9 lotes em Excel → Relatórios/EDOC/  │
│  • Concatena os Excels e remove notas "Rejeitadas"  │
│  • Extrai as Chaves de Acesso (44 dígitos) únicas   │
└────────────────────┬────────────────────────────────┘
                     │ Chaves de Acesso
                     ▼
┌─────────────────────────────────────────────────────┐
│  ETAPA 2 — ZBRMMT416                                │
│  • Abre nova sessão SAP                             │
│  • Cola as chaves e executa a transação             │
│  • Baixa ZIP com XMLs de NF-e → Relatórios/ZBRMMT416│
│  • Extrai o ZIP e parseia cada XML (NFe namespace)  │
│  Campos extraídos por XML:                          │
│    - Chave NF          (infNFe/@Id)                 │
│    - Nome Fornecedor   (emit/xNome)                 │
│    - CNPJ Emissor      (emit/CNPJ)                  │
│    - CNPJ Destinatário (dest/CNPJ)                  │
│    - CFOP              (det/prod/CFOP)              │
│    - Quantidade        (det/prod/qTrib)             │
│    - Valor Total NF    (total/ICMSTot/vNF)          │
│    - Peso Líquido      (transp/vol/pesoL)           │
│  • Merge EDOC ← XML por Chave de Acesso             │
│  • Salva dados_integrados_final.xlsx                │
└────────────────────┬────────────────────────────────┘
                     │ df integrado + Chaves
                     ▼
┌─────────────────────────────────────────────────────┐
│  ETAPA 3 — /VTIN/MDE                                │
│  • Abre nova sessão SAP                             │
│  • Cola as chaves e executa a transação             │
│  • Exporta Excel → Relatórios/VTIN/relatorioVtin.xlsx│
│  • Merge com dados integrados por Chave de Acesso   │
│  • Filtra: mantém apenas NF-e SEM número de doc SAP │
│    (= notas ainda não escrituradas)                 │
│  • Salva dados_final_completo.xlsx                  │
│  • Exibe tabela final na interface Streamlit        │
└─────────────────────────────────────────────────────┘
```

---

## 📊 Colunas do Arquivo Final (`dados_final_completo.xlsx`)

| Coluna | Origem |
|---|---|
| Data de emissão da NF-e | EDOC_COCKPIT |
| Número da nota | /VTIN/MDE |
| Série de dados | EDOC_COCKPIT |
| Chave NF | XML (infNFe/@Id) |
| Nome Fornecedor | XML (emit/xNome) |
| CNPJ Emissor | XML (emit/CNPJ) |
| CNPJ Destinatário | XML (dest/CNPJ) |
| CFOP | XML (det/prod/CFOP) |
| Valor Total NF | XML (total/ICMSTot/vNF) |
| Peso Líquido | XML (transp/vol/pesoL) |

---

## 📦 Pré-requisitos

1. **SAP GUI** instalado (versão 7.60+) e logado
2. **Python 3.7+**
3. **SAP GUI Scripting habilitado**:
   - No SAP Logon: `Opções → Accessibility & Scripting → Scripting`
   - Marcar: "Enable scripting"
4. Arquivo **`base.csv`** na raiz do projeto com coluna `CNPJ` (CNPJs dos fornecedores a filtrar)

### Formato do `base.csv`

```csv
CNPJ
12345678000190
98765432000100
...
```

> CNPJs são automaticamente preenchidos com zeros à esquerda para 14 dígitos.

---

## 🔧 Instalação

```powershell
pip install -r requirements.txt
```

Ou execute diretamente:

```powershell
.\run.bat
```

---

## 🎮 Como Usar

1. **Abra o SAP Logon** e faça login no sistema
2. **Execute a aplicação**:
   ```powershell
   streamlit run main.py
   ```
3. **No navegador**:
   - Confirme que o `base.csv` foi carregado (métricas de registros e CNPJs exibidas)
   - Selecione o **período de busca** (padrão: D-2 até D-1)
   - Clique em **"Executar Automação"**
   - Aguarde — **não interaja com o SAP** durante a execução
4. Ao concluir, a tabela final é exibida na tela e o arquivo `dados_final_completo.xlsx` é salvo em `Relatórios/ZBRMMT416/`

---

## 📂 Estrutura do Projeto

```
fertz-escrituracao/
├── main.py              # Aplicação principal (Streamlit + lógica SAP)
├── base.csv             # CNPJs dos fornecedores a filtrar
├── requirements.txt     # Dependências Python
├── run.bat              # Atalho para execução no Windows
├── .gitignore           # Ignora a pasta Relatórios/
├── README.md            # Este arquivo
└── Relatórios/          # Gerado em tempo de execução (ignorado pelo git)
    ├── EDOC/            # Excels exportados do EDOC_COCKPIT
    ├── ZBRMMT416/       # XMLs de NF-e + Excels consolidados
    └── VTIN/            # Excel exportado do /VTIN/MDE
```

---

## 🛠️ Tecnologias

| Biblioteca | Uso |
|---|---|
| `streamlit` | Interface web |
| `pywin32` | Integração com SAP GUI Scripting |
| `pandas` | Manipulação de dados e merges |
| `openpyxl` / `xlrd` | Leitura/escrita de arquivos Excel |
| `xml.etree.ElementTree` | Parse dos XMLs de NF-e |

---

## 🐛 Troubleshooting

### ❌ "SAP não tem conexões ativas"
- Certifique-se de que o SAP GUI está aberto e com sessão logada antes de executar.

### ❌ "Coluna CNPJ não encontrada"
- Verifique se o `base.csv` possui uma coluna chamada exatamente `CNPJ`.

### ⚠️ Nenhum arquivo Excel encontrado em EDOC
- O período selecionado pode não ter retornado documentos no EDOC_COCKPIT. Amplie o intervalo de datas.

### ⚠️ Nenhum arquivo XML encontrado em ZBRMMT416
- A transação ZBRMMT416 pode não ter gerado o ZIP. Verifique se as chaves de acesso são válidas no ambiente SAP.

### ❌ Erro ao identificar colunas de chave para merge
- As colunas de chave de acesso nos relatórios exportados podem ter nomes diferentes do esperado. Verifique os logs no terminal (`[DEBUG] Colunas disponíveis: ...`).

---

## 🔐 Segurança

- Conecta-se apenas a sessões SAP já autenticadas pelo usuário
- Não armazena nem transmite credenciais
- Utiliza a API oficial do SAP GUI Scripting

---

## 📄 Licença

Uso interno — Louis Dreyfus Company
