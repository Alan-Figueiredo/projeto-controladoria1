# Comparador de Balancete por Centro de Custo (CR) 📊

Este projeto é uma ferramenta de automação desenvolvida em Python para validar a integridade de relatórios financeiros. O objetivo principal é comparar um **Balancete Geral** com diversas **Abas Analíticas**, identificando contas que não foram referenciadas, respeitando a hierarquia de contas e os Centros de Resultado (CR).

## 🚀 Funcionalidades

- **Extração Inteligente**: Identifica contas através da fórmula personalizada `=CONTAS_BALANCETE("codigo")`.
- **Deteção de CR via PROCV**: Mapeia automaticamente quais abas pertencem a quais Centros de Custo através da análise de fórmulas `VLOOKUP/PROCV`.
- **Análise Hierárquica**: Verifica se uma conta "pai" já cobre as suas "filhas", evitando alertas redundantes.
- **Relatórios Formatados**: Gera um arquivo Excel (`.xlsx`) com formatação condicional (cores e bordas) para fácil identificação de pendências.
- **Segurança de Dados**: Opera sobre uma cópia temporária para evitar corrupção do arquivo original.

## 🏗️ Arquitetura e Guia do Projeto

```text
📁 Seu_Projeto/
├── 📄 main.py               # O maestro (ponto de entrada)
├── 📄 requirements.txt      # Dependências do projeto
│
├── 📁 data/                 # Diretórios de dados (I/O)
│   ├── 📁 input/            # Coloque a sua planilha original aqui
│   ├── 📁 output/           # Relatórios gerados salvos aqui
│   └── 📁 temp/             # Cópias de segurança temporárias
│
└── 📁 src/                  # Código-fonte principal
    ├── 📄 config.py         # Configurações globais
    ├── 📁 core/             # Lógica central (Extrator e Transformador)
    ├── 📁 services/         # Serviços de saída (Relatório)
    └── 📁 utils/            # Funções utilitárias

```
---------------------------------------------------------------------------
## 🛠️ PRÉ-REQUISITOS
---------------------------------------------------------------------------
- Python 3.14+
- Microsoft Excel instalado (necessário para processamento via pywin32)
- Dependências: pip install pandas openpyxl pyxlsb pywin32

---------------------------------------------------------------------------
## 📂 COMO CONFIGURAR E USAR
---------------------------------------------------------------------------
1. Configure os caminhos em 'src/config.py'.
2. Adicione sua planilha em 'data/input/'.
3. Execute o comando: python main.py
4. Verifique o resultado em: 'data/output/'.

---------------------------------------------------------------------------
## 📋 STATUS DE PENDÊNCIA NO RELATÓRIO
---------------------------------------------------------------------------
- **🔴 Falta na Aba** : Conta existe no Balancete, mas não foi referenciada.
- **🟡 Parcial**      : Conta "pai" ausente, mas "filhos" cobrem parte do valor.
- **✅ Coberto**      : A conta ou seu ancestral já está validado na aba.

---------------------------------------------------------------------------
## 👨‍💻 AUTOR E NOTAS
---------------------------------------------------------------------------
 - **Desenvolvido por: Alan Cruz de Figueiredo**
- **Requisito de Sistema:** Windows (devido ao motor COM/win32com do Excel)
---------------------------------------------------------------------------
