# Senior VetorRH - Extrator de Apuração de Ponto (PDF para Excel)

Um utilitário de linha de comando (CLI) desenvolvido em Python para automatizar a extração e transformação de relatórios de apuração de ponto exportados do sistema **Senior VetorRH** (`HRAP110.APU`) em planilhas estruturadas do Microsoft Excel.

Este projeto foi criado para otimizar os processos de gestão de manufatura, permitindo a triagem e o tratamento rápido de ocorrências, faltas, horas extras e ajustes de ponto da equipe.

---

## Tecnologias Utilizadas

- **Python 3.10+**
- **pdfplumber**: Extração de texto e parsing de documentos PDF.
- **pandas**: Manipulação e estruturação dos dados.
- **openpyxl**: Formatação e exportação para arquivos `.xlsx`.

---

## Funcionalidades

- **CLI Flexível**: Aceita caminhos personalizados para arquivos de entrada (`.pdf`) e saída (`.xlsx`) via terminal.
- **Parsing e Limpeza Automática**: Identifica e descarta cabeçalhos, rodapés e ruídos de paginação típicos do relatório `HRAP110.APU`.
- **Associação Inteligente**: Mantém a continuidade das marcações e datas para múltiplas ocorrências registradas para o mesmo colaborador.
- **Filtragem por Códigos de Ocorrência**: Processa eventos de ponto com base nos códigos do sistema (`024`, `100`, `115`, `120`, `135`, `200`, `250`, `300`, `301`, `999`, etc.).
- **Planilha Estilizada**:
  - Cabeçalho em tom azul escuro (`#1F4E78`) com fonte branca.
  - Largura ajustada automaticamente para todas as colunas.
  - Congelamento da primeira linha (*Freeze Panes* `A2`) para facilidade de navegação.
  - Filtro automático ativado em todas as colunas.

---

## Estrutura do Repositório

```text
.
├── main.py             # Script principal da aplicação (CLI e parser)
├── requirements.txt    # Dependências do projeto
├── .gitignore          # Regras para ignorar arquivos locais/confidenciais
└── README.md           # Documentação do projeto
```
## Estrutura das Colunas Geradas

A planilha `.xlsx` gerada conta com a seguinte estrutura:

| Coluna | Campo | Descrição |
|---|---|---|
| **A** | `MATRICULA` | Código de matrícula do colaborador no VetorRH |
| **B** | `EMPREGADO` | Nome completo do colaborador |
| **C** | `DIA` | Data da ocorrência (`DD/MM/AAAA`) |
| **D** | `MARCAÇÕES` | Horários das batidas de ponto registradas no dia |
| **E** | `SITUAÇÕES APURADAS` | Código e descrição da situação (ex: *100 HORAS EXTRAS 50%*) |
| **F** | `HORAS` | Quantidade de horas apuradas no formato `HH:MM` |

---

## Como Executar

### 1. Clonar o repositório e preparar o ambiente
```bash
git clone [https://github.com/SEU-USUARIO/SEU-REPOSITORIO.git](https://github.com/SEU-USUARIO/SEU-REPOSITORIO.git)
cd SEU-REPOSITORIO
```
### 2. Instalação das Dependências
```bash
pip install -r requirements.txt
```

### 3. Execução do Script
1. Coloque o relatório em PDF (ex: `HRAP110_TMPFD40.PDF`) na raiz do projeto e execute:
```bash
python main.py
```
   
### 4. Uso Personalizado (CLI)
Você pode especificar os nomes/caminhos dos arquivos de entrada e saída:

```bash
python main.py -i caminho/meu_relatorio.pdf -o relatorio_manufatura.xlsx
```
Parâmetros:

-i / --input: Caminho do arquivo PDF de entrada (Padrão: HRAP110_TMPFD40.PDF).

-o / --output: Caminho do arquivo Excel de saída (Padrão: saida.xlsx).

## Caso de Uso na Gestão da Manufatura

Em ambientes industriais e de manufatura, os relatórios padrão em PDF do VetorRH costumam ter centenas de páginas, o que dificulta a triagem rápida por lideranças e supervisores de turno. 

Ao converter esses dados para Excel:
- A liderança consegue aplicar filtros por **situação** (ex: visualizar apenas faltas não justificadas ou horas extras elevadas).
- Facilita o cruzamento de dados com a escala de turnos da fábrica.
- Agiliza o envio de justificativas de ponto para o setor de Recursos Humanos antes do fechamento do espelho.
