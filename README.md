# Senior VetorRH - Extrator de Apuração de Ponto (PDF para Excel)

Um script em Python desenvolvido para automatizar a extração e transformação de relatórios de apuração de ponto exportados do sistema **Senior VetorRH** (`HRAP110.APU`) em planilhas estruturadas do Microsoft Excel. 

Este projeto foi criado para otimizar os processos de gestão de manufatura, permitindo o tratamento rápido de ocorrências, faltas, horas extras e ajustes de ponto da equipe.

---

## Tecnologias Utilizadas

- **Python 3.12+**
- **pdfplumber**: Extração de texto e marcações de documentos PDF.
- **pandas**: Manipulação e estruturação dos dados em DataFrames.
- **openpyxl**: Formatação estilizada e exportação para arquivos `.xlsx`.

---

## Funcionalidades

- **Parsing de PDFs complexos**: Identifica e ignora cabeçalhos, rodapés e ruídos de impressão das páginas do relatório Senior VetorRH.
- **Associação inteligente de dados**: Mantém a continuidade das marcações e datas para múltiplas ocorrências registradas no mesmo dia/colaborador.
- **Filtragem por Códigos de Situação**: Processa apenas eventos válidos de ponto baseados na tabela interna de situações apuradas (`024`, `100`, `115`, `120`, `135`, `200`, `250`, `300`, `301`, `999`, etc.).
- **Planilha Estilizada Pronta para Uso**:
  - Cabeçalho profissional em tom azul escuro (`#1F4E78`) com fonte branca.
  - Ajuste automático da largura das colunas.
  - Congelamento da primeira linha (*Freeze Panes* `A2`) para facilidade de navegação.
  - Filtro automático habilitado em todas as colunas.

---

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

### 1. Pré-requisitos
Certifique-se de ter o Python instalado na versão 3.10 ou superior.

### 2. Instalação das Dependências
Instale as bibliotecas necessárias utilizando o `pip`:

```bash
pip install pdfplumber pandas openpyxl
```

### 3. Execução do Script
1. Coloque o relatório em PDF (ex: `HRAP110_TMPFD40.PDF`) no mesmo diretório do script.
2. Atualize o nome do arquivo no final do código ou ajuste os caminhos:

```python
texto = extrair_texto("HRAP110_TMPFD40.PDF")
registros = parse_relatorio(texto)
gerar_excel(registros, "saida.xlsx")
```

3. Execute o script em seu terminal ou ambiente de desenvolvimento (ex: Google Colab / Jupyter Notebook):

```bash
python main.py
```

Ao final da execução, será exibida a confirmação no terminal:
```text
1798 linhas geradas para 573 colaboradores
```

---

## Caso de Uso na Gestão da Manufatura

Em ambientes industriais e de manufatura, os relatórios padrão em PDF do VetorRH costumam ter centenas de páginas, o que dificulta a triagem rápida por lideranças e supervisores de turno. 

Ao converter esses dados para Excel:
- A liderança consegue aplicar filtros por **situação** (ex: visualizar apenas faltas não justificadas ou horas extras elevadas).
- Facilita o cruzamento de dados com a escala de turnos da fábrica.
- Agiliza o envio de justificativas de ponto para o setor de Recursos Humanos antes do fechamento do espelho.
