# API de Exportação para Excel
## Visão Geral
Esta API permite gerar arquivos Excel (.xlsx) a partir de um payload JSON. Você pode personalizar o conteúdo, como adicionar títulos, cabeçalhos, dados e aplicar formatações (ex.: cores, bordas, alinhamento).

# Endpoint
## POST /api/v1/export
## Descrição
Gera e retorna um arquivo Excel com base no payload fornecido.

Content-Type
application/json

Response Type
application/vnd.openxmlformats-officedocument.spreadsheetml.sheet (binário)

# Exemplo Simples
## Payload de Exemplo

```
    {
  "filename": "exemplo.xlsx",
  "sheets": [
    {
      "sheetName": "Dados",
      "title": {
        "text": "RELATÓRIO SIMPLES",
        "mergeColumns": 3,
        "style": {
          "fontColor": "#FFFFFF",
          "bgColor": "#4CAF50",
          "bold": true,
          "fontSize": 14,
          "align": "center",
          "valign": "middle"
        }
      },
      "records": [
        ["ID", "Nome", "Valor"],
        ["1", "Produto A", 10.50],
        ["2", "Produto B", 20.75],
        ["3", "Produto C", 15.30]
      ],
      "headerStyle": {
        "bgColor": "#E0E0E0",
        "bold": true,
        "border": true,
        "align": "center"
      },
      "cellStyle": {
        "border": true
      }
    }
  ]
}

```


# Explicação do Payload

filename: Nome do arquivo Excel (ex.: "exemplo.xlsx"). Deve terminar com .xlsx.
sheets: Lista de planilhas a incluir.
sheetName: Nome da planilha (ex.: "Dados").
title: Título da planilha, na primeira linha.
text: Texto do título.
mergeColumns: Número de colunas a mesclar.
style: Estilo do título (cor do texto, fundo, etc.).
records: Dados da planilha (lista de listas).
Primeira linha: Cabeçalho.
Demais linhas: Dados.
headerStyle: Estilo do cabeçalho (primeira linha de records).
cellStyle: Estilo das células de dados.
Resultado Esperado
O arquivo exemplo.xlsx terá uma planilha chamada "Dados":

Linha 1: "RELATÓRIO SIMPLES", mesclado nas 3 colunas (A1:C1), fundo verde (#4CAF50), texto branco, negrito, centralizado, fonte tamanho 14.
Linha 2: Cabeçalho ("ID", "Nome", "Valor"), fundo cinza claro (#E0E0E0), negrito, com bordas, centralizado.
Linhas 3 a 5: Dados com bordas.
Resposta
Status: 200 OK
Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
Content-Disposition: attachment; filename=exemplo.xlsx

# Requisição com JavaScript:

```

import axios from 'axios'

const payload = {
  filename: 'exemplo.xlsx',
  sheets: [
    {
      sheetName: 'Dados',
      title: {
        text: 'RELATÓRIO SIMPLES',
        mergeColumns: 3,
        style: {
          fontColor: '#FFFFFF',
          bgColor: '#4CAF50',
          bold: true,
          fontSize: 14,
          align: 'center',
          valign: 'middle'
        }
      },
      records: [
        ['ID', 'Nome', 'Valor'],
        ['1', 'Produto A', 10.50],
        ['2', 'Produto B', 20.75],
        ['3', 'Produto C', 15.30]
      ],
      headerStyle: {
        bgColor: '#E0E0E0',
        bold: true,
        border: true,
        align: 'center'
      },
      cellStyle: {
        border: true
      }
    }
  ]
}

axios.post('http://localhost:8000/api/v1/export', payload, {
  responseType: 'blob'
})
  .then(response => {
    const url = window.URL.createObjectURL(new Blob([response.data]))
    const link = document.createElement('a')
    link.href = url
    link.setAttribute('download', 'exemplo.xlsx')
    document.body.appendChild(link)
    link.click()
    link.remove()
    window.URL.revokeObjectURL(url)
  })
  .catch(error => {
    console.error('Erro ao baixar o arquivo:', error)
  })

```

# Observações
## Personalizações Avançadas
Use records_format para formatar colunas (ex.: moeda, números).
Use tags para ajustar largura de colunas (set_column) ou altura de linhas (set_row).
Dependências
Certifique-se de ter instalado: fastapi, uvicorn, xlsxwriter.
