# py-excel-api
JSON to excel API made with python

## Installation
```
pip3 install -r requirements.txt
```

## Usage
```
uvicorn main:app
```
The API will be running at `http://localhost:5000`

### Request Example
```js
async function request() {
    const response = await fetch('http://localhost:5000/export', {
        method: 'POST',
        mode: 'cors',
        body: JSON.stringify(dataBody),
        headers: {
            'Content-Type': 'application/json'
        }
    })

    return response
}

request()
    .then(response => response.blob())
    .then(blob => {
        const url = window.URL.createObjectURL(new Blob([blob]));
        const a = document.createElement('a');
        a.href = url;
        a.download = 'arquivo.xlsx';
        document.body.appendChild(a);
        a.click();
        a.remove();
    })
    .catch(error => console.error(error));
```

## Routes
* `export`: `POST` - Receives an object and return a .xlsx

```json
{
  "filename":"Plan1.xlsx",
  "sheets":[
    {
      "sheetName":"planilha1",
      "headers":[
        {
          "blocks":[
            {
              "message":"REGISTRO DE ABASTECIMENTOS",
              "colBegin":"A1",
              "colEnd":"k1",
             "format":{
              "bold":1,
              "border":1,
              "align":"center",
              "bg_color":"#FF0000"
             }
            }
          ]
        },
        {
        "blocks":[
          {
            "message":"PERIODO: 01/06/2024 Á 06/09/2024",
            "colBegin":"A2",
            "colEnd":"C2",
            "format":{
              "bold":1,
              "border":1,
              "align":"center",
              "bg_color":"#FF0000"
            }
          }
        ]
      },
      {
        
      }
      ],
      "recordHeader":[
        {
          "nameHeader":"Posto",
          "formatHeader":{
            "border":1
          }
        },
        {
          "nameHeader":"Data de abetura",
          "formatHeader":{
            
          }
        },
        {
          "nameHeader":"Equipamento",
          "formatHeader":{
            
          }
        },
        {
          "nameHeader":"Tipo de Consumo",
          "formatHeader":{
            
          }
        },
        {
          "nameHeader":"Contador atual",
          "formatHeader":{
            
          }
        },
        {
          "nameHeader":"Contador anterior",
          "formatHeader":{
            
          }
        },
        {
          "nameHeader":"Combustivel",
          "formatHeader":{
            
          }
        },
        {
          "nameHeader":"Quantidade",
          "formatHeader":{
            
          }
        },
        {
          "nameHeader":"Consumo realizado",
          "formatHeader":{
            
          }
        },
        {
          "nameHeader":"Preço Litro",
          "formatHeader":{
            
          }
        },
        {
          "nameHeader":"VLR Total",
          "formatHeader":{
            
          }
        }
      ],
      "recordsFormat":[
        {
          "font_color":"red"
        },
        {
          "font_color":"yellow"
        },
        {
          "font_color":"blue"
        }
      ],
      "records":[
        [
          "a",
          "b",
          "c"
        ],
        [
          "d",
          "e",
          "f"
        ]
      ]
    }
  ]
}
```

