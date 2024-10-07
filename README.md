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
* `exportSuppliersRecords`: `POST` - Receives an object and return a .xlsx formatting materials by suppliers
```json
{
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
              "bg_color":"#FF0000"
            }
          }
        ]
      },
    ],
    "recordHeader":[
      {
        "nameHeader":"Posto",
        "formatHeader":{
          "border":1
        },
      },
      {
        "nameHeader":"Data de abetura",
        "formatHeader":{},
      },
      {
        "nameHeader":"Equipamento",
        "formatHeader":{},
      },
    ],
      "recordsFormat":[
          {},
          {"bg_color": "#FFFF00"},
          {"bg_color": "#FFFF00"},
      ],
      "records":[
      ["Posto 1","01/06/2024","Equipamento A"]
   ]   
  }
}
```

