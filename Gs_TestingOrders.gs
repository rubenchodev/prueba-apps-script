function loadProductByHistoryOrder() {
  
  // obtenemos los datos globales
  var resources = getResources();

  var spreadsheet = SpreadsheetApp.openById(resources.ordersHistorySheetId);
  //throw spreadsheet.getUrl()
  
  var sheet = spreadsheet.getSheetByName("Archivos");
  if(sheet){
    var lastrow = sheet.getLastRow();
    if(lastrow > 1){
      var arrayData = sheet.getRange(2, 2, lastrow - 1, 2).getValues();
      for(var i = 0; i < arrayData.length; i++){
        if(!arrayData[i][1]){
          console.log(arrayData[i][0]);
          var productData = getProductDataByDoc(arrayData[i][0]);

          sheet.getRange((i + 2), 3, 1, 4).setValues([[
            "X",
            productData.importe,
            productData.products,
            productData.columns
          ]]);
          SpreadsheetApp.flush();
        }
      }
    }
  }
}

function getProductDataByDoc(fileId){
  //var fileId = "1Q6gG3OFOua5xpoivAY5jujR9w2zca9cS1T_PEZjTbpw";
  var doc = DocumentApp.openById(fileId);
  var body = doc.getBody();
  var fullText_ = body.getText();
  var resultObject_ = {
    importe: 0,
    products: [],
    columns: 0
  };

  // Obtenemos la información de las filas
  obtenerFilasTablaPorTextoComoJSON("C. Unitario", body, resultObject_);

  var patron = /(\$\d{1,3}(,\d{3})*(\.\d{1,})?)/g;
  var resultado = fullText_.match(patron);

  if (resultado && resultado.length > 0) {
    resultObject_.importe = Number(resultado[0].toString().replace(/\$|\,/g, "")) || 0;
  }
  console.log((resultObject_));
  return resultObject_;
}

function obtenerFilasTablaPorTextoComoJSON(textoBusqueda, cuerpo, resultObject_) {

  // Buscar el texto en el documento
  var elementos = cuerpo.findText(textoBusqueda);
  var tabla = null;

  if (elementos) {
    // Verificar si el texto se encuentra dentro de una tabla
    var elemento = elementos.getElement();
    var parent = elemento;
    while(parent && parent.getType() != DocumentApp.ElementType.TABLE){
      parent = parent.getParent();
    }
    if(parent){
      tabla = parent;
      console.log(parent.getType().toString())
    }
  }
  
  var data = [];

  if (tabla) {
    var filas = tabla.getNumRows();
    if(filas > 1){

      for (var i = 1; i < filas; i++) {
        var filaA_ = tabla.getRow(i);
        var celdas = filaA_.getNumCells();
        resultObject_.columns = celdas;
        if(celdas == 7){
          var filaData = {
            quantity: filaA_.getCell(1).getText(),
            cost: filaA_.getCell(6).getText(),
            name: filaA_.getCell(3).getText(),
            currencyType: filaA_.getCell(4).getText()
          };
          data.push(filaData);
        } else if(celdas == 6){
          var filaData = {
            quantity: filaA_.getCell(0).getText(),
            cost: filaA_.getCell(5).getText(),
            name: filaA_.getCell(2).getText(),
            currencyType: (filaA_.getCell(3).getText() || "MX")
          };
          data.push(filaData);
        }
      }
    }
  }

  resultObject_.products = (data.length > 0 ? JSON.stringify(data) : "");

}

