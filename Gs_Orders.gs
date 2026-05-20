/**
 * Función que permite generar las ordenes de compra respectivas
**/
function generateOrders(ordersData) {
 
  // obtenemos los datos globales
  var resources = getResources(),
      allManufacturerData = [];
  
  // se referencia el libro de administración para adquisiciones
  var spreadsheetAcq = SpreadsheetApp.openById(resources.manufacturerSheetId);
  
  // refernciamos la plantilla de ordenes de compra
  var template = DriveApp.getFileById(resources.ordersTemplateId);
  
  // referenciamos la carpeta donde se almacenaran las nuevas ordenes
  var rootFolder = DriveApp.getFolderById(resources.ordersFolderId);
  
  // obtenemos la fecha actual
  var currentDate = new Date();
  
  // obtenemos la fecha en formuaro "12 de Junio de 2019"
  var dateFormatted = formatDate(currentDate);
  
  // obtenemos el correo del usuario que realiza la creación de la orden de compra
  var userMail = getUserMail();
  
  // obtenemos los datos como nombre y foto
  var userData = getAccountUserPanel(userMail, "");
  
  // definimos la cadena inicial del # de la orden
  var prefix = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "ddMMyy");
  
  // fecha de actualización
  var updateDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");
  
  // obtenemos las propiedades del script
  var scriptProperties = PropertiesService.getScriptProperties();
  
  // obtenemos el prefijo guardado
  var prefixSeved = scriptProperties.getProperty("PrefixOrder");
  
  // se valida si el texto inicial del consecutivo es diferente al actual
  if(prefixSeved != prefix){
    
    // actualizamos el prefix
    scriptProperties.setProperty("PrefixOrder", prefix);
    
    // inicializamos el consecutivo
    scriptProperties.setProperty("ConsecutiveOrder", 1);
    
  }
  
  // variable que guardar los registros nuevos
  var newRecords = [];
  
  // recorremos cada unos de los fabricantes
  for(var key in ordersData){
    
    console.log("hoja")
    console.log(JSON.stringify(spreadsheetAcq.getSheets()))
    console.log(spreadsheetAcq.getSheets())
    console.log(spreadsheetAcq.getUrl())
    
    console.log("manufacturerId")
    console.log(ordersData[key].id)
    // obtenemos todos los datos del fabricante
    allManufacturerData = getManufacturerById(spreadsheetAcq, ordersData[key].id);
    console.log("datos de los fabricantes")
    console.log(allManufacturerData)
    
    // se valida si se tienen datos del fabricante
    if(allManufacturerData){
      
      // obtenemos el # de la orden
      var orderCode = getConsecutiveByOrder(scriptProperties, prefix);
      
      // definimos el nombre del archivo
      var fileName = key + " - orden de compra #" + orderCode;
      
      // generamos una copia del archivo
      var fileCopy = template.makeCopy(fileName, rootFolder);
      
      // agregamos permisos de edición a los usuarios con los vinculos
      fileCopy.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
      
      // abrimos el documento como Google Doc
      var googleDoc = DocumentApp.openById(fileCopy.getId());

      // Variable para obtener los datos de una orden
      var orderObject_ = {
        importe: 0,
        products: []
      };
      
      // insertamos la información a la orden de compra
      insertDataFromOrderFile(googleDoc, allManufacturerData, ordersData[key], dateFormatted, orderCode, userData, orderObject_);
      
      // agregamos el nuevo registro
      newRecords.push([
        orderCode,
        key,
        allManufacturerData.email, 
        updateDate,  
        userMail,
        googleDoc.getUrl(),
        "Generada",
        orderObject_.importe,
        JSON.stringify(orderObject_.products)
      ]);
    }
  }
  
  // validamos si existe la hoja y exsten datos x guardar
  if(newRecords.length > 0){

    // se referencia el libro de historial de órdenes de compra
    var spreadsheetHistory = SpreadsheetApp.openById(resources.ordersHistorySheetId);
    
    // referenciamos la hoja de historial
    var sheetHistory = spreadsheetHistory.getSheetByName("Órdenes");
    
    // validamos si existe la hoja y exsten datos x guardar
    if(sheetHistory){
      
      // agregamos los nuevos datos
      sheetHistory.getRange(sheetHistory.getLastRow() + 1, 1, newRecords.length, newRecords[0].length).setNumberFormat("@STRING@").setValues(newRecords);
      
    }
  }
  //throw "OK"
  // retornamos el historial actualizado
  return getOrdersHistory(spreadsheetHistory);
}

/**
 * Función que permite insertar los datos en la orden de compra
**/
function insertDataFromOrderFile(document, allManufacturerData, ordersData, currentDate, orderCode, userData, orderObject_){
  
  // referenciamos el cuerpo del documento
  var body = document.getBody();
  
  // Reemplazamos la fecha
  body.replaceText("<<Fecha>>", currentDate);
  
  // Reemplazamos el # de la orden
  body.replaceText("<<#ORDEN>>", orderCode);
  
  // Reemplazamos el # de la orden
  body.replaceText("<<#Cotización>>", orderCode);
  
  // Reemplazamos el fabricante
  body.replaceText("<<Fabricante>>", allManufacturerData.manufacturerName);
  
  // Reemplazamos la Razón Social
  body.replaceText("<<Razón Social>>", allManufacturerData.socialReason);
  
  // Reemplazamos el RFC Fabricante
  body.replaceText("<<RFC Fabricante>>", allManufacturerData.rfc);
  
  // Reemplazamos el Nombre de Contacto
  body.replaceText("<<Nombre de Contacto>>", allManufacturerData.contact);
  
  // Reemplazamos Teléfono Oficina
  body.replaceText("<<Teléfono Oficina>>", allManufacturerData.officePhone);
  
  // Reemplazamos el nombre del usuario de adquisiciones
  body.replaceText("<<Usuario Adquisiciones>>", userData.fullName);
  
  // Reemplazamos el nombre del usuario de adquisiciones
  body.replaceText("<<AUTOR>>", userData.fullName);
  
  // refernciamos la tabla de las especies
  var tableDoc = body.getTables()[2],
      row = null, subtotal = 0, importe = 0, ivaArray = [];
  
  // definimos un objeto para realizar la somatoria de iva y del precio
  var resultsObj = {
    "EUR": {subtotal: 0, iva: 0},
    "USD": {subtotal: 0, iva: 0},
    "MX": {subtotal: 0, iva: 0},
  };
  
  // Iva actual
  var tempIva = 0;
  
  // variable para conocer que tipos de cambios se debe trabajar
  var currencyTypeArray = [], exchangeRateArray = [];

  // recorremos cada uno de los concepros
  for(var i = 0; i < ordersData.concepts.length; i++){
 
    // validamos si es el primer especie
    if(i == 0){
      
      // refernciamos la segunda fila de la tabla
      row = tableDoc.getRow(i + 1);
      
    } else {
      
      // sacamos la copia de la fila anterior
      row = tableDoc.appendTableRow(tableDoc.getRow(i).copy());
    }
    
    // obtenemos el iva del producto o concepto
    tempIva = (ordersData.concepts[i].iva / 100);
    tempIva = (Number(ordersData.concepts[i].plist) || 0) * tempIva;
    
    // obtenemos el importe
    //importe = (Number(ordersData.concepts[i].price) * ordersData.concepts[i].quantity) || 0;
    importe = ((Number(ordersData.concepts[i].plist) || 0) * ordersData.concepts[i].quantity) || 0;
    
    // vamos calculando el total
    resultsObj[ordersData.concepts[i].currencyType].subtotal += importe;
    
    // vamos calculando el iva
    //resultsObj[ordersData.concepts[i].currencyType].iva += (importe * (ordersData.concepts[i].iva / 100));
    resultsObj[ordersData.concepts[i].currencyType].iva += tempIva * ordersData.concepts[i].quantity;
    
    // validamos si el iva NO existe
    if(ivaArray.indexOf(ordersData.concepts[i].iva + "%") == -1){
      
      // agregamos el iva en el listado
      ivaArray.push(ordersData.concepts[i].iva + "%");
    }
    
     // validamos si el tipo de moneda no existe
    if(currencyTypeArray.indexOf(ordersData.concepts[i].currencyType) == -1){
      
      // agregamos el tipo de moneda en el listado
      currencyTypeArray.push(ordersData.concepts[i].currencyType);
      
      // agregamos el tipo de cambio de la moneda actual
      exchangeRateArray.push(ordersData.currencyObj[ordersData.concepts[i].currencyType] || "");
      
    }

    // Agregamos el nuevo producto
    orderObject_.products.push({
      //cost: roundDecimal(ordersData.concepts[i].plist, 2),
      cost: roundDecimal(importe, 2),
      quantity: roundDecimal(ordersData.concepts[i].quantity, 2),
      name: ordersData.concepts[i].name,
      currencyType: ordersData.concepts[i].currencyType
    });
    
    // agregamos los datos de cada columna
    row.getCell(0).setText(ordersData.concepts[i].code);
    row.getCell(1).setText(ordersData.concepts[i].quantity);
    row.getCell(2).setText(ordersData.concepts[i].unit);
    row.getCell(3).setText(ordersData.concepts[i].productName);
    row.getCell(4).setText(ordersData.concepts[i].currencyType);
//    row.getCell(5).setText(formatValue(roundDecimal(ordersData.concepts[i].price, 2)));
    row.getCell(5).setText(formatValue(roundDecimal(ordersData.concepts[i].plist, 2)));
//    row.getCell(5).setText(formatValue(roundDecimal(tempUnitCost, 2)));
    row.getCell(6).setText(formatValue(roundDecimal(importe, 2)));
    
  }
  
  // agregamos el yipo de moneda MX
  currencyTypeArray.push("MX");
  
  // Reemplazamos el tipo de moneda
  body.replaceText("<<Moneda>>", String(currencyTypeArray).replace(/,/g, " || "));
  
  // Reemplazamos el Tipo de Cambio
  body.replaceText("<<Tipo de Cambio>>", String(exchangeRateArray).replace(/,/g, " || "));
  
  // validamos si NO existe datos de sutotal de dolares
  if(!resultsObj["USD"].subtotal){
    
    // eliminamos la columna de dolar
    deleteColumnByTag("<<subTotalUSD>>", body);
    deleteColumnByTag("<<ivaUSD>>", body);
    deleteColumnByTag("<<totalUSD>>", body);
    deleteColumnByTag("<<USD>>", body);
  }
  
  // validamos si NO existe datos de sutotal de euros
  if(!resultsObj["EUR"].subtotal){
    
    // eliminamos la columna de dolar
    deleteColumnByTag("<<subTotalEUR>>", body);
    deleteColumnByTag("<<ivaEUR>>", body);
    deleteColumnByTag("<<totalEUR>>", body);
    deleteColumnByTag("<<EUR>>", body);
  }
  
  // validamos si NO existe datos de sutotal de pesos mexicanos
  if(!resultsObj["MX"].subtotal){
    
    // eliminamos la columna de dolar
    deleteColumnByTag("<<subTotalMX>>", body);
    deleteColumnByTag("<<ivaMX>>", body);
    deleteColumnByTag("<<totalMX>>", body);
    deleteColumnByTag("<<MXN>>", body);
  }
  
  // reemplazamos los valos de la tabla de total
  body.replaceText("<<USD>>", "USD");
  body.replaceText("<<EUR>>", "EUR");
  body.replaceText("<<MXN>>", "MXN");
  
  // Reemplazamos el subtotal de los tres tipos de moneda
  body.replaceText("<<subTotalUSD>>", getFormattedValue(resultsObj["USD"].subtotal));
  body.replaceText("<<subTotalEUR>>", getFormattedValue(resultsObj["EUR"].subtotal));
  body.replaceText("<<subTotalMX>>", getFormattedValue(resultsObj["MX"].subtotal));
  
  // Reemplazamos el iva
  body.replaceText("<<iva>>", ivaArray.toString());
  body.replaceText("<<ivaUSD>>", getFormattedValue(resultsObj["USD"].iva));
  body.replaceText("<<ivaEUR>>", getFormattedValue(resultsObj["EUR"].iva));
  body.replaceText("<<ivaMX>>", getFormattedValue(resultsObj["MX"].iva));
  
  // obtenemos los totales en las diferentes monedas
  var totalMx = resultsObj["MX"].subtotal + resultsObj["MX"].iva;
  var totalUsd = resultsObj["USD"].subtotal + resultsObj["USD"].iva;
  var totalEur = resultsObj["EUR"].subtotal + resultsObj["EUR"].iva;
  
  // Reemplazamos el total
  body.replaceText("<<totalUSD>>", getFormattedValue(totalUsd));
  body.replaceText("<<totalEUR>>", getFormattedValue(totalEur));
  body.replaceText("<<totalMX>>", getFormattedValue(totalMx));
  
  // calculamos el total en pesos mexicamos
  var totalEndMX = totalMx;
  
  // validamos si existe un total en dolares
  if(totalUsd){
    
    // convertimos los dolares en pesos mexicanos y los sumamos al total final
    totalEndMX += totalUsd * (Number(ordersData.currencyObj["USD"]) || 0);
  }
  
  // validamos si existe un total en euros
  if(totalEur){
    
    // convertimos los euros en pesos mexicanos y los sumamos al total final
    totalEndMX += totalEur * (Number(ordersData.currencyObj["EUR"]) || 0);
  }

  // Agregamos el valor en pesos mexinos
  orderObject_.importe = totalEndMX;
  
  // agregamos el total en pesos mexicanos
  body.replaceText("<<TOTAL MXN >>", getFormattedValue(totalEndMX));
  
  // agregamos el total en pesos mexicanos en letras
  body.replaceText("<<Importex>>", CifrasEnLetras.convertirNumeroEnLetras(roundDecimal( totalEndMX , 2)));
  
  // referenciamos la cabecera del documento
  var head = document.getHeader();
  
  // Reemplazamos la Razón Social
  head.replaceText("<<Razón Social>>", allManufacturerData.socialReason);
  
  // Reemplazamos el RFC Fabricante
  head.replaceText("<<RFC Fabricante>>", allManufacturerData.rfc);
  
  // Reemplazamos el Nombre de Contacto
  head.replaceText("<<Nombre de Contacto>>", allManufacturerData.contact);
  
  // Reemplazamos el # de la orden
  head.replaceText("<<#Orden>>", orderCode);
  
  // Guardamos los datos del documento
  document.saveAndClose();
}

/**
 * Función que permite eliminar una fila de acuerdo a un tag enviado por parametro
**/
function deleteColumnByTag(tagName, body){
  
  // se realiza la búsqueda 
  var foundTag = body.findText(tagName);
  
  // validamos si encontro algun elemento con ese texto
  if (foundTag != null) {
    
    // variable para referenciar la fila
    var tableCell = null,
        tagElement = foundTag.getElement();
    
    // referenciamos el elemento a validar
    var elemReference = DocumentApp.ElementType.TABLE_CELL;
    
    // se valida si el elemento es un Table_Row
    if(tagElement.getType() != elemReference){
      
      // realizamos un ciclo hasta encontralo
      while(tagElement.getParent() && !tableCell){
       
        // obtenemos el padre del elemento
        tagElement = tagElement.getParent();
        
        // se valida si el elemento es un Table_Row
        if(tagElement.getType() == elemReference){
          
          // cambiamos el estado
          tableCell = tagElement;
          
        }
      }
    }
    
    // validamos si existe un elemento
    if(tableCell){
      
      // eliminamos la respectiva fila
      tableCell.removeFromParent();
    }
  
  }
  
}

/**
 * Función que permite dar formato a un valor de precio
**/
function getFormattedValue(result, decimals, defaultValue){
  
  // si no tiene decimales
  decimals = decimals || 2;
  
  // validamos si el valor es cero o no tiene nada
  if(!result){
    
    // retornamos el valor por defecto o vacio
    return (defaultValue ? defaultValue : "");
    
  }
  
  // retornamos el valor con el formato correspondente
  return formatValue(roundDecimal((result) , decimals));
  
}

/**
 * Función que permite obtener le consecutivo para una orden
**/
function getConsecutiveByOrder(scriptProperties, prefix){
   
   //Obtener un bloqueo de script, porque estamos a punto de modificar un recurso compartido
   var lockService = LockService.getScriptLock();
   
   // Espeamos 30 segundos para que finalicen los procesos anteriores
   lockService.waitLock(30000);
   
   // Referenciamos el consecutivo actual
   var consecutive = Number(scriptProperties.getProperty("ConsecutiveOrder"));
   
   // Damos formato al consecutivo
   var formattedConsecutive = Utilities.formatString('%03d', consecutive);
   
   // Incrementamos en 1 el valor de la propiedad
   scriptProperties.setProperty("ConsecutiveOrder", consecutive + 1);
   
   // Liberamos el bloqueo del servicio
   lockService.releaseLock();
   
   // Retornamos el consecutivo generado
   return "OC" + prefix + formattedConsecutive;
}

/**
 * Función que permite obtener los datos de fabricantes segun su Id
**/
function getManufacturerById(spreadsheet, manufacturerId){
  
  // referenciamos la hoja de fabricantes
  var sheet = spreadsheet.getSheetByName("Fabricantes");
  
  // se valida si existe la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // se valida si existe mas de una fila
    if(lastRow > 1){
      
      // obtenemos los datos
      var arrayData = sheet.getRange(2, 1, lastRow - 1, 24).getDisplayValues();
      
      // variable que determina el listado
      var manufacturerList = [];
      
      // recorremos el listado de fabricantes
      for(var i = 0; i < arrayData.length; i++){
        
        // validamos si el id del fabricante es igual al que se recorre
        if(manufacturerId == arrayData[i][1]){
        
          // retornamos los datos del fabricante
          return {
            id: arrayData[i][1],
            manufacturerName: arrayData[i][4],
            socialReason: arrayData[i][10],
            rfc: arrayData[i][11],
            familiesOffered: arrayData[i][12],
            taxResidence: arrayData[i][13],
            domicileWinery1: arrayData[i][14],
            domicileWinery2: arrayData[i][15],
            domicileWinery3: arrayData[i][16],
            contact: arrayData[i][17],
            officePhone: arrayData[i][18],
            mobilePhone: arrayData[i][19],
            email: arrayData[i][2],
            otherEmail: arrayData[i][3],
            webUrl: arrayData[i][20],
            lastUpdate: arrayData[i][21],
            state: arrayData[i][22],
            code: arrayData[i][23],
            rowIndex: (i + 2)
          };
        }
      }
      
    }
  }
  
  // retornamos un valor vacio
  return null;
}
