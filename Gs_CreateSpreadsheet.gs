/**
 * Función que permite agregar los datos basicos sobre el archivo
**/
function addBasicData(projectSheet, formData, projectName) {

  // agregamos los datos de plazos
  projectSheet.getRange("H3:H5").setValues([
    [formData.issue],
    [formData.validity],
    [formData.execution]
  ]);
  
  // agregamos los datos de razón social del solicitante
  projectSheet.getRange("F7:F10").setValues([
    [formData.client],
    [formData.federalRecord],
    [formData.agent],
    [formData.fiscalRecidence]
  ]);
  
  // agregamos el estado
  projectSheet.getRange("B14").setValue(formData.state);
  
  // agregamos el municipio y localidad
  projectSheet.getRange("F13:F14").setValues([
    [formData.municipality],
    [formData.location]
  ]);
  
  // agregamos los datos de georeferencia
  projectSheet.getRange("I13:I14").setValues([
    [formData.latitude],
    [formData.longitude]
  ]);
  
  // agregamos el nombre del proyecto
  projectSheet.getRange("F16").setValue(projectName);
  
  // agregamos la clasificación del sistemas
  projectSheet.getRange("B18").setValue(formData.systemClasification);
  
  // varable que almacena el formato de la superficie
  var surfaceFormatted = "";
  
  // creamos un array de acuerdo al valor de superficie y el valor de separación
  var surfaceArray = String(formData.surface).split(".");
  
  // recorremos el lstado y le damos formato de 2 dígitos
  surfaceArray.map(function (record){
    
    // valor en numerico
    var valueNumber = Number(record);
    
    // retornamos el valor en formato
    return (valueNumber < 10 ? ("0" + valueNumber) : valueNumber);
  });
  
  // avlidamos si existe es igual a cero
  if(surfaceArray.length == 0){
    
    // asignamos un formato en ceros
    surfaceFormatted = "00-00-00";
    
  } else if(surfaceArray.length == 1){
    
    // asignamos un formato en ceros
    surfaceFormatted = (surfaceArray[0] + "-00-00");
    
  } else {
    
    // asignamos un formato en ceros
    surfaceFormatted = (surfaceArray[0] + "-" + surfaceArray[1] + "-00");
  }
  
  // agregamos los datos de Tipo de sistema, forma emisión, caracteristicas, cultivo, superficie
  projectSheet.getRange("F18:J18").setValues([[
    formData.systemType,
    formData.formEmission,
    formData.features,
    formData.cultive,
    surfaceFormatted
  ]]);
  
}

/**
 * Función que permite agregar los datos de Generación de presupuesto
**/
function insertInformationInSheet(projectSheet, orderProjects) {
  
  // variable para manipular las filas
  var orderCounter = 22,
      orderRange = null,
      lineRange = null,
      currentOrder = [],
      currentLine = [],
      currentConcept = [],
      totalCell = [],
      totalOrder = [],
      totalLines = [];
  
  // definimos la variable de inicio del grupo repectivo
  var startIndexGroupLine = 0;
  
  // recorremos cada una de las ordenes
  for(var i = 0; i < orderProjects.length; i++){
    
    // definimos el orden actual
    currentOrder = orderProjects[i];
    
    // creamos la fila de la nueva orden
    projectSheet.insertRowAfter(orderCounter);
    
    // agregamos el nombre de la orden
    projectSheet.getRange("B" + orderCounter).setValue(currentOrder.orderName);
    
    // obtenemos el rango del total en el orden del proyecto
    orderRange = projectSheet.getRange("F" + orderCounter + ":H" + orderCounter);
    
    // agregamos el primer rango
    totalCell.push("H" + orderCounter);
    
    // inicializamos el array de totales del orden actual
    totalOrder = [];
    
    // agregamos los estilos de la fila de orden
    projectSheet.getRange("B" + orderCounter + ":J" + orderCounter).setBackground("#f5feff").setFontColor("#0000ff").setFontWeight("bold").setFontFamily("Arial").setFontSize(10).setBorder(true, false, true, false, false, false, "#ff9900", null);
    
    // agregamos los estilos de la fila de orden
    projectSheet.getRange("F" + orderCounter).setFontWeight("normal").setFontFamily("Montserrat").setFontSize(9).setFontStyle("italic").setHorizontalAlignment("right");
    
    // alineamos el texto de la columna G de un Orden a la derecha
    projectSheet.getRange("G" + orderCounter).setHorizontalAlignment("center");
    
    // recorremos cada una de las líneas de productos
    for(var j = 0; j < currentOrder.linesProducts.length; j++){
      
      // iniciamos el contador del total de filas de la orden actual
      orderCounter ++;
      
      // se establece el punto inicar del rango
      startIndexGroupLine = orderCounter;
      
      // referenciamos la linea actual
      currentLine = currentOrder.linesProducts[j];
      
      // creamos la fila de la nueva orden
      projectSheet.insertRowAfter(orderCounter);
      
      // obtenemos el rango del total en la línea de producto
      lineRange = projectSheet.getRange("H" + orderCounter + ":I" + orderCounter);
      
      // agregamos el nombre de la lídea del producto
      projectSheet.getRange("C" + orderCounter).setValue(currentLine.lineProductName);
      
      // agregamos el rango de donde va a quedar el todal de la linea actual
      totalOrder.push("I" + orderCounter);
      
      // inicializamos el array de totales del lineas de productos
      totalLines = [];
      
      // agregamos los estilos de la fila de linea de producto
      projectSheet.getRange("C" + orderCounter + ":J" + orderCounter).setBackground("#f7f7f7").setFontColor("#285632").setFontWeight("bold").setFontFamily("Montserrat").setFontSize(10).setBorder(true, false, true, false, false, false, "#ff9900", null);
      
      // recorremos cada uno de las especies
      for(var k = 0; k < currentLine.conceptList.length; k++){
        
        // iniciamos el contador del total de filas de la orden actual
        orderCounter ++;
        
        // referenciamos la especie actual
        currentConcept = currentLine.conceptList[k];
        
        // agregamos el rango de donde va a quedar el total de cada especie
        totalLines.push("J" + orderCounter);
        
        // creamos la fila de la nueva orden
        projectSheet.insertRowAfter(orderCounter);
        
        // agregamos el nombre de la lídea del producto
        projectSheet.getRange("F" + orderCounter + ":O" + orderCounter).setValues([[
          currentConcept.conceptData[1],
          currentConcept.conceptData[4],
          (currentConcept.conceptCount ? currentConcept.conceptCount : 0),
          "=N" + orderCounter + "*O"+ orderCounter,
          "=I" + orderCounter + "*H"+ orderCounter,
          "",
          "",
          "",
          (currentConcept.conceptData[5] ? currentConcept.conceptData[5] : 0),
          1
        ]]);
        
        // restablecemos los estilos 
        projectSheet.getRange("B" + orderCounter + ":J" + orderCounter).setWrap(true).setBackground("#fff").setFontColor("#000").setFontWeight("normal").setFontFamily("Arial").setFontSize(10).setBorder((k == 0 ? true : false), false, false, false, false, false, (k == 0 ? "#ff9900" : null), null);
        
        // agregamos el estilo para que se alinie en el cento la columna G de una especie
        projectSheet.getRange("G" + orderCounter).setHorizontalAlignment("center");
      }
      
      // agregamos la información del total para la linea de producto actual
      lineRange.setValues([["M.N. $", "=" + totalLines.toString().replace(/\,/g, "+")]]);
      
      // agrupamos las líneas de producto
      projectSheet.getRange("A" + (startIndexGroupLine + 1) + ":A" + orderCounter).shiftRowGroupDepth(1);
      
      // se define el nuevo grupo y agrupamos por especie
      projectSheet.getRange("A" + startIndexGroupLine + ":A" + orderCounter).shiftRowGroupDepth(1);
      
    }
    
    // agregamos la información del total para la orden del proyecto
    orderRange.setValues([["SUBTOTAL DE ORDEN", "M.N. $", "=" + totalOrder.toString().replace(/\,/g, "+")]]);
    
    // iniciamos el contador del total de filas de la orden actual
    orderCounter ++;
  }
  
  // Agregamos la formula del total
  projectSheet.getRange("J" + (orderCounter + 2)).setValue("=" + totalCell.toString().replace(/\,/g, "+"));
  
  // agregamos la formula del valor en letras
  projectSheet.getRange("C" + (orderCounter + 3)).setValue("=importex(J" + (orderCounter + 2) + ")");
  
}

/**
 * Función que permite registrar los datos en la hoja principal
**/
function registerGeneralDataFromSheet(sheetsData, spreadsheetId, documentId, formData, projectName, rowIndex, userMail, consecutive, infoCatalogSheetId){
  
  // referenciamos la hoja de cálculo general donde quedan todos los datos de las cotizaciones generadas
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  // obtenemos la hoja de "Cotizaciones"
  var sheet = spreadsheet.getSheetByName("Cotizaciones");
  
  // validamos si la hoja que se guardo los datos es "Proyecto"
  if(sheetsData.sheetName == "Proyecto"){
  
    // validamos si existe la hoja
    if(sheet){
      
      // obtenemos la ultima fila
      var lastRow = sheet.getLastRow();
      
      // obtenemos el nombre del usuario
      var userName = getUserName(userMail);
      
      // se valida si "rowIndex" es igual a -1
      if(rowIndex == -1){
        
        // se define el array a ingresar cuando una solicitud es nueva
        var arrayData = [
          consecutive,
          userMail,
          formData.client,
          formData.issue,
          userName,        
          (formData.state + ", " + formData.municipality + ", " + formData.location),
          formData.systemType,
          formData.total,
          documentId,
          "Activa",
          JSON.stringify(formData),
          sheetsData.sheetName,
          "NO",
          generateRandom(1.007, 1.02, 3),
          "",
          infoCatalogSheetId
        ];
        
        // agregamos los nuevos valores
        sheet.getRange((lastRow + 1) , 1, 1, arrayData.length).setNumberFormat("@STRING@").setValues([arrayData]);
        
      } else {
       
        // se define el array a ingresar cuando una solicitud es nueva
        var arrayData = [
          (formData.state + ", " + formData.municipality + ", " + formData.location),
          formData.systemType,
          formData.total,
          documentId,
          "Activa",
          JSON.stringify(formData),
          sheetsData.sheetName
        ];
        
        // agregamos los nuevos valores de nombre del cliente y fecha de emisisón
        sheet.getRange(rowIndex, 3, 1, 2).setNumberFormat("@STRING@").setValues([[formData.client, formData.issue]]);
        
        // agregamos los nuevos valores sin el consecutivo y el correo del creador
        sheet.getRange(rowIndex, 6, 1, arrayData.length).setNumberFormat("@STRING@").setValues([arrayData]);
        
      }
      
    }
    
  } else {
    
    // se valida si existe la propiedad "rowIndex"
    if(sheetsData.rowIndex && sheetsData.rowIndex != -1){
    
      // obtenemos la hoja de "Administrador de versiones"
      var sheetManage = spreadsheet.getSheetByName("Administrador de versiones");
      
      // actualizamos la información
      sheetManage.getRange("C" + sheetsData.rowIndex).setValue(JSON.stringify(formData));
      
      // validamos si el rowIndex del proyecto es diferente de -1
      if(sheet && rowIndex != -1){
      
        // agregamos los nuevos valores sin el consecutivo y el correo del creador
        sheet.getRange(rowIndex, 12).setNumberFormat("@STRING@").setValue(sheetsData.sheetName);
      }
    }
    
  }
  
}

/**
 * Función que permite obtener las lineas por defecto de una orden
**/
function getDefaultLinesByOrder(spreadsheet, orderName, classification, systemType, formEmission, features, range){
  
  // referenciamos la hoja de "Líneas por defecto"
  var sheetDefaultLines = spreadsheet.getSheetByName("Líneas por defecto");
  
  // validamos si existe la hoja
  if(sheetDefaultLines){
    
    // obtenemos la ultima fila
    var lastRow = sheetDefaultLines.getLastRow();
    
    // se valida que exista mas de 2 filas
    if(lastRow > 2){
      
      // obtenemos los datos de las líneas de producto
      var lineProducts = sheetDefaultLines.getRange(range).getDisplayValues();
      
      //obtenemos los datos
      var arrayData = sheetDefaultLines.getRange(3, 1, lastRow - 2, 4).getValues();
      
      // variable para determinar en que fila se encuentra la conbinación seleccionada
      var rowIndex = -1;
      
      // recorremos cada uno de los items
      for(var i = 0; i < arrayData.length; i++){
        
        // validamos si tiene los mismos valores
        if(arrayData[i][0] == classification && arrayData[i][1] == systemType && arrayData[i][2] == formEmission && arrayData[i][3] == features){
          
          // actualizamos el valor de la fila
          rowIndex = i + 3;
          
          // salimos del ciclo
          break;
          
        }
      }
      
      // validamos si la fila es diferente de -1
      if(rowIndex != -1){
        
        // obtenemos los datos donde estan los chulos segun la linea elegida
        var checksArray = sheetDefaultLines.getRange(String(range).replace(/\d+/g, rowIndex)).getDisplayValues();
        
        // variable para almacenar las lineas de productos que se deben seleccionar
        var linesArray = [];
        
        // validamos si existe datos
        if(checksArray && checksArray.length > 0){
          
          // recorremos los datos para validar si esta chequeado
          checksArray[0].forEach(function(rowData, index){
            
            // se valida que tenga chequeado la linea recorrida
            if(String(rowData[0]).trim() == "✓"){
              
              // agregamos el nombre de la línea
              linesArray.push(lineProducts[0][index]);
            }
            
          });
          
          // retornamos los datos 
          return linesArray;
          
        }
      }
    }
  }
    
  
  // retornamos por defecto un array vacio
  return [];
  
}