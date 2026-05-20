/**
 * Función que permite crear la pestaña de federal si no existe y si existe se procede a eliminar y luego dupplicar a partir de
**/
function createFederal(objectData, chosenFactor){

  // referenciamos el libro
  var spreadsheet = SpreadsheetApp.openById(objectData.documentId);
  
  // obtenemos la hoja por el nombre de "Proyecto"
  var projectSheet = spreadsheet.getSheetByName("Proyecto");
  
  // se valida si existe la pestaña de "Proyecto"
  if(projectSheet){
  
    // obtenemos la hoja por el nombre
    var federalSheet = spreadsheet.getSheetByName("Federal");
    
    // validamos si existe
    if(federalSheet){
      
      // se elimina la pestaña actual
      spreadsheet.deleteSheet(federalSheet);
      
    }
    
    // creamos el duplicado de la hoja
    federalSheet = projectSheet.copyTo(spreadsheet);
    
    // cambiamos el nombre de la pestaña a "Federal"
    federalSheet.setName("Federal");
    
    // validamos si la fila temporal es igual a 22
    if(objectData.lastRow == 22){
      
      // eliminamos solo la fila 22
      federalSheet.deleteRow(22);
      
    } else {
    
      // Removemos las filas ya no utilizadas
      federalSheet.deleteRows(22, objectData.lastRow - 22);
    }
    
    // Valor del costo x hectaria
    var costeH = 0, dearSupport_ = 0;
 
    // convertimos en entero el valor de superficie
    var surfaceValue = Number(objectData.requestData.surface) || 0;

    // Definimos el valor del factor
    var factorValue_ = 30000;

    // Se valida si el tipo de proyecto es "Rehabilitación"
    if(objectData.projectType == "Rehabilitación"){
      factorValue_ = 15000;
    }

    // obtenemos el costo por hectaria
    if(surfaceValue && surfaceValue != 0){

      // se valdia si la superfici es mayor a 5
      if(surfaceValue > 5){
        // Calculamos el estimado de apoyo
        surfaceValue = 5;
      }
      
      // se alla el costo por hectaria
      costeH = objectData.requestData.total / surfaceValue;      
    }

    // obtenemos el multiplacador
    var multiplierValue = roundDecimal((factorValue_ * (chosenFactor + 1)) / costeH, 2);
    
    // agragamos la información de Presupuesto
    insertInfoFederalData(federalSheet, objectData.requestData.orderProjects, multiplierValue);
    
  } else {
  
    // retornamos el error
    throw new Error("Lo sentimos en la cotización actual no existe la hoja <b>Proyecto</b>");
  }
  
}

/**
 * Función que permite generar un valor randomico desde un rango especifico
**/
function generateRandom(min, max, decimalPlaces) {

  // se valida si no existe cantidad de decimales para definir por defecto 2
  decimalPlaces = decimalPlaces || 2;
  
  // se genera el valor randomico
  var randomValue = Math.random() * (max - min) + min;
  
  // se obtiene el valor de acuerdo a la cantidad de decimales
  var power = Math.pow(10, decimalPlaces);
  
  // se retorna la inofrmación con los decimales adecuados
  return Math.floor(randomValue * power) / power;
}

/**
 * Función que permite agregar los datos de Generación de presupuesto de Federal
**/
function insertInfoFederalData(federalSheet, orderProjects, multiplierValue) {
  
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
    federalSheet.insertRowAfter(orderCounter);
    
    // agregamos el nombre de la orden
    federalSheet.getRange("B" + orderCounter).setValue(currentOrder.orderName);
    
    // obtenemos el rango del total en el orden del proyecto
    orderRange = federalSheet.getRange("F" + orderCounter + ":H" + orderCounter);
    
    // agregamos el primer rango
    totalCell.push("H" + orderCounter);
    
    // inicializamos el array de totales del orden actual
    totalOrder = [];
    
    // agregamos los estilos de la fila de orden
    federalSheet.getRange("B" + orderCounter + ":J" + orderCounter).setBackground("#f5feff").setFontColor("#0000ff").setFontWeight("bold").setFontFamily("Arial").setFontSize(10).setBorder(true, false, true, false, false, false, "#ff9900", null);
    
    // agregamos los estilos de la fila de orden
    federalSheet.getRange("F" + orderCounter).setFontWeight("normal").setFontFamily("Montserrat").setFontSize(9).setFontStyle("italic").setHorizontalAlignment("right");
    
    // alineamos el texto de la columna G de un Orden a la derecha
    federalSheet.getRange("G" + orderCounter).setHorizontalAlignment("center");
    
    // recorremos cada una de las líneas de productos
    for(var j = 0; j < currentOrder.linesProducts.length; j++){
      
      // iniciamos el contador del total de filas de la orden actual
      orderCounter ++;
      
      // se establece el punto inicar del rango
      startIndexGroupLine = orderCounter;
      
      // referenciamos la linea actual
      currentLine = currentOrder.linesProducts[j];
      
      // creamos la fila de la nueva orden
      federalSheet.insertRowAfter(orderCounter);
      
      // obtenemos el rango del total en la línea de producto
      lineRange = federalSheet.getRange("H" + orderCounter + ":I" + orderCounter);
      
      // agregamos el nombre de la lídea del producto
      federalSheet.getRange("C" + orderCounter).setValue(currentLine.lineProductName);
      
      // agregamos el rango de donde va a quedar el todal de la linea actual
      totalOrder.push("I" + orderCounter);
      
      // inicializamos el array de totales del lineas de productos
      totalLines = [];
      
      // agregamos los estilos de la fila de linea de producto
      federalSheet.getRange("C" + orderCounter + ":J" + orderCounter).setBackground("#f7f7f7").setFontColor("#285632").setFontWeight("bold").setFontFamily("Montserrat").setFontSize(10).setBorder(true, false, true, false, false, false, "#ff9900", null);
      
      // recorremos cada uno de las especies
      for(var k = 0; k < currentLine.conceptList.length; k++){
        
        // iniciamos el contador del total de filas de la orden actual
        orderCounter ++;
        
        // referenciamos la especie actual
        currentConcept = currentLine.conceptList[k];
        
        // agregamos el rango de donde va a quedar el total de cada especie
        totalLines.push("J" + orderCounter);
        
        // creamos la fila de la nueva orden
        federalSheet.insertRowAfter(orderCounter);
        
        // agregamos el nombre de la lídea del producto
        federalSheet.getRange("F" + orderCounter + ":O" + orderCounter).setValues([[
          currentConcept.conceptData[1],
          currentConcept.conceptData[4],
          (currentConcept.conceptCount ? currentConcept.conceptCount : 0),
          "=TRUNC(N" + orderCounter + "*O" + orderCounter + ", 2)",
          "=I" + orderCounter + "*H"+ orderCounter,
          "",
          "",
          "",
          (currentConcept.conceptData[5] ? currentConcept.conceptData[5] : 0),
          multiplierValue
        ]]);
        
        // restablecemos los estilos 
        federalSheet.getRange("B" + orderCounter + ":J" + orderCounter).setWrap(true).setBackground("#fff").setFontColor("#000").setFontWeight("normal").setFontFamily("Arial").setFontSize(10).setBorder((k == 0 ? true : false), false, false, false, false, false, (k == 0 ? "#ff9900" : null), null);
        
        // agregamos el estilo para que se alinie en el cento la columna G de una especie
        federalSheet.getRange("G" + orderCounter).setHorizontalAlignment("center");
      }
      
      // agregamos la información del total para la linea de producto actual
      lineRange.setValues([["M.N. $", "=(" + totalLines.toString().replace(/\,/g, "+") + ")"]]);
      
      // agrupamos las líneas de producto
      federalSheet.getRange("A" + (startIndexGroupLine + 1) + ":A" + orderCounter).shiftRowGroupDepth(1);
      
      // se define el nuevo grupo y agrupamos por especie
      federalSheet.getRange("A" + startIndexGroupLine + ":A" + orderCounter).shiftRowGroupDepth(1);
      
    }
    
    // agregamos la información del total para la orden del proyecto
    orderRange.setValues([["SUBTOTAL DE ORDEN", "M.N. $", "=(" + totalOrder.toString().replace(/\,/g, '+') + ")"]]);
    
    // iniciamos el contador del total de filas de la orden actual
    orderCounter ++;
  }
  
  // Agregamos la formula del total
  federalSheet.getRange("J" + (orderCounter + 2)).setValue("=(" + totalCell.toString().replace(/\,/g, "+") + ")");
  
  // agregamos la formula del valor en letras
  federalSheet.getRange("C" + (orderCounter + 3)).setValue("=importex(J" + (orderCounter + 2) + ")");
  
}
