/*
* Función encargada de generar el formato de la fecha para el certificado laboral
*/
function formatDateWorkCertificate(date){
  
  // Array con los meses en español
  var months = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  
  // Separamos la fecha en dia, mes, año
  var today = date.getDate();
  var month = date.getMonth();
  var year = date.getFullYear();
  
  return ((today<10?("0"+today):today)+" de "+months[month]+" del "+year);  
}

/*
* Función que permite crear la hoja de calculo con su respectivo catalogo
*/
function createCatalog(orderProjects, projectData, headerDataCatalog) {
   
   // Referenciamos los recursos globales
   var resources = getResources();
  
  // referenciamos la carpeta donde se desea almacenar el catalogo
  var  mainFolder = DriveApp.getFolderById(resources.catalogFolderId);
  
  // Fecha actual
  var newDate = new Date();
  
  // Obtenemos la fecha actual 
  var currentDate = Utilities.formatDate(newDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  
  // Nombre del archivo del catálogo 
  var fileName = "Catálogo del proyecto: " + projectData.id + " - " +  currentDate;
  
  // se crea una copia de la plantilla deel catalogo
  var fileCopy = DriveApp.getFileById(resources.catalogTemplateId).makeCopy(fileName, mainFolder);
  
  // Agregamos los permisos para que cuaquier usuarios dentro de Hidroexpert pueda editar
  fileCopy.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
  
  // se referencia el archivo
  var spreadsheet = SpreadsheetApp.open(fileCopy);
  
  // Validamos que existan especies y la hoja se tenga acceso
  if(orderProjects && spreadsheet){
    
    // Referenciamos a la hoja de catálogo
    var projectCatalogSheet = spreadsheet.getSheetByName("Catálogo");
    
    // Datos del catalogo
    var documentData= {
          orderCounter: 17, 
          orders:[]
        }
        
     // Variable para el total de la orden   
     var totalCell = [];
    
    // Agregamos los localidad, municipio , estado 
    projectCatalogSheet.getRange("C5").setValue(headerDataCatalog.location_catalog);
    
    // Agregamos la superficie
    projectCatalogSheet.getRange("I3").setValue(headerDataCatalog.surface_catalog);
    
    // Agregamos el nombre de la obra civil
    projectCatalogSheet.getRange("C8").setValue(headerDataCatalog.nameHydraulicWork_catalog);
    
    // Agregamos el numero de la concesión 
    projectCatalogSheet.getRange("K5").setValue(headerDataCatalog.concessionNumber_catalog);
    
    // Agregamos el nombre o razón social
    projectCatalogSheet.getRange("C4").setValue(headerDataCatalog.name_catalog);
    
    // Agregamos la fecha 
    projectCatalogSheet.getRange("L2").setValue(formatDateWorkCertificate(newDate));
    
    // Guardamos el numero de laa concesión
    saveConcessionNumber(headerDataCatalog.concessionNumber_catalog, projectData.infoCatalogSheetId);
    
    // Recorremos cada una de las ordenes
    for(var i = 0; i < orderProjects.length; i++){
      
      // agregamos el primer rango
      totalCell.push("I" + documentData.orderCounter);

      // insertamos los datos a la tarjeta
      documentData = insertDataCatalog(projectCatalogSheet, orderProjects[i], documentData);
      
      // validamos si es el ultimo elemento 
      if(i == orderProjects.length-1){
        
        // Agregamos los border azules 
        projectCatalogSheet.getRange("B" + documentData.orderCounter  + ":I" + documentData.orderCounter).setBorder(false, false, true, true, false, false, "#1155cc", null);
        
        // Agregamos los border azules 
        projectCatalogSheet.getRange("K" + documentData.orderCounter  + ":L" + documentData.orderCounter).setBorder(false, true, true, false, false, false, "#1155cc", null);
        
      }
      
    }

    // validamos de que existan las ordenes
    if(documentData && documentData.orders && documentData.orders.length > 0){
      
      // Obtenemos la posicion para insertar la sintesis del catálogo 
      var position = 6 + Number(documentData.orderCounter);
      
      // Recorremos las ordenes
      for(var z = 0; z < documentData.orders.length; z++){
        
        // creamos la fila de la nueva orden
        projectCatalogSheet.insertRowAfter(position);
        
        // Agregamos el simbolo a la orden y el nombre de la orden
        projectCatalogSheet.getRange("A" + position + ":B" + position).setValues([ [documentData.orders[z][0], documentData.orders[z][1]]]);
                                    
        // Agregamos el importe de la orden 
        projectCatalogSheet.getRange("I" + position).setValue("=" + documentData.orders[z][2]);
        
        // Agregamos los border azules 
        projectCatalogSheet.getRange("A" + position  + ":L" + position).setBorder(false, true, false, true, false, false, "#1155cc",  SpreadsheetApp.BorderStyle.DOUBLE);
          
        // validamos si es el ultimo elemento 
        if(z == documentData.orders.length - 1){
          
          // Agregamos los border azules 
          projectCatalogSheet.getRange("A" + position  + ":L" + position).setBorder(false, true, true, true, false, false, "#1155cc",  SpreadsheetApp.BorderStyle.DOUBLE);
          
        }
        
        // Aumentamos la posición
        position++;
      }
    }
    
    
    // Agregamos la formula del total
    projectCatalogSheet.getRange("I" + (position + 2)).setValue("=" + totalCell.toString().replace(/\,/g, "+"));
    
    // agregamos la formula del valor en letras
    projectCatalogSheet.getRange("C" + (position + 3)).setValue("=importex(I" + (position + 2) + ")");
  }
  
  // obligamos guardar los datos en la hoja de cálculo
  SpreadsheetApp.flush();
  
  // Retornamos la url de la hoja creada  con parametros de configuración de imprimir
  return spreadsheet.getUrl() + "?exportFormat=pdf&format=pdf&size=letter&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=true&gridlines=false&fzr=true&gid=GID#gid=78855938";
}


/*
* permite guardar el numero de la concesion 
*/
function saveConcessionNumber(concessionNumber,infoCatalogSheetId){
     
  // Validamos de que exista el id de la hoja de información del catálogo
  if(infoCatalogSheetId){
    
    // Referenciamos la hoja de la información del catálogo
    var spreadSheet = SpreadsheetApp.openById(infoCatalogSheetId);
    
    // Validamos de se tenga acceso a la hoja de parametros de obra y se tengan los ordenes 
    if(spreadSheet){
      
      // Referenciamos a la hoja de obra civil
      var civilWork = spreadSheet.getSheetByName("Obra civil");
      
      // validamos de que se tenga acceso a la hoja de obra civil
      if(civilWork){
        
        // Guardamos el numero de la concesion en la información del catalogo 
        civilWork.getRange("A2").setValue(concessionNumber);
       
      }
    }
  }

}



/*
* Función que permite inserta los datos al catálogo
* @param {Reference} sheetNew Referencia a una hoja de libro de tarjetas
* @param {object} conceptData datos de la especie
*/
function insertDataCatalog(projectCatalogSheet, currentOrder, documentData){
  
  // variable para manipular las filas
  var currentLine = [],
      currentConcept = [],
      orderRange= null,
      startIndexGroupLine = 0,
      lineRange = null,
      totalLines = [],
      totalOrder = [];
  
  // Obtenemos el contador
  var orderCounter = documentData.orderCounter;
  
  // obtenemos el rango del total en el orden del proyecto
  orderRange = projectCatalogSheet.getRange("I" + orderCounter);
  
  // Anexamos la posición del documento
  documentData.orders.push(["├", currentOrder.orderName, "I" + orderCounter]);
  
  // creamos la fila de la nueva orden
  projectCatalogSheet.insertRowAfter(orderCounter);
  
  // agregamos el nombre de la orden
  projectCatalogSheet.getRange("B" + orderCounter).setValue(currentOrder.orderName);
    
  // Agregamos el simbolo a la orden 
  projectCatalogSheet.getRange("A" + orderCounter).setValue("├");
  
  // agregamos los estilos de la fila de orden
  projectCatalogSheet.getRange("A" + orderCounter + ":L" + orderCounter).setBackground("#d0e0e3").setFontColor("#980000").setFontWeight("bold").setFontFamily("Arial").setFontSize(10).setBorder(false, false, false, false, false, false);
  
  // inicializamos el array de totales del orden actual
  totalOrder = [];
    
  // recorremos cada una de las líneas de productos
  for(var j = 0; j < currentOrder.linesProducts.length; j++){
    
    // iniciamos el contador del total de filas de la orden actual
    orderCounter ++;
    
    // se establece el punto inicar del rango
    startIndexGroupLine = orderCounter;
    
    // obtenemos el rango del total en la línea de producto
    lineRange = projectCatalogSheet.getRange("I" + orderCounter ).setNumberFormat("$#,##0.0###");  
    
    // referenciamos la linea actual
    currentLine = currentOrder.linesProducts[j];
    
    // agregamos el rango de donde va a quedar el todal de la linea actual
    totalOrder.push("I" + orderCounter);
    
    // creamos la fila de la nueva orden
    projectCatalogSheet.insertRowAfter(orderCounter);
    
    // agregamos el nombre de la linea del producto
    projectCatalogSheet.getRange("B" + orderCounter + ":C" + orderCounter).setValues([["'+", currentLine.lineProductName]]);
    
    // Agregamos los estilos a la fila de la linea 
    projectCatalogSheet.getRange("B" + orderCounter + ":L" + orderCounter).setBorder(true, false, true, false, false, false, "#1155cc", null).setFontColor("#1155cc").setFontWeight("bold").setBackground("#f7f7f7");
    
    // inicializamos el array de totales del lineas de productos
    totalLines = [];
    
    var tempPrice = null; 
    
    // recorremos cada uno de las especies
    for(var k = 0; k < currentLine.conceptList.length; k++){
      
      // iniciamos el contador del total de filas de la orden actual
      orderCounter ++;
      
      // referenciamos la especie actual
      currentConcept = currentLine.conceptList[k];
      
      // creamos la fila de la nueva orden
      projectCatalogSheet.insertRowAfter(orderCounter);
      
      // agregamos el rango de donde va a quedar el total de cada especie
      totalLines.push("I" + orderCounter);
      
      // agregamos el nombre de la lídea del producto
      projectCatalogSheet.getRange("C" + orderCounter + ":E" + orderCounter).setValues([[
        "Key",
        currentConcept.conceptData[1],
        currentConcept.conceptData[4]
      ]]);
      
      // Agregamos la cantidad
      projectCatalogSheet.getRange("F" + orderCounter).setNumberFormat("#,##0.0###").setValue((currentConcept.conceptCount ? currentConcept.conceptCount : 0));
      
      
      tempPrice = (String(currentConcept.conceptData[5]).replace(/\$|\,/g, "")).replace(/\./g, ",")

      // Agregamos el precio unitario y el importe 
      projectCatalogSheet.getRange("H" + orderCounter + ":I" + orderCounter).setNumberFormat("$#,##0.0###").setValues([[
       tempPrice ,
        "=F" + orderCounter +"*H" +orderCounter
      ]]);
      
      // restablecemos los estilos 
      projectCatalogSheet.getRange("C" + orderCounter + ":L" + orderCounter).setWrap(true).setBackground("#fff").setFontColor("#000").setFontWeight("normal").setFontFamily("Arial").setFontSize(10);
      
      // agregamos el estilo para que se alinie en el cento la columna G de una especie
      projectCatalogSheet.getRange("J" + orderCounter).setBorder(false, true, false, true, false, false, "#1155cc", null);;
    }
    
    // agrupamos las líneas de producto
    projectCatalogSheet.getRange("A" + (startIndexGroupLine + 1) + ":A" + orderCounter).shiftRowGroupDepth(1);
    
    // se define el nuevo grupo y agrupamos por especie
    projectCatalogSheet.getRange("A" + startIndexGroupLine + ":A" + orderCounter).shiftRowGroupDepth(1);
    
    // agregamos la información del total para la linea de producto actual
    lineRange.setNumberFormat("$#,##0.0###").setValue( "=" + totalLines.toString().replace(/\,/g, "+")).setFontColor("#1155cc").setFontWeight("bold");
      
  }
  
  // agregamos la información del total para la orden del proyecto
  orderRange.setNumberFormat("$#,##0.0###").setValue("=" + totalOrder.toString().replace(/\,/g, "+"));
  
  //Aumentamos el contador de las filas
  orderCounter++;
  
  // Reemplazamos la cantidad
  documentData.orderCounter = orderCounter;
  
  // retornamos el contador de las filas
  return documentData;
}


/***************************************************Creación de Tarjetas*********************************************************************************/

/*
* Función que permite crear la hoja de calculo con sus respectivas tarjetas 
*/
function createCards(orderProjects, projectData) {

  // Referenciamos los recursos globales
  var resources = getResources();  
  
  // Referenciamos la carpeta donde se desea almacenar las tarjetas
  var  mainFolder = DriveApp.getFolderById(resources.cardFolderId);
  
  // Obtenemos la fecha actual 
  var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  
  // Nombre del archivo de tarjetas  
  var fileName = "Tarjetas de especies para el proyecto: " + projectData.id + " - " +  currentDate;
  
  // Se crea una copia de la plantilla de las tarjetas
  var fileCopy = DriveApp.getFileById(resources.cardTemplateId).makeCopy(fileName, mainFolder);
  
  // Agregamos los permisos para que cuaquier usuarios dentro de Hidroexpert pueda editar
  fileCopy.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
  
  // Se referencia el archivo
  var spreadsheet = SpreadsheetApp.open(fileCopy);
  
  // Variables para manipular los filas 
  var currentOrder = [],
      currentLine= [],
      currentConcept = [];
  
  // Obtenemos los datos de las especies costos 
  var costsConcepts = getCostsConcepts(projectData.infoCatalogSheetId);
  
  // Referenciamos la pantiilla de la tarjeta
  var templatedSheet = spreadsheet.getSheetByName("_Tarjeta");
  
  // Validamos que existan especies y la hoja se tenga acceso
  if(orderProjects && spreadsheet){
    
    // Recorremos cada una de las ordenes
    for(var i = 0; i < orderProjects.length; i++){
      
      // Obtenemos los datos de la orden 
      currentOrder = orderProjects[i];
      
      // recorremos cada una de las líneas de productos
      for(var j = 0; j < currentOrder.linesProducts.length; j++){

        // Obtenemos los datos de la  linea
        currentLine = currentOrder.linesProducts[j];

        // recorremos cada uno de las especies
        for(var k = 0; k < currentLine.conceptList.length; k++){
        
          // Obteneomos los datos de la especie 
          currentConcept = currentLine.conceptList[k].conceptData;
          
          // Validamos de que vengan lineas
          if(costsConcepts && costsConcepts[currentLine.lineProductName]){
            
            // Creamos la copia de la plantilla 
            var sheetNew = templatedSheet.copyTo(spreadsheet);
         
            // Cambiamos el nombre a la plantilla
            sheetNew.setName("O" + ( i + 1 ) + "L" + ( j + 1 ) + "T" + ( k + 1 ));
            
            // insertamos los datos a la tarjeta
            insertDataCard(sheetNew, currentConcept, costsConcepts[currentLine.lineProductName][currentConcept[1]], resources, projectData.id,  currentLine.conceptList[k].conceptCount);
          }
        }
      }
    }
    // Eliminamos la hoja platilla
    spreadsheet.deleteSheet(templatedSheet);
  }
  
  // se guardan todos los cambios de la hoja de cálculo
  SpreadsheetApp.flush();
  
  // Retornamos la url de la hoja creada 
  return convertSpreadsheetToPdf(spreadsheet, mainFolder);
}

/*
* Función que permite insertar los datos de las especies a la tarjeta
*/
function insertDataCard(cardSheet, conceptData, cost, resources, projectId, conceptCount){

  //Validamos de que se tenga la referencia a la hoja y los datos de la especie
  if(cardSheet && conceptData){
    
    // Agregamos el analisis del concpeto (Nombre, unidad, cantidad)
    cardSheet.getRange("D10:F10").setNumberFormat("@STRING@").setValues([[conceptData[1], conceptData[4], conceptCount ]]);
    
    // Agregamos el analisis del concpeto (P.unitario, importe) Con el formato de numero 
    cardSheet.getRange("G10:H10").setNumberFormat("$#,##0.0###"); //.setValues([[conceptData[5], ( conceptCount * conceptData[5] )]]);
    
    // Posición donde se empiezan a insertar los items
    var cardData = {
       position: 13,
       totalCost: [],
       totalCard: []
     };
    
    // Validamos de que vengan los costos de la especie 
    if(cost){
    
      // insertamos los costos para el grupo materiales 
      cardData = insertGroupCosts(cardData, cost.materials, cardSheet);
      
      // insertamos los costos para el grupo mano de obra 
      cardData = insertGroupCosts(cardData, cost.workforce, cardSheet);
      
      // insertamos los costos para el grupo equipos  
      cardData = insertGroupCosts(cardData, cost.teams, cardSheet);
      
      // insertamos los costos para el grupo de herramientas
      cardData = insertGroupCosts(cardData, cost.tools, cardSheet);
      
      // insertamos los costos para el grupo mano de auxiliares
      cardData = insertGroupCosts(cardData, cost.auxiliaries, cardSheet);
      
      // insertamos los costos para el grupo de fletes
      cardData = insertGroupCosts(cardData, cost.lading, cardSheet);
      
      // insertamos los costos para el grupo mano de subcontratos 
      //cardData = insertGroupCosts(cardData, cost.subcontracts, cardSheet);

      // establecemos la fila donde esta el total de costos directos
      cardData.position = cardData.position - 1;
      
      // agregamos la información del total el grupo 
      cardSheet.getRange("H" + (cardData.position)).setValue("=" + cardData.totalCost.toString().replace(/\,/g, "+"));
      
      //agregamos la posición
      cardData.totalCard.push("H" + (cardData.position));
      
      //agregamos los datos de costos de la especie
      addConceptCosts(cardSheet, cardData, resources, conceptData[5]);
    }
  }
}



/*
*+ Función que me permite agregar los costos a la tarjeta
* param {spreadShett}  cardSheet referencia a la tarjeta que se esta elaborando
* param {objectData}  cardData datos de la tarjeta
*/
function addConceptCosts(cardSheet, cardData, resources, total, targetHtml){
 
  // validamos de que venga los costos 
  if(cardData && cardData.totalCost[0]){
    
    // obtenemsos el costos total
    var totalCost = 0;
    
    // Validamos si es para vizulización de una tarjeta
    if(targetHtml){
      // Obtenemos el total de la posición
      totalCost = cardData.totalCost[0];
      
    }else{
    
      //Obtenemos el total
      totalCost = cardSheet.getRange(cardData.totalCost[0]).getDisplayValues();
    }
    
    // Validamos si vienen los recursos globales
    if(!resources){
    
      // Referenciamos los recursos globales
      resources = getResources();  
    }
    
    // 
    //Referenciamos la hoja de parametros de obra
    var spreadSheet = SpreadsheetApp.openById(resources.workParameters);
    
    // Obtenemos la hoja de configuraciones 
    var configSheet = spreadSheet.getSheetByName("Configuraciones");
    
    // Obtenemos los datos 
    var arrayData = configSheet.getRange(2, 1, 5, 2).getDisplayValues();
    
    // validamos de que vengan datos
    if(arrayData){
    
      // Referenciamos la hoja de costos indirectos
      var costSheet = spreadSheet.getSheetByName("Costos indirectos");
      
      // Creamos la copia de la plantilla 
      var sheetNew = costSheet.copyTo(spreadSheet);

      // insertamos el valor de mano de obra
      sheetNew.getRange(arrayData[0][1]).setValue(totalCost);
      
      // Obtenemos los costos indirectos oficina 
      var indirectOffice = sheetNew.getRange(arrayData[1][1]).getDisplayValue();
      
      // Obtenemos los costos indirectos campo  
      var indirectField = sheetNew.getRange(arrayData[2][1]).getDisplayValue();
      
      // Referenciamos la hoja de finaciamiento
      var financingSheet = spreadSheet.getSheetByName("Financiamiento");
      
      // Referenciamos la hoja de utilidad 
      var utilitySheet = spreadSheet.getSheetByName("Utilidad");
      
      // Obtenemos el financiamiento
      var financing = financingSheet.getRange(arrayData[3][1]).getDisplayValue();
      
      // Obtenemos la utilidad
      var utility = utilitySheet.getRange(arrayData[4][1]).getDisplayValue();
      
      // Obtenemos el total de la posición
      var object = {
        financing: String(financing ? financing: 0).replace(/\%/g, ""),
        utility:  String(utility ? utility: 0).replace(/\%/g, ""),
        indirectField: indirectField.replace(/\%/g, ""),
        indirectOffice: indirectOffice.replace(/\%/g, ""),
        indirect: (Number(indirectOffice.replace(/\%/g, "")) + Number(indirectField.replace(/\%/g, ""))),
        additionalCharge: "0.5025",
        impost: "16"
      };
      
      // Validamos si es para vizulización de una tarjeta
      if(targetHtml){
      
        // Eliminamos la de costos indirectos
        spreadSheet.deleteSheet(sheetNew);
        
        // Retornamos los valores
        return object;
      }
      
      // Aumentamos la posición para ubicarnos en la celda de precios indirectos
      cardData.position = cardData.position + 1;
      
      // Agregamos la suma de costos directos 
      cardSheet.getRange("F" + (cardData.position)).setNumberFormat("##0.0###").setValue("=F"+ (cardData.position + 1) + "+F"+ (cardData.position + 2)); 
      
      // Aumentamos la posición para ubicarnos en la posición de precios de oficina
      cardData.position = cardData.position + 1;
      
      // Agregamos los costos de indirectos de oficina 
      cardSheet.getRange("F" + cardData.position).setNumberFormat("##0.0###").setValue(String(object.indirectOffice).replace(/\./g, ","));

      // Agregamos los costos de indirectos de campo 
      cardSheet.getRange("F" + (cardData.position + 1)).setNumberFormat("##0.0###").setValue(String(object.indirectField).replace(/\./g, ","));

      // Agregamos el finaciamiento
      cardSheet.getRange("F" + (cardData.position + 2) ).setNumberFormat("##0.0###").setValue(String(object.financing).replace(/\./g, ","));
      
      // Agregamos la utilidad
      cardSheet.getRange("F" + (cardData.position + 3)).setNumberFormat("##0.0###").setValue(String(object.utility).replace(/\./g, ","));
      
      // Agregamos el total de sobre costos
      cardSheet.getRange("H" + (cardData.position + 6)).setNumberFormat("$#,##0.0###").setValue("=G" + (cardData.position - 1) + "+G" + (cardData.position + 2) + "+G" + (cardData.position + 3) + "+G" + (cardData.position + 4) + "+G" + (cardData.position + 5));
      
       // Establecemos el precio unitario
      cardSheet.getRange("D" + (cardData.position + 9)).setNumberFormat("$#,##0.0###").setValue("=importex(H" + (cardData.position + 8) + ")");
      
      // Eliminamos la de costos indirectos
      spreadSheet.deleteSheet(sheetNew);
    }
  }
}


/*
* Función que permite insertar un grupo a la tarjeta (cardData, cost.materials, cardSheet)
*/
function insertGroupCosts(cardData, middleObj, cardSheet, st){
  
  // validamos de que vengan items en el grupo 
  if(middleObj && middleObj.recordsObj && middleObj.recordsObj.length > 0){
    
    // Obtenemos los datos de cada uno de los items
    var arrayData = middleObj.recordsObj;
    
    // Posiciones para el total del grupo 
    var totalGroup = [];
    
    // Recorremos cada uno de los items
    for(var i = 0; i < arrayData.length; i++){
      
      // creamos la fila de la nueva orden
      cardSheet.insertRowAfter(cardData.position); 
      
      // Aumentamos la posición
      cardData.position ++;
      
      // referenciamos el rango para agregar la clave, descripción y unidad
      cardSheet.getRange(cardData.position, 3, 1, 4)
      .setWrap(true)
      .setVerticalAlignment("middle")
      .setFontColor("#000")
      .setFontWeight("normal")
      .setFontFamily("Arial")
      .setFontSize(10)
      .setHorizontalAlignment("right")
      .setValues([[arrayData[i].key, arrayData[i].description, arrayData[i].unity, arrayData[i].quantity.replace(/\./g, ",")]]);
              
      cardSheet.getRange(cardData.position, 7, 1, 2)
      .setWrap(true)
      .setVerticalAlignment("middle")
      .setFontColor("#000")
      .setFontWeight("normal")
      .setFontFamily("Arial")
      .setFontSize(10)
      .setNumberFormat("$#,##0.0###")
      .setHorizontalAlignment("right")
      .setValues([[arrayData[i].unitCost.replace(/\,/g, "").replace(/\./g, ","), ("=F" + cardData.position + "*G" + cardData.position)]])
      
      // Ajustamos los estilos de la primera columna 
      cardSheet.getRange(cardData.position, 3, 1, 1).setWrap(false).setHorizontalAlignment("center").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      
      // Ajustamos los estilos de la segunda columna 
      cardSheet.getRange(cardData.position, 4, 1, 3).setHorizontalAlignment("center");
      
      // agregamos el rango de donde va a quedar el total de cada  grupo 
      totalGroup.push("H" + cardData.position);
    }
    
    // Agregamos la posición del total del grupo 
    cardData.totalCost.push("H" + (cardData.position + 1)); 
    
    // agregamos la información del total el grupo 
    cardSheet.getRange("H" + (cardData.position + 1)).setNumberFormat("$#,##0.0###").setValue("=" +  totalGroup.toString().replace(/\,/g, "+"));
  }
    
  //Aumentamos 4 para el proximo item
  cardData.position = cardData.position + 4;
  
  // Retornamos la posicion
  return cardData;
}


/**
 * Función que permite convertir una hoja de cálculo en un archivo PDF
 * @param {Class} spreadsheet Referncia de la hoja de cálculo
 * @param {Folder} rootFolder Carpeta donde se almacena el archivo generado
 **/
function convertSpreadsheetToPdf(spreadsheet, rootFolder) {

 // se define el nombre del pdf
  var pdfName = spreadsheet.getName();
  
  // se define la url base
  var url_base = spreadsheet.getUrl().replace(/edit$/, "");
  
  // se concadena los parametros para luego exportar el archivo
  var url_ext = "export?exportFormat=pdf&format=pdf"   //export as pdf
  // Imprimir el libro completo, si desea imprimir una hoja en particular se usa el parametro (&gid)
  + "&id=" + spreadsheet.getId()
  // Los siguientes parámetros son opcionales...
  + "&size=letter"      // Tamaño de papel
  + "&portrait=true"    // Orientación si este parametro es false significa que la orientación es horizontal
  + "&fitw=true"        // Ajustal ancho del papel
  + "&sheetnames=false&printtitle=false&pagenumbers=false"  // ocultar encabezados y pies de página opcionales
  + "&gridlines=false"  // Oculta cuadriculas
  + "&fzr=false";       // No repetir encabezados de fila (filas congeladas) en cada página
  
  // se definen las opciones de autorización
  var options = {
    headers: {
      "Authorization": "Bearer " +  ScriptApp.getOAuthToken(),
    }
  }
  
  // se realiza la petición para convertir el archivo en PDF  
  var response = UrlFetchApp.fetch(url_base + url_ext, options);
  
  // se valida que la petición fue exitosa
  if(response.getResponseCode() == 200){
    
    // se obtiene el blob y se define un nombre
    var newBlob = response.getBlob().setName(pdfName + ".pdf");
    
    // se crea el archivo en la carpeta respectiva
    var pdfFile = rootFolder.createFile(newBlob);
    
    // retornamos la url del archivo
    return pdfFile.getUrl();
    
  }
  
  // se retorna por defecto null
  return null;
  
}
