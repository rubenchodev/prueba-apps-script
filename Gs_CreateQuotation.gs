/*
* Permiete eliminar una version del proyecto
*/
function deleteVersionProject(deleteNameVersion, versionDataDelete, requestData, newOptionSelected){
    
  // Validamos de que se tenga los datos
  if(requestData && requestData.documentId){
  
     // Referencamos el documentos 
     var spreadSheet = SpreadsheetApp.openById(requestData.documentId);
     
     // Referenciamos la hoja para eliminar
     var sheetProject = spreadSheet.getSheetByName(deleteNameVersion);
     
     // validamos de que se tenga la refencia de la hoja 
     if(sheetProject){
     
       // Cambiamos el nombre de la hoja 
       sheetProject.setName(deleteNameVersion + "_Eliminada");
     }
     
     // Referenciamos los recursos globales 
     var resources = getResources();
      
     // Referenciamos el libro de las solictudades 
     var requestSpreadSheet = SpreadsheetApp.openById(resources.requestSheetId);
     
     // Referenciamos la hoja de cotizciones 
     var quatitionSheet = requestSpreadSheet.getSheetByName("Cotizaciones");
     
     // Referenciamos la hoja de administracion de solicitudes
     var adminVersionSheet = requestSpreadSheet.getSheetByName("Administrador de versiones");
     
     // Validamos de que se tenga acceso a la hoja 
     if(quatitionSheet && newOptionSelected && adminVersionSheet && versionDataDelete && versionDataDelete.rowIndex){
     
      // Referenciamos la fila y hacemos el cambio de la hoja principal 
      quatitionSheet.getRange(requestData.rowIndex, 12).setValue(newOptionSelected);
      
      // Referenciamos la celda y cambiamos el nombre 
      adminVersionSheet.getRange(versionDataDelete.rowIndex, 1).setValue(requestData.consecutive + "_Eliminada");
      
     }
  }
}


/**
 * Función que permite realizar el duplicado de una pestaña en blanco
**/
function duplicateTabSelectedBlank(objectData){
  
  // referenciamos el libro
  var spreadsheet = SpreadsheetApp.openById(objectData.documentId);
  
  // se obtiene la hoja segun la seleccionda por el usuario
  var sheetSelected = spreadsheet.getSheetByName(objectData.sheetName);
  
  // se valida si existe la hoja
  if(sheetSelected){
  
    // Obtenemos la cantidad de pestañas
    var sheetCount = spreadsheet.getSheets().length;
      
    // obtenemos la hoja con el nombre "Proyecto V" + objectData.sheetsCount
    var sheetReference = spreadsheet.getSheetByName("Proyecto V" + sheetCount -1);
    
    // validamos si ya existe una hoja con el nombre "Proyecto V" + objectData.sheetsCount
    if(sheetReference){
      
      // retornamos el error al momento de intentar crear el duplicado
      throw new Error("Se esta intentando crear una hoja con el nombre: <b>Proyecto V" + objectData.sheetsCount + "</b>, sin embargo se detecta que ya fue creada manualmente. Para poder continuar por favor renombre la hoja que se creo manualmente.");
      
    } else {
      
      // se obtiene los recursos de hojas de cálculo
      var resources = getResources();
      
      // referenciamoas la hoja plantila
      var spreadsheetTemplate = SpreadsheetApp.openById(resources.templateId);
      
      // se referncia la hoja de proyecto
      var sheetProject = spreadsheetTemplate.getSheetByName("Proyecto");
      
      // se valida si no existe la información
      if(!sheetProject){
        
        // Mostramos el error
        throw new Error("Lo sentimos no existe la hoja para realizar una copia de la versión.");
      }
      
    
      // creamos el duplicado de la hoja
      sheetReference = sheetProject.copyTo(spreadsheet);
      
      // cambiamos el nombre de la pestaña a "Proyecto V" + objectData.sheetsCount
      sheetReference.setName("Proyecto V" + sheetCount);
      
      // Variable para obtener la información de la solicitud
      var structureData = {};
      
      // Obtenemos los datos de la pestaña
      try{
      
        // Convertimos en objeto los datos de la versión seleccionada
        structureData = JSON.parse(objectData.structure);
        
      }catch(e){
        structureData = {};
      }
      
      // obtenemos el nombre del proyecto de acuerdo a los datos de clasificación del sistema, tipo, etc
      var projectName = getProjectName(resources.settingSheetId, structureData);
      
      // agregamos los datos basicos
      addBasicData(sheetReference, structureData, projectName);
      
      // referenciamos la hoja de solicitudes
      var requestSpreadsheet = SpreadsheetApp.openById(resources.requestSheetId);
      
      // referenciamos la hoja de "Cotizaciones" & "Administrador de versiones"
      var cotSheet = requestSpreadsheet.getSheetByName("Cotizaciones");
      var manageSheet = requestSpreadsheet.getSheetByName("Administrador de versiones");
      
      // se valida si existe la hoja de "Administrador de versiones"
      if(cotSheet && manageSheet){
        
        // cambiamos el estado en la columna M a "SI"
        cotSheet.getRange("M" + objectData.rowIndex).setValue("SI");
        
        // obtenemos la ultima fila y le aumentamos en 1
        var lastRow = manageSheet.getLastRow() + 1;
        
        // agregamos la información de la referencia
        manageSheet.getRange(lastRow, 1, 1, 3).setNumberFormat("@STRING@").setValues([[
          objectData.consecutive,
          ("Proyecto V" + sheetCount),
          objectData.structure
        ]]);
        
        // retornamos los datos necesarios par actualizar el select de versiones
        return {
          newSheetName: ("Proyecto V" + sheetCount),
          newRowIndex: lastRow,
        };
        
      } else {
        
        // retornamos el error informando que no existe la hoja de "Cotizaciones" & "Administrador de versiones"
        throw new Error("Lo sentimos no encontramos una de las siguientes hojas 'Cotizaciones' o 'Administrador de versiones'");
        
      }
    }
  } else {
    
    // retornamos el mensaje respectivo
    throw new Error("No es posible encontrar la hoja: <b>" + objectData.sheetName + "</b>");
    
  }
  
}

/**
 * Función que permite realizar el duplicado de una pestaña
**/
function duplicateTabSelected(objectData){
//  throw JSON.stringify(objectData)
  // referenciamos el libro
  var spreadsheet = SpreadsheetApp.openById(objectData.documentId);
  
  // se obtiene la hoja segun la seleccionda por el usuario
  var sheetSelected = spreadsheet.getSheetByName(objectData.sheetName);
  
  // se valida si existe la hoja
  if(sheetSelected){
    
    // Obtenemos la cantidad de pestañas
    var sheetCount = spreadsheet.getSheets().length;
    
    // obtenemos la hoja con el nombre "Proyecto V" + objectData.sheetsCount
    var sheetReference = spreadsheet.getSheetByName("Proyecto V" + sheetCount -1);
    
    // validamos si ya existe una hoja con el nombre "Proyecto V" + objectData.sheetsCount
    if(sheetReference){
      
      // retornamos el error al momento de intentar crear el duplicado
      throw new Error("Se esta intentando crear una hoja con el nombre: <b>Proyecto V" + sheetCount + "</b>, sin embargo se detecta que ya fue creada manualmente. Para poder continuar por favor renombre la hoja que se creo manualmente.");
      
    } else {
      
      // creamos el duplicado de la hoja
      sheetReference = sheetSelected.copyTo(spreadsheet);
      
      // cambiamos el nombre de la pestaña a "Proyecto V" + objectData.sheetsCount
      sheetReference.setName("Proyecto V" + sheetCount);
      
      // se obtiene los recursos de hojas de cálculo
      var resources = getResources();
      
      // referenciamos la hoja de solicitudes
      var requestSpreadsheet = SpreadsheetApp.openById(resources.requestSheetId);
      
      // referenciamos la hoja de "Cotizaciones" & "Administrador de versiones"
      var cotSheet = requestSpreadsheet.getSheetByName("Cotizaciones");
      var manageSheet = requestSpreadsheet.getSheetByName("Administrador de versiones");
      
      // se valida si existe la hoja de "Administrador de versiones"
      if(cotSheet && manageSheet){
        
        // cambiamos el estado en la columna M a "SI"
        cotSheet.getRange("M" + objectData.rowIndex).setValue("SI");
        
        // obtenemos la ultima fila y le aumentamos en 1
        var lastRow = manageSheet.getLastRow() + 1;
        
        // agregamos la información de la referencia
        manageSheet.getRange(lastRow, 1, 1, 3).setNumberFormat("@STRING@").setValues([[
          objectData.consecutive,
          ("Proyecto V" + sheetCount),
          objectData.structure
        ]]);
        
        // retornamos los datos necesarios par actualizar el select de versiones
        return {
          newSheetName: ("Proyecto V" + sheetCount),
          newRowIndex: lastRow,
        };
        
      } else {
        
        // retornamos el error informando que no existe la hoja de "Cotizaciones" & "Administrador de versiones"
        throw new Error("Lo sentimos no encontramos una de las siguientes hojas 'Cotizaciones' o 'Administrador de versiones'");
        
      }
    }
  } else {
    
    // retornamos el mensaje respectivo
    throw new Error("No es posible encontrar la hoja: <b>" + objectData.sheetName + "</b>");
    
  }
  
}

/**
 * Función que permite obtener las solicitudes por el años seleccionado
**/
function getProjectByYear(userRole, year){
  
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // se obtiene el correo del usuario
  var userMail = getUserMail();
  
  // retornamos los datos
  return filterRequestByMail(userMail, resources.requestSheetId, userRole, year);
}

/**
 * Función que permite obtener los datos del panel de creación
**/
function getCreatePanelData(userRole){
  
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // se obtiene el correo del usuario
  var userMail = getUserMail();
  
  // referenciamos la hoja de cálculo de configuración
  var settingSpreadsheet = SpreadsheetApp.openById(resources.settingSheetId);
  
  // retornamos los datos
  return {
    ratingList: getRatingList(settingSpreadsheet),
    comunNameList: getComunNameList(settingSpreadsheet),
    cultivationList: getCultivationList(settingSpreadsheet),
    projectClassList: getProjectClassList(settingSpreadsheet),
    requestData: filterRequestByMail(userMail, resources.requestSheetId, userRole),
    priorityProductLines: getPriorityProductLines(settingSpreadsheet, userMail)
  };
}

/**
 * Función que permite obtener el listado de nombres comunes
**/
function getComunNameList(spreadsheet) {
  
  // obtenemos la hoja de segun parametro
  var sheet = spreadsheet.getSheetByName("Nombre del proyecto");
  
  // validamos si existe la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
      
      // retornamos los datos
      var allData = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      
      // variable para almacenar los datos de la lista de nombres comunes
      var comunNameList = [];
      
      // recorremos cada uno de los registros
      for(var i = 0; i < allData.length; i++){
        
        // agregamos los datos en el objeto para el campo de busqueda
        comunNameList.push({
          clasification: allData[i][0], // Clasificación
          systemType: allData[i][1], // Tipo de sistema
          emmisionForm: allData[i][2], // Forma de emisión
          features: allData[i][3], // Caracteristicas del emisor
          comunName: allData[i][5], // Nombre comun
        });
        
      }
      
      // retornamos los datos
      return comunNameList;
      
    }
  }
  
  // retornamos por defecto un array vacio
  return [];
}

/**
 * Función que permite generar la cot. respectiva
**/
function createQuotation(formData, consecutiveClient, sheetsData, documentId, rowIndex, lastRowTemp, userRole, year){
 
  // variable para guardar la referencia del archivo de la plantilla
  var spreadsheet = null,
      mainFolder = null,
      message = "Cotización generada con éxito.",
      consecutive = null,
      action = "creation",
      infoCatalogSheetId = null;
  
  // referenciamos los datos globales
  var resources = getResources(), projectSheet = null;

  // Agregamos la carpeta principal de almacenamiento de cotizaciones
  formData.folderId = resources.projectMainFolderId;
  
  // obtenemos el nombre del proyecto de acuerdo a los datos de clasificación del sistema, tipo, etc
  var projectName = getProjectName(resources.settingSheetId, formData);
  
  // variable para definir el nombre del archivo
  var fileName = "@PP: " + formData.client + ": " + formData.systemType + " " + formData.surface + " has, " + formData.location + ", " + formData.municipality + " - " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  
  // validamos si existe el id del documento lo que significa que se va a modificar la cot
  if(documentId){
    
    // Cambiamos el nombre del archivo
    DriveApp.getFileById(documentId).setName(fileName);

    // referenciamos el libro
    spreadsheet = SpreadsheetApp.openById(documentId);
    
    // cambiamos un mensaje
    message = "Cotización actualizada con éxito.";
    
    // cambiamos el estado 
    action = "modification";
    
    // referenciamos el consecutivo
    consecutive = consecutiveClient;
    
    // referenciamoas la hoja plantila
    var spreadsheetTemplate = SpreadsheetApp.openById(resources.templateId);
    
    // se referncia la hoja de proyecto
    var sheetProject = spreadsheetTemplate.getSheetByName("Proyecto");
    
    // se valida si no existe la información
    if(!sheetProject){
      
      // Mostramos el error
      throw new Error("Lo sentimos no existe la hoja para realizar una copia de la versión.");
    }
    
    // referenciamos la hoja de segun la versión seleccionda.
    projectSheet = spreadsheet.getSheetByName(sheetsData.sheetName);
    
    // variable para referenciar la hoja que se desea eliminar
    var sheetToDelete = null;
    
    // se valida si existe la hoja
    if(projectSheet){
      
      // Referenciamos la hoja a eliminar
      sheetToDelete = projectSheet.setName("currentSheetDelete");
      
    }
    
    // creamos el duplicado de la hoja principal
    projectSheet = sheetProject.copyTo(spreadsheet).setName(sheetsData.sheetName);
    
    // se valida si existe la hoja
    if(sheetToDelete){
      
      // removemos la hoja
      spreadsheet.deleteSheet(sheetToDelete);      
    }
    
  } else {
    
    // referenciamos la carpeta donde se desea almacenar la cotización
    mainFolder = DriveApp.getFolderById(formData.folderId);
    
    // se crea una copia de la plantilla
    var fileCopy = DriveApp.getFileById(resources.templateId).makeCopy(fileName, mainFolder);
    
    // Agregamos los permisos para que cuaquier usuarion dentro de Hidroexpert pueda editar
    // fileCopy.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
    
    // se referencia el archivo
    spreadsheet = SpreadsheetApp.open(fileCopy);
    
    // obtenemos el consecutivo
    consecutive = "PR" + getConsecutive("PR_Consecutive");
    
    // Sacamos una copia de la plantilla de la hoja de catálogo de cuenta
    var infoCatalogSheet  = DriveApp.getFileById(resources.infoCatalogSheetId).makeCopy("Información del catálogo de especie para el proyecto " + consecutive, mainFolder);
    
    // obtenemos el Id del archivo de información del catálogo
    // infoCatalogSheetId = infoCatalogSheet.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT).getId();
    infoCatalogSheetId = infoCatalogSheet.getId();
    
    // referenciamos la hoja de segun la versión seleccionda.
    projectSheet = spreadsheet.getSheetByName(sheetsData.sheetName);
    
  }
  
  // agregamos los datos basicos
  addBasicData(projectSheet, formData, projectName);
    
  // insertamos los datos en la hoja inicial
  insertInformationInSheet(projectSheet, formData.orderProjects);
  
  // obtenemos el correo del usuario
  var userMail = getUserMail();
  
  // registramos la colección de patrones del proyecto
  registerProjectPatterns(resources, formData, userMail, consecutive);
 
  // adicionamos el registro en la hoja principal
  registerGeneralDataFromSheet(sheetsData, resources.requestSheetId, spreadsheet.getId(), formData, projectName, rowIndex, userMail, consecutive, infoCatalogSheetId);

  // se retorna la respuesta
  return {
    action: action,
    message: message,
    requestData: filterRequestByMail(userMail, resources.requestSheetId, userRole, year)
  };
}

/**
 * Función que permite registrar los datos de la colección de patrones del proyecto
**/
function registerProjectPatterns(resources, formData, userMail, consecutive){
  
  // referenciamos el libro de patrones
  var spreadsheetPatterns = SpreadsheetApp.openById(resources.patternsSheetId);
  
  // referenciamos la hoja de "Líneas" & "Orden"
  var lineSheet = spreadsheetPatterns.getSheetByName("Líneas");
  var orderSheet = spreadsheetPatterns.getSheetByName("Orden");
  
  // se referencia solo la información del presupuesto
  var orderProjects = formData.orderProjects,
      currentOrder = [],
      currentLine = [],
      currentConcept = [],
      resumeObj = {};
  
  // se recorre el presupuesto para obtener la información respectiva
  // recorremos cada una de las ordenes
  for(var i = 0; i < orderProjects.length; i++){
    
    // definimos el orden actual
    currentOrder = orderProjects[i];
    
    // definimos el nombre del orden
    resumeObj[currentOrder.orderName] = {
      total: 0,
      percentage: 0,
      lines: {}
    };
    
    // variable para guardar el total del orden
    var orderTotal = 0;
    
    // recorremos cada una de las líneas de productos
    for(var j = 0; j < currentOrder.linesProducts.length; j++){
      
      // referenciamos la linea actual
      currentLine = currentOrder.linesProducts[j];
      
      // variable para guardar el total de as líneas
      var lineTotal = 0;
      
      // recorremos cada uno de los especie
      for(var k = 0; k < currentLine.conceptList.length; k++){
        
        // referenciamos el especie actual
        currentConcept = currentLine.conceptList[k];
        
        // validamos si tiene datos del valor por unidad
        if(currentConcept.conceptData[5]){
          
          // agregamos el valor del especie
          lineTotal += Number(currentConcept.conceptCount) * Number(currentConcept.conceptData[5]);
        }
        
      }
      
      // agregamos la información de la linea
      resumeObj[currentOrder.orderName].lines[currentLine.lineProductName] = {
        total: roundDecimal(lineTotal, 2),
        percentage: 0
      };
      
      // vamos sumando el total de order
      orderTotal += lineTotal;
      
    }
    
    // se recorre cada una de las línea y se calcula el porcentaje de las lineas
    for(var key in resumeObj[currentOrder.orderName].lines){
      
      // actualizamos el porcentaje de la linea
      resumeObj[currentOrder.orderName].lines[key].percentage = roundDecimal((resumeObj[currentOrder.orderName].lines[key].total * 100) / orderTotal, 2);
    }
    
    // actualizamos el total del orden
    resumeObj[currentOrder.orderName].total = roundDecimal(orderTotal, 2);
    
    // actualizamos el porcentaje del valor total
    resumeObj[currentOrder.orderName].percentage = roundDecimal((orderTotal * 100) / formData.total, 2);
    
  }
  
  //definimos el objeto 
  var orderIndex = {
    "1. Sistema de riego": 13,
    "2. Obra civil": 14,
    "3. Equipo mecanico y electrico": 15,
    "4. Servicios complementarios": 16,
  };
  
  // se define el objeto de indices de lineas
  var lineIndex = {
    "E.1.1. Cabezal de Riego": 16,
    "E.1.2. Elementos de Medición": 17,
    "E.1.3. Filtración": 18,
    "E.1.4. Fertiriego": 19,
    "E.1.5. Sistemas de Automatización": 20,
    "E.1.6. Linea de Conducción": 21,
    "E.1.7. Línea Secundaria": 22,
    "E.1.8. Lineas Portalaterales": 23,
    "E.1.9. Emisores": 24,
    "E.1.10. Lineas Colectoras y Valvulas de Lavado": 25,
    "E.1.11. Accesorios de Automatización en Válvulas": 26,
    "E.1.12. Excavación y Relleno de Zanja": 27,
    "E.1.13. Línea de Seccionamiento": 28,
    "E.1.14. Seguridad del Sistema": 29,
    "E.2.1. Excavación y Movimiento de Tierras": 30,
    "E.2.2. Terraplenes": 31,
    "E.2.3. Captación": 32,
    "E.2.4. Carcamos": 33,
    "E.2.5. Descarga": 34,
    "E.2.6. Protección de las Instalaciones": 35,
    "E.2.7. Línea de Succión": 36,
    "E.3.1. Accesorios que integran el equipo mecanico y Eléctrico": 37,
    "E.3.2. Grupo Motor - Bomba": 38,
    "E.3.3. Bastidor Soporte - Bomba": 39,
    "E.3.4. Medidores de Gasto y Manometros": 40,
    "E.3.5. Válvulas de Operación y Especiales": 41,
    "E.3.6. Múltiple de Descarga: Tuberia y Piezas Especiales": 42,
    "E.4.1. Capacitación acerca del sistema de Riego": 43,
    "E.4.2. Calibración y puesta en marcha del sistema": 44,
    "E.4.3. Servicio de mantenimiento preventivo": 45,
    "E.4.4. Prueba del sistema de riego": 46
  }
  
  // se valida si existen las hoja de "Orden"
  if(orderSheet && lineSheet){
    
    // formamos el array para guardar el dato de orden
    var orderArray = [
      consecutive,
      userMail,
      formData.client,
      formData.issue,
      formData.location,
      formData.municipality,
      formData.state,
      formData.systemClasification,
      formData.systemType,
      formData.formEmission,
      formData.features,
      formData.total,
      "",
      "",
      "",
      "",
      ""
    ];
    
    // formamos el array para guardar el dato de lineas
    var lineArray = [
      consecutive,
      userMail,
      formData.client,
      formData.issue,
      formData.surface,
      formData.cultive,
      formData.location,
      formData.municipality,
      formData.state,
      formData.systemClasification,
      formData.systemType,
      formData.formEmission,
      formData.features,
      formData.total,
      "",
      "","","","","","","","","","",
      "","","","","","","","","","",
      "","","","","","","","","","",""
    ];
    
    // se recorre la información del orden
    for(var key in resumeObj){
    
      //Agregamos el registro en el array
      orderArray[orderIndex[key]] = resumeObj[key].percentage;
      
      // se recorre los datos de lineas
      for(var keyLine in resumeObj[key].lines){
        
        //Agregamos el registro en el array
        lineArray[lineIndex[keyLine]] = resumeObj[key].lines[keyLine].percentage;
        
      }
    }
    
    // agregamos los datos de orden y linea
    lineSheet.appendRow(lineArray);
    orderSheet.appendRow(orderArray);
    
  }
  
}

/**
 * Función que permite obtener el nombre del proyecto
**/
function getProjectName(fileId, formData){
  
  // referenciamos la hoja de cálculo de valores configurables
  var spreadsheet = SpreadsheetApp.openById(fileId);
  
  // obtenemos la hoja de "Nombre del proyecto"
  var sheet = spreadsheet.getSheetByName("Nombre del proyecto");
  
  // validamos si existe la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
      
      // retornamos los datos de los clientes
      var allData = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      
      // recorremos cada uno de los registros y los filtramos por usuario
      for(var i = 0; i < allData.length; i++){
        
        // validamos que valor de clasificasión, tipo de sistema, forma de emisión y caracteristicas del emisor
        if(allData[i][0] == formData.systemClasification && allData[i][1] == formData.systemType && allData[i][2] == formData.formEmission && allData[i][3] == formData.features){
          
          // retornamos 
          return allData[i][4];
          
        }
        
      }
      
    }
  }
  
  // retornamos por defecto un array vacio
  return "";
  
}

/**
 * Función que permite obtener los respectivos estados
**/
function getStateList(fileId) {
  
  // Obtenemos la cache
  var scriptCache = CacheService.getScriptCache();
  
  // Obtenemos la propiedad de la cache
  var stateCache = scriptCache.get("stateCache");
  
  // variable para guardar los municipios
  var allData = [];
  
  // validamos la existencia de la información en cache
  if(stateCache == null){
  
    // referenciamos la hoja de cálculo de estados y localidades
    var spreadsheet = SpreadsheetApp.openById(fileId);
    
    // obtenemos la hoja de "Estados"
    var stateSheet = spreadsheet.getSheetByName("Estados");
    
    // validamos si existe la hoja
    if(stateSheet){
      
      // obtenemos la ultima fila
      var lastRow = stateSheet.getLastRow();
      
      // validamos que exista mas de 1 fila
      if(lastRow > 1){
        
        // retornamos los datos de los clientes
        allData = stateSheet.getRange(2, 1, lastRow - 1, 4).getValues();
        
        // Guardamos los datos de la variable en cache
        scriptCache.put("stateCache", JSON.stringify(allData), 21600);
        
      }
    }
  } else {
    
    // obtenemos la matriz para manipularla desde la cache
    allData = JSON.parse(stateCache);
    
  }
  
  // variable para almacenar los datos del estado
  var stateData = [];
  
  // recorremos cada uno de los registros y los filtramos por usuario
  for(var i = 0; i < allData.length; i++){
    
    // agregamos los datos en el objeto para el campo de busqueda
    stateData.push({
      stateName: allData[i][0],
      startRow: Number(allData[i][1]),
      endRow: Number(allData[i][2])
    });
    
  }
  
  // retornamos los datos
  return stateData;
}

/**
 * Función que permite obtener los municipios del estado seleccionado
**/
function getMunicipalityList(startRow, endRow) {

  //restamos dos posiciones a las filas
  startRow -= 2;
  endRow -= 2;
  
  // Obtenemos la cache
  var scriptCache = CacheService.getScriptCache();
  
  // Obtenemos la propiedad de la cache
  var municipalityCache = scriptCache.get("municipalityCache");
  
  // variable para guardar los municipios
  var allData = [];
  
  // validamos la existencia de la información en cache
  if(municipalityCache == null){
  
    // se obtiene los recursos de hojas de cálculo
    var resources = getResources();
    
    // referenciamos la hoja de cálculo de estados y localidades
    var spreadsheet = SpreadsheetApp.openById(resources.stateSheetId);
    
    // obtenemos la hoja de "Municipios"
    var municipalitySheet = spreadsheet.getSheetByName("Municipios");
    
    // validamos si existe la hoja
    if(municipalitySheet){
      
      // obtenemos la ultima fila
      var lastRow = municipalitySheet.getLastRow();
      
      // retornamos los datos de los clientes
      allData = municipalitySheet.getRange("A2:C" + lastRow).getValues();
      
      // Guardamos los datos de la variable en cache
      scriptCache.put("municipalityCache", JSON.stringify(allData), 21600);
        
    }
  
  } else {
    
    // obtenemos la matriz para manipularla desde la cache
    allData = JSON.parse(municipalityCache);
    
  }
 
  
  // variable para almacenar los datos del estado
  var municipalityData = [];
  
  // recorremos cada uno de los registros y los filtramos por usuario
  for(var i = startRow; i < endRow; i++){
 
    // agregamos los datos en el objeto para el campo de busqueda
    municipalityData.push({
      municipalityName: allData[i][0],
      startRow: Number(allData[i][1]),
      endRow: Number(allData[i][2])
    });
    
  }
  
  // retornamos los datos
  return municipalityData;
  
}

/**
 * Función que permite obtener las localidades del municipio seleccionado
**/
function getLocationList(startRow, endRow) {
  
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // referenciamos la hoja de cálculo de estados y localidades
  var spreadsheet = SpreadsheetApp.openById(resources.stateSheetId);
  
  // obtenemos la hoja de "Localidades"
  var locationSheet = spreadsheet.getSheetByName("Localidades");
  
  // validamos si existe la hoja
  if(locationSheet){
    
    // retornamos los datos de los clientes
    var allData = locationSheet.getRange("A" + startRow + ":A" + endRow).getValues();
      
    // variable para almacenar los datos del estado
    var locationData = [];
      
    // recorremos cada uno de los registros y los filtramos por usuario
    for(var i = 0; i < allData.length; i++){
      
      // agregamos los datos en el objeto para el campo de busqueda
      locationData.push({
        locationName: allData[i][0]
      });
      
    }
    
    // retornamos los datos
    return locationData;
      
  }
  
  // retornamos por defecto un array vacio
  return [];
}

/**
 * Función que permite obtener el listado de clasificación del sistema
**/
function getRatingList(spreadsheet) {

  // obtenemos la hoja de segun parametro
  var stateSheet = spreadsheet.getSheetByName("Clasificación");
  
  // validamos si existe la hoja
  if(stateSheet){
    
    // obtenemos la ultima fila
    var lastRow = stateSheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
      
      // retornamos los datos
      var allData = stateSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      
      // variable para almacenar los datos de las configuraciones
      var settingsData = [];
      
      // recorremos cada uno de los registros
      for(var i = 0; i < allData.length; i++){
        
        // agregamos los datos en el objeto para el campo de busqueda
        settingsData.push(allData[i][0]);
        
      }
      
      // retornamos los datos
      return settingsData;
      
    }
  }
  
  // retornamos por defecto un array vacio
  return [];
}

/**
 * Función que permite obtener el listado de las configuraciones segun parametro
**/
function getSettingList(sheetName, compareValue, clasificationValue, systemTypeValue, formEmissionValue, feautureValue, range) {
  
  // validamos si la hoja es "Líneas de productos"
  if(sheetName == "Líneas de productos"){
    
    // retornamos los datos vacios
    return getInfoProductLine(compareValue, clasificationValue, systemTypeValue, formEmissionValue, feautureValue, range);
    
  }
  
  // Obtenemos la cache
  var scriptCache = CacheService.getScriptCache();
  
  // referenciamos la propiedad segun el nombre de la hoja
  var sheetNameCache = scriptCache.get(sheetName + "_Cache"),
      allData = [];
  
  // validamos si es null
  if(sheetNameCache == null){
    
    // se obtiene los recursos de hojas de cálculo
    var resources = getResources();
  
    // referenciamos la hoja de cálculo de configuración
    var spreadsheet = SpreadsheetApp.openById(resources.settingSheetId);
    
    // obtenemos la hoja de segun parametro
    var sheet = spreadsheet.getSheetByName(sheetName);
    
    // validamos si existe la hoja
    if(sheet){
      
      // obtenemos la ultima fila
      var lastRow = sheet.getLastRow();
      
      // validamos que exista mas de 1 fila
      if(lastRow > 1){
        
        //Definimos la cantidad de columnas
        var columnsCounter = 2;
        
        // se valida si la pestaña es Orden
        if(sheetName == "Orden"){
          
          // colocamos la cantidad de columnas a 3
          columnsCounter = 3;
          
        }
        
        // retornamos los datos
        allData = sheet.getRange(2, 1, lastRow - 1, columnsCounter).getValues();
        
        // Guardamos los datos de la variable en cache
        scriptCache.put(sheetName + "_Cache", JSON.stringify(allData), 21600);
        
      }
    }
    
  } else {
      
     // obtenemos los datos almacenados en cache
     allData = JSON.parse(sheetNameCache);
     
  }
  
  // variable para almacenar los datos de las configuraciones
  var settingsData = [];
  
  // recorremos cada uno de los registros
  for(var i = 0; i < allData.length; i++){
    
    // validamos si el valor de la relación es igual al valor de comparación
    if(allData[i][1] == compareValue){
      
      // se valida si la pestaña es Orden
      if(sheetName == "Orden"){
      
        // agregamos los datos en el objeto para el campo de busqueda
        settingsData.push({
          name: allData[i][0],
          range: allData[i][2]
        });
        
      } else {
        
        // agregamos los datos en el objeto para el campo de busqueda
        settingsData.push(allData[i][0]);
        
      }
    }
    
  }
  
  // retornamos por defecto un array vacio
  return settingsData;
  
}

/**
 * Funciónn que permite obtene los datos de las líneas de producto
**/
function getInfoProductLine(compareValue, clasificationValue, systemTypeValue, formEmissionValue, feautureValue, range){
  
  
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // referenciamos la hoja de cálculo de configuración
  var spreadsheet = SpreadsheetApp.openById(resources.settingSheetId);
  
  // obtenemos la hoja de segun parametro
  var sheet = spreadsheet.getSheetByName("Líneas de productos");
  
  // validamos si existe la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
    
      // retornamos los datos vacios
      return {
        arrayData: orderingProductLines(sheet, compareValue, lastRow),
        defaultLines: getDefaultLinesByOrder(spreadsheet, compareValue, clasificationValue, systemTypeValue, formEmissionValue, feautureValue, range)
      };
    
    }
  }
  
  // retornamos los datos vacios
  return {
    arrayData: [],
    defaultLines: []
  };
  
}

/**
 * Función que permite ordenar las líneas de producto segun la configuración del usuario
**/
function orderingProductLines(sheet, compareValue, lastRow){
  
  // obtenemos el correo del usuario activo
  var userMail = getUserMail();
  
  // retornamos los datos
  var arrayData = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues(),
    columnIndex = 2;
  
  // validamos si existe algun dato
  if(arrayData.length > 0){
    
    // validamos que exista mas de 3 columnas
    if(arrayData[0].length > 3){
      
      // Recorremos la primera columna y validamos si el usuario esta en alguna de las columnas
      for(var col = 3; col < arrayData[0].length; col++){
        
        // validamos si el correo pertenece a uno de las columnas
        if(userMail == arrayData[0][col]){
          
          // definimos en que columna se encuentra (TIPO MATRIZ 0-n)
          columnIndex = col;
          
          // salimos del ciclo
          break;
          
        }
      }
    }

    // variable para almacenar los datos de las configuraciones
    var newArray = [],
        returnArray = [];
    
    // recorremos cada uno de los registros
    for(var i = 0; i < arrayData.length; i++){
      
      // validamos si el valor de la relación es igual al valor de comparación
      if(arrayData[i][1] == compareValue){
        
        // agregamos los datos en el objeto para el campo de busqueda
        newArray.push({
          name: arrayData[i][0],
          index: Number(arrayData[i][columnIndex])
        });
        
        // agregamos un valor vacio
        returnArray.push("");
        
      }
    }
    
    // recorremos el nuevo array y lo ordenamos
    newArray.forEach(function(objData){
      
      // agregamos el valor en la posición especifica
      returnArray[objData.index - 1] = objData.name;
      
    });
    
    // retornamos el array
    return returnArray;    
  }
  
  // retornamos un array vacio
  return [];
  
}

/**
 * Función que permite obtener el listado de los cultivos
**/
function getCultivationList(spreadsheet) {
  
  // obtenemos la hoja de "Cultivo"
  var stateSheet = spreadsheet.getSheetByName("Cultivo");
  
  // validamos si existe la hoja
  if(stateSheet){
    
    // obtenemos la ultima fila
    var lastRow = stateSheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
      
      // retornamos los datos de los cultivos
      var allData = stateSheet.getRange(2, 1, lastRow - 1, 3).getValues();
      
      // variable para almacenar los datos de las configuraciones
      var cultiveData = [];
      
      // recorremos cada uno de los registros
      for(var i = 0; i < allData.length; i++){
        
        // agregamos los datos en el objeto para el campo de busqueda
        cultiveData.push(allData[i][2] == "SIN CLASIFICAR" ? allData[i][1] : (allData[i][1] + " - " + allData[i][2]));
        
      }
      
      // retornamos los datos
      return cultiveData;
      
    }
  }
  
  // retornamos por defecto un array vacio
  return [];
}

/**
 * Función que permite obtener el listado de clasificación del sistema
**/
function getProjectClassList(spreadsheet) {
  
  // obtenemos la hoja de segun parametro
  var stateSheet = spreadsheet.getSheetByName("Clase del proyecto");
  
  // validamos si existe la hoja
  if(stateSheet){
    
    // obtenemos la ultima fila
    var lastRow = stateSheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
      
      // retornamos los datos
      var allData = stateSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      
      // variable para almacenar los datos de las clase del proyecto
      var projectClassData = [];
      
      // recorremos cada uno de los registros
      for(var i = 0; i < allData.length; i++){
        
        // agregamos los datos en el objeto para el campo de busqueda
        projectClassData.push(allData[i][0]);
        
      }
      
      // retornamos los datos
      return projectClassData;
      
    }
  }
  
  // retornamos por defecto un array vacio
  return [];
}

/**
 * Función que permite obtener todos las especies
**/
function getConceptList(spreadsheet, settingSheetId) {

  // validamos si NO existe el id del archivo
  if(!spreadsheet && !settingSheetId){
    
    // obtenemos el id de la función global de recursos
    settingSheetId = getResources().settingSheetId;
    
  }
  
  // validamos si no viene información del libro
  if(!spreadsheet){
  
    // referenciamos la hoja de cálculo de estados y localidades
    spreadsheet = SpreadsheetApp.openById(settingSheetId);
  }
  
  // obtenemos la hoja de segun parametro
  var sheet = spreadsheet.getSheetByName("Conceptos");
  
  // validamos si existe la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
    
      // retornamos los datos
      return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
      
    }
  }
  
  // retornamos por defecto un array vacio
  return [];
}
