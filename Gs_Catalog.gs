/**
 * Función que permite almacenar la información del catálogo asociado a una especie
**/
function saveInfoCatalog(beforeData, allCatalogData, projectData){

  // se valida que exista el archivo donde se debe almacenar la información del catálogo
  if(projectData.infoCatalogSheetId){
  
    // se referencia la hoja donde se va almacenar
    //se referencia la hoja de calculo donde se va almacenar la información del catálogo
    var spreadsheet = SpreadsheetApp.openById(projectData.infoCatalogSheetId);
    
    // se obtiene el nombre de la especie y la linea
    var lineName = beforeData.lineProductName,
        currentConceptName = beforeData.conceptObj.conceptData[1];
    
    // definimos un objeto donde se especifica el nombre de la hoja que debe referenciar
    var sheetObject = {
      lading: "Fletes",
      teams: "Equipos",
      materials: "Materiales",
      auxiliaries: "Auxiliares",
      workforce: "Mano de obra",
      subcontracts: "Subcontratos",
      tools: "Herramientas"
    };
    
    // se definen las variables que se usan en el siguiente ciclo
    var sheetName = "", sheet = null, lastRow = 0;
    
    // Validamos de que existan la propieda de costos
    if(!beforeData.conceptObj.cost){
    
      // Inicializamos el objeto de los costos
      beforeData.conceptObj.cost = {
        auxiliaries: {},
        lading: {}, // Fletes 
        materials: {},
        subcontracts: {},
        teams: {},
        tools: {},
        workforce: {}
      };
    }
    
    // recorremos cada una de las propiedad del los datos del catálogo
    for(var key in allCatalogData){
      
      // se obtiene el nombre de la hoja
      sheetName = sheetObject[key];
      
      // se valida que exista el nombre de la hoja
      if(sheetName){
        
        // obtenemos la hoja
        sheet = spreadsheet.getSheetByName(sheetName);
        
        // se valida que la pestaña exista
        if(sheet){
        
          // eliminamos las filas
          deleteRowByLineOrConcept(sheet, lineName, currentConceptName);
        
          // agregamos la propiedad vacia
          beforeData.conceptObj.cost[key] = [];
          
          // se valida si existen nuevos registros para almacenar
          if(allCatalogData[key].records.length > 0){
            
            // obtenemos la ultima fila + 1
            lastRow = sheet.getLastRow() + 1;
            
            // agregamos los nuevos registros
            sheet.getRange(lastRow, 1, allCatalogData[key].records.length, allCatalogData[key].records[0].length).setNumberFormat("@STRING@").setValues(allCatalogData[key].records);
            
            // recorremos cada unod de los datos
            beforeData.conceptObj.cost[key] = allCatalogData[key];
          }
        
          // validamos si vienen los datos generales 
          if(allCatalogData[key].generalParameters != null){
            
            // agregamos Los datos generales 
            sheet.getRange("B1").setNumberFormat("@STRING@").setValue(jsonToString(allCatalogData[key].generalParameters));
            
          }
          
        }
      }
    }
    
    // referenciamos la hoja de "Conceptos"
    var conceptSheet = spreadsheet.getSheetByName("Conceptos");
    
    // validamos si existe
    if(conceptSheet){
      
      // eliminamos el registro actual sobre el material almacenado
      deleteRowByLineOrConcept(conceptSheet, lineName, currentConceptName);
      
      // agregamos los nuevos registros
      conceptSheet.getRange(conceptSheet.getLastRow() + 1, 1, 1, 2).setNumberFormat("@STRING@").setValues([[lineName, currentConceptName]]);
    }
   
    // retornamos el objeto actualizado
    return beforeData;
  
  } else {
    
    // Se muestra un error en la interfaz 
    throw "Lo sentimos no hemos encontrado la hoja de cálculo para almacenar la información del catálogo asociada a la especie seleccionada.";
  }
}

/**
 * Permite convertir texto en objeto manipulable
**/
function jsonToString(value){
  
  // variable para guardra el valor
  var stringData = "";
  
  // encerramos en try-cath para evitar errores al convertir el texto en objeto manipulable
  try {
                    
    // se convierte el texto en objeto manipulable
    stringData = JSON.stringify(value);
    
  } catch(e){
    
    // inicializamos la variable en null
    stringData = "";
  }
  
  // Retornamos el objeto
  return stringData;
  
}

/**
 * Función que permite eliminar las filas de una determinada hoja segun su linea y especies
**/
function deleteRowByLineOrConcept(sheet, lineName, currentConceptName){
  
  // Variable que se usan para eliminar las filas
  var lastRow = 0, arrayData = [];
  
  // obtenemos la ultima fila diligenciada
  lastRow = sheet.getLastRow();
  
  // se valida si existe mas de 1 fila
  if(lastRow > 1){
   
    // obtenemos los datos de la hoja
    arrayData = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    
    // recorremos cada uno de los registros
    for(var i = (arrayData.length - 1); i >= 0; i--){
      
      // se valida si la linea y la especie recorrido son lo mismos que a los que se actualiza en el momento
      if(arrayData[i][0] == lineName && arrayData[i][1] == currentConceptName){
        
        // removemos la fila
        sheet.deleteRow(i + 2);  
        
      }
    }
  }
  
}

/**
 * Función que obtiene los datos por defecto para el catálogo
**/
function getGeneralDataFromCatalog(){
  
  // obtenemos los datos globales
  var resources = getResources();
  
  // Referenciamos a la hoja de calculo 
  var spreadsheet = SpreadsheetApp.openById(resources.workParameters);
  
  // se define el objeto que se desea retornar
  var returnObject = {};
  
  // obtenemos los datos de la hoja de mano de obra
  returnObject.manpowerData = getDataFromSettingCatalog(spreadsheet, "Mano de obra", 9, {key: 0, description: 1, unit: 2, unitCost: 5, participation: 8});
  
  // obtenemos los datos de la hoja de Herramientas
  returnObject.kitData = getDataFromSettingCatalog(spreadsheet, "Herramientas", 6, {key: 0, description: 1, unit: 2, quantity: 3, unitCost: 4, participation: 3});
  
  // obtenemos los datos de la hoja de Equipos
  returnObject.equipmentData = getDataFromSettingCatalog(spreadsheet, "Equipos", 9, {key: 0, description: 1, unit: 3, quantity: 5, unitCost: 4, participation: 5, type: [6, 7, 8]});
  
  // obtenemos los datos de la hoja de Auxiliares
  returnObject.auxiliaryData = getDataFromSettingCatalog(spreadsheet, "Auxiliares", 4, {key: 0, description: 1, unit: 2, unitCost: 3});
  
  // obtenemos los datos de la hoja de fletes
  returnObject.lading = getDataFromSettingCatalog(spreadsheet, "Fletes", 4, {key: 0, description: 1, unit: 2, unitCost: 3});
  
  // Obtenemos la lista de unidades de medida
  returnObject.unitList = getDataMeasureUnits(resources, spreadsheet);
  
  // Obtenemos la lista de unidades de medida de la estimación 
  returnObject.unitListUME = getDataUnitsUME(resources, spreadsheet);
  
  returnObject.vehicles = getVehicles(resources, spreadsheet);

  // retornamos los datos
  return returnObject;
  
}


/*
* permite obtener los vehiculos  
*/
function getVehicles(resources, spreadsheet){

  // Validamos si no vienen los recursos 
  if(!resources){
    // obtenemos los datos globales
    resources = getResources();
  }
  
  // validamos si no viene la hoja 
  if(!spreadsheet){
  
    // Referenciamos a la hoja de calculo 
    spreadsheet = SpreadsheetApp.openById(resources.workParameters);
  
  }
  
  // Referenciamos la hoja de vehiculos para obtener los datos 
  var sheet = spreadsheet.getSheetByName("Vehículos");
  
  // datos para retornar 
  var resultArrayData = [];
  
  // se valida que exista la hoja 
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
      
      // Obtenmos los datos
      var arrayData = sheet.getRange(3, 1, lastRow-2, 20).getValues();
      
      // Recorremos los datos 
      for(var i = 0; i < arrayData.length; i++){
        
        // validamos de que se encuentre el elemento 
        if(arrayData[i][0]){
            
          // Agregamos el elemento 
          resultArrayData.push({
             ladingId: arrayData[i][0] ,
             middle: arrayData[i][1], // medio
             mark: arrayData[i][2], // Marca
             model: arrayData[i][3], // Modelo
             year: arrayData[i][4], // Año 
             version: arrayData[i][4], // Version
             carDetails: arrayData[i][6], // Detalles del carroo
             trasnmission: arrayData[i][7], // Tramisión
             fuel: arrayData[i][8], // Combustibes
             capFuel: arrayData[i][9], // Capacidad de combustible
             cylinders: arrayData[i][10], // Cilindros
             power: arrayData[i][11], // Potencia
             size: arrayData[i][12], // tamaño
             unit: arrayData[i][13], // Unidad
             city: arrayData[i][14], // Ciudad
             way: arrayData[i][15], // Carretera
             combined: arrayData[i][16], // Combinado
             tight: arrayData[i][17], // Ajustado
             lKM: arrayData[i][18], // L/km
             
          });
        }
      }
    }
  }

  // Retornamos los datos 
  return resultArrayData;
  
}


/*
* Permite obtener el listado de unidadesde medida 
*/
function getDataUnitsUME(resources, spreadsheet){
  
  // Validamos si no vienen los recursos 
  if(!resources){
    // obtenemos los datos globales
    resources = getResources();
  }
  
  // validamos si no viene la hoja 
  if(!spreadsheet){
  
    // Referenciamos a la hoja de calculo 
    spreadsheet = SpreadsheetApp.openById(resources.workParameters);
  
  }
  
  // Refereciamos la pestaña de unidades de medida
  var sheet = spreadsheet.getSheetByName("Constantes");
  
  // datos para retornar 
  var resultArrayData = [];
  
  // Objeto con los datos 
  var resultObjData = {};
  
  // Datos temporales del item 
  var tempData = {};
  
  // se valida que exista la hoja 
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
      
      // Obtenmos los datos
      var arrayData = sheet.getRange(3, 1, lastRow-2, 4).getValues();
      
      // Recorremos los datos 
      for(var i = 0; i < arrayData.length; i++){
        
        // validamos de que se encuentre el elemento 
        if(arrayData[i][0]){
          
          // Obtenemos los datos del item 
          tempData = {
             name: (arrayData[i][0] + " (" +arrayData[i][1] + ")"),
             nameEquipment: (arrayData[i][0] + " (" +arrayData[i][3] + ")"),
             fuelName: arrayData[i][0], 
             cmeUnit: arrayData[i][1],
             kuml: arrayData[i][2],
             kumlUnit: arrayData[i][3]
          }
          
          // Agregamos los datos del item 
          resultObjData[arrayData[i][0]]= tempData; 
          
          // Agregamos el elemento 
          resultArrayData.push(tempData);
        }
      }
    }
  }

  // Retornamos los datos 
  return {
    resultArrayData: resultArrayData,
    resultObjData: resultObjData
  };
}



/*
*Permite obtener el listado de unidadesde medida 
*/
function getDataMeasureUnits(resources, spreadsheet){
  
  // Validamos si no vienen los recursos 
  if(!resources){
    // obtenemos los datos globales
    resources = getResources();
  }
  
  // validamos si no viene la hoja 
  if(!spreadsheet){
  
    // Referenciamos a la hoja de calculo 
    spreadsheet = SpreadsheetApp.openById(resources.workParameters);
  
  }
  
  // Refereciamos la pestaña de unidades de medida
  var sheet = spreadsheet.getSheetByName("Unidades de medida");
  
  // datos para retornar 
  var resultArrayData = [];
  
  // se valida que exista la hoja 
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
      
      // Obtenmos los datos
      var arrayData = sheet.getRange(2, 1, lastRow-1, 1).getValues();
      
      // Recorremos los datos 
      for(var i = 0; i < arrayData.length; i++){
        
        // validamos de que se encuentre el elemento 
        if(arrayData[i][0]){
            
          // Agregamos el elemento 
          resultArrayData.push({
             name: arrayData[i][0]
          });
         
        }
      }
    }
  }

  // Retornamos los datos 
  return resultArrayData;
}



/**
 * Función que obtiene los datos generales para el catálogo
**/
function getDataFromSettingCatalog(spreadsheet, sheetName, columns, referenceColumns){
  
  // referenciamos la hoja
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  // se valida que exista la hoja 
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
      
      // retornamos los datos de los clientes
      var arrayData = sheet.getRange(2, 1, lastRow - 1, columns).getValues();
      
      // variable para almacenar los clientes
      var returnData = {
        data: {},
        allData: {}
      };
      
      // variable para refernciar el registro actual
      var record = null,
          tempObject = {};
      
      // recorremos cada uno de los registros y los filtramos por usuario
      for(var i = 0; i < arrayData.length; i++){
        
        // se refencia el registro actual
        record = arrayData[i];
        
        // validamos que exista una descripción
        if(record[referenceColumns.description]){
        
          // agregamos los datos en el objeto para el campo de busqueda
          returnData.data[record[referenceColumns.description]] = null;
          
          // se define el objeto temporal
          tempObject = {
            key: record[referenceColumns.key],
            description: record[referenceColumns.description],
            unit: record[referenceColumns.unit],
            unitCost: record[referenceColumns.unitCost],
            quantity: (referenceColumns.quantity ? record[referenceColumns.quantity] : 1),
            participation: (referenceColumns.participation ? record[referenceColumns.participation] : "")
          };
          
          // validamos si existe la opción "type"
          if(referenceColumns.type){
            
            // se define un array temporal
            var tempArray = [];
            
            // realizamos un ciclo de acuerdo a la cantidad de tipo que se debe obtener
            for(var j = 0; j < referenceColumns.type.length; j++){
              
              // agregamos cada uno de los datos
              tempArray.push(record[referenceColumns.type[j]]);
              
            }
            
            // creamos la propiedad en el objeto temporal
            tempObject.type = tempArray;
          }
          
          // agregamos los datos en el objeto general
          returnData.allData[record[referenceColumns.description]] = tempObject;
        }
      }
      
      // retornamos los datos
      return returnData;
    }
  }
  
  // se retorna null por defecto
  return null;
  
}


/*
* Función que permtie traer los costos del proyecto 
* @parms {string} idProject id del proyecto 
*/
function getCostsConcepts(idProject) {

  // Información de los costos registrados por cada uno de las especies 
  var costObject = {};
 
  // Referenciamos a la hoja de calculo 
  var spreadsheetProject = SpreadsheetApp.openById(idProject);
  
  // Validamos de que se tenga acceso a la hoja 
  if(spreadsheetProject){
    
    // Referenciamos la hoja de conceptos 
    var sheetConcepts = spreadsheetProject.getSheetByName("Conceptos");
    
    // Validamos el acceso a la hoja de conceptos 
    if(sheetConcepts){
      
      // obtenemos la ultima fila seleccionada
      var lastRowConcepts = sheetConcepts.getLastRow();
      
      // Validamos de que se tenga fila de conceptos 
      if(lastRowConcepts > 1){
      
        // Obtenemos los  datos de los conceptos
        var arrayConcepts = sheetConcepts.getRange(2, 1, lastRowConcepts - 1, 2).getDisplayValues();
        
        // Recorremos las especies 
        for(var i = 0; i < arrayConcepts.length; i++){
          
          // Validamos si no  existe elemento en le  
          if( !costObject[arrayConcepts[i][0]]){
            
            // Creamos la propiedad del elmento 
            costObject[arrayConcepts[i][0]] = {};
          }
          
          // Validamos que tenga especies
          if (arrayConcepts[i][1]){
          
            // Agregamos la especie al array 
            costObject[arrayConcepts[i][0]][arrayConcepts[i][1]] = {
              materials: {recordsObj: [], generalParameters: null},
              workforce: {recordsObj: [], generalParameters: null},
              teams: {recordsObj: [], generalParameters: null},
              tools: {recordsObj: [], generalParameters: null},
              auxiliaries: {recordsObj: [], generalParameters: null},
              lading: {recordsObj: [], generalParameters: null},
              subcontracts: {recordsObj: [], generalParameters: null}
             } ;
          }
        }
        
        // Obtenemos los datos de los materiales
        costObject = getDataCosts(spreadsheetProject, "Materiales", "materials", costObject, 3, 9);

        // Obtenemos los datos de la mano de obra 
        costObject = getDataCosts(spreadsheetProject, "Mano de obra", "workforce", costObject, 3, 10);
        
        // Obtenemos los datos de los Equipos
        costObject = getDataCosts(spreadsheetProject, "Equipos", "teams", costObject, 2, 10);

        // Obtenemos los datos de los fletes
        costObject = getDataCosts(spreadsheetProject, "Fletes", "lading", costObject, 3, 8);
        
        // Obtenemos los datos de las Herramientas
        costObject = getDataCosts(spreadsheetProject, "Herramientas", "tools", costObject, 2, 8);
        
        // Obtenemos los datos de los auxiliares
        costObject = getDataCosts(spreadsheetProject, "Auxiliares", "auxiliaries", costObject, 2, 8);
        
        // Obtenemos los datos de los subcontratos
        costObject = getDataCosts(spreadsheetProject, "Subcontratos", "subcontracts", costObject, 2, 8);
        
      }
    }
  }

  // Retornamos los costos de las especies 
  return costObject;
  
}


/*
* Función que permite obtener los datos de materiales o mano de obra o Equipos o Herramientas o auxiliares o fletes o subcontratos 
* @param {sheet}: spreadsheetProject: Referecia a la hoja de los datos del proyecto
* @param {string}: sheetByName: Nombre de la hoja donde se van a consultar los datos 
* @param {string}: property: nombre de la propiedad donde se deben guardar los datos en el objeto
* @param {string}: conceptsCostsObjt: Objeto con los datos de las especies del proyecto 
*/
function getDataCosts(spreadsheetProject, sheetByName, property, costObject, startRow, columnsCount){

  // Referenciamos la hoja 
  var sheet = spreadsheetProject.getSheetByName(sheetByName);
  
  // Validamos el acceso a la hoja 
  if(sheet){
    
    // obtenemos la ultima fila seleccionada
    var lastRow = sheet.getLastRow();
    
    // Validamos de que se tenga fila 
    if(lastRow > 1){
      
      // Obtenemos los  datos
      var arrayData = sheet.getRange(startRow, 1, lastRow - 1, columnsCount).getDisplayValues();

      // Datos del item 
      var objectData = {};
      
      // variable para guardar los datos generales
      var generalParameters = null;
      
      // se valida si la fila de inicio es 3
      if(startRow == 3){
       
        // agregamos la información de los datos generales
        generalParameters = sheet.getRange("B1").getValues();
   
        // validamos si existen datos
        if(generalParameters){
            
          // se convierte el texto en objeto manipulable
          generalParameters = stringToObjectJson(generalParameters);
            
        }
      }
      
      // Recorremos los datos de la hoja requerida
      for(var x = 0; x < arrayData.length; x++){
      
        // validamos de que vengan los costos 
        if(costObject){
          
          /// Recoremos las lineas
          for (var objectKey in costObject){
        
            // Validamos si es la linea
            if(costObject[arrayData[x][0]] && costObject[arrayData[x][0]] == costObject[objectKey]){
        
              //validamos de que la linea venga con especies 
              for (var objeckeyConcept in costObject[objectKey]){
            
                // Validamos por el nombre de la especie 
                if(costObject[arrayData[x][0]][arrayData[x][1]] && costObject[arrayData[x][0]][arrayData[x][1]] == costObject[objectKey][objeckeyConcept]){
                 
                  // se valida si la fila de inicio es 3
                  if(startRow == 3 && !costObject[objectKey][objeckeyConcept][property].generalParameters && generalParameters){
                   
                    // agregamos la información de los datos generales
                    costObject[objectKey][objeckeyConcept][property].generalParameters = generalParameters;
                    
                  }
                  
                  // Creamos el objeto con los datos 
                  objectData = {
                    key: arrayData[x][2],
                    description: arrayData[x][3],
                    unity: arrayData[x][4],
                    quantity: arrayData[x][5],
                    unitCost: arrayData[x][6],
                    total: arrayData[x][7],
                    rowIndex: (x + 2)
                  }
                 
                  // Validamos si es de mano de obra para obtener la participación 
                  if(property == "workforce"){
                
                    // Agregamos la participación 
                    objectData.participation = arrayData[x][8];
                    
                    // agregamos la información de la propiedad de avance
                    objectData.avanceData = stringToObjectJson(arrayData[x][9]);
                                      
                  } else if(property == "materials"){ // Si es materiales
                    
                    // agregamos la información de la propiedad de avance
                    objectData.avanceData = stringToObjectJson(arrayData[x][8]);
                      
                  } else if(property == "teams"){ // Si es equipos
                  
                    // Agregamos los datos de fletes
                    objectData.arrayLading = stringToObjectJson(arrayData[x][9]);
                    
                    // agregamos la información de la propiedad de avance
                    objectData.avanceData = stringToObjectJson(arrayData[x][8]);
                   
                  }
                  
                  
                  // Agregamos los datos registrados
                  costObject[objectKey][objeckeyConcept][property].recordsObj.push(objectData);
                }
              }
            }
          }
        }
      
      }
    }
  }
  
  // Retornamos los datos de la especie
  return costObject;
}

/**
 * Permite convertir texto en objeto manipulable
**/
function stringToObjectJson(value){
  
  // variable para guardra el valor
  var jsonData = null;
  
  // encerramos en try-cath para evitar errores al convertir el texto en objeto manipulable
  try {
                    
    // se convierte el texto en objeto manipulable
    jsonData = JSON.parse(value);
    
  } catch(e){
    
    // inicializamos la variable en null
    jsonData = null;
  }
  
  // Retornamos el objeto
  return jsonData;
  
}

/*
* Función que permite obtener la mano de obra para cada lines registrada
*/
function getListCostsConcepts(orderProjects, projectData, userRole, year) {

  var returnObject = { 
    orderProjects: orderProjects,
    performanceData: {}
  };

  // Validamos de que vengan ordenes
  if(orderProjects && orderProjects.length > 0){
   
    // variable que determinar si es necesario consultar información del catálogo 
    var isConsult = true;
    
    // Referenciamos los recursos globales
    var resources = getResources();

    // validamos si no viene la hoja para la información del catalogo
    if(projectData && !projectData.infoCatalogSheetId){
      
      // Creamaos la hoja para la información del catálogo 
      projectData.infoCatalogSheetId = createInfoCatalogSheet(projectData, resources);
      
      // Establecemos de que no se debe consultar
      isConsult = false;
      
      // se obtiene el correo del usuario
      var userMail = getUserMail();
      
      // Obtenemos las solicitudes
      returnObject.requestData = filterRequestByMail(userMail, resources.requestSheetId, userRole, year);
      
      // Retornamos los datos actualizados del proyectos
      returnObject.projectData = projectData;
    }
    
    var costosConcepts = {};
    
    // Validamos si requiere consulta 
    if(isConsult){
 
      // Consultamos los datos registrados
      costosConcepts = getCostsConcepts(projectData.infoCatalogSheetId);
    }
    
    // Obtenemos los datos de rendimiento por concepto
    returnObject.performanceData = getPerformanceByConcept(resources);
    
    var conceptName = "";
    var lineName = {};
    
    // Obtenemos la mano de obra 
    var laborByLine = getLaborByLine(resources);
    
    // Recorremos cada una de las ordenes
    for(var i = 0; i < orderProjects.length; i++){
      
      // recorremos cada una de las líneas de productos
      for(var j = 0; j < orderProjects[i].linesProducts.length; j++){
       
         // Agregmaos 
         lineName = orderProjects[i].linesProducts[j].lineProductName;
         
        // Insertamos la manos de obra 
        orderProjects[i].linesProducts[j].workforceDefault = laborByLine[lineName]; 
         
        // Validamos si requiere agregamos los datos almacenados
        if(isConsult){
          
          // Validamos de que vengan especies
          if(orderProjects[i].linesProducts[j].conceptList && orderProjects[i].linesProducts[j].conceptList.length > 0){
          
            // Recorremos las especies 
            for(var z = 0; z < orderProjects[i].linesProducts[j].conceptList.length; z++){
              
              // Agregamos el nombre de la especie
              conceptName = orderProjects[i].linesProducts[j].conceptList[z].conceptData[1];

              /// Validamos de que esxista la linea y la especie 
              if(costosConcepts[lineName] && costosConcepts[lineName][conceptName]){
               
                // Agregamos los datos de los costos del proyecto 
                orderProjects[i].linesProducts[j].conceptList[z].cost = costosConcepts[lineName][conceptName];
              
              }
            }
          }
        }
      } 
    }
  }
//  throw JSON.stringify(projectData)
  
  // Retornamos los valores de la solcitu d
  return returnObject;
}

/**
 * Función que permite obtener el rendimiento de cada concepto
**/
function getPerformanceByConcept(resources){
  
  // obtenemos la lista de conceptos
  var conceptList = getConceptList(null, resources.settingSheetId);
  
  // definimos el objeto que vamos a retornar
  var returnObject = {};
  
  // recorremos cada uno de los conceptos
  for(var i = 0; i < conceptList.length; i++){
    
    // se valida si no existe la propiedad de departamento
    if(!returnObject[conceptList[i][0]]){
      returnObject[conceptList[i][0]] = {};
    }
    
    // se valida si no existe la propiedad de familia
    if(!returnObject[conceptList[i][0]][conceptList[i][2]]){
      returnObject[conceptList[i][0]][conceptList[i][2]] = {};
    }
    
    // se valida si no existe la propiedad de palabra clave
    if(!returnObject[conceptList[i][0]][conceptList[i][2]][conceptList[i][3]]){
      returnObject[conceptList[i][0]][conceptList[i][2]][conceptList[i][3]] = {};
    }
    
    // se valida si no existe la propiedad de Nombre del concepto
    if(!returnObject[conceptList[i][0]][conceptList[i][2]][conceptList[i][3]][conceptList[i][1]]){
      returnObject[conceptList[i][0]][conceptList[i][2]][conceptList[i][3]][conceptList[i][1]] = (Number(conceptList[i][33]) || 0);
    }
    
  }
  
  // retornamos el onjeto
  return returnObject;
}

/**
 * Función que crear el archivo donde se almacena la información del catálogo
**/
function createInfoCatalogSheet(projectData, resources){
  
  // referenciamos la hoja de cálculo general donde quedan todos los datos de las cotizaciones generadas
  var spreadsheet = SpreadsheetApp.openById(resources.requestSheetId);
  
  // Obtenemos los padres de la hoja la cotización
  var parents = DriveApp.getFileById(projectData.documentId).getParents();
  
  // se define la variable para la referencia del padre
  var mainFolder = null;
  
  // se valida si tiene padres
  if(parents.hasNext()){
    
    // se referencia el primer padre
    mainFolder = parents.next();
    
  } else {
    
    // referenciamos la carpeta raiz
    mainFolder = DriveApp.getFolderById(resources.forderDefaultId);
    
  }
  
  // Sacamos una copia de la plantilla de la hoja de catálogo de cuenta
  var infoCatalogSheet  = DriveApp.getFileById(resources.infoCatalogSheetId).makeCopy("Información del catálogo de especies para el proyecto " + projectData.id, mainFolder);
  
  // obtenemos el Id del archivo de información del catálogo
  var infoCatalogSheetId = infoCatalogSheet.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT).getId();
  
  // obtenemos la hoja de "Cotizaciones"
  var sheet = spreadsheet.getSheetByName("Cotizaciones");
  
  // validamos si existe la hoja
  if(sheet){
    
    // referenciamos la columna P donde se almacena la información del archivo donde se almacane la info de catálogo
    sheet.getRange("P" + projectData.row).setNumberFormat("@STRING@").setValue(infoCatalogSheetId);
    
  }
  
  // retornamos el id del archivo
  return infoCatalogSheetId;
}


/*
* Función que permite traer la mano de obra por linea
*/
function getLaborByLine(resources){

  // referenciamos El archivo de parametros de obra 
  var spreadsheet = SpreadsheetApp.openById(resources.workParameters);
  
  // Validamos de se tenga acceso a la hoja de parametros de obra y se tengan los ordenes 
  if(spreadsheet){
    
    // Referenciamos a la hoja de mano de obra por linea
    var workParametersSheet = spreadsheet.getSheetByName("Mano de obra por linea");
    
    // validamos de que se tenga acceso a la hoja de mano de obra por linea
    if(workParametersSheet){
      
      // Obtenemos la cantidad de filas 
      var lastRow = workParametersSheet.getLastRow();
      
      // Objetos para guardar los datos por cada linea
      var objectLine = {},
          objectConcept = {},
          currentRow = [];
      
      // Validamos de que se tenga filas diligenciadas
      if(lastRow && lastRow > 3){
        
        // Obtenemos los datos 
        var arrayData = workParametersSheet.getRange(1, 1, lastRow, workParametersSheet.getLastColumn()).getDisplayValues();

        // Obtenemos los numerales de la lines
        var lineNumber = arrayData[0];
        
        // Obtenemos los nombres de las lines
        var lineName = arrayData[1];
        
        // Recoremos las lineas para crear el objeto 
        for(var z = 3; z < lineNumber.length; z++){
          
          //Agregamos la lines
          objectLine["E."+lineNumber[z] + " " + lineName[z]] = [];
        }
        
        // Recorremos cada uno de los registros
        for(var i = 3; i < arrayData.length; i++){
          
          // Obtenemos los datos fila 
          currentRow = arrayData[i];
          
          // Validamos de se tenga el nombre de la persona
          if(currentRow && (currentRow[0] || currentRow[1])){
            
            // Reiniamos los datos del objecto de la especie
            objectConcept = {
              position: currentRow[0],
              responsible: currentRow[1]
            };
            
            // Recorremos cada uno de los datos de la fila 
            for(var y = 3; y < currentRow.length; y++){
             
              // validamos si existe un dato 
              if(currentRow[y] && currentRow != ""){
                
                // Agregamos el porcentaje
                objectConcept.percent = Number(String(currentRow[y]).replace(/%/g, ""));
                
                // Agregamos la persona al objeto
                objectLine["E."+lineNumber[y] + " " + lineName[y]].push(JSON.parse(JSON.stringify(objectConcept)));
                
              }
            }
          }
        }
        
        // Retornamos la mano de obra 
        return objectLine;
      }
    }
  }
  
  // Retornamos vacio por si no se obtienen datos
  return false;
}


/*
* Función que permite obtener los valores por defecto de una especie 
* @param {sheet}: spreadSheet: Referecia a la hoja de parametros de obra 
* @param {string}: sheetByName: Nombre de la hoja donde se van a consultar los datos por defecto 
* @param {string}: property: nombre de la propiedad donde se deben guardar los datos en el objeto
* @param {object}: costsConcept: Objeto con los datos de los por defecto que tiene la especie para mano de obra
* @param {array}: arrayList: array con el listado de id de los  items para concultar e cada uno de los parametros de mano de obra 
*/
function getDataConcept(spreadSheet, sheetByName, property, costsConcept, arrayList ){
  
  // Referenciamos la hoja 
  var sheet = spreadSheet.getSheetByName(sheetByName);
  
  // Validamos el acceso a la hoja 
  if(sheet){
  
    // obtenemos la ultima fila seleccionada
    var lastRow = sheet.getLastRow();
    
    // Validamos de que se tenga fila 
    if(lastRow > 1){
      
      // Obtenemos los  datos
      var arrayData = sheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
      
      // Recorremos los datos
      for(var x = 0; x < arrayData.length; x++){
        
        // Validamos si no existe la propiedad
        if(!costsConcept[property]){
          
          // Creamos la proiedad en el array 
          costsConcept[property] = [];
          
        }
        
        // Validamos de  que se debe ingresar este parametro a la mano de obra del concpeto 
        if(arrayList.indexOf(arrayData[x][0]) != -1){
        
          //Agregamos los datos 
          costsConcept[property].push({
            key: arrayData[x][0],
            description: arrayData[x][1],
            unity: arrayData[x][2],
            unitCost: arrayData[x][3],
          });
        }
      }
    }
  }
  
  //Retornamos los costos por defecto del concpeto por esa propiedad
  return costsConcept;
}


/*
* Permite obtener el numero de catalogo
*/
function getConcessionNumber(infoCatalogSheetId){

   // Numero de la concesion
   var concessionNumber = "";
   
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
       
         // Obtenemos los datos 
         var concessionData = civilWork.getRange("A2").getDisplayValue();
         
         // Validamos de que se tenga los datos de la concession
         if(concessionData){
         
           // Adicionamos el numero de la concesión que se encuentra guardado 
           concessionNumber = concessionData;
           
         }
       }
     }
   }

   // Retornamos el numero de concesion obtenido 
   return concessionNumber;
 
}
