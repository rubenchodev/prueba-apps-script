/*
* Permite guardar la actuakziación de precio de la especie por parte de Hidro
*/
function saveUpdatePriceSpecieForHidroCompetitor(formData, manufacturerData){
  
  // variable a retornar
  var returnObject = {
    success: false,
    message: "La especie no puede ser actualizada, por favor intenta de nuevo."
  };
  
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources(), spreadSheet = null;
  
  try{
    // referenciamos la hoja solicitudes de proyectos
    spreadSheet = SpreadsheetApp.openById(resources.competitorsPricingSheetId);
  }catch(e){}
  
  // validamos de que se tenga acceso a la hoja 
  if(spreadSheet){
     
    // fecha actual del sistema
    var currentDate = new Date();
    var timeZone = Session.getScriptTimeZone();   
   
    // validamos si no existe el index, es decir que el fabricanto no esta registrado
    if(!manufacturerData.index){
       
      // fecha actual si caracteres especiales
      var currentDateTrim = Utilities.formatDate(currentDate, timeZone, "ddMMyyyyHHmmss");
       
      // Creamos el código de actualización
      manufacturerData.codePack = currentDateTrim + Utilities.getUuid();
      
      // Referenciamos la hoja de Fabricante/Solicitudes
       var mainSheet = spreadSheet.getSheetByName("Competidores/Solicitudes");
       
       // Validamos de que se tenga acceso a la hoja Precios/Fabricantes
       if(mainSheet){
         
         // Obtenemos la ultima fila
        var mainLastRow = mainSheet.getLastRow();
        
        //Actualizamos el registro en la hoja
        mainSheet.getRange(mainLastRow + 1, 1, 1, 3).setValues([[manufacturerData.manufacturerId, manufacturerData.codePack, "Pendiente"]]);
       }
     }
     
     // Referenciamos la hoja de Precios / Fabricantes
     var sheet = spreadSheet.getSheetByName("_Precios/Competidores");
     
     // Validamos de que se tenga acceso a la hoja Precios/Fabricantes
     if(sheet){
       
       // obtenemos la fila
       var rowIndex = manufacturerData.index;
       
       // validamos si es un registro nuevo
       if(manufacturerData.codePack){
         
         // obtenemos la fecha actual en "MM/dd/yyyy HH:mm:ss"
         var currentDateStr = Utilities.formatDate(currentDate, timeZone, "dd/MM/yyyy HH:mm:ss");
         
         // obtenemos la ultima fila y sumamos uno
         rowIndex = sheet.getLastRow() + 1;

         // Agregamos la información actual
         sheet.getRange(rowIndex, 1, 1, 6).setValues([[
           manufacturerData.manufacturerId,
           manufacturerData.codePack,
           manufacturerData.manufacturerName,
           manufacturerData.conceptName,
           manufacturerData.unit,
           currentDateStr
         ]]);
       }
      
       // se valida que exista la fila
       if(rowIndex){
         
         // Guardamos los tipos de cambios y el valor de unidad
         sheet.getRange(rowIndex, 10, 1, 3).setValues([[formData.dolar, formData.euro, formData.unit]]);
         sheet.getRange(rowIndex, 8).setValue(formData.productName);
         sheet.getRange(rowIndex, 16, 1, 5).setValues([[formData.currency, formData.discountApplied, formData.quantity, formData.suggestedPrice, formData.iva]]);
         sheet.getRange(rowIndex, 21).setValue("=S"+rowIndex+"*(T"+rowIndex+"/100 + 1)");
         sheet.getRange(rowIndex, 22).setValue(formData.total);
         sheet.getRange(rowIndex, 30).setNumberFormat("@STRING@").setValue(Utilities.formatDate(currentDate, timeZone, "dd/MM/yyyy HH:mm:ss"));
         
         // actualizamos el estado y el mensaje
         returnObject.success = true;
         returnObject.message = "El precio del concepto: <b>" + formData.productName + "</b> ha sido registrado correctamente.";
         returnObject.data = getAllDataCompetitors();
       }
     }
  }
  
  // retornamos el valor
  return returnObject;
}


/*
* Permite guardar los datos del tipo de cambio 
*/
function updateCurrencyTypeHexpCompetitors(currencyTypeData){
   
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // fecha actual del sistema
  var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
  
  // referenciamos la hoja de cálculo de estados y localidades
  //var spreadsheet = SpreadsheetApp.openById(getResources().settingSheetId);
  var spreadsheet = SpreadsheetApp.openById(getResources().competitorsPricingSheetId);
  
  // Referenciamos la hoja de tipo de cambio 
  var currencyTypeSheet = spreadsheet.getSheetByName("Tipo de cambio");
  
  // Obtenemos el correo del usuario activo 
  var activeUserMail = getUserMail();
  
  // Validamos de que se tenga acceso a al hoja de tipo de cambio 
  if(currencyTypeSheet){
    var lastRow = currencyTypeSheet.getLastRow() + 1;
    // referenciamos el rango y guardamos los datos 
    currencyTypeSheet.getRange(lastRow, 1, 1, 9).setNumberFormat("@STRING@").setValues([[
      '',
      currentDate, 
      //activeUserMail,
      currencyTypeData.usdFixHexp,
      currencyTypeData.usdHexp,
      currencyTypeData.usdHexp,
      0,
      0,
      currencyTypeData.eurHexp,
      0
    ]]);
    
    currencyTypeSheet.getRange(lastRow,6).setFormula("=MAX(C"+lastRow+":E"+lastRow+")");
    currencyTypeSheet.getRange(lastRow,7).setFormula("=F"+lastRow+"+0.001");
    currencyTypeSheet.getRange(lastRow,9).setFormula("=H"+lastRow+"+0.001");
     
    // Retornamos los datos de la ctualización
    return {
      currentDate: currentDate, 
      //activeUserMail: activeUserMail,
      usdHexp: currencyTypeData.usdHexp,
      eurHexp: currencyTypeData.eurHexp
    }
  }
    
}

/*
* permite consultar los datos de tipo de cambio actuales 
*/
function getCurrencyTypeDataCompetitors(){

  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // referenciamos la hoja de cálculo de estados y localidades
  //var spreadsheet = SpreadsheetApp.openById(getResources().settingSheetId);
  var spreadsheet = SpreadsheetApp.openById(getResources().competitorsPricingSheetId);
  
  
  // Referenciamos la hoja de tipo de cambio 
  var currencyTypeSheet = spreadsheet.getSheetByName("Tipo de cambio");
  
  // Validamos de que se tenga acceso a al hoja de tipo de cambio 
  if(currencyTypeSheet){
    var lastRow = currencyTypeSheet.getLastRow();
    // Consultamos los datos 
    var arrayData = currencyTypeSheet.getRange(3, 1, 1, 19).getValues();
    var ultimaFila = currencyTypeSheet.getRange(lastRow, 1, 1, 9).getValues();
    console.log('arrayData',{
       currentDate: ultimaFila[0][1] ? ultimaFila[0][1]: 0 , 
      //activeUserMail: arrayData[0][1] ? arrayData[0][1]: 0 ,
      usdHexp: ultimaFila[0][5] ? ultimaFila[0][5]: 0 ,
      eurHexp: ultimaFila[0][7] ? ultimaFila[0][7]: 0 
     })
    // Retornamos los datos 
    return {
      id: ultimaFila[0][0] ? ultimaFila[0][0]: 0 , 
      currentDate: ultimaFila[0][1] ? ultimaFila[0][1]: 0 , 
      //activeUserMail: arrayData[0][1] ? arrayData[0][1]: 0 ,
      usdHexp: ultimaFila[0][5] ? ultimaFila[0][5]: 0 ,
      eurHexp: ultimaFila[0][7] ? ultimaFila[0][7]: 0 
     }
    /*return {
      currentDate: arrayData[0][0] ? arrayData[0][0]: 0 , 
      //activeUserMail: arrayData[0][1] ? arrayData[0][1]: 0 ,
      usdHexp: arrayData[0][2] ? arrayData[0][2]: 0 ,
      eurHexp: arrayData[0][3] ? arrayData[0][3]: 0 
     }*/
  }  
}



/**
 * Función que permite obtener los datos del usuario activo
**/
function getUserDataFromCompetitor() {
  // se obtiene el correo del usuario
  var userMail = getUserMail();
  /*console.log('getUserDataFromCompetitor',{
    userData: getAccountUserPanel(userMail, "Usuario"),
    conceptList: getConceptListByPricesCompetitors(),
    allDataConcepts: getConceptList(),
    currencyTypeData: getCurrencyTypeDataCompetitors()
  });*/
  const datos =  {
    userData: getAccountUserPanel(userMail, "Usuario"),
    conceptList: getConceptListByPricesCompetitors(),
    allDataConcepts: getConceptList(),
    currencyTypeData: getCurrencyTypeDataCompetitors()
  };
  //console.log('getUserDataFromCompetitor', datos);
  return JSON.stringify(datos);
  //return datos;
  // retornamos el los datos
 /* return {
    userData: getAccountUserPanel(userMail, "Usuario"),
    conceptList: getConceptListByPricesCompetitors(),
    allDataConcepts: getConceptList(),
    currencyTypeData: getCurrencyTypeDataCompetitors()// Datos del tipo de cambio para hidroExpert 
  };*/
  
}


/**
 * Función que permite obtener todos las especies
 * <por alguna razon esta funcion no se usa>
**/
function getConceptListByPricesCompetitors() {
  var resources = getResources();
  var libro = SpreadsheetApp.openById(resources.competitorsPricingSheetId);
  // Referenciamos la hoja de conceptos 
  var hoja = libro.getSheetByName("Precios");
  // validamos si existe la hoja
  if(hoja){
    
    // obtenemos la ultima fila
    var lastRow = hoja.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
      return hoja.getRange(2, 1, lastRow - 2,38/* hoja.getLastColumn() */).getDisplayValues();
      // retornamos los datos
      
    }
  }
  
  // retornamos por defecto un array vacio
  return [];
}

/**
 * Función que permite obtener los datos de paquetería y fabricantes
**/
function getAllDataCompetitors(){
  
  /*return {
    parcels: [],
    manufacturers: [],
    dataListSltSearch: [],
    projectsData: [],
    prueba: 'hola',
    pricesAnalysis: [] // Se debe de verificar esta información
  }*/
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // referenciamos la hoja de parcelas
  var spreadsheet = SpreadsheetApp.openById(resources.competitorsPricingSheetId);
  //throw spreadsheet.getUrl()
  // referenciamos la hoja fabricantes
  var spreadsheetCompetitor = SpreadsheetApp.openById(resources.competitorsSheetId);
  //throw JSON.stringify( getAnalysisDataByPrices(spreadsheet))
  // retornamos los datos de fabricantes y paquetería

  return {
    //parcels: getParcelsListCompetitors(spreadsheetManufacturer),
    competitors: getCompetitorsList(spreadsheetCompetitor),
    dataListSltSearch: getListDataCompetitors(resources),
    pricesCompetitors: getAllDataPriceCompetitors(),
    //projectsData: getProjectsCompetitors(resources),
    pricesAnalysis: getAnalysisDataByPrices(spreadsheet) // Se debe de verificar esta información
  };
  
}


/*
* Permite obtener la lista de departamentos, nombre de especie, familia, palabra clave
*/
function getListDataCompetitors(resources){
  
  // Validamos si no viene los recursos 
  if(!resources){
    
    // se obtiene los recursos de hojas de cálculo
    var resources = getResources();
  }
  
  // referenciamos la hoja configuración 
  var spreadsheet = SpreadsheetApp.openById(resources.settingSheetId);

  // Referenciamos la hoja de conceptos 
  var speciesSheet = spreadsheet.getSheetByName("Conceptos");
  
  // Objeto con los datos de retorno 
  var objectData ={
     speciesNameListSlt: {}, // Array del listado de todas las especies
     departamentListObj: {},  // Departamentos 
     departamentListObjSlt: {},  // Departamentos 
     familyListObjSlt: {}, // Listadod e familias
     keywordListObjSlt: {},// Listado de las palabras clave
     familyListObj: {}, // Listadod e familias
     keywordListObj: {}// Listado de las palabras clave
  } 
  
  // Array temporal para el listados de conceptos
  var tempArrySpecieName  = []; 
  
  // Validamos de que se tenga acceso a la hoja de conceptos
  if(speciesSheet){
    
    // Obtenemos la ultima fila que tiene registros:
    var lastRow = speciesSheet.getLastRow();
    
    // se valida si existe mas de una fila
    if(lastRow > 1){
    
      // obtenemos los datos
      var arrayData = speciesSheet.getRange(2, 1, lastRow - 1, 5).getDisplayValues();
   
      // Recorremos cada uno de los registro obtenidos para formar el objeto que se requiere
      for(var i = 0; i < arrayData.length; i++){
         
         // validamos de que la unidad sea diferente a grupo y que el concepto no se encuentre
         if(arrayData[i][4].toUpperCase().trim() != "GRUPO" && (tempArrySpecieName.indexOf(arrayData[i][1]) == -1)){
           
           // Agregamos la especie
           tempArrySpecieName.push(arrayData[i][1]);
           
           // agregamos la especies
           objectData.speciesNameListSlt[arrayData[i][1]] = null;
           
           // Validamos que exista el nombre del departamento en la celda
           if(arrayData[i][0]){
           
             // Validamos si existe el departamento
             if(!objectData.departamentListObj[arrayData[i][0]]){
               
               // Creamos la propiedad de departamento
               objectData.departamentListObj[arrayData[i][0]] = {
                 familyListObj: {}
               };
               
               // Creamos la propiedad de departamento para el campo de selección
               objectData.departamentListObjSlt[arrayData[i][0]] = null;
             }
           
             // Validamos que extia el nombre de la familia en la celda
             if(arrayData[i][2]){
             
               // Validamos si existe la familia
               if(!objectData.departamentListObj[arrayData[i][0]].familyListObj[arrayData[i][2]]){
                 
                 // Creamos la propiedad de familia
                 objectData.departamentListObj[arrayData[i][0]].familyListObj[arrayData[i][2]] = {
                    keywordListObj:{}
                 };
                 
                 // Creamos la propiedad de familia para el campo de selección
                 objectData.familyListObjSlt[arrayData[i][2]] = null;
                 
                 // validamos si no existe el objeto de la familia 
                 if(!objectData.familyListObj[arrayData[i][2]]){
                   // Agregamos la propiedad 
                   objectData.familyListObj[arrayData[i][2]] = {};
                 }
               }
               
               // Validamos de que exista el valor de la palabra clave 
               if(arrayData[i][3]){
                 
                 // Validamos si existe la propiedd de palabra clave
                 if(!objectData.departamentListObj[arrayData[i][0]].familyListObj[arrayData[i][2]].keywordListObj[arrayData[i][3]]){
                   
                   // Creamos la propiedad de palabra clave
                   objectData.departamentListObj[arrayData[i][0]].familyListObj[arrayData[i][2]].keywordListObj[arrayData[i][3]] = [];
                   
                   // Creamos la propiedad de palabra clave para el campo de selección
                   objectData.keywordListObjSlt[arrayData[i][3]] = null;
                   
                   // Validamos de que la palabra clave no se encuentre como propiedad del objeto 
                   if(!objectData.keywordListObj[arrayData[i][3]]){
                   
                     // Creamos la propiedad
                     objectData.keywordListObj[arrayData[i][3]] = [];
                   
                   }
                   
                 }
                 
                 // Validamos de que exista la propiedad
                 if(!objectData.familyListObj[arrayData[i][2]][arrayData[i][3]]){
                 
                   // Creamos las propiedad para las familias
                   objectData.familyListObj[arrayData[i][2]][arrayData[i][3]] = [];
                 }
                 
                 // Insertamos las espcies
                 objectData.keywordListObj[arrayData[i][3]].push(arrayData[i][1]);
                 
                 // Insertamos las espcies
                 objectData.familyListObj[arrayData[i][2]][arrayData[i][3]].push(arrayData[i][1]);
                 
                 // Agregamos el concepto a la palabra clave
                 objectData.departamentListObj[arrayData[i][0]].familyListObj[arrayData[i][2]].keywordListObj[arrayData[i][3]].push(arrayData[i][1]);
                 
               }
             
             }
           }
           
         }
      }
    }
  }

  // Retornamos los datos 
  return objectData;
  
}


//
///*
//* Permite eliminar los simbilos para poder comparar
//*/
//function escapeRegExp(string) { return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string }

/**
 * Función que permite obtener los datos de paquetería
**/
function getParcelsListCompetitors(spreadsheet){
  
  // referenciamos la hoja de Paquetería
  var sheet = spreadsheet.getSheetByName("Paqueteria");
  
  // se valida si existe la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // se valida si existe mas de una fila
    if(lastRow > 1){
      
      // obtenemos los datos
      var arrayData = sheet.getRange(2, 1, lastRow - 1, 14).getDisplayValues();
      
      // variable que determina el listado
      var manufacturerList = [];
      
      // recorremos el listado de paquetería
      for(var i = 0; i < arrayData.length; i++){
        
        // agregamos los paquetería al array
        manufacturerList.push({
          id: arrayData[i][0],
          commerce: arrayData[i][1],
          businessName: arrayData[i][2],
          branch: arrayData[i][3],
          clientNumber: arrayData[i][4],
          rfc: arrayData[i][5],
          contact: arrayData[i][6],
          officePhone: arrayData[i][7],
          mobilePhone: arrayData[i][8],
          email: arrayData[i][9],
          otherEmail: arrayData[i][10],
          lastUpdate: arrayData[i][11],
          state: arrayData[i][12],
          code: arrayData[i][13],
          rowIndex: (i + 2)
        });
        
      }
      
      //retornamos los datos
      return manufacturerList;
      
    }
  }
  
  // retornamos un valor vacio
  return [];
}

/**
 * Función que permite obtener los datos de fabricantes
**/
function getCompetitorsList(spreadsheet){
  
  // referenciamos la hoja de fabricantes
  var sheet = spreadsheet.getSheetByName("_Competidores");
  
  // Objetos con los datos del objeto
  var objDataCompetitor = {
    competitorList: [],
    //manufacturerListObj: {},
    //manufacturerFamilies: {}
  }
  
  // se valida si existe la hoja
  if(sheet){    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();    
    // se valida si existe mas de una fila
    if(lastRow > 1){      
      // obtenemos los datos
      var arrayData = sheet.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
      //var arrayData = sheet.getRange(2, 1, lastRow - 1, 18).getDisplayValues();
      
      // variable que determina el listado
      var manufacturerList = [];      
      // objeto con el nombre de los fabricantes
      var manufacturerListObj = [], families = null, currFamily;      
      // recorremos el listado de competidores
      for(var i = 0; i < arrayData.length; i++){         
        // agregamos los competidores al array    
        objDataCompetitor.competitorList.push({
          status: arrayData[i][0],
          id: arrayData[i][1],
          competitorName: arrayData[i][2],          
          socialReason: arrayData[i][3],
          rfc: arrayData[i][4],
          presence: arrayData[i][5],
          address: arrayData[i][6],
          nameContact: arrayData[i][7],
          officePhone: arrayData[i][8],
          mobilePhone: arrayData[i][9],
          webUrl: arrayData[i][10],
          created_at: arrayData[i][11],
          lastUpdate: arrayData[i][12],
          rowIndex: (i + 2)
         });        
      }
    }
  }  
  // retornamos los datos
  return objDataCompetitor;
}


/*
* Función que permite guardar un fabricante o un producto 
* @param {Object} objectDataMP Objeto con los datos para la creción
*/
function createCompetitorOrParcel(objectDataMP){
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();  
  // referenciamos la hoja fabricantes
  var spreadsheet = SpreadsheetApp.openById(resources.competitorsSheetId);
  //var spreadsheet = SpreadsheetApp.openById(resources.acquisitionSheet);
  //validamos que existan los datos para la creación 
  if(objectDataMP){    
    // Validmos si se debe crea un fabricante 
    //if(objectDataMP.creationType.typeEnglish == "manufacturers"){
      
      // Guardamos los datos del fabricatne 
      //return saveDataCompetitor(objectDataMP, spreadsheet.getSheetByName("_Competidores"), spreadsheet);
      
      // Validamos si se debe crear un paquete 
    //}else{
      
      // Guardamos los datos del paquete
     // return saveDataParcelCompetitor(objectDataMP, spreadsheet.getSheetByName("Paqueteria"), spreadsheet);
      
   // } 
  }
    
}


/*
* Función que guarda la información del competidor
* @params {Objeto} objectDataMP Datos para guardar el competidor
* @params {reference} sheet Referencia a la hoja de competidor
*/
function saveDataCompetitor_OLD(objectDataMP, sheet, spreadsheet){
   
   // Validamos de que se tenga acceso a la hoja
   if(sheet){
    
     // Obtenemos la ultima fila 
     var lastRow = sheet.getLastRow();
     
     // Guardamos los datos
     sheet.getRange(lastRow +1, 1, 1, 13).setNumberFormat("@STRING@").setValues([[
          (lastRow +1),
          objectDataMP.nameMP,
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          objectDataMP.contactName,
          "",
          "",
          objectDataMP.email
        
      ]]);
     
   }
   
   // Retornamos la lista de fabricantes
   return getCompetitorsList(spreadsheet).competitorList;
}



/*
* Función que guarda la información del paquete
* @params {Objeto} objectDataMP Datos para guardar el paquete
* @params {reference} sheet Referencia a la hoja de paquete
*/
function saveDataParcelCompetitor(objectDataMP, sheet,spreadsheet){

    // Validamos de que se tenga acceso a la hoja
   if(sheet){
    
     // Obtenemos la ultima fila 
     var lastRow = sheet.getLastRow();
     
     // Guardamos los datos 
     sheet.getRange(lastRow +1, 1, 1, 10).setNumberFormat("@STRING@").setValues([[
          (lastRow +1),
          objectDataMP.nameMP,
          "",
          "",
          "",
          "",
          objectDataMP.contactName,
          "",
          "",
          objectDataMP.email
         
      ]]);
     
   }
   
   // Retornamos la lista de paqueterias
   return getParcelsListCompetitors(spreadsheet);
}


/*
* Funcioón que permite obtener los proyectos con sus respectivas especies
*
*/
function getProjectsCompetitors(resources){
  
  // Validamos si no viene los recursos
  if(!resources){
  
    // se obtiene los recursos de hojas de cálculo
    var resources = getResources();
  }
  
  // referenciamos la hoja solicitudes de proyectos
  var spreadSheet = SpreadsheetApp.openById(resources.requestSheetId);
 
  // validamos de que se tenga acceso a la hoja 
  if(spreadSheet){
     
     // Referenciamos la hoja de cotizaciones
     var quotationsSheet = spreadSheet.getSheetByName("Cotizaciones");
     
     // Validamos de que se tenga acceso a la hoja de cotizaciones
     if(quotationsSheet){
        
        // Obtenemos la ultima final donde se tenga registros
        var lastRow = quotationsSheet.getLastRow();
        
        // Validamos de que se tenga registros
        if(lastRow > 1){
          
          // Datos para el select
          var objectData = {};
          
          // Proyectos para verificar
          var arrayVerify = [];
          
          // Variable para el nombre del proyecto
          var nameProject = "",
              objectProjects = {},
              structureTemp = [], sheetsList = [];
          
          // obtenemos la información de las versiones
          var versionsData = getVersionsObj(spreadSheet);
          
          // Obtenmos los datos de los proyectos 
          var arrayData = quotationsSheet.getRange(2, 1, lastRow, 16).getDisplayValues();
          
          // Recorremos cada uno de los proyectos
          for(var i = 0; i < (arrayData.length - 1); i++){
            
            // se valida que el proyecto este activdo
            if(arrayData[i][9] == "Activa"){
            
              // Establecemos el nombre del proyecto
              nameProject =  arrayData[i][0].trim() + " " + arrayData[i][2].trim(); 
              
              // Adiconamos el datos al objeto del select
              objectData[nameProject] = null;
              
              // Agregamos el nombre del proyecto al array
              arrayVerify.push(nameProject);
              
              // definimos la estructura inicial
              structureTemp = {
                "Proyecto": {
                  structure: arrayData[i][10]
                }
              };
              
              // agregamos la primera versión
              sheetsList = ["Proyecto"];
           
              // se valida si se tiene un versionado en la solicitud
              if(arrayData[i][12] == "SI"){
                
                // obtenemos los datos de la propiedad
                var relationArray = versionsData[arrayData[i][0]];
                
                // se valida si existe datos de la relación actual
                if(relationArray){
                  
                  // se recorre cada una de las versiones
                  for(var j = 0; j < relationArray.length; j++){
                    
                    // agregamos la información de la versión
                    structureTemp[relationArray[j].sheetName] = relationArray[j];
                    
                    // agregamos el nuevo nombre de la hoja
                    sheetsList.push(relationArray[j].sheetName);
                  }
                }
              }
              
              // Objeto con los datos del projecto 
              objectProjects[nameProject] = {
                 projectId: arrayData[i][0],
                 projectName: arrayData[i][2],
                 projectData: arrayData[i][10],
                 structures: structureTemp,
                 sheetsList: sheetsList,
                 sheetDefault: arrayData[i][11]
              }
            }
          }

          // Retornamos los datos para el proyecto 
          return {
            objectData: objectData,
            objectProjects: objectProjects,
            arrayVerify: arrayVerify
          };
        }
     } 
  }
}




/*
* Función que permite actualizar el tipo de cambio del fabricante
*/
function updateCurrencyTypeCompetitor(manufacturerData){

 // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // referenciamos la hoja solicitudes de proyectos
  var spreadSheet = SpreadsheetApp.openById(resources.competitorsPricingSheetId);
  
  // validamos de que se tenga acceso a la hoja 
  if(spreadSheet){
     
     // Referenciamos la hoja de Precios / Fabricantes
     var manufacturerPricesSheet = spreadSheet.getSheetByName("_Precios/Competidores");
     
     // Validamos de que se tenga acceso a la hoja Precios/Fabricantes
     if(manufacturerPricesSheet){
        
        // Obtenemos la ultima final donde se tenga registros
        var lastRow = manufacturerPricesSheet.getLastRow();
        
        // Validamos de que se tenga registros
        if(lastRow > 1){
           
          // Obtenmos los datos de los proyectos 
          var arrayDataManufacturers = manufacturerPricesSheet.getRange(2, 1, lastRow -1, 10).getDisplayValues();
          
          // fecha actual del sistema
          var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");
          
          // obtenemos el usuario activo 
          var userEmail = getUserMail();
          
          // Recorremos cada uno de los registros
          for( var i = 0; i < arrayDataManufacturers.length; i++){
              
              // Validamos por el nombre del fabricante 
              if(manufacturerData.manufacturerName == arrayDataManufacturers[i][2]){
               
                // Actualizamos el tipo de cambio para dorales y euros
                manufacturerPricesSheet.getRange( i + 2, 10, 1, 2).setNumberFormat("@STRING@").setValues([[manufacturerData.newUSD.trim().replace(/,/g, ""), manufacturerData.newEUR.trim().replace(/,/g, "")]]);
                
                // Actualizamos la fecha de actualización y responsable
                manufacturerPricesSheet.getRange( i + 2, 30, 1, 2).setNumberFormat("@STRING@").setValues([[currentDate,userEmail]]);
              }
           }
           
          // Retornamos el listado de precios
          return getAllDataCompetitors(); //getAnalysisDataByPrices(spreadSheet);
        }
     }
   }

} 


/**
 * Función que pemite obtener el listado de familias disponibles hasta la fecha
**/
function getFamiliesListCompetitors(){

  // Se invoca la función donde se referencia las variables globales de la aplicación
  var resources = getResources();
  
  // referenciamos el libro de configuración del hidrocotizador
  var spreadsheet = SpreadsheetApp.openById(resources.settingSheetId);
  
  // referenciamos la hoja de conceptos
  var sheet = spreadsheet.getSheetByName(resources.conceptSheetName);
  
  // se valida si existe la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // se valida si existe mas de una fila
    if(lastRow > 1){
      
      // obtenemos los datos
      var arrayData = sheet.getRange(2, 3, lastRow - 1, 1).getDisplayValues();
      
      // variable que determina el listado
      var familiesList = [];
      
      // recorremos el listado de fabricantes
      for(var i = 0; i < arrayData.length; i++){
        
        // se valida si la familia NO existe
        if(arrayData[i][0] && familiesList.indexOf(arrayData[i][0]) == -1){
          
          // agregamos la familia a la lista
          familiesList.push(arrayData[i][0]);
          
        }
      }

      // retornamos el listado
      return familiesList;
    }
  }
  
  // retornamos un listado vacio
  return [];
}


/**
 * Permite actualizar las familias de un fabricante
 */
function updateAssignFamiliesCompetitors(formObject) {
  
  // Se invoca la función donde se referencia las variables globales de la aplicación
  var resources = getResources();
  
  // Se declara una variable para referenciar el lbro donde estan los fabricantes
  var sheet, recordArray, spreadsheet = SpreadsheetApp.openById(resources.competitorsSheetId);
  
  // referenciamos la hoja de fabricantes
  var sheet = spreadsheet.getSheetByName(resources.competitorsSheetName);
  
  // Se valida la acción a ejecutar
  if (sheet) {
    
    // Se inicializa el array con los siguientes valores
    var familiesData = [
      String(formObject.familiesOffered), // Familias ofertadas
    ];
      
    // Se valida si la hoja y el array de datos existen
    if (familiesData.length > 0) {
      
      // Se guardan los datos finales
      sheet.getRange(formObject.rowIndex, 13, 1, familiesData.length).setNumberFormat("@STRING@").setValues([familiesData]);
      
    } else {
      // Se genera una excepción para finalizar el recorrido
      throw "Lo sentimos no es posible almacenar la información, intenta nuevamente y si persiste el error por favor comunicate con el administrador.";
    }
  } else {
    
    // retornamos un mensaje de error
    throw "La hoja 'Competidores' no existe.";
  }
  
    
  // retornamos los datos de fabricantes y paquetería
  return getCompetitorsList(spreadsheet).competitorList;
}

/* ===================== Funciones de Alejandro ================================ */
/*
* Función que guarda la información del competidor
* @params {Objeto} objectDataMP Datos para guardar el competidor
* @params {reference} sheet Referencia a la hoja de competidor
*/
function saveDataCompetitor(objectDataMP){
  var currentDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/M/yyyy HH:mm:ss");
  var resources = getResources();  
  // referenciamos la hoja fabricantes
  var spreadsheet = SpreadsheetApp.openById(resources.competitorsSheetId);
  var sheet = spreadsheet.getSheetByName("_Competidores");
  //var spreadsheet = SpreadsheetApp.openById(resources.acquisitionSheet);
  //validamos que existan los datos para la creación 
  if(objectDataMP){
   // Validamos de que se tenga acceso a la hoja
   if(sheet){    
     // Obtenemos la ultima fila 
     var lastRow = sheet.getLastRow();     
     // le damos forma a la hoja
     var arrayData = [
       objectDataMP.status,
       Utilities.getUuid(),//lastRow,       
       objectDataMP.competitorName,
       objectDataMP.socialReason,
       objectDataMP.rfc,
       objectDataMP.presence,
       objectDataMP.address,
       objectDataMP.nameContact,
       objectDataMP.officePhone,
       objectDataMP.mobilePhone,
       objectDataMP.webUrl,       
       currentDateStr,
       currentDateStr     
     ];
  
       // guardamos los datos base del usuario
     sheet.getRange(lastRow+1, 1, 1, 13).setNumberFormat("@STRING@").setValues([arrayData]);     
     // Guardamos el nombre de contacto
     //sheet.getRange(lastRow, 16).setNumberFormat("@STRING@").setValue(objectDataMP.contactName);    
   }  
  } 
  // Retornamos la lista de competidores
  return getCompetitorsList(spreadsheet).competitorList;
}

/*
* Función que actualizar la información del competidor
* @params {Objeto} objectDataMP Datos para guardar el competidor
*/
function updateDataCompetitor(objectDataMP){
  var currentDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/M/yyyy HH:mm:ss");
  var resources = getResources();  
  // referenciamos la hoja fabricantes
  var spreadsheet = SpreadsheetApp.openById(resources.competitorsSheetId);
  var sheet = spreadsheet.getSheetByName("_Competidores");
  //var spreadsheet = SpreadsheetApp.openById(resources.acquisitionSheet);
  //validamos que existan los datos para la creación 
  if(objectDataMP){
   // Validamos de que se tenga acceso a la hoja
   if(sheet){    
     // le damos forma a la hoja
     var arrayData = [   
       objectDataMP.competitorName,
       objectDataMP.socialReason,
       objectDataMP.rfc,
       objectDataMP.presence,
       objectDataMP.address,
       objectDataMP.nameContact,
       objectDataMP.officePhone,
       objectDataMP.mobilePhone,
       objectDataMP.webUrl   
     ];
     var listado = getCompetitorsList(spreadsheet).competitorList;
     var rowUpdate = listado.filter((competitor) => competitor.id == objectDataMP.id );
     if( rowUpdate.length > 0 ){
       sheet.getRange(rowUpdate[0].rowIndex, 3, 1, 9).setNumberFormat("@STRING@").setValues([arrayData]);      
       sheet.getRange(rowUpdate[0].rowIndex,1).setNumberFormat("@STRING@").setValues([[objectDataMP.status]]);
       sheet.getRange(rowUpdate[0].rowIndex,13).setNumberFormat("@STRING@").setValues([[currentDateStr]]); 
     }
      // guardamos los datos base del competidor
     
   }  
  } 
  // Retornamos la lista de competidores
  return getCompetitorsList(spreadsheet).competitorList;
}


/*
* Función que guarda la información de Precios/Competidores
* @params {Objeto} objectDataMP Datos para guardar el Precios/Competidores
* @params {reference} sheet Referencia a la hoja de Precios/Competidores
*/
function saveDataPrecioCompetitor(objectDataMPS){
  var currentDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
  var resources = getResources();  
  // referenciamos la hoja fabricantes
  var spreadsheet = SpreadsheetApp.openById(resources.competitorsPricingSheetId);
  var sheet = spreadsheet.getSheetByName("_Precios/Competidores");
 
  //var spreadsheet = SpreadsheetApp.openById(resources.acquisitionSheet);
  //validamos que existan los datos para la creación 
  if(objectDataMPS.length > 0){
  
   // Validamos de que se tenga acceso a la hoja
   if(sheet){
     // Obtenemos la ultima fila 
     
     for( i = 0; i< objectDataMPS.length; i++){
       var lastRow = sheet.getLastRow() + 1;
       var objectDataMP = objectDataMPS[i];
       // le damos forma a la hoja
      
       var arrayData = [
         //Hidro
         objectDataMP.idCompetidor,//A
         objectDataMP.concepto,//B
         objectDataMP.unidad,//C
         objectDataMP.cantidadSolicitada,//D         
         objectDataMP.codigoProducto,//E
         objectDataMP.conceptoCompetidor,//F
         objectDataMP.tipoCambioUSD,//G
         objectDataMP.tipoCambioEUR,//H
         objectDataMP.unidadCompetidor,//I
         objectDataMP.cantidadEquivalenteCompetidor,//J
         objectDataMP.equivalenciaUnidad,//K
         objectDataMP.cantidadMinima,//L
         objectDataMP.precioLista,//M
         objectDataMP.moneda,//N
         objectDataMP.descuentoAplicado,//O
         objectDataMP.cantidadCotizada,//P
         objectDataMP.precioUnitarioOfertado,//Q 
         objectDataMP.iva,//R
         objectDataMP.precioAdquisicion,//S
         objectDataMP.importeCotizado,//T
         '0',//(objectDataMP.cantidadEquivalenteCompetidor*objectDataMP.precioAdquisicion)/objectDataMP.cantidadSolicitada,//objectDataMP.precioXUnidadCompetidor,//U
         objectDataMP.precioMXNxUnidadHEXP,//V
         currentDateStr,//W
         getUserMail(),//X
         objectDataMP.estatus//Y
       ];
       
       var listado = getPriceCompetitorsList(spreadsheet);
       var rowUpdate = [];
       var rowsUpdate = listado.filter((priceCompetitor) => (priceCompetitor.idCompetidor == objectDataMP.idCompetidor && priceCompetitor.concepto == objectDataMP.concepto ) );
       var rowNueva = true;
       
       if(rowsUpdate.length > 0){
         rowUpdate = rowsUpdate[rowsUpdate.length - 1];
         var dias = diffDias( rowUpdate.fechaActualizacion );
         //console.log('rowUpdate',dias,rowUpdate);
         if(dias < 15){
           rowNueva = false;
         }
       }
       
       if(rowNueva){//nueva si no existe o  tiene mas de 15 dias sin actualizacion
         sheet.getRange(lastRow, 1, 1, 25).setNumberFormat("@STRING@").setValues([arrayData]);
          
         sheet.getRange(lastRow,7)
         .setFormula("=('Tipo de Cambio'!$Q$2)");
         sheet.getRange(lastRow,8)
         .setFormula("=('Tipo de Cambio'!$S$2)");
         
         sheet.getRange(lastRow,21)
         .setFormula("=(J"+lastRow+"*S"+lastRow+")/D"+lastRow+"");
         sheet.getRange(lastRow,22)
         .setFormula("=IF(N"+lastRow+"=\"USD\",U"+lastRow+"*'Tipo de Cambio'!$Q$2,IF(N"+lastRow+"=\"EUR\",'Tipo de Cambio'!$S$2*U"+lastRow+",U"+lastRow+"))");
         
         for(j = 0; j < rowsUpdate.length; j++){
           sheet.getRange(rowsUpdate[i].rowIndex,25).setNumberFormat("@STRING@").setValues([['Inactivo']]);    
         }
         
       }else{ // actualiza el ultimo registro
         sheet.getRange(rowUpdate.rowIndex, 1, 1, 25).setNumberFormat("@STRING@").setValues([arrayData]);
        
         sheet.getRange(rowUpdate.rowIndex,7)
         .setFormula("=('Tipo de Cambio'!$Q$2)");
         sheet.getRange(rowUpdate.rowIndex,8)
         .setFormula("=('Tipo de Cambio'!$S$2)");
         
         sheet.getRange(rowUpdate.rowIndex,21)
         .setFormula("=(J"+rowUpdate.rowIndex+"*S"+rowUpdate.rowIndex+")/D"+rowUpdate.rowIndex+"");
         
         sheet.getRange(rowUpdate.rowIndex,22)
         .setFormula("=IF(N"+rowUpdate.rowIndex+"=\"USD\",U"+rowUpdate.rowIndex+"*'Tipo de Cambio'!$Q$2,IF(N"+rowUpdate.rowIndex+"=\"EUR\",'Tipo de Cambio'!$S$2*U"+rowUpdate.rowIndex+",U"+rowUpdate.rowIndex+"))");
         
         for(j = 0; j < rowsUpdate.length - 1; j++){
           sheet.getRange(rowsUpdate[i].rowIndex,25).setNumberFormat("@STRING@").setValues([['Inactivo']]);         
         }
       }
         
      }
       
       
    }  
  } 
  // Retornamos la lista de competidores
  return getPriceCompetitorsList(spreadsheet);
}

function getAllDataPriceCompetitors(idsCompetidores = []){
  var resources = getResources();  
  var spreadsheet = SpreadsheetApp.openById(resources.competitorsPricingSheetId);
  //var sheet = spreadsheet.getSheetByName("_Precios/Competidores");
  return getPriceCompetitorsList(spreadsheet, idsCompetidores);

}

function getAllDataPriceHidroCompetitorsBak(col){
  var column = col;
  var resources = getResources();  
  //console.log('cantida de competidores')
  //var sheet = spreadsheet.getSheetByName("_Precios/Competidores");
  return getPriceHidroCompetitorsList(spreadsheet, column);

}

function getListaValoresCompetidores(){
  var resources = getResources();
  var lista = [];
  var spreadsheet = SpreadsheetApp.openById(resources.competitorsSheetId);
  var sheet = spreadsheet.getSheetByName("Lista de valores");
  if(sheet){
    var lastRow = sheet.getLastRow();
    if(lastRow > 1){
      arrayData = sheet.getRange(1, 1, lastRow - 1, 38 ).getDisplayValues();
      for(var i = 1; i < arrayData.length; i++){        
        lista.push(arrayData[i][0]);
      }
    }
  }
  
  return lista;
}

function getAllDataPriceHidroCompetitors(filtros){
  var resources = getResources();

  var spreadsheet = SpreadsheetApp.openById(resources.competitorsPricingSheetId);
  var sheet = spreadsheet.getSheetByName("Precios");

  var spreadsheetCompetidores = SpreadsheetApp.openById(resources.competitorsSheetId);
  
  var competidoresList = getCompetitorsList(spreadsheetCompetidores).competitorList;
  var idsCompetidores = [];
  if(filtros.competidor.length > 0){
    competidoresList = competidoresList.filter(
      (competidor) =>  {
        if(filtros.competidor.indexOf( competidor.competitorName ) >= 0){
          idsCompetidores.push(competidor.id);
          return competidor
        }        
      }
    );     
  }
    
  var pricesCompetitorList = getPriceCompetitorsList(spreadsheet, idsCompetidores );
  
  var listaCompetidores={};
  var arrayData = [];
  var dataReturn = [];// = [{nombre: 'Nombre de la especie',unidad: 'Unidad', competidores : {}}];
  
  for(i = 0; i < competidoresList.length; i++){
    if(competidoresList[i].status == 'Activo'){
      listaCompetidores[ `${competidoresList[i].id}` ] = {'nombre':competidoresList[i].competitorName, especie: false};
    }        
  }
  
  dataReturn.push({nombre: 'Nombre de la especie',unidad: 'Unidad', 'listaCompetidores':listaCompetidores});
  if(sheet){
    
    var lastRow = sheet.getLastRow();
    if(lastRow > 1 && pricesCompetitorList.length > 0){
               
      arrayData = sheet.getRange(1, 1, lastRow - 1, 38 ).getDisplayValues();
      //Recorrer especies/precios
      for(var i = 1; i < arrayData.length; i++){
            
        if(filtros){
        
          if(filtros.departamento.length > 0){
            if(filtros.departamento.indexOf(arrayData[i][0]) < 0 ){
              continue;
            }
          }
          if(filtros.familia.length > 0){
            if(filtros.familia.indexOf(arrayData[i][2]) < 0 ){
              continue;
            }
          }
          if(filtros.palabraClave.length > 0){
            if(filtros.palabraClave.indexOf(arrayData[i][3]) < 0 ){
              continue;
            }
          }
          if(filtros.especie.length > 0){
            if(filtros.especie.indexOf(arrayData[i][1]) < 0 ){
              continue;
            }
          }
        }
        //buscar la especie en competidores
        var resultadoBusqueda = pricesCompetitorList.filter(
          (precioCom) => (precioCom.concepto == arrayData[i][1] && precioCom.estatus == 'activo') 
        );
        
        //si encuentra la especie en un competidor
        
        var arregloPreciosCompetidores = {};
        for(k = 0; k < competidoresList.length; k++){
          if(competidoresList[k].status == 'Activo'){
             arregloPreciosCompetidores[ `${competidoresList[k].id}` ] = {'nombre':competidoresList[k].competitorName, especie: false};
           }        
        }
                
        var precios = [];
        if(resultadoBusqueda.length > 0){
         for(var j = 0; j < resultadoBusqueda.length; j++){
            precios.push(resultadoBusqueda[j].precioMXNxUnidadHEXP);
            arregloPreciosCompetidores[`${resultadoBusqueda[j].idCompetidor}`].especie= {
              'conceptoCompetidor':resultadoBusqueda[j].conceptoCompetidor,
              'unidad' : resultadoBusqueda[j].unidadCompetidor,
              'tipoCambioUSD': resultadoBusqueda[j].tipoCambioUSD,
              'tipoCambioEUR': resultadoBusqueda[j].tipoCambioEUR,
              'precioLista' : resultadoBusqueda[j].precioMXNxUnidadHEXP,
              'moneda' : resultadoBusqueda[j].moneda,
              'index' : resultadoBusqueda[j].rowIndex
            };          
          }
          //console.log('arregloPreciosCompetidores:',arregloPreciosCompetidores)
          dataReturn.push({
            'nombreEspecie': arrayData[i][1],
            'unidad': arrayData[i][4],
            'precioMax' : precios.length > 1 ? Math.max(...precios) : '',
            'precioMin' : precios.length > 1 ? Math.min(...precios) : '',
            'listaCompetidores':arregloPreciosCompetidores,//arregloPreciosCompetidores,
            'rowIndex': (i + 2)
          });
        }       
      }
    }
  }
  //console.log('Lista precios hidro, mas competidores', dataReturn);
  return  dataReturn ;
}
/**
 * Función que permite obtener los datos de precios/Competidores
**/
function getPriceCompetitorsList(spreadsheet, idsCompetidores = []){
  var sheet = spreadsheet.getSheetByName("_Precios/Competidores");
  var pricesCompetitor = [];  
  if(sheet){
    var lastRow = sheet.getLastRow();    
    if(lastRow > 1){
      var arrayData = sheet.getRange(2, 1, lastRow - 1, 25).getDisplayValues();
     
      for(var i = 0; i < arrayData.length; i++){
        if(arrayData[i][24] == 'Activo' || arrayData[i][24] == 'activo'){
          if(idsCompetidores.length > 0){
            if(idsCompetidores.indexOf(arrayData[i][0]) < 0 ){
              continue;
            }
          }
          pricesCompetitor.push({
            idCompetidor: arrayData[i][0],
            concepto: arrayData[i][1],
            unidad: arrayData[i][2],
            cantidadSolicitada: arrayData[i][3],
            codigoProducto: arrayData[i][4],
            conceptoCompetidor: arrayData[i][5],
            tipoCambioUSD: arrayData[i][6],
            tipoCambioEUR: arrayData[i][7],            
            unidadCompetidor: arrayData[i][8],
            cantidadEquivalenteCompetidor: arrayData[i][9],
            equivalenciaUnidad: arrayData[i][10],
            cantidadMinima: arrayData[i][11],
            precioLista: arrayData[i][12],
            moneda: arrayData[i][13],            
            descuentoAplicado: arrayData[i][14],
            cantidadCotizada: arrayData[i][15],
            precioUnitarioOfertado: arrayData[i][16],
            iva: arrayData[i][17],            
            precioAdquisicion: arrayData[i][18],
            importeCotizado: arrayData[i][19],            
            precioXUnidadCompetidor: arrayData[i][20],
            precioMXNxUnidadHEXP: arrayData[i][21],            
            fechaActualizacion: arrayData[i][22],
            correoAutorizo: arrayData[i][23],
            estatus: arrayData[i][24],
            rowIndex: (i + 2)
           });
        }                
      }
    }
  }
  //console.log('return pricesCompetitor', pricesCompetitor);
  return pricesCompetitor;
}

function diffDias(fecha){
  var dt1 = new Date(),
    dt2 = new Date(fecha);
  var t1 = dt1.getTime(),
      t2 = dt2.getTime();

  var diffInDays = Math.floor((t1-t2)/(24*3600*1000));
  return diffInDays;
}


//FUNCION PARA HACER TEST
function testFuncion(){
  //console.log('Test');
  //var dias = diffDias('09/20/2021 17:37:19');
  //console.log("test Funcion")
  var datos = getAllDataPriceCompetitors([]);
 /* 
  var datos= getAllDataPriceHidroCompetitors({
    "departamento": [],
    "familia": [],
    "especie": [],//["ABRAZADERA !ACERO  ø4 Pulg SIN FIN"],
    "palabraClave": [],
    "competidor": []
});*/
console.log('datos', datos)
return false;
  //console.log('Test diffDias', dias);
  //return false;
  /*return saveDataPrecioCompetitor({
    idCompetidor: '3',
    concepto: 'ABRAZADERA !ACERO  ø4 Pulg SIN FIN',
    unidad: 'M',
    codigoProducto: '123456',
    conceptoCompetidor: 'XXXX',
    tipoCambioUSD: 'XXXX',
    tipoCambioEUR: 'XXXX',
    unidadCompetidor: 'XXXX',
    equivalenciaUnidad: 'XXXX', 
    cantidadMinima: 'XXXX',
    precioLista: 30,
    moneda: 'XXXX',
    descuentoAplicado: 'XXXX',
    cantidadCotizada: 'XXXX',
    precioUnitarioOfertado: 'XXXX',
    iva: 'XXXX',
    precioAdquisicion: 'XXXX',
    importeCotizado: 'XXXX',
    estatus: 'activo'    
  });*/
  /*return updateDataCompetitor({
    id: '3',
    status: 'Activo',
    competitorName: 'DeWalt S.A. de C.V.',
    socialReason: 'DeWalt S.A. de C.V.',
    rfc: '',
    presence: 'TAMAULIPAS,MICHOACÁN,NAYARIT',
    address: 'Central de Abastos, Jesús Díaz Barriga 60, Elías Pérez Ávalos, 58218 Morelia, Mich.',
    nameContact: 'Rogelio',
    officePhone: '443 333 4812',
    mobilePhone: '834 123 654 78',
    webUrl: 'https://dwalt.com/',
  });*/
  
  
}


//Test actualiacion de precios fabricantes
function saveUpdatePriceSpecieForHidro(formData, manufacturerData){
  
  // variable a retornar
  var returnObject = {
    success: false,
    message: "La especie no puede ser actualizada, por favor intenta de nuevo."
  };
  
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources(), spreadSheet = null;
  
  try{
    // referenciamos la hoja solicitudes de proyectos
    spreadSheet = SpreadsheetApp.openById(resources.acquisitionSheet);
  }catch(e){}
  
  // validamos de que se tenga acceso a la hoja 
  if(spreadSheet){
     
    // fecha actual del sistema
    var currentDate = new Date();
    var timeZone = Session.getScriptTimeZone();
    
    // validamos si no existe el index, es decir que el fabricanto no esta registrado
    if(!manufacturerData.index){
       
      // fecha actual si caracteres especiales
      var currentDateTrim = Utilities.formatDate(currentDate, timeZone, "ddMMyyyyHHmmss");
       
      // Creamos el código de actualización
      manufacturerData.codePack = currentDateTrim + Utilities.getUuid();
      
      // Referenciamos la hoja de Fabricante/Solicitudes
       var mainSheet = spreadSheet.getSheetByName("Fabricante/Solicitudes");
       
       // Validamos de que se tenga acceso a la hoja Precios/Fabricantes
       if(mainSheet){
         
         // Obtenemos la ultima fila
        var mainLastRow = mainSheet.getLastRow();
        
        //Actualizamos el registro en la hoja
        mainSheet.getRange(mainLastRow + 1, 1, 1, 3).setValues([[manufacturerData.manufacturerId, manufacturerData.codePack, "Pendiente"]]);
       }
     }
     
     // Referenciamos la hoja de Precios / Fabricantes
     var sheet = spreadSheet.getSheetByName("Precios/Fabricantes");
     
     // Validamos de que se tenga acceso a la hoja Precios/Fabricantes
     if(sheet){
       
       // obtenemos la fila
       var rowIndex = manufacturerData.index;
       
       // validamos si es un registro nuevo
       if(manufacturerData.codePack){
         
         // obtenemos la fecha actual en "MM/dd/yyyy HH:mm:ss"
         var currentDateStr = Utilities.formatDate(currentDate, timeZone, "dd/MM/yyyy HH:mm:ss");
         
         // obtenemos la ultima fila y sumamos uno
         rowIndex = sheet.getLastRow() + 1;

//commentamoslo fuera para probar Chris Meli 4/3/2022
//estaba antes CUrrentdateStr en el siguente array           manufacturerData.cantidadSolicitada,


         // Agregamos la información actual
         sheet.getRange(rowIndex, 1, 1, 7).setValues([[
           manufacturerData.manufacturerId,
           manufacturerData.codePack,
           manufacturerData.manufacturerName,
           manufacturerData.conceptName,
           manufacturerData.unit,
           currentDateStr,
           (formData ? formData.code : "")
         ]]);
       }
      
       // se valida que exista la fila
       if(rowIndex){
          var userEmail = getUserMail();

          // Guardamos los tipos de cambios y el valor de unidad
         sheet.getRange(rowIndex, 10, 1, 3).setValues([[formData.dolar, formData.euro, formData.unit]]);
         sheet.getRange(rowIndex, 8).setValue(formData.productName);
         sheet.getRange(rowIndex, 7).setValue(formData.code || "");
         sheet.getRange(rowIndex, 16, 1, 5).setValues([[formData.currency, formData.discountApplied, formData.quantity, formData.suggestedPrice, formData.iva]]);         
         sheet.getRange(rowIndex, 22).setValue(formData.total);
         sheet.getRange(rowIndex, 30).setNumberFormat("@STRING@").setValue(Utilities.formatDate(currentDate, timeZone, "dd/MM/yyyy HH:mm:ss"));
         sheet.getRange(rowIndex, 21).setFormula(`=S${rowIndex}*(T${rowIndex}/100 + 1)`);
         sheet.getRange(rowIndex, 22).setFormula(`=U${rowIndex}*R${rowIndex}`);
         sheet.getRange(rowIndex, 31).setValue(userEmail);


         // Guardamos los tipos de cambios y el valor de unidad
         //sheet.getRange(rowIndex, 8).setValue(formData.codigoProducto);
         //sheet.getRange(rowIndex, 9).setValue(formData.productName);
         //sheet.getRange(rowIndex, 11, 1, 3).setValues([[formData.dolar, formData.euro, formData.unit]]);
         //
        // sheet.getRange(rowIndex, 17, 1, 5).setValues([[formData.currency, formData.discountApplied, formData.quantity, formData.suggestedPrice, formData.iva]]);         
        // sheet.getRange(rowIndex, 23).setValue(formData.total);
        // sheet.getRange(rowIndex, 33).setNumberFormat("@STRING@").setValue(Utilities.formatDate(currentDate, timeZone, "dd/MM/yyyy HH:mm:ss"));
         
         // actualizamos el estado y el mensaje
         returnObject.success = true;
         returnObject.message = "El precio del concepto: <b>" + formData.productName + "</b> ha sido registrado correctamente.";
         returnObject.data = getAllDataAcquisition();
       }
     }
  }
  
  // retornamos el valor
  return returnObject;
}
