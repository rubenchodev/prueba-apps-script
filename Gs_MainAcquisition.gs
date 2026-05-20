/*
* Permite guardar la actuakziación de precio de la especie por parte de Hidro
*/
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
         
         // Guardamos los tipos de cambios y el valor de unidad
         sheet.getRange(rowIndex, 10, 1, 3).setValues([[formData.dolar, formData.euro, formData.unit]]);
         sheet.getRange(rowIndex, 8).setValue(formData.productName);
         sheet.getRange(rowIndex, 7).setValue(formData.code || "");
         sheet.getRange(rowIndex, 16, 1, 5).setValues([[formData.currency, formData.discountApplied, formData.quantity, formData.suggestedPrice, formData.iva]]);         
         sheet.getRange(rowIndex, 22).setValue(formData.total);
         sheet.getRange(rowIndex, 30).setNumberFormat("@STRING@").setValue(Utilities.formatDate(currentDate, timeZone, "dd/MM/yyyy HH:mm:ss"));
         
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


/*
* Permite guardar los datos del tipo de cambio 
*/
function updateCurrencyTypeHexp(currencyTypeData){
   
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // fecha actual del sistema
  var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
  
  // referenciamos la hoja de cálculo de estados y localidades
  var spreadsheet = SpreadsheetApp.openById(getResources().settingSheetId);
  
  // Referenciamos la hoja de tipo de cambio 
  var currencyTypeSheet = spreadsheet.getSheetByName("Tipo de cambio");
  
  // Obtenemos el correo del usuario activo 
  var activeUserMail = getUserMail();
  
  // Validamos de que se tenga acceso a al hoja de tipo de cambio 
  if(currencyTypeSheet){
    
    // referenciamos el rango y guardamos los datos 
    currencyTypeSheet.getRange(2, 1, 1, 4).setNumberFormat("@STRING@").setValues([[
      currentDate, 
      activeUserMail,
      currencyTypeData.usdHexp,
      currencyTypeData.eurHexp
    ]]); 
     
    // Retornamos los datos de la ctualización
    return {
      currentDate: currentDate, 
      activeUserMail: activeUserMail,
      usdHexp: currencyTypeData.usdHexp,
      eurHexp: currencyTypeData.eurHexp
    }
  }
    
}

/*
* permite consultar los datos de tipo de cambio actuales 
*/
function getCurrencyTypeData(){

  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // referenciamos la hoja de cálculo de estados y localidades
  var spreadsheet = SpreadsheetApp.openById(getResources().settingSheetId);
  
  // Referenciamos la hoja de tipo de cambio 
  var currencyTypeSheet = spreadsheet.getSheetByName("Tipo de cambio");
  
  // Validamos de que se tenga acceso a al hoja de tipo de cambio 
  if(currencyTypeSheet){
    
    // Consultamos los datos 
    var arrayData = currencyTypeSheet.getRange(2, 1, 1, 4).getValues();
    
    // Retornamos los datos 
    return {
      currentDate: arrayData[0][0] ? arrayData[0][0]: 0 , 
      activeUserMail: arrayData[0][1] ? arrayData[0][1]: 0 ,
      usdHexp: arrayData[0][2] ? arrayData[0][2]: 0 ,
      eurHexp: arrayData[0][3] ? arrayData[0][3]: 0 
     }
  }  
}


/**
 * Función que permite enviar los correo a los de paquetería seleccionados
**/
function sendMailFromParcels(parcelsList) {

  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // fecha actual del sistema
  var currentDate = new Date();
  
  // fecha actual si caracteres especiales
  var currentDateTrim = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "ddMMyyyyHHmmss");
  
  // referenciamos la hoja paquetería
  //var spreadsheet = SpreadsheetApp.openById(resources.acquisitionSheet);
  var spreadsheet = SpreadsheetApp.openById(resources.manufacturerSheetId);
  
  // referenciamos la hoja de paquetería
  var sheet = spreadsheet.getSheetByName("Paqueteria"),
      uuid = "";
  
  // se valida si existe la hoja
  if(sheet){
    
    // Recorremos los registros 
    for(var i = 0; i < parcelsList.length; i++){
      
      // se valida si existe el correo
      if(parcelsList[i].email){
      
        // generamos el correo
        uuid = currentDateTrim + Utilities.getUuid();
        
        // se procede a enviar el correo al usuario
        sendMailFromUpdateDate(parcelsList[i], uuid, resources.webUrl, "P");
        
        // referenciamos el rango
        sheet.getRange(parcelsList[i].rowIndex, 13, 1, 2).setNumberFormat("@STRING@").setValues([[
          "Enviado para actualizar",
          uuid
        ]]);
      }
    }
  }
  
  // retornamos los datos de fabricantes y paquetería
  return getParcelsList(spreadsheet);
}

/**
 * Función que permite enviar los correo a los fabricantes seleccionados
**/
function sendMailFromManufacturers(manufacturerList) {

  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // fecha actual del sistema
  var currentDate = new Date();
  
  // fecha actual si caracteres especiales
  var currentDateTrim = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "ddMMyyyyHHmmss");
  
  // referenciamos la hoja fabricantes
  var spreadsheet = SpreadsheetApp.openById(resources.manufacturerSheetId);
  //var spreadsheet = SpreadsheetApp.openById(resources.acquisitionSheet);
  
  // referenciamos la hoja de fabricantes
  var sheet = spreadsheet.getSheetByName("Fabricantes"),
      uuid = "";
  
  // se valida si existe la hoja
  if(sheet){
    
    // Recorremos los registros 
    for(var i = 0; i < manufacturerList.length; i++){
      
      // se valida si existe el correo
      if(manufacturerList[i].email){
      
        // generamos el correo
        uuid = currentDateTrim + Utilities.getUuid();
        
        // se procede a enviar el correo al usuario
        sendMailFromUpdateDate(manufacturerList[i], uuid, resources.webUrl, "F");
        
        // referenciamos el rango
        sheet.getRange(manufacturerList[i].rowIndex, 23, 1, 2).setNumberFormat("@STRING@").setValues([[
          "Enviado para actualizar",
          uuid
        ]]);
        /*sheet.getRange(manufacturerList[i].rowIndex, 17, 1, 2).setNumberFormat("@STRING@").setValues([[
          "Enviado para actualizar",
          uuid
        ]]);*/
      }
    }
  }
  
  // retornamos los datos de fabricantes y paquetería
  return getManufacturerList(spreadsheet).manufacturerList;
}

/**
 * @desc: Función que envia la notificación una vez se aprueba o rechaza una solicitud
 * @author: Rubén Sánchez
 */
 function sendMailFromUpdateDate(manufacturerData, uuid, webUrl, identifier){
   
   //Referenciamos el contenido de la plantilla
   var contentNotification = HtmlService.createHtmlOutputFromFile("Html_MailTemplateUpdateData").getContent();
   
   //Insertamos el nombre del usuario
   //contentNotification = contentNotification.replace("|*USER_NAME*|", manufacturerData.contact);
   
   //Insertamos el vínculo del formulario de actualización
   if(identifier == "F"){
     contentNotification = contentNotification.replace("|*LINK*|", "https://script.google.com/a/macros/hidro.expert/s/AKfycby5t2PiMMKidE2j49LTaf41FeyuVmxLDmbaltLf5NXLQEvX6fI/exec");
   } else {
     contentNotification = contentNotification.replace("|*LINK*|", webUrl + "?" + identifier + uuid);
   }
   
   // se procede enviar el correo electronico
   MailApp.sendEmail({
     to: manufacturerData.email,
     subject: "Agrocity: Solicitud de actualización de datos",
     htmlBody: contentNotification,
     name: "Adquisiciones"
   });
}

/**
 * Función que permite obtener los datos del usuario activo
**/
function getUserDataFromAcquisition() {
  // se obtiene el correo del usuario
  var userMail = getUserMail();

  // referenciamos la hoja de cálculo de estados y localidades
  var spreadsheetPrice_ = SpreadsheetApp.openById(getResources().acquisitionSheet);
  
  // retornamos el los datos
  return {
    userData: getAccountUserPanel(userMail, "Usuario"),
    conceptList: getConceptListByPrices(spreadsheetPrice_),
    allDataConcepts: getConceptList(), 
    currencyTypeData: getCurrencyTypeData(),// Datos del tipo de cambio para hidroExpert 
    conceptListFilter: getConceptListByFilter_(spreadsheetPrice_)
  }
  
}

/**
 * Función que permite obtener todos las especies
 * <por alguna razon esta funcion no se usa>
**/
function getConceptListByFilter_(spreadsheet) {

  // validamos si no viene información del libro
  if(!spreadsheet){
  
    // referenciamos la hoja de cálculo de estados y localidades
    spreadsheet = SpreadsheetApp.openById(getResources().acquisitionSheet);
  }
  
  // obtenemos la hoja de segun parametro
  var sheet = spreadsheet.getSheetByName("Precios/Fabricantes");
  
  // validamos si existe la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
      
      // retornamos los datos
      return sheet.getRange(2, 7, lastRow - 1, 2).getDisplayValues();      
    }
  }
  
  // retornamos por defecto un array vacio
  return [];
}


/**
 * Función que permite obtener todos las especies
 * <por alguna razon esta funcion no se usa>
**/
function getConceptListByPrices(spreadsheet) {

  // validamos si no viene información del libro
  if(!spreadsheet){
  
    // referenciamos la hoja de cálculo de estados y localidades
    spreadsheet = SpreadsheetApp.openById(getResources().acquisitionSheet);
  }
  
  // obtenemos la hoja de segun parametro
  var sheet = spreadsheet.getSheetByName("Precios");
  
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

/**
 * Función que permite obtener los datos de paquetería y fabricantes
**/
function getAllDataAcquisition(){
  /*return {
    parcels: [],
    manufacturers: [],
    dataListSltSearch: [],
    projectsData: [],
    pricesAnalysis: [] // Se debe de verificar esta información
  }*/
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // referenciamos la hoja de parcelas
  var spreadsheet = SpreadsheetApp.openById(resources.acquisitionSheet);
  //throw spreadsheet.getUrl()
  // referenciamos la hoja fabricantes
  var spreadsheetManufacturer = SpreadsheetApp.openById(resources.manufacturerSheetId);
  
  // retornamos los datos de fabricantes y paquetería
  return {
    parcels: getParcelsList(spreadsheetManufacturer),
    manufacturers: getManufacturerList(spreadsheetManufacturer),
    dataListSltSearch: getListData(resources),
    projectsData: getProjects(resources),
    pricesAnalysis: getAnalysisDataByPrices(spreadsheet) // Se debe de verificar esta información
  };
  
}


/*
* Permite obtener la lista de departamentos, nombre de especie, familia, palabra clave
*/
function getListData(resources){
  
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
function getParcelsList(spreadsheet){
  
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
function getManufacturerList(spreadsheet){
  
  // referenciamos la hoja de fabricantes
  var sheet = spreadsheet.getSheetByName("Fabricantes");
  
  // Objetos con los datos del objeto
  var objDataManufactures = {
    manufacturerList: [],
    manufacturerListObj: {},
    manufacturerFamilies: {}
  }
  
  // se valida si existe la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // se valida si existe mas de una fila
    if(lastRow > 1){
      
      // obtenemos los datos
      var arrayData = sheet.getRange(2, 1, lastRow - 1, 24).getDisplayValues();
      //var arrayData = sheet.getRange(2, 1, lastRow - 1, 18).getDisplayValues();
      
      // variable que determina el listado
      var manufacturerList = [];
      
      // objeto con el nombre de los fabricantes
      var manufacturerListObj = [], families = null, currFamily;
      
      // recorremos el listado de fabricantes
      for(var i = 0; i < arrayData.length; i++){
        
        // agregamos los fabricantes al array
        objDataManufactures.manufacturerList.push({
          id: arrayData[i][1],
          manufacturerName: arrayData[i][4],
          password: arrayData[i][5],
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
        });
        /*objDataManufactures.manufacturerList.push({
          id: arrayData[i][0],
          manufacturerName: arrayData[i][1],
          socialReason: arrayData[i][2],
          rfc: arrayData[i][3],
          familiesOffered: arrayData[i][4],
          taxResidence: arrayData[i][5],
          domicileWinery1: arrayData[i][6],
          domicileWinery2: arrayData[i][7],
          domicileWinery3: arrayData[i][8],
          contact: arrayData[i][9],
          officePhone: arrayData[i][10],
          mobilePhone: arrayData[i][11],
          email: arrayData[i][12],
          otherEmail: arrayData[i][13],
          webUrl: arrayData[i][14],
          lastUpdate: arrayData[i][15],
          state: arrayData[i][16],
          code: arrayData[i][17],
          rowIndex: (i + 2)
        });*/
        
        // Validamos de que la propiedad en el objeto no se encuentre
        if(arrayData[i][4] && !manufacturerListObj[arrayData[i][4]]){
          
          // Agregamos la proiedad al objeto
          objDataManufactures.manufacturerListObj[arrayData[i][4]] = null;
          
          // obtenemos la lista de las familias ofertadas
          families = String(arrayData[i][12]).trim().split(",");
          
          // recorermocas cada una de las familias
          for(var j = 0; j < families.length; j++){
            
            // familia recorrida
            currFamily = String(families[j]).trim();
            
            // validamos que la familia sea diferente de vacio
            if(currFamily){
            
              // validamos si la familia no existe en el objeto para agregarla
              if(!objDataManufactures.manufacturerFamilies[currFamily]){
                objDataManufactures.manufacturerFamilies[currFamily] = {};
              }
              
              // agregamos el fabricante a la lista
              objDataManufactures.manufacturerFamilies[currFamily][arrayData[i][4]] = {
                manufacturerId: arrayData[i][1],
                manufacturerName: arrayData[i][4]
              };
            }
          }
        } 
      }
    }
  }
  console.log(objDataManufactures);
  // retornamos los datos
  return objDataManufactures;
}

/*
* Permite actualizar las familias ofertadas
* @param {Object} objectDataMP Objeto con los datos para la creción
*/
function updateOfertFamilies(recordObject_, newFamilies_){
  
  // Variable a retornar
  var returnObject_ = {
    success: false,
    message: "No fue posible actualizar las familias ofertadas."
  };

  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // referenciamos la hoja fabricantes
  var spreadsheet = SpreadsheetApp.openById(resources.manufacturerSheetId);

  // Obtenemos la hoja
  var sheet_ = spreadsheet.getSheetByName("Fabricantes");

  // Validamos de que se tenga acceso a la hoja
  if(sheet_){

    // Actualizamos la familia
    sheet_.getRange("M" + recordObject_.rowIndex).setNumberFormat("@STRING@").setValue(newFamilies_);
    SpreadsheetApp.flush();

    // Actualizmos los valores
    returnObject_.success = true;
    returnObject_.message = "Se han actualizado las familias ofertadas para el fabricante <b>" + recordObject_.manufacturerName + "</b>.";
    returnObject_.list = getManufacturerList(spreadsheet).manufacturerList;
  }
  
  // Retornamos los datos de respuesta
  return returnObject_;
}


/*
* Función que permite guardar un fabricante o un producto 
* @param {Object} objectDataMP Objeto con los datos para la creción
*/
function createManufacturerOrParcel(objectDataMP){

  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // referenciamos la hoja fabricantes
  var spreadsheet = SpreadsheetApp.openById(resources.manufacturerSheetId);
  //var spreadsheet = SpreadsheetApp.openById(resources.acquisitionSheet);

  //validamos que existan los datos para la creación 
  if(objectDataMP){
    
    // Validmos si se debe crea un fabricante 
    if(objectDataMP.creationType.typeEnglish == "manufacturers"){
      
      // Guardamos los datos del fabricatne 
      return saveDataManufacturer(objectDataMP, spreadsheet.getSheetByName("Fabricantes"), spreadsheet);
      
      // Validamos si se debe crear un paquete 
    }else{
      
      // Guardamos los datos del paquete
      return saveDataParcel(objectDataMP, spreadsheet.getSheetByName("Paqueteria"), spreadsheet);
      
    } 
  }
    
}


/*
* Función que guarda la información del fabricante
* @params {Objeto} objectDataMP Datos para guardar el fabricante
* @params {reference} sheet Referencia a la hoja de fabricantee
*/
function saveDataManufacturer_OLD(objectDataMP, sheet, spreadsheet){
   
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
   return getManufacturerList(spreadsheet).manufacturerList;
}

/*
* Función que guarda la información del fabricante
* @params {Objeto} objectDataMP Datos para guardar el fabricante
* @params {reference} sheet Referencia a la hoja de fabricantee
*/
function saveDataManufacturer(objectDataMP, sheet, spreadsheet){
   
   // Validamos de que se tenga acceso a la hoja
   if(sheet){
    
     // Obtenemos la ultima fila 
     var lastRow = sheet.getLastRow() + 1;
     
     // le damos forma a la hoja
     var arrayData = [
       'Activo',
       lastRow,       
       objectDataMP.email,
       '',
       objectDataMP.nameMP,
       objectDataMP.password
     ];
     
     // guardamos los datos base del usuario
     sheet.getRange(lastRow, 1, 1, 6).setNumberFormat("@STRING@").setValues([arrayData]);
     
     // Guardamos el nombre de contacto
     sheet.getRange(lastRow, 16).setNumberFormat("@STRING@").setValue(objectDataMP.contactName);
     
     // se procede a enviar un correo electronico al usuario que se crea
     sendMailNotifyAccess(objectDataMP);
     
   }
   
   // Retornamos la lista de fabricantes
   return getManufacturerList(spreadsheet).manufacturerList;
}

/**
 * Función que permite enviar un correo al usuario indicando cuales son los datos de acceso
 */
 function sendMailNotifyAccess(manufacturerData){
   
   //Referenciamos el contenido de la plantilla
   var contentNotification = HtmlService.createHtmlOutputFromFile("Html_MailTemplateAccess").getContent();
   
   // definimos la variable para el nombre del fabricante
   var manufacturerName = 'Fabricante';
   
   // se valida si existe el nombre del fabricante con la propiedad "nameMP"
   if(manufacturerData.nameMP) {
     
     // se temoa este valor cuando se crea un nuevo fabricante
     manufacturerName = manufacturerData.nameMP;
     
   } else if(manufacturerData.manufacturerName){
     
     // se temoa este valor cuando se envia el correo desde la lista de fabricantes
     manufacturerName = manufacturerData.manufacturerName;
   }
   
   //Agregamos el nombre del fabricante
   contentNotification = contentNotification.replace("|*MANUFACTER_NAME*|", manufacturerName);
   
   //Agregamos el correoo de acceso
   contentNotification = contentNotification.replace("|*EMAIL*|", manufacturerData.email);
   
   //Agregamos la contraseña
   contentNotification = contentNotification.replace("|*PASSWORD*|", manufacturerData.password);
   
   // se procede enviar el correo electronico
   MailApp.sendEmail({
     to: manufacturerData.email,
     subject: "Agrocity: Datos de acceso portal de fabricantes",
     htmlBody: contentNotification,
     name: "Adquisiciones"
   });
}

/*
* Función que guarda la información del paquete
* @params {Objeto} objectDataMP Datos para guardar el paquete
* @params {reference} sheet Referencia a la hoja de paquete
*/
function saveDataParcel(objectDataMP, sheet,spreadsheet){

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
   return getParcelsList(spreadsheet);
}


/*
* Funcioón que permite obtener los proyectos con sus respectivas especies
*
*/
function getProjects(resources){
  
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
function updateCurrencyType(manufacturerData){

 // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // referenciamos la hoja solicitudes de proyectos
  var spreadSheet = SpreadsheetApp.openById(resources.acquisitionSheet);
  
  // validamos de que se tenga acceso a la hoja 
  if(spreadSheet){
     
     // Referenciamos la hoja de Precios / Fabricantes
     var manufacturerPricesSheet = spreadSheet.getSheetByName("Precios/Fabricantes");
     
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
          return getAllDataAcquisition(); //getAnalysisDataByPrices(spreadSheet);
        }
     }
   }

} 


/**
 * Función que pemite obtener el listado de familias disponibles hasta la fecha
**/
function getFamiliesList(){

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
function updateAssignFamilies(formObject) {
  
  // Se invoca la función donde se referencia las variables globales de la aplicación
  var resources = getResources();
  
  // Se declara una variable para referenciar el lbro donde estan los fabricantes
  var sheet, recordArray, spreadsheet = SpreadsheetApp.openById(resources.manufacturerSheetId);
  
  // referenciamos la hoja de fabricantes
  var sheet = spreadsheet.getSheetByName(resources.nameUsers);
  
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
    throw "La hoja 'Fabricantes' no existe.";
  }
  
    
  // retornamos los datos de fabricantes y paquetería
  return getManufacturerList(spreadsheet).manufacturerList;
}