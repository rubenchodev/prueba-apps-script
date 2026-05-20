/**
 * Función que obtiene el id de los archivos base de cotizaciones, configuración y clientes
 * @return  getResources().ordersHistorySheetId
 */
function getResources(){
  
  // retornamos la información de ids de hojas
  return {
    requestSheetId: "1dQqhGgL7U-nV7ikUPmwOO6J7FLmhouUzOn6BBs-52ek", // OK
    renovationSheetId: "15YfkeWoC3_EtROuo3TRz5j_S5QdvYTSTnWGqnnH1Uw8", // OK
    clientSheetId: "18aVLMJAoybn9iHL0ii2iJdDmcBwjissyn1QGgi4NkwE", // OK
    stateSheetId: "1dsz76FUS7p0Lmd5kQaabJN0XSpscJmzJaelEByhiFfY",  // hoja de estados, municipios y localidades - OK
    settingSheetId: "1o_uevNLuwxrOm_hiCxPd0Gj5FB5gPWYMKw8-yCG08Do", // hoja de datos parametrizables - OK
    templateId: "14CX3rZwZ64tTGrrCenHFLZ-IVLceS2UnYwTv_Pd3bjw", // plantilla de la cot. - OK
    templateRenovationId: "1_L9mtuZwznjdpv3ST3U4AHTao7uTrybwIHLV5wabngI", // plantilla de la cotización de refacción - OK
    userFileId: "1Mt2PBHItAOhco8zZwvyB9gu83rDMq47_GUaq3XjrfZA", // hoja donde se parametriza los usuarios (FALTA) - OK
    patternsSheetId: "1VOLpB_Vsms_PID_j9O8H5qaLqweRw_RQse_UTF7N9jQ", // libro donde se almacena los patrones del proyecto - OK
    federalTemplate: "11qMngopPW7OrqnPxVuLWmoc3oH2AFRPcEU5xlXrb_VE", // Plantilla para crear hoja de Federal - OK
    statalTemplateId: "1hTVKr1wNgZ8GcaPpqzLt1jUrsAeh7DcLJLk8ZjYg8LY", // Plantilla para crear hoja de Federal - OK
    statalFolderId: "1B4IbuYEhZMNkrHchkfnWCkW6q4nP55OU", // ID de la carpeta estatal - OK
    technificationSheetId: "1ifHVSrddnGQ9eMFlYaLtcPrdbqdj-9APB94BPJHpOqI", // Id de la hoja de tecnificación - OK
    
    acquisitionSheet: "1wbPPZ0DLKwZtGTZGYDM8UYGN1pIRuegemx9tkJJ_x0E", // Hoja de fabricantes y paquetería - OK
    
    ordersHistorySheetId: "1nEQAKymGW6O3oo8Eg-KActF0wH0oMFQmYkEsLYJ0cSc", // Id de la hoja de historial de órdenes - OK
    
    //ordersTemplateId: "1AzstoE-spZf-QG1_1pwXROgoP3KJTjg0woW_FJPMRtA", // plantilla en Google Doc de una Orden
    ordersTemplateId: "1EuJfiP3OFkKS5tlml4qxh_6Tvp0nm9zUiWXM4UZ2BNw", // plantilla en Google Doc de una Orden - OK
    
    ordersFolderId: "14gBE6i2MY8xmnC693qzh88bGiQrxafRH", // Carpeta donde se almacenarán todas las órdenes - OK
    ordersSentFolderId: "1Hzx1Se_qn_eo_-EmpXmHCZvBwGMKdMqC", // Carpeta donde se almacenan las ordenes enviadas - OK
    webUrl:"https://script.google.com/macros/s/AKfycbzMabp2eYJrI2d2yzX1YUJXhUCINoLwJ1kiLYVCNaZ7VmLhG2g_OLD/exec", // Url base del formulario de actualización de precios - OK
    
    infoCatalogSheetId: "1xszvNRP6pnsxpYSel_20ynWuAQnha_nyEIDIyPr1X7k", // Id de la hoja donde se almacena la información del catálogo - OK
    workParameters: "1fZo9TA7u7jbzOYIS1a66hz1UBuHfKqgV673qRybwwAQ", // Id de la Hoja de parametros de obra - OK
    cardTemplateId: "1m0zprDZ76T0xV6umTZ-hnTpEmkXQbC_NQnmPMQ9i-6U", // Id de la plantilla de las tarjetas - OK
    cardFolderId: "1X0fOPholmKzZuTIaKVZCAZ-mtoBBaDAm", // Id de  carpeta donde se alojaran las tarjetas - OK
    catalogTemplateId: "16_CInmuHxS2p4mdfoVNsCmVDuUuLPEHm8k-J3lC0-oU", // Id de la plantilla para el catálogo - OK
    catalogFolderId:"1ElES3EGyE0lr0rDtNkOk9K8zehG5jneU", // Id de  carpeta donde se alojaran los catálogo - OK
    
    forderDefaultId: "1HO_WcZGO9NvPALb8tdP4WBTajFHpyLMH", // Es para almacenar aquellos archivoS que x algun motivo no tiene o no se encontro en donde guardar - OK
    validationSheetId: "1Xd0fDTPRzQ3FUCYxkd5H4qaXySmlZBi_xYlKgXDK3pE", // hoja de los datos de las validaciones - OK
    imageFolderId: "1DgKUU54nQ5ssurGirNyqnuzAns_ngvIG", // carpeta donde se almacenan las imagenes de cada una de las especies -OK
    
    manufacturerSheetId: "1KYx8q4iAnE74R76bGU7wLGzUW3uDf8SVlA6aJ2m1YQA", // Id de la hoja donde se administran la lista de fabricantes - OK
    
    nameUsers: "Fabricantes", // Nombre de la pestaña de usuarios
    
    conceptSheetName: "Conceptos", // Nombre de la hoja de conceptos
    
    
    // Configuración para el panel de la competencia.
    competitorsSheetId: "1bNMKoE7idoqvkyzuBFpwfIuXt53zfhpvGxST3ei1Wxg", // Id de la hoja donde se administran la lista de competidores - OK
    competitorsPricingSheetId: "1ETkiDSpNPHrgPJFymGjX5MZtGOHD4M5X3taUmuock3A", // Id del libro de precios de competidores - OK
    competitorsSheetName: "Competidores", // Nombre de la pestaña de competidores
    competitorsPricingSheetName: "Precios/Competidores", // Nombre de la pestaña de usuarios

    // Datos de variables por defecto en una cotización de REFACCIÓN
    renovationTemplateId: "1zOQNZIGeu47mR5NTCSCev4fNQqmQHdCkx03mhTK5s24", // - OK
    renovationFolderId: "1v7_1a-wH-545OpTeFeNvBa5d7T2a-Zpa", // - OK

    // Variable para las hoja de configuración de cruceros
    crossingSheetId: "1LWOafPa9WMhGmnJIP6ijEStU9EkwIbjZ3zBoTdNn4D8",
    diameterSheetName: "Diámetros",
    diameterValveSheetName: "Diámetros de válvula",
    refConceptSheetName: "Referencia de conceptos",
    irrigationSheetName: "Sistema de riego",
    qualifySheetName: "Campos del cuantificador",
    thongSheetName: "Cintas",
    microSprinklerSheetName: "Microaspersores",
    dropperSheetName: "Goteros",

    // Carpeta base de cotizaciones de proyectos
    projectMainFolderId: "1nVMyn97tBslVJblxmwGq7YF46O7CyVh0",
  };
}
///**
// * Función que obtiene el id de los archivos base de cotizaciones, configuración y clientes    *****<<<<<AMBIENTE DE SINOVA>>>>>>*********
// * @return  getResources().ordersHistorySheetId 
// */
//function getResources(){
//  
//  // retornamos la información de ids de hojas
//  return {
//    requestSheetId: "1GeyoXBjGh8fBFiUAtquAJK-nnC3RJfMmxieQ2NC95Y8",
//    renovationSheetId: "1Oe-uCN4fNoeeoKMTuoR73WOaj6eLVG8TXCtC1-sWD-c",
//    clientSheetId: "1zmK8-vb-oj0o5E_ejYOyQMjQIsLyL8PglRMqYXq-3fA",
//    stateSheetId: "1c3jFLErZJIEt2v7pGyJWj0nlS50-53tUypAv6nr3zBU",  // hoja de estados, municipios y localidades
//    settingSheetId: "1S0ns_bAdd-mciPBDgq2n6dQ89dZ6yDg1FiYNk5RwHLU", // hoja de datos parametrizables
//    templateId: "1Do21BXgONbm7qdsAASTXHdOpOTtdhf9Ave8USLCYgw0", // plantilla de la cot.
//    templateRenovationId: "1Osv2-dwKnjVzDhZJpHDP0tGVcAHaaAKc4JbqOfl__2w", // plantilla de la cotización de refacción
//    userFileId: "1EsrtW6V9XA9Z6mg0vYpWFoxAUwKILxnvhmagEDRB9ss", // hoja donde se parametriza los usuarios (FALTA)
//    patternsSheetId: "1oHn9l4Am-_GRqMHLORgTv3FPcqlonXa3PFAKdwaRh1g", // libro donde se almacena los patrones del proyecto
//    federalTemplate: "1zoRpSLU2hmGp6GWWyzd6T5vjbFe8GUkRpajvZc4FUtA", // Plantilla para crear hoja de Federal
//    technificationSheetId: "1oxS_5qZX2Zz_19XBQlZquwBFyEPFHxVhCWc-MLZWqqI", // Id de la hoja de tecnificación
//    acquisitionSheet: "14oy0rY-1aqKmK0SlJG3i3EOyfraSOcFs-Uui4rjRuw4", // Hoja de fabricantes y paquetería
//    ordersHistorySheetId: "1AToVn0LADVLl7CAaCSOR79o4b9WditVeOHIByP-naj4", // Id de la hoja de historial de órdenes
//    ordersTemplateId: "1YTO-JNnSxEXF1iB1lM0cjoUOrWnU5xHAeRraBsL_EpI", // plantilla en Google Doc de una Orden
//    ordersFolderId: "1dXHDMA2bhsh0nr3_e9E7gbiOBoYTMnee", // Carpeta donde se almacenarán todas las órdenes
//    ordersSentFolderId: "1Sl38oVCNZ_o1aAKkoPRkRFrNIFS186D5", // Carpeta donde se almacenan las ordenes enviadas
//    webUrl:"https://script.google.com/a/sinova.co/macros/s/AKfycbxzaGGyHs3XBm_d7YLdiHbkviwz7IoFMCDGdutrd0aR/dev", // Url base del formulario de actualización de precios
//    
//    infoCatalogSheetId: "14aqc4fOwxuamqWH07G_2DNHe2v3FBA9qrVesJzHLLXA", // Id de la hoja donde se almacena la información del catálogo
//    workParameters: "1Bl8H3Knfu7t4wjU32IOYFCDywyKAx5NntK5crB-lG8s", // Id de la Hoja de parametros de obra
//    cardTemplateId: "1LR2C09LAMCG_f2Yh5Ii8zU5TWpSGlibVHjsHHI38hfM", // Id de la plantilla de las tarjetas
//    cardFolderId: "1O5-MpfHQ3Pva4gJeogzB3YjWW3nPS4NN", // Id de  carpeta donde se alojaran las tarjetas
//    catalogTemplateId: "1TXjp-mEYZiycslOzWJe7WsJWIrGGZjPGnNGmbe-vKuA", // Id de la plantilla para el catálogo
//    catalogFolderId:"1wNhfdvKtFPIkuCUut4uF6vGXIQIKlW48", // Id de  carpeta donde se alojaran los catálogo
//    
//    forderDefaultId: "1kwgocoONIfWHRZzE1xKa947zP30XcMzT", // Es para almacenar aquellos archivoS que x algun motivo no tiene o no se encontro en donde guardar
//    validationSheetId: "18hRygrhYr0i7H3DOkJ3G48gnzZ-c9R7z0-iJZWiYLw8", // hoja de los datos de las validaciones
//    imageFolderId: "1fVl1knuZPZcqXPwGW4qN26ISENe0mcKv", // carpeta donde se almacenan las imagenes de cada una de las especies
//    
//    manufacturerSheetId: "1urKs9uFaCXjKRAYENsDWwcJjiz5KDuYgFGvDUdCU8Is" // Id de la hoja donde se administran la lista de fabricantes
//  };
//}

/**
 * Función que consulta el correo del usuario activo
 * @return 
 */
function getUserMail(){
//  return "anteproyecto@hidro.expert"
//return "carlos.almanza@sinova.co"
//  return "suministros@hidro.expert";
//  return "soporte@hidro.expert";
  // retornamos el correo de usuario activo
  return Session.getActiveUser().getEmail();
  
}

/**
 * Función que permite obtener la url del script 
 */
function getUrlScript(){
  
  // retornamos la url del script
  return "https://script.google.com/macros/s/AKfycbxAvZVBYcnb9kvpnEQsz1H2H3rjsA3Bzz9RS9W7XoE/dev";
  return "https://script.google.com/a/hidro.expert/macros/s/AKfycbzu4bt-7GmSKWMwU202sLaALbqlTCdXd5UXDgEXeemYdsop_H5u/exec";
  
}

/**
 * Función que consulta la información del usuario en el Panel de Administración del dominio.
 * @param {String} email: Correo electrónico del usuario que se desea consultar.
 * @return 
 */
function getAccountUserPanel(email, defaultName){

  // definimos el objeto que deseamos retornar
  var userObject = {
    email: email, // Nombre del usuario por defecto
    photoUrl: "http://ssl.gstatic.com/accounts/ui/avatar_2x.png",
    name: defaultName,
    fullName: defaultName
  };
  
  // encerramos en un try - catch la consulta de la información del usuario
  try {
    
    // Se consulta los datos del usuario en el Panel de Administración, mediante el correo electrónico "email".
    // Adicionalmente se define los parámetros del objeto
    var response = AdminDirectory.Users.get(email, {
      projection: "full", // Conjunto de campos a retornar (todos los campos asociados al usuario)
      viewType: "domain_public" // Vista de datos públicos del dominio que otros usuarios pueden consultar
    });
    
    // Se valida si existe la propiedad del nombre del usuario "name"
    if (response.name) {
    
      // Se asigna a la propiedad "name" del objeto el nombre completo, o solo el nombre del usuario
      userObject.name = response.name["givenName"] || email;
      
      // Se asigna a la propiedad "name" del objeto el nombre completo, o solo el nombre del usuario
      userObject.fullName = response.name["fullName"] || email;
    }
    
    // Se valida si existe la propiedad de la imagen de usuario "thumbnailPhotoUrl"
    if (response.thumbnailPhotoUrl) {
      
      // agregamos el nombre del usuario
      userObject.photoUrl = response.thumbnailPhotoUrl;
      
    }
  } catch (ex) {
    
    // mostramos en console el error
    saveErrorLog(ex);
  }
  
  // retornamos por defecto el objeto base
  return userObject;
}

/**
 * Función que consulta la información del usuario en el Panel de Administración del dominio.
 * @param {String} retorna el nombre de usuario
 * @return 
 */
function getUserName(email){
  
  // encerramos en un try - catch la consulta de la información del usuario
  try {
    
    // Se consulta los datos del usuario en el Panel de Administración, mediante el correo electrónico "email".
    // Adicionalmente se define los parámetros del objeto
    var response = AdminDirectory.Users.get(email, {
      projection: "full", // Conjunto de campos a retornar (todos los campos asociados al usuario)
      viewType: "domain_public" // Vista de datos públicos del dominio que otros usuarios pueden consultar
    });
    
    // Se valida si existe la propiedad del nombre del usuario "name"
    if (response && response.name) {
      
      // agregamos el nombre del usuario
      return response.name.fullName;
      
    }
  } catch (ex) {
    
    // mostramos en console el error
    saveErrorLog(ex);
    
  }
  
  // retornamos por defecto vacio
  return "";
}

/**
 * Función que permite validar si el usuario se encuentra registrado y a la ves obtiene el rol del mismo
**/
function validateUserAccess(userMail){
  
  // validamos is no existe un correo
  if(!userMail){
    
    // obtenemos el correo del usuario
    userMail = getUserMail();
    
  }
  
  // variable de respuesta
  var responseObject = {
    isValid: false,
    userRole: "",
    onlyOneRole: true,
    state: ""
  };
  
  // obtenemos los datos de recursos
  var resources = getResources();
  
  // referenciamos la hoja de usuarios
  var userSpreadsheet = SpreadsheetApp.openById(resources.userFileId);
  
  // referenciamos la hoja de usuarios y roles
  var userSheet = userSpreadsheet.getSheetByName("Usuarios y roles");
  
  // validamos si existe la hoja
  if(userSheet){
    
    // obtenemos la ultima fila
    var lastRow = userSheet.getLastRow();
    
    // se valida si existe mas de una fila
    if(lastRow > 1){
      
      // definimos un objeto para determinar el rol en ingles
      var rolesObj = {
        "Ventas": "sales",
        "Ingeniería": "engineer",
        "Dirección ejecutiva": "executiveDirection",
        "Entorno":"environment"
      };
      
      // obtenemos los datos
      var arrayData = userSheet.getRange(2, 1, lastRow - 1, 3).getValues();
      
      // recorremos el listado de correo parametrizados
      for(var i = 0; i < arrayData.length; i++){
      
        // Validamos si el usuario es igual al que se esta recorreiendo
        if(userMail == arrayData[i][0]){
          
          // Cambiamos el estado y definimos el nombre del rol          
          responseObject.isValid = true;
          responseObject.userRole = rolesObj[arrayData[i][1]];
          
          // salimos del cilco
          break;
          
        }
      }
    }
  }
  
  // referenciamos la hoja de adquisiciones
  var acquisitionSheet = userSpreadsheet.getSheetByName("Adquisiciones");
  
  // validamos si existe la hoja
  if(acquisitionSheet){
    
    // obtenemos la ultima fila
    var lastRowAcquisition = acquisitionSheet.getLastRow();
    
    // se valida si existe mas de una fila
    if(lastRowAcquisition > 1){
      
      // obtenemos los datos
      var userList = acquisitionSheet.getRange(2, 1, lastRowAcquisition - 1, 1).getValues();
      
      // recorremos el listado de correo parametrizados
      for(var i = 0; i < userList.length; i++){
  
        // Validamos si el usuario es igual al que se esta recorreiendo
        if(userMail == userList[i][0]){
          
          // se valida si aun no tiene el nombre del usuario
          if(!responseObject.isValid){
            
            // Cambiamos el estado y definimos el nombre del rol          
            responseObject.isValid = true;
            responseObject.userRole = "acquisition";
            
          } else {
            
            // cambiamos el estado para determinar que no solo tiene un rol
            responseObject.onlyOneRole = false;
            
          }
          
          // salimos del cilco
          break;
          
        }
      }
    }
  }
    
  // referenciamos la hoja de entorno
  var environmentSheet = userSpreadsheet.getSheetByName("Entorno");
  
  // validamos si existe la hoja
  if(environmentSheet){
    
    // obtenemos la ultima fila
    var lastRowEnv = environmentSheet.getLastRow();
    
    // se valida si existe mas de una fila
    if(lastRowEnv > 1){
      
      // obtenemos los datos
      var userListEnv = environmentSheet.getRange(2, 1, lastRowEnv - 1, 2).getValues();
      
      // recorremos el listado de correo parametrizados
      for(var i = 0; i < userList.length; i++){
  
        // Validamos si el usuario es igual al que se esta recorreiendo
        if(userMail == userList[i][0]){
          
          // se valida si aun no tiene el nombre del usuario
          if(!responseObject.isValid){
            
            // Cambiamos el estado y definimos el nombre del rol          
            responseObject.isValid = true;
            responseObject.userRole = "environment";
            responseObject.state = userList[i][1];
            
          } else {
            
            // cambiamos el estado para determinar que no solo tiene un rol
            responseObject.onlyOneRole = false;
            
          }
          
          // salimos del cilco
          break;
          
        }
      }
    }
  }
  
  // retornamos un objeto con la información requerida
  return responseObject;
}
/**
* Función encargada de establecer un formato al consecutivo de una solicitud
* {String} consecutive Número del consecutivo
*/
function formatConsecutive(consecutive){
 
 // variable que guarda los cero que hacen falta para completar el consecutivo
 var stringCeros = "";
 
 // recorremos la cantidad de ceros que hace falta
 for(var i = 0; i < (5 - String(consecutive).length); i++) stringCeros += "0";
 
 // retornamos el consecutivo
 return stringCeros + consecutive;
}
 
 
/**
* Función encargada de obtener el valor del consecutivo, incrementarlo en uno y retornarlo
* {String} propertyName Nombre de la propiedad
*/
function getConsecutive(propertyName){
 
 // validamos si existe la propiedad
 if(!propertyName){
   
   // asignamos la propiedad del consecutivo del contrato
   propertyName = "DA_Consecutive";
 }
 
 // obtenemos las propiedades del script
 var scriptProperties = PropertiesService.getScriptProperties();
 
 //Obtener un bloqueo de script, porque estamos a punto de modificar un recurso compartido
 var lockService = LockService.getScriptLock();
 
 // Espeamos 30 segundos para que finalicen los procesos anteriores
 lockService.waitLock(30000);
 
 // Referenciamos el consecutivo actual
 var consecutive = Number(scriptProperties.getProperty(propertyName));
 
 // Damos formato al consecutivo
 var formattedConsecutive = formatConsecutive(consecutive);
 
 // Incrementamos en 1 el valor de la propiedad
 scriptProperties.setProperty(propertyName, consecutive + 1);
 
 // Liberamos el bloqueo del servicio
 lockService.releaseLock();
 
 // Retornamos el consecutivo generado
 return formattedConsecutive;
}

/**
 * Función que permite convertir una fecha new Date() a un formato "12 de Junio de 2019"
**/
function formatDate(date){
  
  // Listado de meses
  var months = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  
  // retornamos la fecha
  return date.getDate() + " de " + months[date.getMonth()] + " de " + date.getFullYear();
}