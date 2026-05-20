
/**
* Carga el menu de ejecución en la hoja de calculo para las funciones de carga y
* ejecución de celdas activas en el evento de abrir el documento."
*/
function doGet(request) {    
  
  // variable para administarr la interfaz que se desea mostrar
  var htmlService = "";
  
  // obtenemos el correo del usuario
  var userMail = getUserMail();
  
  // obtenemos los datos de validación
  var userAccess = validateUserAccess(userMail);
  
  console.log(userAccess);
  // Titulo de la pestaña del hidrocotizador
  var title = "HidroCotizador";
  
  // validamos si el usuario esta en el listado de usuario parametrizados
  if(userAccess.isValid){
    
    // se valida si el rol es "acquisition"
    if(userAccess.userRole == "acquisition" || (request.queryString == "4cqu151t10n" && userAccess.onlyOneRole == false)){
      
      // referenciamos el archivo base
      htmlService = HtmlService.createTemplateFromFile("Html_MainAcquisition");
      
      // Establecemos el nombre de la pestaña 
      title = "Adquisiciones";
      
    } else if(userAccess.userRole == "environment" || (request.queryString == "3nv1r0nm3n7" && userAccess.onlyOneRole == false)){
      
      // referenciamos el archivo base
      htmlService = HtmlService.createTemplateFromFile("Html_MainEnvironment");
      
      // Establecemos el nombre de la pestaña 
      title = "Entorno";
      
    } else if(userAccess.userRole == "environment" || (request.queryString == "3r1ck" && userAccess.onlyOneRole == false)){
      
      // referenciamos el archivo base
      htmlService = HtmlService.createTemplateFromFile("Html_MainEnvironment_Erick");
      
      // Establecemos el nombre de la pestaña 
      title = "Entorno";
      
    } else {
      // referenciamos el archivo base
      htmlService = HtmlService.createTemplateFromFile("Html_Main");
      
      // Establecemos el nombre de la pestaña 
      title = "Presupuesto";
    }
    
    // agregamos el scriplet con el rol del usuario
    htmlService.userRole = userAccess.userRole;
    
    // agregamos el scriplet el estado si pertenece a uno o varios roles
    htmlService.onlyOneRole = userAccess.onlyOneRole;
    
    // agregamos el scriplet el link del usuario
    htmlService.link = getUrlScript();
    
  } else {
    
    // referenciamos el archivo de no acceso
    htmlService = HtmlService.createTemplateFromFile("Html_DenyAccess");
    
    // agregamos el scriplet con la cuenta de usuario
    htmlService.userMail = userMail;
    
  }
  
  // retornamos la interfaz
  return htmlService.evaluate()
  .setTitle(title)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setFaviconUrl("https://agrocity.mx/cdn/shop/files/favicon-agrocity_32x32.png")
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
* @desc: Función encargada de obtener el contenido de un archivo HTML del proyecto.
* @params: {String} fileName: nombre del archivo HTML.
* @params: {Boolean} isJSCode: es código JavaScript entre etiquetas <script>.
* @params: {Boolean} evaluate: quiere decir que se requiere primero evaluar el codigo antes de obtener el mismo.
* @params: {Array} [jsFile]: Array de archivos JavaScript a incluir.
* @return: {String} Retorna el contenido del archivo. @example: includeHtmlFile('StyleSheet.html') @example: includeHtmlFile('JSModule.html', true)
*/
function includeHtmlFile(fileName, isJsCode, evaluate, params) {
  
  // verificamos si la variable "fileName" es un string la convertimos en un array
  if(typeof(fileName) == "string"){
  
    fileName = [fileName];
  }
  
  // creamos la variable para alojar el html
  var htmlCode = "";
  
  // recorremos cada uno de los archivos a loscuales se debe ontener la el contenido
  for(var i = 0; i < fileName.length; i++){
    
    // verificamos  si el parametro "evaluate" es verdadero para evaluar el codigo y de lo contrario solo obtiene el contenido
    if (evaluate) {
      
      // obtenemos el contenido del archivo
      var htmlCodeTemp = HtmlService.createTemplateFromFile(fileName[i]);
      
      // verificamos si trae un nombre de archivo
      if(params){
        // agragmos el nombre del archivo como un scriplets
        htmlCodeTemp.params = params;
        
      }
      
      //se evalua el contenido y retorna el contenido del archivo
      htmlCode += htmlCodeTemp.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).getContent();
      
    }else{
    	
      //solo hacemos referencia al archivo y obtenemos el contenido
      htmlCode += HtmlService.createHtmlOutputFromFile(fileName).setSandboxMode(HtmlService.SandboxMode.IFRAME).getContent();
      
    }
    
    // verificamos  si el contenido que se obtuvo es de tipo Javascript
    if (isJsCode) {
      
      // eliminamos los tags "<script>,</script>"
      htmlCode = htmlCode.replace(/(<script>|<\/script>)/g, "");
      htmlCode = htmlCode.replace(/(^[ \t]*\n)/g, "");
    }
  }
  
  // Se retorna el codigo
  return htmlCode;
}

/**
 * Función que permite obtener los datos base de la herramienta
**/
function getClientList(fileId){

  // referenciamos la hoja de cálculo de clientes
  var spreadsheet = SpreadsheetApp.openById(fileId);
  
  // obtenemos la hoja de "Clientes"
  var clientSheet = spreadsheet.getSheetByName("Clientes");
  
  // validamos si existe la hoja
  if(clientSheet){
    
    // obtenemos la ultima fila
    var lastRow = clientSheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
      
      // retornamos los datos de los clientes
      var allData = clientSheet.getRange(2, 1, lastRow - 1, 7).getValues();
      
      // variable para almacenar los clientes
      var returnData = {
        fieldData: {},
        clients: {}
      };
      
      // recorremos cada uno de los registros y los filtramos por usuario
      for(var i = 0; i < allData.length; i++){
        
        // agregamos los datos en el objeto para el campo de busqueda
        returnData.fieldData[allData[i][0]] = null;
        
        // agregamos los datos en el objeto general
        returnData.clients[allData[i][0]] = {
          id: i,
          text: allData[i][0],
          federalRecord: allData[i][1],
          agent: allData[i][2],
          residence: allData[i][3],
          saleContact: allData[i][4],
          telephone: allData[i][5],
          email: allData[i][6]
        };
        
      }
      
      // retornamos los datos
      return returnData;
      
    }
  }
  
  // retornamos por defecto un array vacio
  return {
    fieldData: {},
    clients: {}
  };
  
}

/**
 * Función que permite obtener los datos base de la herramienta
**/
function filterRequestByMail(userMail, fileId, userRole, year){
 
  // Se valida si no existe el años
  if(!year){
    year = new Date().getFullYear();
  }
  
  // referenciamos la hoja de cálculo de solicitudes de cotizaciones
  var spreadsheet = SpreadsheetApp.openById(fileId);
  
  // obtenemos la hoja de "Cotizaciones"
  var quotationSheet = spreadsheet.getSheetByName("Cotizaciones");
  
  // validamos si existe la hoja
  if(quotationSheet){
    
    // obtenemos la ultima fila
    var lastRow = quotationSheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
      
      // obtenemos todos los datos de la hoja
      var allData = quotationSheet.getRange(2, 1, lastRow - 1, 16).getValues();
      
      // variable para almacenar la solicitudes
      var requestData = [],
          requestObj = {},
          relationsObj = null,
          relationArray = [];
      
      // recorremos cada uno de los registros y los filtramos por usuario
      //for(var i = 0; i < allData.length; i++){
      for(var i = allData.length - 1; i >= 0; i--){
      
        // validamo si el correo el usuario activo es el mismo de la cotización recorrida en el momento
        if((allData[i][1] == userMail || userRole == "executiveDirection") && allData[i][9] == "Activa" && allData[i][3].indexOf(year) != -1){
          
          // se valida si la propiedad "relationsObj" no tiene valor
          if(relationsObj == null){
        
            // obtenemos los datos de relación de versiones
            relationsObj = getVersionsObj(spreadsheet);            
          }
       
          // obtenemos el objeto
          requestObj = {
            id: allData[i][0],
            name: allData[i][2],
            date: allData[i][3],
            applicant: allData[i][4],
            location: allData[i][5],
            typeIrrigation: allData[i][6],
            proposal: allData[i][7],
            documentId: allData[i][8],
            structures: {
              "Proyecto": {
                structure: allData[i][10]
              }
            },
            mainSheetName: allData[i][11],
            diplicateState: allData[i][12],
            sheets: ["Proyecto"],
            chosenFactor: Number(allData[i][13]),
            federalFileId: allData[i][14],
            infoCatalogSheetId: allData[i][15],
            row: Number(i) + 2
          };
          
          // se valida si se tiene un versionado en la solicitud
          if(requestObj.diplicateState == "SI"){
            
            // obtenemos los datos de la propiedad
            relationArray = relationsObj[requestObj.id];
            
            // se valida si existe datos de la relación actual
            if(relationArray){
              
              // se recorre cada una de las versiones
              for(var j = 0; j < relationArray.length; j++){
                
                // agregamos la información de la versión
                requestObj.structures[relationArray[j].sheetName] = relationArray[j];
                
                // agregamos el nuevo nombre de la hoja
                requestObj.sheets.push(relationArray[j].sheetName);
              }
            }
          }
          
          // agregamos el registro
          requestData.push(requestObj);
          
        }
      }
      
      // retornamos los datos
      return requestData;
      
    }
  }
  
  // retornamos por defecto un array vacio
  return [];
  
}

/**
 * Función que permite obtener la estructura en objeto sobre la información de las versiones creadas para un proyecto
**/
function getVersionsObj(spreadsheet){
  
  // obtenemos la hoja de "Administrador de versiones"
  var sheet = spreadsheet.getSheetByName("Administrador de versiones");
  
  // variable para guardar el objeto
  var objectData = {};
  
  // validamos si existe la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
      
      // obtenemos todos los datos de la hoja
      var allData = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
      
      // recorremos cada uno de los datos 
      for(var i = 0; i < allData.length; i++){
        
        // validamos si la propiedad no existe
        if(!objectData[allData[i][0]]){
          
          // inicializamos la propiedad
          objectData[allData[i][0]] = [];
          
        }
        
        // agregamos el nuevo dato
        objectData[allData[i][0]].push({
          sheetName: allData[i][1],
          structure: allData[i][2],
          rowIndex: (i + 2)
        });
        
      }
    }
    
  }
  
  // retornamos el objeto
  return objectData;
  
}

/**
* Función que permite registrar los errores en la consola
**/
function saveErrorLog(ex) {

  // Registramos el error en la consola
  console.error("%s Archivo: (%s) Línea: %d", ex, ex.fileName, ex.lineNumber);
}

/**
* Función encargada de obtener un Token de autorización
* @return {String} Retorna un script con el token de autorización
*/
function getOAuthToken() {
  //Retornamos un token de autorización
  return ScriptApp.getOAuthToken();
}

/**
 * Función que relaciona los valores requeridos en el uso del plugin Google Drive Picker.
 * @return {Object} Retorna un objeto con los parámetros del plugin Google Drive Picker.
 */
function getPickerInfo(module, paramsData, type) {
  // Se hace referencia a la carpeta raiz (Mi unidad) en Google Drive
  var rootFolder = DriveApp.getRootFolder();
  
  // definimos la carpeta inicial de busqueda
  var folderId = rootFolder.getId();
  
  // validamos si el módulo es renovación(Refacción)
  if(module == "renovation"){
    
    // validamos si es carpeta
    if(type != "folder"){
      
       // establecemos como carpeta de busqueda se encuentran las plantillas
       folderId = "1tiTqKvgh6YwmCSicamstW9RNSUPZRgGj"; //OJO
      
    }
  }
  
  // Se retorna un objeto con la siguiente información
  return {
    token: getOAuthToken(),// Token de acceso de OAuth 2.0 para el usuario efectivo.
    folderId: folderId, // ID de la carpeta donde se consultarán las plantillas
    //developerKey: "AIzaSyATuM7ZeY12KXWQa2_hvrldE-I88-5VM58", // Clave de API para el DrivePicker
    developerKey: "AIzaSyBVz0lzxiNeQtoDH1W5pZ9aAO8ilqxm_mE", // Clave de API para el DrivePicker
    dialogDimensions: {width: 1000, height: 600}, // Dimensiones (ancho y alto) de la ventana de selección de archivos,
    type: type, // que tipo de archivo debe mostrar
    paramsData: paramsData // datos como el id del contenedor, text y nombre de la propiedad 
  };
}