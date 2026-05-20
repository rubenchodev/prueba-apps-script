/*
* Función que permite eliminar un archivo
* @param {String} documentId Id del archivo que deseamos eliminar
* @param {String} row Fila donde esta ubicada la cotización
*/
function deleteQuotation(documentId, row, module){
  
  // encerramos esta sentencia en un try-catch
  try{
    
    // se obtiene los recursos de hojas de cálculo
    var resources = getResources();
    
    // validamos desde que tipo de módulo se realizo la eliminación
    if(module == "renovation"){
      
      // se procede a cambiar el estado de la cotización de una refacción
      changeStateQuotation(documentId, row, resources.renovationSheetId, "Solicitudes", 10);
      
    } else {
      
      // se procede a cambiar el estado de la cotización de un proyecto
      changeStateQuotation(documentId, row, resources.requestSheetId, "Cotizaciones", 10);
      
    }
  
    // eliminamos el respectivo archivo
    DriveApp.getFileById(documentId).setTrashed(true);
    
  } catch(e) {
    
    // registramos el mensaje de error en la consola
    saveErrorLog(e);
    
  }
  
}

/**
 * Función que permite cambiar el estado de la cotización
**/
function changeStateQuotation(documentId, row, spreadsheetId, sheetName, column){
  
  // referenciamos la hoja de cálculo de solicitudes de cotizaciones
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  // obtenemos la hoja de "Cotizaciones"
  var quotationSheet = spreadsheet.getSheetByName(sheetName);
  
  // validamos si existe la hoja
  if(quotationSheet){
    
    // referenciamos el rango y cambiamos el estado
    quotationSheet.getRange(row, column).setValue("Eliminada");
  }
}

/**
 * Función que permite obtener los datos generales de los 2 tipos de cotizaciones
**/
function getGeneralData(){
  
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // se obtiene el correo del usuario
  var userMail = getUserMail();
  
  // referenciamos la hoja de cálculo de configuración
  var settingSpreadsheet = SpreadsheetApp.openById(resources.settingSheetId);
  
  // retornamos los datos
  return {
    clientsData: getClientList(resources.clientSheetId),
    stateList: getStateList(resources.stateSheetId),
    conceptList: getConceptList(settingSpreadsheet),
    userData: getAccountUserPanel(userMail, "Usuario")
  };
}
