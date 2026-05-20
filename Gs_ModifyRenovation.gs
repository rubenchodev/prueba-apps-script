/**
 * Función que permite obtener los datos de una cot. de refacción
**/
function getInfoUpdatedRenovation(documentId, currentData) {
  
  // se deja el objeto manipulable
  currentData = JSON.parse(currentData);
  
  // referenciamos la hoja de cálculo de valores configurables
  var spreadsheet = SpreadsheetApp.openById(documentId);
  
  // obtenemos la hoja de "Nombre del proyecto"
  var sheet = spreadsheet.getSheetByName("Proyecto");
  
  // validamos si existe la hoja
  if(sheet){
    
    // definimos la variable que determina la fila actual
    var currentRowIndex = 15;
    
    // se recorre cada uno de los valores 
    for(var i = 0; i < currentData.conceptList.length; i++){
      
      // actualizamos el valor de la cantidad de elementos
      currentData.conceptList[i].conceptCount = Number(sheet.getRange("H" + currentRowIndex).getValue());
      
      // actualizamos el valor de inventario 
      currentData.conceptList[i].inventory = sheet.getRange("C" + currentRowIndex).getValue();
      
      // se aumenta el contador de la fila actual
      currentRowIndex ++;
      
    }
  }
  
  // retornamos el array actualizado
  return {
    currentData: currentData,
    lastRow: currentRowIndex
  };
}

/**
* Función que se ejecuta al final del d{ia para resetear el consecutivo
**/
function resetConsecutive(){
  
  // obtenemos las propiedades del script
  var scriptProperties = PropertiesService.getScriptProperties();
  
  // iniciamos el consecutivo en 1
  scriptProperties.setProperty("Consecutive", "1");
  
}