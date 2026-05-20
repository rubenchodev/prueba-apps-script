/**
 * Función que permite obtener los datos de una cot.
**/
function getInformationeUpdated(documentId, currentData, sheetName) {
  
  // se deja el objeto manipulable
  currentData = JSON.parse(currentData);
  
  // referenciamos la hoja de cálculo de valores configurables
  var spreadsheet = SpreadsheetApp.openById(documentId);
  
  // obtenemos la hoja de "Nombre del proyecto"
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  // validamos si existe la hoja
  if(sheet){
    
    // definimos la variable que determina la fila actual
    var currentRowIndex = 22,
        currentCounter = 0;
    
    // se recorre cada uno de los valores 
    for(var i = 0; i < currentData.orderProjects.length; i++){
      
      // se recorre cada una de las lineas
      for(var j = 0; j < currentData.orderProjects[i].linesProducts.length; j++){
        
        // se aumenta el contador de la fila actual
        currentRowIndex ++;
        
        // se recorre cada especie
        for(var k = 0; k < currentData.orderProjects[i].linesProducts[j].conceptList.length; k++){
          
          // se aumenta el contador de la fila actual
          currentRowIndex ++;
          
          // obtenemos el contador
          currentCounter = Number(sheet.getRange("H" + currentRowIndex).getValue());
          
          // validamos si es un numero
          if(!isNaN(currentCounter)){
          
            // actualizamos el valor de la cantidad de elementos
            currentData.orderProjects[i].linesProducts[j].conceptList[k].conceptCount = currentCounter;
          }
          
        }
        
      }
      
      // se aumenta el contador de la fila actual
      currentRowIndex ++;
      
    }
  }
  
  // retornamos el array actualizado
  return {
    currentData: currentData,
    lastRow: currentRowIndex,
    
  };
}

/**
 * Función que permite agregar una nueva especie
**/
function addNowConcept(objectData){
  
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // referenciamos la hoja de cálculo de solicitudes de cotizaciones
  var spreadsheet = SpreadsheetApp.openById(resources.settingSheetId);
  
  // obtenemos la hoja de "Conceptos"
  var sheet = spreadsheet.getSheetByName("Conceptos");
  
  // validamos si existe la hoja
  if(sheet){
    
    // obtenemos el correo del usuario
    var userMail = getUserMail();
    
    // obtenemos la última fila
    var lastRow = sheet.getLastRow();
    
    // referenciamos el rango y agregamos los datos
    sheet.getRange(lastRow + 1, 1, 1, 6).setValues([[
      objectData.departament,
      objectData.specie,
      objectData.family,
      objectData.keyWord,
      objectData.unit,
      objectData.price
    ]]);
    
    // almacenamos la información del correo que realiza la creación de la nueva especie
    sheet.getRange("AI" + (lastRow + 1)).setValue(userMail);
  
    // referenciamos la hoja de usuarios
    var userSpreadsheet = SpreadsheetApp.openById(resources.userFileId);
    
    // referenciamos la hoja de "Adquisiciones"
    var userSheet = userSpreadsheet.getSheetByName("Adquisiciones");
    
    // obtenemos la ultima fila
    var userLastRow = sheet.getLastRow();
    
    // validamos si existe mas de una fila
    if(userLastRow > 1){
      
      // obtenemos los correo separados por coma
      var emails = userSheet.getRange(2, 1, userLastRow - 1, 1).getValues();
      
      // validamos si existen correos
      if(emails){
        
        //Referenciamos el contenido de la plantilla
        var contentNotification = HtmlService.createHtmlOutputFromFile("Html_MailGeneralTemplate").getContent();
        
        //Insertamos el valor de especies
        contentNotification = contentNotification.replace("|*SPECIES*|", objectData.specie);
        
        //Insertamos el valor de unidad
        contentNotification = contentNotification.replace("|*UNID*|", objectData.unit);
        
        //Insertamos el valor de precio sugerido
        contentNotification = contentNotification.replace("|*SUGGESTED_PRICE*|", objectData.price);
        
        //Insertamos la fila donde esta ubicado la especie
        contentNotification = contentNotification.replace("|*ROW*|", lastRow + 1);
        
        //Insertamos la url del módulo de adquisiciones
        contentNotification = contentNotification.replace("|*SCRIPT_URL*|", getUrlScript());
        
        //Enviamos el correo
        MailApp.sendEmail({
          to: emails.toString(),
          subject: "Nueva especie creada",
          htmlBody: contentNotification,
          name: "Hidrocotizador"
        });
      
      }
    }
  }
}


