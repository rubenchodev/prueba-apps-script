/**
 * Función encargada de enviar el correo al a´rea de adquisiciones recordandole que debe actualizar los tipos de cambios
**/
function sendRecodatory() {
  
  // obtenemos los recursos globales
  var resources = getResources();
  
  // se referencia la hoja de usuarios
  var spreadsheet = SpreadsheetApp.openById(resources.userFileId);
  
  // se referencia la hoja de "Adquisiciones"
  var sheet = spreadsheet.getSheetByName("Adquisiciones");
  
  // se valida que exista la hoja
  if(sheet){
    
    // se obtiene la ultima fila
    var lastRow = sheet.getLastRow();
    
    // se valida que exista mas de una fila
    if(lastRow > 1){
      
      // obtenemos los datos
      var arrayData = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      
      // variable para determinar en que fila esta 
      var currentRow = -1;
      
      // recorremos cada una de los registros
      for(var i = 0; i < arrayData.length; i++){
        
        // se valida si tiene el texto "Enviar"
        if(arrayData[i][1] == "Enviar"){
          
          // se valida que exista un valor
          if(arrayData[i][0]){
          
            // se procede a enviar un correo electronico
            sendMailReminder(arrayData[i][0], resources.acquisitionSheet);
            
          }
          
          // definimos la fila donde esta el texto enviar
          currentRow = i + 2;
          
          // salimos del ciclo
          break;
        }
      }
      
      // se valida que la fila actual sea diferente de -1
      if(currentRow != -1){
        
        // se elimina el texto actual
        sheet.getRange(currentRow, 2).setValue("");
        
      }
      
      // se valida si la fila actual es diferente de -1
      if(currentRow != -1 && currentRow < lastRow){
        
        // se aumenta en uno la fila donde se debe cambiar el dato
        currentRow ++;
        
      } else {
        
        // se define la fila 2 por defecto
        currentRow = 2;
        
      }
      
      // se agrega el nuevo texto
      sheet.getRange(currentRow, 2).setValue("Enviar");
    }
  }
}


/**
 * Función que envia un recordatorio al área de aquisiciones para que proceda en la actualización del tipo de cambio
 * @author: Rubén Sánchez
 */
 function sendMailReminder(email, fileId){
   
   //Referenciamos el contenido de la plantilla
   var contentNotification = HtmlService.createHtmlOutputFromFile("Html_MailTemplateReminder").getContent();
   
   //Insertamos el vínculo del formulario de actualización
   contentNotification = contentNotification.replace("|*FILE_ID*|", fileId);
   
   // se procede enviar el correo electronico
   MailApp.sendEmail({
     to: email,
     subject: "Agrocity: Recordatorio actualización de tipo de cambios",
     htmlBody: contentNotification,
     name: "Adquisiciones"
   });
}