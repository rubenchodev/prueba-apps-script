/**
 * Permite obtener la lista de unidades
**/


/**
 * Función que permite guardar la solicitud de actualización de precios
**/
function sendUpdatePrices(sendArrayData){
  
  // obtenemos los datos de recursos
  var resources = getResources();
  
  // referenciamos la hoja fabricantes
  var spreadSheet = SpreadsheetApp.openById(resources.acquisitionSheet);

  var spreadSheetManufacturer = SpreadsheetApp.openById(resources.manufacturerSheetId);  
  
  // se valida si existe la hoja de cálculo
  if(spreadSheet){
    
    // fecha actual del sistema
    var currentDate = new Date();
    
    // fecha actual si caracteres especiales
    var currentDateTrim = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "ddMMyyyyHHmmss");
    
    // obtenemos la fecha actual en "MM/dd/yyyy HH:mm:ss"
    var currentDateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
    
    // Creamos el código de actualización
    var codePack = currentDateTrim + Utilities.getUuid();
    
    // Referenciamos la hoja de precios
    var preciesSheet = spreadSheet.getSheetByName("Precios");
    
    // llamamos la función que procesa cada uno de las especies
    var processAllConceptsObject = processAllConcepts(sendArrayData, codePack, spreadSheet, spreadSheetManufacturer, resources);
    
    // inicializamos el array de id del fabricante
    var arrayManufacturerIds = [];
   
    // Recorremos cada uno de los fabricantes
    for(var key in processAllConceptsObject){
      // obtenemos el id de cada fabricante
      arrayManufacturerIds.push(processAllConceptsObject[key]["manufacturerId"]);
    } 
    
    // Validamos de que exista la hoja
    if(preciesSheet){
      
      // recorremos cada uno de las especies
      for(var i = 0; i < sendArrayData.length; i++){
        
        // validamos si existe el "rowIndex"
        if(sendArrayData[i].rowIndex){
        
          //insertamos el valor en la hoja de cálculo 
          preciesSheet.getRange(sendArrayData[i].rowIndex, 9, 1, 5).setNumberFormat("@STRING@").setValues([[currentDateStr, sendArrayData[i].quantity, codePack, "", String(arrayManufacturerIds)]]);
        }
      }
    }
    
    return processAllConceptsObject;
    
  } else {
    
    // retornamos un error
    throw new Error("Lo sentimos la hoja 'Fabricantes' para adquisiciones no existe.");
    
  }
}

/**
 * Función que permite enviar el listado de especies que desea solicitar la actualización de precios
**/
function automaticUpdateOfConcepts(){
    
  // obtenemos los datos de recursos
  var resources = getResources();
  
  // referenciamos la hoja fabricantes
  var spreadSheet = SpreadsheetApp.openById(resources.acquisitionSheet);

  var spreadSheetManufacturer = SpreadsheetApp.openById(resources.manufacturerSheetId);  
  
  // se valida si existe la hoja de cálculo
  if(spreadSheet && spreadSheetManufacturer){
    
     // fecha actual si caracteres especiales
    var currentDateTrim = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "ddMMyyyyHHmmss");
    
    // Creamos el código de actualización
    var codePack = currentDateTrim + Utilities.getUuid();
    
    // Obtenemos las especies para actualizar 
    var arrayConcepts = getConceptsUpdate(spreadSheet, codePack);
    
    // llamammos la función que procesa cada uno de las especies
    processAllConcepts(arrayConcepts, codePack, spreadSheet, spreadSheetManufacturer, resources);
    
  }
}

/**
 * Función que procesa el listado de especies y los envia a los fabricantes respectivos
**/
function processAllConcepts(arrayConcepts, codePack, spreadSheet, spreadSheetManufacturer, resources){
  
  // Objeto con los datos de los fabricantes
  var objectData = {};
  
  // validamos de que se tengan especies 
  if(arrayConcepts && arrayConcepts.length > 0){
  
    // Hacemos referencia a la hoja de fabricantes 
    var manufacturersSheet = spreadSheetManufacturer.getSheetByName("Fabricantes");
    
    // Obtenemos la ultima fila
    var lastRow = manufacturersSheet.getLastRow();
    
    // Obtenemos los datos de los fabricantes
    var manufacturersData = manufacturersSheet.getRange(2, 1, lastRow , 19).getDisplayValues();
    
    // recorremos las especies para actualizar 
    for(var i = 0; i < manufacturersData.length-1; i++){
      
      // Validamoos si tiene familias 
      if(manufacturersData[i][12]){
        
        // obtenemos la familias del fabricante en un array
        var manufacturerfamilies = manufacturersData[i][12].split(",");
        
        // Recorremos los fabricantes
        for(var j = 0; j < arrayConcepts.length; j++){
          
          //validamos si la especie tiene la familia del fabricante 
          if(manufacturerfamilies.indexOf(arrayConcepts[j].conceptFamily) != -1){
            
            // Validamos si el fabricante no existe
            if(objectData[manufacturersData[i][4]]){
            
              // Agregamos la especie 
              objectData[manufacturersData[i][4]].concepts.push(arrayConcepts[j])
              
            }else{
              
              // Validamos de que tenga el nombre del fabricante
              if(manufacturersData[i][1]  && manufacturersData[i][4]){
                
                // Agregamos los datos del fabricante 
                objectData[manufacturersData[i][4]] = {
                  concepts: [arrayConcepts[j]],
                  manufacturerId: manufacturersData[i][1],
                  email: manufacturersData[i][2],
                  state: false // Indicando de que no se ha enviado la notificación 
                }
              }
            }
          }
        }
      }
    }
  }
  
  // validamos de que se deban enviar las notificaciones 
  if(Object.keys(objectData).length > 0){
    
    // Referenciamos la hoja de fabricantes/ solicitudes
    var manufacturersRequests =  spreadSheet.getSheetByName("Fabricante/Solicitudes");
    
    // Obtenemos la url del aplicativo 
    var webUrl =  resources.webUrl;
     
    // obtenemos la fecha actual en "MM/dd/yyyy HH:mm:ss"
    var currentDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
    
  
    // Recorremos cada uno de los fabricantes
    for(var key in objectData){
      
      // Obtenemos la ultima fila
      var lastRow = manufacturersRequests.getLastRow();
      
      //Actualizamos el registro en la hoja
      manufacturersRequests.getRange(lastRow + 1, 1, 1, 3).setValues([[objectData[key].manufacturerId, codePack, "Pendiente"]]);
      
      // Envimaos las notificaciones
      objectData[key].state = sendNotification(objectData[key], webUrl, codePack, currentDateStr);   
    }
  }
  
  // Retornamoslos datos del cliente 
  return objectData
}


/*
* Función que permite enviar la notificaciones de actualización de precios de especies
* 
*/
function sendNotification(manufacturersDataObj, webUrl, codePack, currentDateStr){

  // Manejamos el error por si no se envia el correo a un fabricante 
  try{
 
     //Referenciamos el contenido de la plantilla
     var contentNotification = HtmlService.createHtmlOutputFromFile("Html_MailTemplateUpdateConcepts").getContent();
    
     //Insertamos la tabla de las especies
     contentNotification = contentNotification.replace("|*CONCEPT_LIST*|", getHtmlConcepts(manufacturersDataObj.concepts));
     
     //Insertamos el vínculo del formulario de actualización
     //contentNotification = contentNotification.replace("|*LINK*|",  webUrl + "?uuid=" + manufacturersDataObj.manufacturerId + "_" + codePack);
     contentNotification = contentNotification.replace("|*LINK*|",  "https://script.google.com/a/macros/hidro.expert/s/AKfycby5t2PiMMKidE2j49LTaf41FeyuVmxLDmbaltLf5NXLQEvX6fI/exec");
    
     // se procede enviar el correo electronico
     sendEmailDelegation_(manufacturersDataObj.email, ("Agrocity: Solicitud de actualización de precios de conceptos - " + currentDateStr), contentNotification);
     /*MailApp.sendEmail({
       to: manufacturersDataObj.email,
       subject: ("Agrocity: Solicitud de actualización de precios de conceptos - " + currentDateStr),
       htmlBody: contentNotification,
       name: "Adquisiciones"
     });*/
     
     //retornamos true si se envio
     return true;
   
   }catch(error){
     
      // Retornamos false 
      return false;
   }
}


/*
* Funcion que permite obtener los el html para las especies
*/
function getHtmlConcepts(conceptList){
   
  // Variable para el html la tabla de especies
  var html = "<table width='100%' border='1' cellpadding='4' cellspacing='0' style='margin-top: 10px;border-collapse:collapse;font-family:sans-serif;font-size:13px;line-height:20px;border: 1px solid #cccccc;' bordercolor='#CCCCCC'>"
    +"<tr>"
    +"<th width='80%' bgcolor='#d1d1d1' style='padding: 5px;'>Concepto</th>"
    +"<th  width='20%' bgcolor='#d1d1d1' style='padding: 5px;text-align: center;'>Unidad</th>"
    +"<th  width='20%' bgcolor='#d1d1d1' style='padding: 5px;text-align: center;'>Cantidad</th>"
    +"</tr>";
  
  // Recorremos cada uno de las especies
  for(var i= 0; i < conceptList.length; i++){
    
    // agregamos la fila
    html += " <tr>"
    +"<td  style='padding: 5px;border: 1px solid #ddd;'>" + conceptList[i].conceptName + "</td>"
    +"<td style='padding: 5px;border: 1px solid #ddd;text-align: center;'>" + conceptList[i].unit + "</td>"
    +"<td style='padding: 5px;border: 1px solid #ddd;text-align: center;'>" + conceptList[i].quantity + "</td>"
    " </tr>";
    
  }
  
  // retornamos la tabla creada
  return html += "</table>";
}

/*
* Función que permite obtener las especies que se van a actualziar
*/
function getConceptsUpdate(spreadSheet, code) {

    // Referenciamos la hoja de precios
    var preciesSheet = spreadSheet.getSheetByName("Precios");
    
    // Validamos de que exista la hoja
    if(preciesSheet){
      
      // Obtenemosla ultima fila 
      var lastRow = preciesSheet.getLastRow();
      
      // valdiamos de que se tengan registros 
      if(lastRow > 1){
        
        // Obtenemos la fecha actual
        var newDate = new Date();
       
        // obtenemos la fecha actual 
        var currentDate = Utilities.formatDate(newDate, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
        
        // Obtenemos los milsegundos de la fecha actual
        var currentDateMilli =  newDate.getTime();
        
        // Obtenemos los datos 
        var arrayData = preciesSheet.getRange(2, 1, lastRow, 12).getDisplayValues();
       
        // Contador para la  cantidad de registro         
        var cont = 1;
        
        // Milisegundos de la fecha a comparar
        var millisecondsDate = 0;
        
        // Diferencia entre fechas
        var diffDate = null;
        
        // Estado si se debe actualizar
        var stateUpdate = true,
            date = "";
        
        // Datos de las especies a actualizar 
        var arrayConcepts = [];
        
        // Recorremos cada uno de los datos 
        for(var i = 0; i < arrayData.length && cont <= 20; i++){
          
          // Estabelcmeos de que se debe actualizar el registro 
          stateUpdate = true;
        
          //validamos de que exista una fecha
          if(arrayData[i][8]){
            
            // Creamos la fecha con el formato requerido
            date = new Date(arrayData[i][8]);
            
            // obtenemos los milisegndos de la fecha 
            millisecondsDate = date.getTime();
            
            // Obtenemos la diferencia en dias
            diffDate =  ((currentDateMilli - millisecondsDate) / 3600000 ) / 24;
            
            //validamos la fecha del registro
            if(diffDate < 45){
            
              // Establecemos de que no se debe agregar el registro  
              stateUpdate =  false;
                
            }
          }
          
          // Validamos si se debe agregar el registro 
          if(stateUpdate){
            
            // Agregamos los datos 
            arrayConcepts.push({   
               conceptName: arrayData[i][1],
               conceptFamily: arrayData[i][2],
               unit:arrayData[i][4],
               quantity: 1,
               
            });
            
            //insertamos el valor en la hoja de cálculo 
            preciesSheet.getRange(i + 2, 9, 1, 3).setNumberFormat("@STRING@").setValues([[currentDate, 1, code]]);
            
            // Aumentamos el contador 
            cont++;
          }
        }
        
        // Retornamos las especies que se van a actualizar 
        return arrayConcepts;
        
      } 
    }
}