/**
 * Función que permite obtener la información de los tipos de concexión
**/
function getTypeConection(){
  
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // Referenciamos la hoja de cálculo de donde se encuentran las validaciones
  var spreadsheet = SpreadsheetApp.openById(resources.validationSheetId);
  
  // obtenemos la hoja de "Grupos"
  var sheet = spreadsheet.getSheetByName("Validaciones");
  
  // validamos si existe la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // se valida que exista mas de 3 fias
    if(lastRow > 3){
      
      // retornamos los datos
      var arrayData = sheet.getRange(4, 10, lastRow - 3, 1).getValues();
      
      // array para retornar
      var returnArray = [];
      
      // recorremos cada uno de los registros
      for(var i = 0; i < arrayData.length; i++){
        
        // validamos que el valor sea diferente a vacio
        if(arrayData[i][0]){
           
           // lo agregamos al listado
           returnArray.push(arrayData[i][0]);
        }
      }
      
      // retornamos los datos
      return returnArray;
      
    }
  }
  
  // retornamos el listado vacio
  return [];
}


/**
 * Función que permite crear un nuevo grupo
**/
function createNewGroup(conceptObj, formData, lastRowSheetGroup, type, groupName){

  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // Referenciamos la hoja de cálculo de configuración
  var settingSpreadsheet = SpreadsheetApp.openById(resources.settingSheetId);
  
  // obtenemos la hoja de "Grupos"
  var groupSheet = settingSpreadsheet.getSheetByName("Grupos");
  
  // variable que se tiene para validar el nombre actual del grupo
  var currentGroupName = null;
  
  // validamos si existe la hoja
  if(groupSheet){
    
    // obtenemos la ultima fila
    var groupLastRow = groupSheet.getLastRow();
    
    // validamos si existe mas de una fila
    if(groupLastRow > 1 && !lastRowSheetGroup){
    
      // obtenemos la información
      var groupData = groupSheet.getRange(2, 1, groupLastRow - 1, 1).getValues();
      
      // recorremos cada uno de los grupos y validamo sis este ya existe
      for(var i = 0; i < groupData.length; i++){
        
        // se valida si existe el grupo
        if(groupName == groupData[i][0]){
          
          // Mostramos excepción
          throw new Error("El grupo ya existe por favor valida los datos.");
            
        }
      }
    }
      
    // variable del costo
    var cost = 0;
    
    // recorremos cada uno de las especies
    for(var row = 0; row < conceptObj.length; row++){
   
      // agregamos la sumatoria de costo de cada producto
      cost += (conceptObj[row].conceptData[5] ? (Number(String(conceptObj[row].conceptData[5]).replace(/\$|\,/g, "")) * Number(String(conceptObj[row].quantity).replace(/\,/g, "."))) : 0);
      
      
    }
    
    // se valida si no es un número 
    if(isNaN(cost)){
      
      // definimos el valor de cero
      cost = 0;
    }
    
    // Agregamos el precio sugerido
    formData.currentPrice = "$" + formatValue(roundDecimal(cost, 4));
    
    // validamos si existe la fila del grupo
    if(lastRowSheetGroup){  
      
      // obtenemos el nombre actual del grupo
      currentGroupName = groupSheet.getRange(lastRowSheetGroup, 1).getValue();
      
      // guardamos la informaicón del grupo
      groupSheet.getRange(lastRowSheetGroup, 1, 1, 3).setNumberFormat("@STRING@").setValues([[groupName, JSON.stringify(conceptObj), JSON.stringify(formData)]]);
      
    } else {
      
      // guardamos la informaicón del grupo
      groupSheet.getRange(groupLastRow + 1, 1, 1, 3).setNumberFormat("@STRING@").setValues([[groupName, JSON.stringify(conceptObj), JSON.stringify(formData)]]);
    }
    
    // referencamos la hoja de conceptos
    var conceptSheet = settingSpreadsheet.getSheetByName("Conceptos");
    
    // se valida que la hoja exista
    if(conceptSheet){
    
      // variable para guardar el array del nuevo grupo
      var nowGroupArray = [
        "",
        groupName,
        "",
        "",
        "GRUPO",
        roundDecimal(cost, 3)
      ];
      
      // validamos si es una actualización del grupo
      if(lastRowSheetGroup){
        
        // se obtiene la ultima fila de la hoja
        var lastRowConcepts = conceptSheet.getLastRow();
        
        // se valida que exista al menos 2 filas e la hoja
        if(lastRowConcepts > 1){
        
          // obtenemos el listado de nombre de la especie de la lista
          var conceptList = conceptSheet.getRange(2, 2, lastRowConcepts - 1, 1).getValues()
       
          // recorremos cada uno de las especies de la hoja de cálculo
          for(var j = conceptList.length - 1; j >= 0; j--){
           
            // validamos si la especie aparece en el listado
            if(currentGroupName == conceptList[j][0]){
             
              // almacenamos la nueva información
              conceptSheet.getRange((j + 2), 1, 1, nowGroupArray.length).setValues([nowGroupArray]);
              
              // salimos del ciclo
              break;
              
            }
          }
        }
      
      } else {
        
        // agregamos un nuevo grupo como especie
        conceptSheet.appendRow(nowGroupArray);
        
      }
      
      // retornamos el listado actualizado la especie
      return {
        conceptUpdated: getConceptList(settingSpreadsheet),
        newGroupName: groupName,
        nowGroupArray:nowGroupArray,
        type: type
      };
      
    } else {
      
      // retornamos un mensaje de error
      throw "La hoja de <b>Conceptos</b> no existe.";
      
    }
    
  } else { // mensaje que se muestra por si no existe la hoja de grupos 
  
    // mostramos un mensaje por defecto
    throw "La hoja de <b>Grupos</b> no existe.";
  }
}


/**
 * Función que permite obtener los datos de configuración de un grupo seleccionado
**/
function getInfoByGroupName(conceptData){

  // obteneos el nombre del grupo
  var currentGroupName = String(conceptData[1]).trim();
  
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // Referenciamos la hoja de cálculo de configuración
  var settingSpreadsheet = SpreadsheetApp.openById(resources.settingSheetId);
  
  // obtenemos la hoja de "Grupos"
  var groupSheet = settingSpreadsheet.getSheetByName("Grupos");
  
  // validamos si existe la hoja
  if(groupSheet){
    
    // obtenemos la ultima fila
    var groupLastRow = groupSheet.getLastRow();
    
    // validamos si existe mas de una fila
    if(groupLastRow > 1){
    
      // obtenemos los datos
      var arrayData = groupSheet.getRange(2, 1, groupLastRow - 1, 3).getValues();
      
      // recorremos cada uno de los datos
      for(var i = 0; i < arrayData.length; i++){
       
        // validamos si el nombre del grupo existe
        if(currentGroupName == String(arrayData[i][0]).trim()){
          
          // retornamos los datos
          return {
            groupObj: JSON.parse(arrayData[i][1]),
            rowIndex: (i + 2),
            groupName: currentGroupName,
            aditionalData: JSON.parse(arrayData[i][2])
          };
        }
      }
    }
  }
  
  //  retornamos por defecto null
  return null;
  
}

/**
 * Función que permite obtener todos los grupos registrados hasta la fecha
**/
function getGroupList(){
  
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // Referenciamos la hoja de cálculo de configuración
  var settingSpreadsheet = SpreadsheetApp.openById(resources.settingSheetId);
  
  // obtenemos la hoja de "Grupos"
  var groupSheet = settingSpreadsheet.getSheetByName("Grupos");
  
  // validamos si existe la hoja
  if(groupSheet){
    
    // obtenemos la ultima fila
    var groupLastRow = groupSheet.getLastRow();
    
    // validamos si existe mas de una fila
    if(groupLastRow > 1){
    
      // obtenemos los datos
      var arrayData = groupSheet.getRange(2, 1, groupLastRow - 1, 3).getValues();
      
      // variable para guardar cada uno de los datos de los grupos
      var groupList = [];
      
      // recorremos cada uno de los datos
      for(var i = 0; i < arrayData.length; i++){
          
        // retornamos los datos
        groupList.push({
          groupObj: JSON.parse(arrayData[i][1]),
          rowIndex: (i + 2),
          groupName: arrayData[i][0],
          aditionalData: JSON.parse(arrayData[i][2])
        });
        
      }
      
      // retornamos la lista de grupos
      return groupList;
    }
  }
  
  //  retornamos la lista vacia
  return [];
  
}
