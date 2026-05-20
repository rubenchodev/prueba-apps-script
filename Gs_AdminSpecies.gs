/*
* Función que permite actualizar las imágenes asociadas a las especies
*/
function updateMultipleImages(rowIndexList, imageArrayData){
  
  // obtenemos los recursos globales  
  var resources =  getResources();

  // Referenciamos la hoja de configuración
  var spreadsheet = SpreadsheetApp.openById(resources.settingSheetId);  
      
  // se obtiene la hoja de conceptos
  var conceptSheet = spreadsheet.getSheetByName("Conceptos");
  
  // se valida que exista la hoja
  if(conceptSheet){
    
    // Validamos si vienen imágenes 
    if(imageArrayData){
       
      // referenciamos la carpeta donde se deben almacenar las imagenes de especies
      var rootFolder = DriveApp.getFolderById(resources.imageFolderId);
      
      // variable para guardar la información temporal de las imagenes
      var imageTempData = [];
      
      // recorremos cada una de las imagenes
      for(var i = 0; i < imageArrayData.length; i++){
        
        // agregamos al listado la imagen correspondiente
        imageTempData.push({
          id: (imageArrayData[i].id ? imageArrayData[i].id : uploadFileToDrive(imageArrayData[i], rootFolder)),
          name: imageArrayData[i].name
        });
        
      }
      
      // actualizamos la columna 36 de imagenes
      imageTempData = JSON.stringify(imageTempData);
      
      // recorremos cada uno de los registros para actualizar
      for(var j = 0; j < rowIndexList.length; j++){
        
        // Guardamos la información de las imágenes
        conceptSheet.getRange(rowIndexList[j], 38).setNumberFormat("@STRING@").setValue(imageTempData.length == 0 ? "" : imageTempData);
      }
    }
  }
}

/*
* Función que permite crear o modificar una especie
*/
function createNewSpecie(startSpecieFormData, endSpecieFormData, imageArrayData, rowIndex){
  
  // obtenemos los recursos globales
  var resources =  getResources();
  
  // Referenciamos la hoja de configuración
  var spreadsheet = SpreadsheetApp.openById(resources.settingSheetId);
  
  // se obtiene la hoja de conceptos
  var conceptSheet = spreadsheet.getSheetByName("Conceptos");
  
  // se valida que exista la hoja
  if(conceptSheet){
    
    // Validamos si vienen imágenes
    if(imageArrayData){
      
      // referenciamos la carpeta donde se deben almacenar las imagenes de especies
      var rootFolder = DriveApp.getFolderById(resources.imageFolderId);
      
      // variable para guardar la información temporal de las imagenes
      var imageTempData = [];
      
      // recorremos cada una de las imagenes
      for(var i = 0; i < imageArrayData.length; i++){
        
        // agregamos al listado la imagen correspondiente
        imageTempData.push({
          id: (imageArrayData[i].id ? imageArrayData[i].id : uploadFileToDrive(imageArrayData[i], rootFolder)),
          name: imageArrayData[i].name
        });
        
      }
      
      // actualizamos la columna 36 de imagenes
      endSpecieFormData[2] = imageTempData.length > 0 ? JSON.stringify(imageTempData) : "";
    }
    
    // creamos la variable donde se guarda el correo a quien se debe enviar el correo
    var userMailNotify = {};
    
    // validamos si no existe el rowIndex
    if(!rowIndex){
      
      // obtenemos la ultima fila y le sumamos 1
      rowIndex = conceptSheet.getLastRow() + 1;
      
      // obtenemos los correo de Ingeniería y Dirección ejecutiva
      userMailNotify = getMailToNotify(resources);
      
    } else {
      
      // obtenemos el correo de la persona que creo inicialmente la especie
      var afterData = conceptSheet.getRange("A" + rowIndex + ":AI" + rowIndex).getValues();
      
      // validamos si existe un correo
      if(String(afterData[0][34]).trim()){
        
        // enviamos un correo al usuario notificando que se ha realizado la actualización de la especie respectiva
        userMailNotify = {
          mail: String(afterData[0][34]).trim().toLowerCase(),
          afterData: afterData
        }
        
        // borramos el correo actual
        conceptSheet.getRange("AI" + rowIndex).clearContent();
        
      }
    }
    
    // almacenamos la información de a especie y atributos
    conceptSheet.getRange(rowIndex, 1, 1, startSpecieFormData.length).setNumberFormat("@STRING@").setValues([startSpecieFormData]);
    
    // almacenamos la información de código, descripción e imágenes
    conceptSheet.getRange(rowIndex, 36, 1, endSpecieFormData.length).setNumberFormat("@STRING@").setValues([endSpecieFormData]);
    
    // se valida si existe correo para enviar la notificación
    if(userMailNotify && userMailNotify.mail){
    
      // enviamos un correo al usuario notificando que se ha realizado la actualización de la especie respectiva
      sendMailBeforeUpdate(userMailNotify, startSpecieFormData);
    }
    
  }
}

/**
 * Función que permite obtener los correo de Ingeniería y Dirección ejecutiva
**/
function getMailToNotify(resources){
  
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
      
      // obtenemos los datos
      var arrayData = userSheet.getRange(2, 1, lastRow - 1, 2).getValues();
      
      // lista de correos
      var mailsList = [];
      
      // recorremos el listado de correo parametrizados
      for(var i = 0; i < arrayData.length; i++){
      
        // Validamos si el usuario es igual al que se esta recorreiendo
        if("Ingeniería" == arrayData[i][1] || "Dirección ejecutiva" == arrayData[i][1]){
          
          // Se valida que no se agreguen correo repetidos
          if(mailsList.indexOf(arrayData[i][0]) == -1){
            
            // agregamos el correo al listado
            mailsList.push(arrayData[i][0]);
          }
        }
      }
      
      // retornamos el listado
      return {
        mail: mailsList.toString()
      };
    }
  }
  
  // retornamos un objeto con la información requerida
  return {
    mail: ""
  };  
}

/**
 * Función que envia la notificación una vez se actualice la información de la especie
 * @param: {string} userMailNotify: Información como correo y datos anteriores
 * @author: Rubén Sánchez
 */
 function sendMailBeforeUpdate(userMailNotify, startSpecieFormData){
   
   // variable para contener los textos respectivos
   var title = "Creación de nueva especie",
       content = "Te informamos que se ha creado una nueva especie denominada:<b><br><center>" + startSpecieFormData[1] + "</center>";
   
   // validamos si existen datos anteriores
   if(userMailNotify.afterData){
     
     // ajustamos el texto del título
     title = "Actualización de datos de una especie";
     
     // actualizamos el texto del contenido del correo
     content = "Te informamos que se actualizo la especie que inicial fue creada y denominada: <b>" + userMailNotify.afterData[0][1] + "</b><br><br>Los nuevos datos son:<br><br><b>Especie:</b> " + startSpecieFormData[1] + "<br><b>Costo sugerido:</b> " + startSpecieFormData[5];
     
   }
   
   //Referenciamos el contenido de la plantilla
   var contentNotification = HtmlService.createHtmlOutputFromFile("Html_MailTemplateBasic").getContent();
   
   //Insertamos el nombre del usuario
   contentNotification = contentNotification.replace("|*TITLE_MAIL*|", title);
   
   //Insertamos el vínculo del archivo
   contentNotification = contentNotification.replace("|*TEXT_CONTENT*|", content);
   
   //Enviamos el correo
   MailApp.sendEmail({
     to: userMailNotify.mail,
     subject: "Creación o actualización de una especie",
     htmlBody: contentNotification,
     name: "Adquicisiones"
   });
}

/**
 * Función que permite subir un archivo a Drive utilizando el base64 del mismo
**/
function uploadFileToDrive(imageData, rootFolder) {

  // capturamos el error al momento de crear el archivo
  try{
    
    // convertirmos el texto en un array de 2 posiciones de acuerdo al caracter(,)
    var splitBase = imageData.base64.split(','),
        // obtenemos el tipo de archivo eliminando el texto "data:" de la primera posicion del array anterior, eje: data:image/png;base64
        mimeType = splitBase[0].split(';')[0].replace('data:','');
    
    // Utilizando la clase Utilities decodificamos el texto en Base64 a bytes
    var byteCharacters = Utilities.base64Decode(splitBase[1]);
    
    // generamos un objeto binario o mas conocido como blob, utilizando el servicio Utilities
    var blob = Utilities.newBlob(byteCharacters, mimeType);
    
    // le asignamos un nombre
    blob.setName(imageData.name);
    
    // Creamos el archivo de imagen, la almacenamos en la carpeta de imágenes y referenciamos el id.
    var imageId = rootFolder.createFile(blob).getId();
    
    // retornamos el id de la imagen
    return imageId;
    
  }catch(e){
    throw e
  }
  
  // Retornamos por defecto vacio
  return "";
}


/*
* Función que permite obtener los datos para los selects 
*/
function getDataSelects(){
  
  // obtenemos el Id de la hoja de validaciones 
  var validationSheetId =  getResources().validationSheetId;
  
  // Referenciamos la hoja de validaciones
  var spreadSheet = SpreadsheetApp.openById(validationSheetId);
  
  // Validamos de que se tenga acceso a la hoja
  if(spreadSheet){
    
    // Objeto  para los datos de los select de formulario de creación o edicion 
    var validationsListObj = {
      department: getDepartmentData(spreadSheet),     // Departamento
      family: getFamilyData(spreadSheet),         //Familia
      keyword: getKeywordData(spreadSheet),        // Palabra clave 
      gender: [],         // Genero
      nominalDiameter: {},// Diametro nominal (Entrada)y (Salida)
      connectionType: [], // Tipo de conexxión (Entrada) y (Salida)
      material: [],       // material
      variety: getVarietyData(spreadSheet),         // Variadad
      pieceLength: [],    // Longitud de la pieza
      height: [],         // Altura
      adjustmentType: [], // Tipo de ajuste
      sizeOrCapacity: [], // Tamaño o capacidaad
      width: [],          // Ancho
      caliber: [],        // Calibre 
      model: [],          // Modelo
      flow: [],           // Caudal
      separation: [],     // Separación
      unit: [],           // unidad
      currencySpecie: [],  // Moneda
      documentRD:  getDocumentRD(spreadSheet),  // Cedula RD/espesor  
     
    }
    
    // Referecciamos la pestaña de validaciones 
    var validationSheet = spreadSheet.getSheetByName("Validaciones");
    
    // Validamos de se tenga acceso a la pestaña de  validaciones
    if(validationSheet){
      
      // Obtenmsos la ultima fila de la hoja diligenciada
      var lastRow = validationSheet.getLastRow();
      
      // validamos de que existan registros en la hoja
      if(lastRow > 3){
        
        // Obtenemos los datos de la hoja de cálculo 
        var arrayData = validationSheet.getRange(4, 1, lastRow, 37).getDisplayValues();
        
        // Recorremos cada registro para establecer los item de cada select
        for(var i = 0; i < arrayData.length; i++){
          
          // Validamos si viene los datos de generos y diametro
          if(arrayData[i][2] ){
            
            // Validamos si existe la propiedad en el objeto de genero 
            if(!validationsListObj.nominalDiameter[arrayData[i][2]]){
              
              // Creamos la propiedad
              validationsListObj.nominalDiameter[arrayData[i][2]] = [];
              
              // Insertamos los genero
              validationsListObj.gender.push([arrayData[i][2]]);
            }
            
            // Agregamos el item de los datos diametros
            validationsListObj.nominalDiameter[arrayData[i][2]].push([arrayData[i][3],arrayData[i][4]]);
          }
          
          // Validamos si viene los datos de tipo de conexión
          if(arrayData[i][8]){
            
            // Agregamos el item de tipo de conexión
            validationsListObj.connectionType.push([arrayData[i][8], arrayData[i][9]]);
          }
          
          // Validamos si viene datos de material
          if(arrayData[i][0] && arrayData[i][1]){
            
            // Agregamos el item de material
            validationsListObj.material.push([arrayData[i][0],arrayData[i][1]]);
          }
          
          // Validamos si viene los datos de longitud de pieza
          if(arrayData[i][6] && arrayData[i][7] ){
            
            // Agregamos el item de longitud de pieza
            validationsListObj.pieceLength.push([arrayData[i][6],arrayData[i][7]]);
          }
          
          // Validamos si viene los datos de altura
          if(arrayData[i][10] && arrayData[i][11]){
            
            // Agregamos el item de altura
            validationsListObj.height.push([arrayData[i][10], arrayData[i][11]]);
          }
          
          // Validamos si viene los datos de tipo de ajuste
          if(arrayData[i][15] && arrayData[i][16]){
            
            // Agregamos el item de tipo de ajuste
            validationsListObj.adjustmentType.push([arrayData[i][15], arrayData[i][16], arrayData[i][17]]);
          }
          
          // Validamos si viene los datos de Tamaño o capacidad
          if(arrayData[i][20] && arrayData[i][21]){
            
            // Agregamos el item de tamaño o capacidad
            validationsListObj.sizeOrCapacity.push([arrayData[i][20], arrayData[i][21]]);
          }
          
          // Validamos si viene los datos de Ancho
          if(arrayData[i][22] && arrayData[i][23] && arrayData[i][24]){
            
            // Agregamos el item de Ancho
            validationsListObj.width.push([arrayData[i][22], arrayData[i][23], arrayData[i][24]]);
          }
          
          // Validamos si viene los datos de calibre
          if(arrayData[i][30] && arrayData[i][31]){
            
            // Agregamos el item de calibre
            validationsListObj.caliber.push([arrayData[i][30], arrayData[i][31]]);
          }
          
          // Validamos si viene los datos de Modelo
          if(arrayData[i][25] && arrayData[i][26] && arrayData[i][27]){
            
            // Agregamos el item de Modelo
            validationsListObj.model.push([arrayData[i][25], arrayData[i][26], arrayData[i][27]]);
          }
          
          // Validamos si viene los caudal
          if(arrayData[i][32] && arrayData[i][33] && arrayData[i][34]){
            
            // Agregamos el item caudal
            validationsListObj.flow.push([arrayData[i][32], arrayData[i][33], arrayData[i][34]]);
          }
          
          // Validamos si viene los datos de separación
          if(arrayData[i][28] && arrayData[i][29]){
            
            // Agregamos el item de separación
            validationsListObj.separation.push([arrayData[i][28], arrayData[i][29]]);
          }
          
          // Validamos si viene los datos unidad
          if(arrayData[i][35]){
            
            // Agregamos el item de unidad
            validationsListObj.unit.push([arrayData[i][35]]);
          }
          
          // Validamos si viene los datos de moneda
          if(arrayData[i][36]){
            
            // Agregamos el item de moneda
            validationsListObj.currencySpecie.push([arrayData[i][36]]);
          }
          
        }
        
        // Retornamos los datos
        return validationsListObj;
      } 
    }
  }
  
  // Retornamos falso si  se pueden obtener los datos
  return false;
}

/*
* Función que permite obtener las diferentes cédulas
*/
function getDocumentRD(spreadSheet){

  // Objeto para los datos de las cedulas
  var objectData = {};
  
  // Referecciamos la pestaña de Rd
  var documentSheet = spreadSheet.getSheetByName("RD");
  
  // Validamos de que se tenga acceso a la hoja de cedulas RD
  if(documentSheet){
   
    // obtenemos la ultima fila 
    var lastRowDocument = documentSheet.getLastRow();
    
    // Validamos de que existan Cedulas
    if(lastRowDocument > 2){
      
      // Obtenemos los datos 
      var arrayDataDocument = documentSheet.getRange(3, 2, lastRowDocument, 4).getDisplayValues();
     
      // Recorremos los datos de las cedulas
      for(var j = 0; j < arrayDataDocument.length; j++){
        
        // Validamos de que no se encuentren vacias las celdas 
        if(arrayDataDocument[j][2] && arrayDataDocument[j][3]){
        
          // Validamos si no existe el item en el objeto
          if(!objectData[arrayDataDocument[j][2]]){
              
              // Creamos la propiedad en el objeto con un array vacio
              objectData[arrayDataDocument[j][2]] = [];
            }
            
            // Agregamos el item de la cedula
            objectData[arrayDataDocument[j][2]].push([arrayDataDocument[j][0].replace(/\./g, ""), arrayDataDocument[j][3]]);
          }         
        }
        
    }
  }

  // Retornamos los datos obtenidos 
  return objectData;

}

/*
* Función que permite obtener la variedad 
*/
function getVarietyData(spreadSheet){
    
  // Array para los datos de la familia 
  var objectData = {};
  
  // Referecciamos la pestaña de Relaciones
  var varietySheet = spreadSheet.getSheetByName("Relaciones");
  
  // Validamos de que se tenga acceso a la hoja de Relaciones
  if(varietySheet){
    
    // obtenemos la ultima fila 
    var lastRowVariety = varietySheet.getLastRow();
    
    // Validamos de que existan Relaciones
    if(lastRowVariety > 2){
      
      // Obtenemos los datos 
      var arrayDataVariety = varietySheet.getRange(2, 1, lastRowVariety, 3).getDisplayValues();
      
      // Recorremos los datos de las palabras claves
      for(var j = 0; j < arrayDataVariety.length; j++){
        
        // Validamos si no existe el item en el objeto
        if(!objectData[arrayDataVariety[j][1]]){
         
         // Creamos la propiedda en el objeto con un array vacio
         objectData[arrayDataVariety[j][1]] = [];
        }
        
        // Agregamos el item de la palabra clave
        objectData[arrayDataVariety[j][1]].push([arrayDataVariety[j][0], arrayDataVariety[j][2]]);
      }
    }
  }
  
  // Retornamos los datos obtenidos 
  return objectData;

}


/*+
* Función que permite obtener los datos de llas palabras claves
*/
function getKeywordData(spreadSheet){
   
  // Array para los datos de la familia 
  var objectData = {};
  
  // Referecciamos la pestaña de palabras claves
  var keywordSheet = spreadSheet.getSheetByName("Palabra clave");
  
  // Validamos de que se tenga acceso a la hoja de palabras claves
  if(keywordSheet){
    
    // obtenemos la ultima fila 
    var lastRowKeyword = keywordSheet.getLastRow();
    
    // Validamos de que existan palabras claves
    if(lastRowKeyword > 4){
      
      // Obtenemos los datos 
      var arrayDataKeyword = keywordSheet.getRange(2, 1, lastRowKeyword, 3).getDisplayValues();
      
      // Recorremos los datos de las palabras claves
      for(var j = 0; j < arrayDataKeyword.length; j++){
        
        // Validamos si no existe el item en el objeto
        if(!objectData[arrayDataKeyword[j][0]]){
         
         // Creamos la propiedda en el objeto con un array vacio
         objectData[arrayDataKeyword[j][0]] = [];
        }
        
        // Agregamos el item de la palabra clave
        objectData[arrayDataKeyword[j][0]].push([arrayDataKeyword[j][1]]);
      }
    }
  }
  
  // Retornamos los datos obtenidos 
  return objectData;
}

/*+
* Función que permite obtener los datos de llas palabras claves
*/
function getKeywordArrayData(spreadSheet){
   
  // Array para los datos de la familia 
  var arrayDataKeyword = []
  
  // Referecciamos la pestaña de palabras claves
  var keywordSheet = spreadSheet.getSheetByName("Palabra clave");
  
  // Validamos de que se tenga acceso a la hoja de palabras claves
  if(keywordSheet){
    
    // obtenemos la ultima fila 
    var lastRowKeyword = keywordSheet.getLastRow();
    
    // Validamos de que existan palabras claves
    if(lastRowKeyword > 4){
      
      // Obtenemos los datos 
      arrayDataKeyword = keywordSheet.getRange(2, 1, lastRowKeyword, 3).getDisplayValues();
      
      return arrayDataKeyword
    }
  }
  
  // Retornamos los datos obtenidos 
  return arrayDataKeyword;
}


/*+
* Función que permite obtener los datos de los departamentos
*/
function getDepartmentData(spreadSheet){

  // Referecciamos la pestaña de departamentos
  var departamentSheet = spreadSheet.getSheetByName("Departamentos");
  
  // Validamos de que se tenga acceso a la hoja de departamentos
  if(departamentSheet){
    
    // obtenemos la ultima fila 
    var lastRowDepartment = departamentSheet.getLastRow();
    
    // Validamos de que existan departamentos
    if(lastRowDepartment > 4){
      
      // Obtenemos los datos 
      var arrayDataDepartment = departamentSheet.getRange(5, 1, lastRowDepartment, 29).getDisplayValues();
      
      // retornamos los datos del departamento
      return arrayDataDepartment;
    }
  }
  
  // Retornamos vacio si no tenemos datos 
  return [];
}


/**
 * Función que permite obtener los datos de la familia
**/
function getFamilyData(spreadSheet){
   
  // Array para los datos de la familia 
  var arrayData = [];
  
  // Referecciamos la pestaña de Familias
  var familySheet = spreadSheet.getSheetByName("Familias");
  
  // Validamos de que se tenga acceso a la hoja de Familias
  if(familySheet){
    
    // obtenemos la ultima fila 
    var lastRowFamily = familySheet.getLastRow();
    
    // Validamos de que existan Familias
    if(lastRowFamily > 4){
      
      // Obtenemos los datos 
      var arrayDataSheet = familySheet.getRange(2, 1, lastRowFamily, 4).getDisplayValues();
      
      // Recorremos los datos de las familias
      for(var j = 0; j < arrayDataSheet.length; j++){
        
        // Agregamos el item de familias
       arrayData.push([arrayDataSheet[j][0],arrayDataSheet[j][1], arrayDataSheet[j][2], arrayDataSheet[j][3]]);
      }
    }
  }
  
  // Retornamos los datos obtenidos 
  return arrayData;
}