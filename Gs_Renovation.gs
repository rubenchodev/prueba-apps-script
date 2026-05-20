/**
 * Función que permite obtener las cotizaciones de refacciones
**/
function filterRenovationsByMail(userRole){
  
  // se obtiene el correo del usuario
  var userMail = getUserMail();
  
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // referenciamos la hoja de cálculo de solicitudes de cotizaciones
  var spreadsheet = SpreadsheetApp.openById(resources.renovationSheetId);
  
  // obtenemos la hoja de "Solicitudes"
  var renovationSheet = spreadsheet.getSheetByName("Solicitudes");
  
  // validamos si existe la hoja
  if(renovationSheet){
    
    // obtenemos la ultima fila
    var lastRow = renovationSheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
      
      // obtenemos todos los datos de la hoja
      var allData = renovationSheet.getRange(2, 1, lastRow - 1, 11).getValues();
      
      // variable para almacenar la solicitudes
      var requestData = [];
      
      // recorremos cada uno de los registros y los filtramos por usuario
      for(var i = 0; i < allData.length; i++){
      
        // validamo si el correo el usuario activo es el mismo de la cotización recorrida en el momento
        if((allData[i][0] == userMail || userRole == "executiveDirection") && allData[i][9] == "Activa"){
          
          // agregamos el registro
          requestData.push({
            quotationNumber: allData[i][1],
            client: allData[i][2],
            issueDate: allData[i][3],
            validateDate: allData[i][4],
            author: allData[i][5],
            importe: allData[i][6],
            documentId: allData[i][7],
            googleDoc: allData[i][8],
            structure: allData[i][10],
            row: Number(i) + 2
          });
          
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
 * Función que permite generar la cotización de refacciones
**/
function generateCotRenovation(formData, documentId, rowIndex, lastRowTemp, consecutive, userRole){
  
  // variable para guardar la referencia del archivo de la plantilla
  var spreadsheet = null,
      document = null,
      mainFolder = null,
      message = "Cotización de refacciones generada con éxito.";
  
  // referenciamos los datos globales
  var resources = getResources();

  // Validamos si no existe la plantilla
  if(!formData.templateId){
    formData.templateId = resources.renovationTemplateId; // HEXP - Plantilla #1
  }

  // Validamos si no existe la carpeta
  if(!formData.folderId){
    formData.folderId = resources.renovationFolderId; // Proyectos-Cotizaciones
  }
  
  // obtenemos la fecha actual
  var currentDate = new Date(), fileName = "";
  console.log(JSON.stringify(formData));
  
  // referenciamos la carpeta donde se desea almacenar la cotización
  mainFolder = DriveApp.getFolderById(formData.folderId);
  
  // validamos si existe el id del documento lo que significa que se va a modificar la cot
  if(documentId){
    
    // referenciamos el libro
    spreadsheet = SpreadsheetApp.openById(documentId);
    
    // Variable para administrar el contador de archivos Google doc
    var counterFiles = 0,
        currentDoc = null;
        
    // se establece el nombre del archivo
    fileName = "@PR" + consecutive + ". " + formData.client;
    
    // obtenemos la cantidad de archivos actuales en la carpeta de tipo Google Doc
    var googleDocFiles = mainFolder.getFilesByType(MimeType.GOOGLE_DOCS);
    
    // recorremos el iterador de archivos
    while(googleDocFiles.hasNext()){
      
      // Referenciamos el archivo recorrido
      currentDoc = googleDocFiles.next();
      
      // se valida si el nombre es el mismo de la cotización actual
      if(currentDoc.getName().indexOf(fileName) != -1){
      
        // aumentamos el contador
        counterFiles ++;
      }
      
    }
    
    // validamos si el contador es mayor a cero(0)
    if(counterFiles > 0){
      
      // ajustamos el nombre para agregarla la versión
      fileName += " V" + (counterFiles + 1); 
      
    }
    
    // cambiamos un mensaje
    message = "Cotización de refacción actualizada con éxito.";
    
  } else {
    
    // generamos el consecutiva 
    consecutive = generateConsecutive(currentDate);
    
    // variable para definir el nombre del archivo
    fileName = "@PR" + consecutive + ". " + formData.client;
    
    // se crea una copia de la plantilla de refacciones
    var fileCopy = DriveApp.getFileById(resources.templateRenovationId).makeCopy(fileName, mainFolder);
    
    // Agregamos los permisos para que cuaquier usuarion dentro de Hidroexpert pueda editar
    // fileCopy.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
    
    // se referencia el archivo
    spreadsheet = SpreadsheetApp.open(fileCopy);
    
  }
  
  // crearmos un documento a partir de una plantilla
  var fileByTemplate = DriveApp.getFileById(formData.templateId).makeCopy(fileName, mainFolder);
  
  // Agregamos los permisos para que cuaquier usuarion dentro de Hidroexpert pueda editar
  // fileByTemplate.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
  
  // creamos el documento plantilla
  document = DocumentApp.openById(fileByTemplate.getId());
  
  // obtenemos el correo del usuario activo
  var userMail = getUserMail();
  
  // obtenemos los datos publicos del usuario activo
  var userData = getAccountUserPanel(userMail, "Usuario");
  
  // referenciamos la hoja de "Proyecto"
  var projectSheet = spreadsheet.getSheetByName("Proyecto");
  // throw JSON.stringify(userData)
  // validamos si existe el id del documento lo que significa que se va a modificar la cot
  if(documentId){
    
    // borramos el contenido de especies
    projectSheet.getRange("A15:J" + lastRowTemp).clearContent();
    
    // se valida si la fila temporal es igual a 16
    if(lastRowTemp == 16){
      
      // se elimina una sola fila
      projectSheet.deleteRow(16);
      
    } else {
    
      // Removemos las filas ya no utilizadas
      projectSheet.deleteRows(16, lastRowTemp - 16);
    }
    
  }
  
  // agregamos los datos basicos de la renovación
  addRenovationBasicData(projectSheet, formData, consecutive, userData);
  
  // ordenamos las especies de tal manera que primero muestre los de Euro, Dolar y finalmenet la moneda Mexicana
  formData.conceptList.sort(function(a, b){
    return a.index - b.index;
  });
  
  // agregamos las especies a la cotización
  var conceptInfoObj = addConceptByRenovation(projectSheet, formData.conceptList, formData.euValue, formData.usdValue, formData.deliveryTime);
  
  // agregamos los datos al documento
  insertDataByGoogleDoc(document, conceptInfoObj, formData, consecutive, userData);
  
  // registramos la información de la refacción en la hoja principal
  registerGeneralDataRenovation(resources.renovationSheetId, spreadsheet.getId(), formData, consecutive, rowIndex, userData, document);
  
  // guardamos los datos del documento
  document.saveAndClose();
  
  // validamos si se requiere enviar el correo electrónico
  if(formData.sendMail && formData.email){
    
    // asignamos permisos al documento de edición al usuario respectivo
    DriveApp.getFileById(document.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.COMMENT);
    
    // se llama la función que enviar el correo electronico
    sendMail(document.getId(), formData.client, formData.email, consecutive, userData);
  }
  
  // se retorna la respuesta
  return {
    message: message,
    requestData: filterRenovationsByMail(userRole)
  };
  
}

/**
 * @desc: Función que envia la notificación una vez se aprueba o rechaza una solicitud
 * @param: {string} subject: asunto
 * @param: {string} recipient: correo(s) del destinatario
 * @param: {object} mailOptions: [greeting: "saludo al destinatario", message: "mensaje del correo"]
 * @author: Felipe López
 */
 function sendMail(fileId, userName, userMail, consecutive, userData){
   
   // referenciamos el archivo adjunto
   var doc = DocumentApp.openById(fileId);
   
   //Referenciamos el contenido de la plantilla
   var contentNotification = HtmlService.createHtmlOutputFromFile("Html_MailTemplate").getContent();
   
   //Insertamos el nombre del usuario
   contentNotification = contentNotification.replace("|*USER_NAME*|", userName);
   
   //Insertamos el vínculo del archivo
   contentNotification = contentNotification.replace("|*LINK*|", doc.getUrl());
   
   //Insertamos el consecutivo de la solicitud
   contentNotification = contentNotification.replace("|*NUMBER*|", consecutive);
   
   //Insertamos el nombre de la persona que es el agente de ventas
   contentNotification = contentNotification.replace("|*SALES_AGENT*|", (userData && userData.fullName) ? userData.fullName : "No definido");
   
   //Enviamos el correo
   MailApp.sendEmail({
     to: userMail,
     subject: ("Agrocity: Cotización de refacción número " + consecutive),
     htmlBody: contentNotification,
     name: "Cotización de refacción"
   });
}

/**
 * Función que permite generar el consecutivo de la cotización de refacción
**/
function generateConsecutive(currentDate){
  
  // obtenemos las propiedades del script
  var scriptProperties = PropertiesService.getScriptProperties();
  
  // obtenemos la propiedad del consecutivo
  var consecutiveCustom = (scriptProperties.getProperty("Consecutive") || "1").replace(".0", "");
  
  // variable para guardar el consecutivo
  var consecutive = currentDate.getDate() < 10 ? ("0" + currentDate.getDate()) : currentDate.getDate();
  
  // agregamos el mes
  consecutive += (currentDate.getMonth() + 1) < 10 ? ("0" + (currentDate.getMonth() + 1)) : (currentDate.getMonth() + 1);
  
  // agregamos el mes
  consecutive += String(currentDate.getFullYear()).substring(2, 4);
  
  // agregamos el consecutivo 
  consecutive += (consecutiveCustom.length < 3 ? Utilities.formatString('%03d', consecutiveCustom) : consecutiveCustom);
  
  // convertimos en numero el valor
  consecutiveCustom = Number(consecutiveCustom);
  
  // guardamos el nuevo consecutivo
  scriptProperties.setProperty("Consecutive", (consecutiveCustom + 1));
  
  // retornamos el consecutivo
  return "" + consecutive;  
}

/**
 * Función que permite agregar los datos basicos sobre el archivo
**/
function addRenovationBasicData(projectSheet, formData, consecutive, userData) {
  
  // agregamos el consecutivo y el nombre de quien creo la cotización  
  projectSheet.getRange("E2:E3").setValues([[consecutive],[userData.fullName]]);
  
  // agregamos los datos de plazos
  projectSheet.getRange("H3:H5").setValues([
    [formData.issue],
    [formData.validity],
    [formData.state + ", " + formData.municipality + ", " + formData.location]
  ]);
  
  // agregamos los datos de razón social del solicitante
  projectSheet.getRange("F8:F10").setValues([
    [formData.client],
    [formData.federalRecord],
    [formData.fiscalRecidence]
  ]);
  
  // agregamos los datos del contacto de ventas
  projectSheet.getRange("G8:G9").setValues([
    [formData.salesContact],
    [formData.phone]
  ]);
  
  // agregamos el correo del contacto de ventas
  projectSheet.getRange("I9").setValue(formData.email);
  
}

/**
 * Función que permite agregar las especies en la cotización de refaciones
**/
function addConceptByRenovation(projectSheet, conceptList, euValue, usdValue, deliveryTimeList){
  
  // definimos la fila inicial
  var startRow = 14,
      currentConcept = [],
      exchangeRateObj = {
        "EU": [],
        "USD": [],
        "MN": []
      };
  
  // variable para obtener las refacciones que no estan en el inventario
  var deliveryTimeArray = [],
      conceptsArray = [],
      counterConcepts = 0;
  
  // recorremos cada uno de las especies
  for(var i = 0; i < conceptList.length; i++){
    
    // creamos la fila de la nueva orden
    if(i > 0)
      projectSheet.insertRowAfter(startRow);
    
    // aumentamos el contador
    startRow ++;
    
    // obtenemos los datos de las especies actual
    currentConcept = conceptList[i].conceptData;
    
    // agregamos la moneda y el inventario
    projectSheet.getRange("A" + startRow + ":C" + startRow).setValues([[(i + 1), currentConcept[6], conceptList[i].inventory]]);
    
    // agregamos la descripción
    projectSheet.getRange("F" + startRow).setValue(currentConcept[1]).setWrap(true);
    
    // agregamos la unidad
    projectSheet.getRange("G" + startRow).setValue(currentConcept[4]).setHorizontalAlignment("center");
    
    // agregamos la cantidad, C. unitario e Importe
    projectSheet.getRange("H" + startRow + ":O" + startRow).setValues([[conceptList[i].conceptCount, "=N" + startRow + "*O" + startRow, "=I" + startRow + "*H" + startRow, "", "", "", currentConcept[7], "1"]]).setHorizontalAlignment("right");
    
    // obtenemos el tipo de moneda
    var moneyType_ = (["MXN", "MN"].indexOf(currentConcept[6]) != -1) ? "MN" : currentConcept[6];

    // validamos que el tipo de moneda sea valido
    if(exchangeRateObj["" + moneyType_]){
    
      // grardamos el rango segun el tipo de moneda
      exchangeRateObj["" + moneyType_].push("J" + startRow);
    }
    
    // validamos si la referencia pertenece a una de las que no estan en el inventario
    if(deliveryTimeList[conceptList[i].reference]){
      
      // agregamos los datos de la posición de la refación y el tiempo
      deliveryTimeArray.push([
        (i + 1),
        deliveryTimeList[conceptList[i].reference]
      ]);
      
    }
  }
  
  // obtenemos los datos actuales de especies
  conceptsArray = projectSheet.getRange("A15:J" + startRow).getDisplayValues();
  
  // agregamos la formula de contador de productos
  projectSheet.getRange("H" + (startRow + 4)).setValue("=SUM(H15:H" + startRow + ")");
  
  // obtenemos la cantidad
  counterConcepts = projectSheet.getRange("H" + (startRow + 4)).getDisplayValue();
  
  // se define la fila del subtotal de euros
  var subTotalEU = (startRow + 7);
  
  // se define la fila del subtotal de Dolares
  var subTotalUSD = (startRow + 8);
  
  // se define la fila del subtotal de pesos mexicanos
  var subTotalMN = (startRow + 9);
  
  // Agregamos la formula del total en euros
  projectSheet.getRange("H" + subTotalEU + ":J" + subTotalEU).setValues([[euValue, (exchangeRateObj["EU"].length > 0 ? "=" + exchangeRateObj["EU"].toString().replace(/\,/g, "+") : "0"), ("=I" + subTotalEU + "*H" + subTotalEU)]]);
  
  // Agregamos la formula del total en Dolares
  projectSheet.getRange("H" + subTotalUSD + ":J" + subTotalUSD).setValues([[usdValue, (exchangeRateObj["USD"].length > 0 ? "=" + exchangeRateObj["USD"].toString().replace(/\,/g, "+") : "0"), ("=I" + subTotalUSD + "*H" + subTotalUSD)]]);
  
  // Agregamos la formula del total en Pesos mexicanos
  projectSheet.getRange("I" + subTotalMN + ":J" + subTotalMN).setValues([[(exchangeRateObj["MN"].length > 0 ? "=" + exchangeRateObj["MN"].toString().replace(/\,/g, "+") : "0"), ("=I" + subTotalMN + "*H" + subTotalMN)]]);
  
  // Agregamos el total
  projectSheet.getRange("J" + (startRow + 10)).setValue("=J" + subTotalEU + "+J" + subTotalUSD + "+J" + subTotalMN);
  
  // agregamos la formula de total en letras
  projectSheet.getRange("C" + (startRow + 11)).setValue("=importex(J" + (startRow + 10) + ")");
  
  // obtenemos los datos de subtotal y total
  var totalArray = projectSheet.getRange("H" + subTotalEU + ":J" + subTotalMN).getDisplayValues();
  
  // retornamos la matriz de las refacciones que se necesitan programar la entrega en una fecha posterior
  return {
    deliveryTimeArray: deliveryTimeArray,
    conceptsArray: conceptsArray,
    counter: counterConcepts,
    importe:{
      number: projectSheet.getRange("J" + (startRow + 10)).getDisplayValue(),
      letter: projectSheet.getRange("C" + (startRow + 11)).getDisplayValue()
    },
    totalArray: totalArray
  };
}

/**
 * Función que permite registrar los datos en la hoja principal de solicitudes de refacción
**/
function registerGeneralDataRenovation(spreadsheetId, documentId, formData, consecutive, rowIndex, userData, googleDoc){
  
  // referenciamos la hoja de cálculo general donde quedan todos los datos de las cotizaciones generadas
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  // obtenemos la hoja de "Solicitudes"
  var sheet = spreadsheet.getSheetByName("Solicitudes");
  
  // validamos si existe la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // agregamos los nuevos valores
    sheet.getRange(((rowIndex != -1) ? rowIndex : (lastRow + 1)) , 1, 1, 11).setValues([[
      userData.email,
      consecutive,
      formData.client,
      formData.issue,
      formData.validity,
      userData.fullName,
      formData.total,
      documentId,
      googleDoc.getId(),
      "Activa",
      JSON.stringify(formData)
    ]]);
    
  }
  
}

/**
 * Función que permite obtener todos las especies
**/
function getPricesList() {
  
  // referenciamos la hoja de cálculo de configuración
  var spreadsheet = SpreadsheetApp.openById(getResources().settingSheetId);
  
  // obtenemos la hoja "Descripción precios"
  var sheet = spreadsheet.getSheetByName("Descripción precios");
  
  // validamos si existe la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
      
      // retornamos los datos
      return sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      
    }
  }
  
  // retornamos por defecto un array vacio
  return [];
}

/**
 * Función que permite el envio de correo al área de adquisiciones
**/
function notifyAcquisitions(cotFileId, consecutive){
  
  // obtenemos la referencia de los archivos principales
  var resources = getResources();
  
  // referenciamos la hoja de usuarios
  var spreadsheet = SpreadsheetApp.openById(resources.userFileId);
  
  // referenciamos la hoja de "Adquisiciones"
  var sheet = spreadsheet.getSheetByName("Adquisiciones");
  
  // validamos si existe la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // validamos si existe mas de una fila
    if(lastRow > 1){
      
      // obtenemos los correo separados por coma
      var emails = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      
      // validamos si existen correos
      if(emails){
        
        // capturamos el error que puede que surga en al momento de enviar la notificación
        try{
        
          // asignamos permisos al documento de edición al usuario respectivo
          DriveApp.getFileById(cotFileId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.COMMENT);
          
          // referenciamos el archivo adjunto
          var docSheet = SpreadsheetApp.openById(cotFileId);
          
          //Referenciamos el contenido de la plantilla
          var contentNotification = HtmlService.createHtmlOutputFromFile("Html_MailTemplateNotify").getContent();
          
          //Insertamos el vínculo del archivo
          contentNotification = contentNotification.replace("|*LINK*|", docSheet.getUrl());
          
          //Enviamos el correo
          MailApp.sendEmail({
            to: emails.toString(),
            subject: ("Aprobación de la cotización de refacción número " + consecutive),
            htmlBody: contentNotification,
            name: "Cotización de refacción"
          });
        
        } catch(e){
          
          // mostramos un mensaje de error
          throw "No tienes acceso a la cotización de refacción."
          
        }
        
        // retornamos un mensaje
        return {
          message: "Notificación enviada satisfactoriamente."
        };
        
      }
    }
  }
  
  // retornamos un objeto por defecto
  return {
    message: "No existe el área de adquisiciones para enviar la respectiva notificación"
  };
  
}

/**
 * Función que permite obtener la información de las lineas de productos y su prioridad
**/
function getPriorityProductLines(settingSpreadsheet, userMail) {
  
  // validmoa si no existe el parametro de configuración
  if(!settingSpreadsheet){
  
    // referenciamos la hoja de cálculo de configuración
    settingSpreadsheet = SpreadsheetApp.openById(getResources().settingSheetId);
    
    // obtenemos el correo del usuario activo
    userMail = getUserMail();
  }
  
  // obtenemos la hoja "Líneas de productos"
  var sheet = settingSpreadsheet.getSheetByName("Líneas de productos");
  
  // definimos la variable para guardar los datos validos del usuario
  var returnObj = {
    data: [],
    column: -1,
    exist: false
  };
  
  // validamos si existe la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // validamos que exista mas de 1 fila
    if(lastRow > 1){
      
      // retornamos los datos
      var arrayData = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
      
      // validamos si existe algun dato
      if(arrayData.length > 0){
        
        // validamos que exista mas de 3 columnas
        if(arrayData[0].length > 3){
        
          // Recorremos la primera columna y validamos si el usuario esta en alguna de las columnas
          for(var col = 3; col < arrayData[0].length; col++){
            
            // validamos si el correo pertenece a uno de las columnas
            if(userMail == arrayData[0][col]){
              
              // definimos en que columna se encuentra (TIPO MATRIZ 0-n)
              returnObj.column = col;
              
              // Definimos que si existe el usuario
              returnObj.exist = true;
              
              // salimos del ciclo
              break;
              
            }
          }
        }
        
        // ahora recorremos cada uno de las solicitudes
        for(var row = 1; row < arrayData.length; row++){
          
          // agregamos los datos al listado
          returnObj.data.push([
            arrayData[row][1],
            arrayData[row][0],
            arrayData[row][2],
            (returnObj.exist ? arrayData[row][returnObj.column] : "")
          ]);
          
        }
      
      }
    }
  }
  
  // retornamos el objeto
  return returnObj;
}

/**
 * Función que permite guardar la configuración del usuario
**/
function savePriorityProductLine(column, arrayData){
  
  // referenciamos la hoja de cálculo de configuración
  var settingSpreadsheet = SpreadsheetApp.openById(getResources().settingSheetId);
  
  // obtenemos el correo del usuario activo
  var userMail = getUserMail();
  
  // obtenemos la hoja "Líneas de productos"
  var sheet = settingSpreadsheet.getSheetByName("Líneas de productos");
  
  // validamos si la columna es igual a -1
  if(column == -1){
    
    // obtenemos la ultima columna y le aumentamos en 1
    column = sheet.getLastColumn() + 1;
    
    // almacenamos el correo del usuario
    sheet.getRange(1, column).setValue(userMail);
    
  } else {
    
    // aumentamos en 1 la columna
    column ++;
    
  }
  
  // almacenamos los datos de prioridad
  sheet.getRange(2, column, arrayData.length, 1).setValues(arrayData);
  
  // retornamos la nueva columna
  return column - 1;
  
}