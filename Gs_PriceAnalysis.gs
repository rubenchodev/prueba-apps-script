/**
 * Función que permite obtener le análisis de los precios regsitrados hasta la fecha
**/
function getAnalysisDataByPrices(spreadsheet) {
  
  // validamos si no existe el libro como parametro
  if(!spreadsheet){
    
    // referenciamos la hoja adquisiciones
    spreadsheet = SpreadsheetApp.openById(getResources().acquisitionSheet);
  }
  
  // referenciamos la hoja de "Precios/Fabricantes"
  var sheet = spreadsheet.getSheetByName("Precios/Fabricantes");
  
  var costConcept = null;
  
  // Objeto con los datos de los conceptos
  var objectData = {
    manufacturerObj: {},
    speciesObj:{}
  };
  
  // se valida que exista la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // se valida si existe mas de una fila
    if(lastRow > 1){
      
      // obtenemos los datos
      var arrayData = sheet.getRange(2, 1, lastRow - 1, 31).getDisplayValues();
      
      // variable que almacena cada uno de los registros que debe mostrar
      var specieDataObj = {}, objectDataSpecie = {};
     
      // recorremos cada uno de los registros
      for(var i = (arrayData.length -1); i > 0 ; i--){
        
        // Validamos is existe la propiedad 
        if(!objectData.speciesObj[arrayData[i][3]]){
          
          //Creamos la especie con al información principal del concepto
          objectData.speciesObj[arrayData[i][3]] = {
            specieName: arrayData[i][3],
            manufacturersObj: {},
            unitRequested: arrayData[i][4],
            quantity: arrayData[i][17].split("|")[0]  
          };
        }
        
        // Validamos que no se encuentre el fabricante 
        if(!objectData.speciesObj[arrayData[i][3]].manufacturersObj[arrayData[i][2]]){
          
          // Obtenemos el costo de la especie por el fabricante
          costConcept = verifyCost(arrayData[i][18], arrayData[i][17]);
          
          // Obtenemos los datos del fabricante
          objectData.speciesObj[arrayData[i][3]].manufacturersObj[arrayData[i][2]] = {
            packageCode: arrayData[i][1],
            productNameHidro: arrayData[i][3],
            exchangeEur: (Number(arrayData[i][10]) || 0),
            manufacturerName: arrayData[i][2],
            manufacturerId: arrayData[i][0],
            priceUnit: arrayData[i][18],
            priceUnitAP: arrayData[i][18],
            price: arrayData[i][14],
            unit: arrayData[i][11],
            equivalency: arrayData[i][12],
            currencyType: arrayData[i][15],
            index: (i+2),
            code: arrayData[i][6],
            discountApplied: (Number(arrayData[i][16]) || 0),
            quantity: (Number(arrayData[i][17]) || 1),
            iva: (Number(arrayData[i][19]) || 0),
            productName: arrayData[i][7],
            updateDate: String(arrayData[i][29]),
            updateUserMail: arrayData[i][30],
            exchangeRateUsd: (Number(arrayData[i][9]) || 0),
            exchangeRateEur: (Number(arrayData[i][10]) || 0)
          }
          
          // Validamos si se tiene que realizar la conversiónn a dolares
          if(arrayData[i][16] == "USD"){
            
            // Agragamos el valor en pesos mexicanos
            objectData.speciesObj[arrayData[i][3]].manufacturersObj[arrayData[i][2]].priceUnit = costConcept * (Number(arrayData[i][9]) || 0);
            
            // Validamos si se tiene que realizar la conversiónn a euros  
          }else if(arrayData[i][16] == "EUR"){
            
            // Agragamos el valor en pesos mexicanos
            objectData.speciesObj[arrayData[i][3]].manufacturersObj[arrayData[i][2]].priceUnit = costConcept * (Number(arrayData[i][10]) || 0);
            
          }
        }
       
        // validamos si no existe la propiedad del fabricante
        if(!objectData.manufacturerObj[arrayData[i][2]]){
        
          // Creamos la propiedad 
          objectData.manufacturerObj[arrayData[i][2]] = {
           manufacturerData: {
            packageCode: arrayData[i][1],
            manufacturerName: arrayData[i][2],
            manufacturerId: arrayData[i][0],
            priceUnit: arrayData[i][18],
            priceUnitAP: arrayData[i][18],
            price: arrayData[i][14],
            unit: arrayData[i][11],
            equivalency: arrayData[i][12],
            currencyType: arrayData[i][15],
            index: (i+2),
            code: arrayData[i][6],
            discountApplied: (Number(arrayData[i][16]) || 0),
            quantity: (Number(arrayData[i][17]) || 1),
            iva: (Number(arrayData[i][19]) || 0),
            productName: arrayData[i][7],
            updateDate: String(arrayData[i][29]),
            updateUserMail: arrayData[i][30],
            exchangeRateUsd: (Number(arrayData[i][9]) || 0),
            exchangeRateEur: (Number(arrayData[i][10]) || 0)
           },
           specieArray:[]
          };
        }
        
        // insertamos los datos de la especie el el objeto de fabricantes
        objectData.manufacturerObj[arrayData[i][2]].specieArray.push(arrayData[i][3]);
        
      }
    }
  }
//throw JSON.stringify(objectData)
  // retornamos un objeto vacio
  return objectData;  
}




/**
 * Función que permite obtener le análisis de los precios regsitrados hasta la fecha
**/
function getAnalysisDataByPrices2(spreadsheet) {
  
  // validamos si no existe el libro como parametro
  if(!spreadsheet){
    
    // referenciamos la hoja adquisiciones
    spreadsheet = SpreadsheetApp.openById(getResources().acquisitionSheet);
  }
  
  // referenciamos la hoja de "Precios/Fabricantes"
  var sheet = spreadsheet.getSheetByName("Precios/Fabricantes");
  
  // Objeto con los datos de los conceptos
  var objectData = {
    manufacturerObj: {},
    speciesObj:{}
  };
  
  // se valida que exista la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // se valida si existe mas de una fila
    if(lastRow > 1){
      
      // obtenemos los datos
      var arrayData = sheet.getRange(2, 1, lastRow - 1, 31).getDisplayValues();
      
      // variable que almacena cada uno de los registros que debe mostrar
      var specieDataObj = {};
     
      // recorremos cada uno de los registros
      for(var i = 0; i < arrayData.length; i++){
        
        // Agregamos la información al array principal
        specieDataObj = { 
        
            manufacturerId: arrayData[i][0],
            packageId: arrayData[i][1],
            manufacturerName: arrayData[i][2],
            specieNameClient: arrayData[i][7],
            unitRequested: arrayData[i][4],
            requestDate: arrayData[i][5],
            productCode: arrayData[i][6],
            imageName: arrayData[i][8],
            imageId: arrayData[i][9],
            exchangeRateUsd: arrayData[i][10],
            exchangeRateEur: arrayData[i][11],
            unitManufacturer: arrayData[i][12],
            equivalent: arrayData[i][13],
            minimumAmount: arrayData[i][14],
            pList: arrayData[i][15],
            currency: arrayData[i][16],
            appliedDiscount: arrayData[i][17],
            quotedAmount: arrayData[i][18],
            offeredUnitP: arrayData[i][19],
            iva: arrayData[i][20],
            importQuoted: arrayData[i][21],
            deliveryTime: arrayData[i][22],
            deliveryUnit: arrayData[i][23],
            updateDate: arrayData[i][29],
            updateUserMail: arrayData[i][30],
            rowIndex: (i+2)
        }; 
        
       
        // validamos si no existe la especie en el objeto
        if(!objectData.speciesObj[arrayData[i][3]]){
        
          // Creamosla propiedad he ingresamos los datos 
          objectData.speciesObj[arrayData[i][3]] = specieDataObj;
          
        }
        
        // validamos si no existe la propiedad del fabricante
        if(!objectData.manufacturerObj[arrayData[i][2]]){
          // Creamos la propiedad 
          objectData.manufacturerObj[arrayData[i][2]] = {};
        }
        
        // insertamos los datos de la especie el el objeto de fabricantes
        objectData.manufacturerObj[arrayData[i][2]][arrayData[i][3]] = specieDataObj;
      }
    }
  }
  
  // retornamos un objeto vacio
  return objectData;  
}




/*
* Función que permite verificar el costo real del producto  
*/
function verifyCost(priceUnit, quatity){
   
   // Obtenemos la cantidad 
   var arrayQuatity = quatity.split("|");
   
   // Obtenemosel precio unitario ofertado
   var conceptQuatity = Number(priceUnit) || 0;
   
   //Validamos si trae dos posiciones
   if(arrayQuatity.length > 1){
     
     // Obtenemos el precio unitario respecto al precio segun la unidad solicitada
     conceptQuatity = (conceptQuatity * (Number(arrayQuatity[1]) || 0)) / (Number(arrayQuatity[0]) || 0);
   }

   // Retornamos el precio del producto 
   return  conceptQuatity;

}

/*
* función que permite obtener las ordenes de compra
*/
function getOrdersHistory(spreadsheet){
  
  // Validamos si no viene la referencia a la hoja
  if(!spreadsheet){
  
    // referenciamos la hoja adquisiciones
    spreadsheet = SpreadsheetApp.openById(getResources().ordersHistorySheetId);
  }
  //throw spreadsheet.getUrl()
  // referenciamos la hoja de "Órdenes"
  var sheet = spreadsheet.getSheetByName("Órdenes");
  
  // se valida que exista la hoja
  if(sheet){
    
    // obtenemos la ultima fila
    var lastRow = sheet.getLastRow();
    
    // se valida si existe mas de una fila
    if(lastRow > 1){
      
      // obtenemos los datos
      var arrayData = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
      
      // Array para las ordenes de compra
      var arrayOrders = [];
      
      // recorremos cada uno de los registros
      for(var i = arrayData.length - 1; i >= 0; i--){
        
        // validamos si el estado es diferente de "Eliminada"
        if(arrayData[i][6] != "Eliminada"){
          
          // Agregamos los datos de la orden al array 
          arrayOrders.push({
            orderId: arrayData[i][0],
            manufacturerName: arrayData[i][1],
            manufacturerEmail: arrayData[i][2],
            createDate: arrayData[i][3],
            creatorEmail: arrayData[i][4],
            url: arrayData[i][5],
            state: arrayData[i][6],
            importe: getImporteByOrder_(arrayData[i][7]),
            productData: getProductData_(arrayData[i][8]),
            index: (i + 2)
          });
        
        }
      }
      //throw JSON.stringify(arrayOrders)
      // retornamos el listado de ordenes
      return arrayOrders;
    }
  }
  
  // retornamos un objeto vacio
  return  [];
  
}

/**
 * Permite obtener los datos 
 */
function getProductData_(productData_){
  
  try{
    return JSON.parse(productData_);
  } catch(ex){}

  // retornamos por defecto un null
  return null;
}

/**
 * Permite obtener el importe de la orden de compra
 */
function getImporteByOrder_(currentValue_){

  // Se valida si no existe el importe
  if(!currentValue_) return 0;

  // Reemplazamos la coma por el punto
  currentValue_ = currentValue_.toString().replace(/\,/g, ".");

  // Retornamos el importe
  return (Number(currentValue_) || 0);
}

/*
* Función que permite eliminar la orden de compra
*/
function deleteHistoryOrder(orderData){
  
  // Obtenemos los recursos globales 
  var resources = getResources();
  
  // Hacemos referencia a la hoja de historial de ordenes de compra
  var spreadsheet = SpreadsheetApp.openById(resources.ordersHistorySheetId);
  
  // validamos si existen correos y esta el archivo de la orden de compra 
  if(orderData.index && spreadsheet){
    
    // referenciamos la hoja de "Órdenes"
    var orderSheet = spreadsheet.getSheetByName("Órdenes");
    
    // se valida que exista la hoja
    if(orderSheet){
      
      // Cambiamos el estado de la orden a Enviada
      orderSheet.getRange(orderData.index, 7, 1, 1).setValue("Eliminada");
    }
  }
  
  // Retornamos el listado de la ordenes de compra 
  return getOrdersHistory(spreadsheet);
  
}

/*
* Función que permite enviar la orden de compra al fabricante 
*/
function sendOrderToManufacturer(orderData){
  
  // Obtenemos los recursos globales 
  var resources = getResources()
  
  // Hacemos referencia a la hoja de historial de ordenes de compra
  var spreadsheet = SpreadsheetApp.openById(resources.ordersHistorySheetId);
  
  // Obtenemos la referencia al archivo en pdf generado 
  var orderPdf = createPdf(orderData.url, resources.ordersSentFolderId);
  
  // validamos si existen correos y esta el archivo de la orden de compra 
  if(orderData.manufacturerEmail && orderPdf){
    
    // obtenemos el correo del usuario que realiza la creación de la orden de compra
    var userMail = getUserMail();
    
    // obtenemos los datos como nombre y foto
    var userData = getAccountUserPanel(userMail, "");
    
    //Referenciamos el contenido de la plantilla
    var contentNotificationOrder = HtmlService.createHtmlOutputFromFile("Html_MailTemplateOrder").getContent();
    
    //Insertamos el nombre del cliente  |*SALES_AGENT*|
    contentNotificationOrder = contentNotificationOrder.replace("|*USER_MANUFACTURER*|", orderData.manufacturerName);
    
    // Insertamos el nombre del usuario activo 
    contentNotificationOrder = contentNotificationOrder.replace("|*SALES_AGENT*|", userData.fullName);
    
    //Enviamos el correo
    MailApp.sendEmail({
      to: orderData.manufacturerEmail,
      subject: ("Orden número: " + orderData.orderId),
      htmlBody: contentNotificationOrder,
      name: "Hidrocotizador",
      attachments: orderPdf
    });
    
    // Validamos que se tenga acceso la hoja de ordenes
    if(spreadsheet){
      
      // referenciamos la hoja de "Órdenes"
      var orderSheet = spreadsheet.getSheetByName("Órdenes");
      
      // se valida que exista la hoja
      if(orderSheet){
      
        // Cambiamos el estado de la orden a Enviada
        orderSheet.getRange(orderData.index, 7, 1, 1).setValue("Enviada");
      }
    }
  }
  
  // Retornamos el listado de la ordenes de compra 
  return getOrdersHistory(spreadsheet);
  
}


/*
* Función que permite crear el pdf y alojarlo en la carpeta
*/
function createPdf(url, ordersSentFolderId){
    
    // Obtenemos el id del archivo
    var orderId = url.match(/[-\w]{25,}/);
    
    // Referencimos la orden para enviarla    
    var file = DriveApp.getFileById(orderId);
    
    // Hacemos referencia a la carpeta donde se va a guarder la orden 
    var folder = DriveApp.getFolderById(ordersSentFolderId);
    
    // referenciamos la orden de compra en pdf 
    var filePdf = file.getAs('application/pdf');
    
    // Creamos la orden en formato pdf
    folder.createFile(filePdf); 
    
    // Retornamos la referencia de la orden en pdf 
    return filePdf;
}



