/**
 * Función que permite obtener los datos de la pestañana federal siempre y cuando la tenga
 * @author Rubén Sánchez
**/
function getFederalData(fileId, attempts_){
  attempts_ = attempts_ || 1;
 
  // referenciamos la hoja de la cotización
  var spreadsheet = SpreadsheetApp.openById(fileId);
  
  // referenciamos la hoja de Federal
  var federalSheet = spreadsheet.getSheetByName("Federal");
  
  // Si existe la hoja
  if(federalSheet){
    
    // obtenemos la ultima fila
    var lastRow = federalSheet.getLastRow();
    
    // se valida que existan mas de 25 filas
    if(lastRow > 25){
      
      // obtenemos los datos de orden, linea y especies
      var arrayData = federalSheet.getRange(22, 2, lastRow - 25, 9).getDisplayValues();

      // Intentamos obtener el valor válido de importeLetter
      var importeLetter_ = federalSheet.getRange(lastRow, 3).getDisplayValue();

      // se valida si el importe es valido
      if(isValidImporteLetter_(importeLetter_)){
      
        // retornamos el objeto con los datos
        return {
          concepts: arrayData,
          importe: federalSheet.getRange(lastRow - 1, 10).getDisplayValue(),
          //importeLetter: federalSheet.getRange(lastRow, 3).getDisplayValue()
          importeLetter: importeLetter_
        }
      } else {

        // Se valida si la cantidad de intentos es menor que 3
        if(attempts_ < 3){
          attempts_++;

          // Esperamos 500 milisegundos
          Utilities.sleep(500);

          // retornamos la respuesta
          return getFederalData(fileId, attempts_);
        } else {
          throw "La formula del importe en letras no se finalizó de ejecutar, por favor intenta nuevamente.";
        }
      }
      
    }
    
  } else {
    
    // retornamos el mensaje de que no existe el federal
    throw "Por favor se revise el presupuesto editable, al parecer la hoja federal no ha sido generada";
  }
  
  // se retorna
  return null;  
}

/**
 * Obtiene el valor del importe
 */
function getValidImporteLetter_(sheetFederal_, fileId, lastRow, maxAttempts) {
  var baseSleepTime = 500; // Tiempo base en milisegundos
  var importeLetter;
  
  for (var attempt = 0; attempt < maxAttempts; attempt++) {
    console.log("Intento: " + attempt)
    // Se valida si el intento es mayor a cero
    if(attempt > 0){
      // Reabrimos la hoja de cálculo en cada intento
      var spreadsheet = SpreadsheetApp.openById(fileId);
      sheetFederal_ = spreadsheet.getSheetByName("Federal");
    }
    
    // Obtenemos el valor de importeLetter
    importeLetter = sheetFederal_.getRange(lastRow, 3).getDisplayValue();
    
    if (isValidImporteLetter_(importeLetter)) {
      return importeLetter;
    }
    
    // Calculamos el tiempo de espera exponencial
    var sleepTime = baseSleepTime;// * Math.pow(2, attempt);
    Utilities.sleep(sleepTime);
  }
  
  // Si después de todos los intentos no se obtuvo un valor válido, lanzamos un error
  throw "La formula del importe en letras no se finalizó de ejecutar, por favor intenta nuevamente.";
}

/**
 * Permite validar el importe
 */
function isValidImporteLetter_(importeLetter) {
  console.log("importeLetter: " + importeLetter)
  // Verificamos si el valor es válido (no es un error de fórmula)
  var invalidValues = ["#NAME?", "#N/A", "#DIV/0!", "#VALUE!", "#REF!", "#NUM!", "#NULL!"];
  return importeLetter && invalidValues.indexOf(importeLetter) === -1;
}

/**
* Función que permite validar si existe el tag y agregar las respectivas imagens
* @author Rubén Sánchez
*/
function initCreationStatalFile(imageList, requestData, estimateData, userRole, year) {

  // referenciamos el archivo de la cotización
  var projectSheet = DriveApp.getFileById(requestData.documentId);
  
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  var mainFolder = DriveApp.getFolderById(resources.statalFolderId);
  
  // obtenemos las ubicaciones
  /*var parents = projectSheet.getParents();
  
  // variable para referenciar el padre
  var mainFolder = null;
  
  // se valida si existe algun padre
  if(parents.hasNext()){
    
    // referenciamos la carpeta donde esta ubicado el archivo
    mainFolder = parents.next();
    
  } else {
  
    // se referencia la carpeta de mi unidad
    mainFolder = DriveApp.getRootFolder();  
  }*/
  
  // se obtiene la fecha actual
  var currentDate = new Date();
  
  // fecha en formato DD de MES de YYYY
  var stringDate = formatStringDate(currentDate);
  
  // referenciamos la estructura actual
  var currentStructure = JSON.parse(requestData.structures[requestData.mainSheetName].structure);
  
  // se define el nombre del archivo
  var fileName = "Cotización estatal: " + currentStructure.client + " - " + stringDate;
  
  // se realiza una copia del documento plantilla del estatal
  var federalFile = DriveApp.getFileById(resources.statalTemplateId).makeCopy(fileName, mainFolder);
  
  // se referencia el documento de Federal
  var document = DocumentApp.openById(federalFile.getId());
  
  // se referencia el cuerpo
  var body = document.getBody();
  
  // obtenemos el childIndex del parrafo donde se debe agregar las imagenes
  var conceptChildIndexArray = getChildIndexByTagImage(body, "<<PRESUPUESTO_DEL_SISTEMA>>");
  
  // validamos si es diferente a -1
  if(conceptChildIndexArray != null){
  
    // recorremos cada una de las imagenes
    for(var i = 0; i < imageList.length; i++){
      
      // generamos el blob a partir del base 64
      imageList[i] = base64toBlob(imageList[i], "conceptImage_" + i);
      
    }
    
    // recorremos la cantida de tags encontrados
    for(var j = 0; j < conceptChildIndexArray.length; j++){
      
      // agregamos las imagensen la ubicación respectiva
      insertImageToChildIndex(body, imageList, conceptChildIndexArray[j]);
      
    }
    
    // remplazamos el tag de "<<PRESUPUESTO_DEL_SISTEMA>>"
    body.replaceText("<<PRESUPUESTO_DEL_SISTEMA>>", "");
  }

  // Insertamos las lineas de producto
  insertProductLineTableData_(body, "<<Líneas de Producto>>", requestData.productLines);
  
  // agregamos los datos del cotizador al documneto
  insertFederalDataInDocument(document, body, requestData, currentStructure, estimateData);
  
  // otenemos los datos de tecnificación
  var technificationData = getDataTechnification(currentStructure.federalRecord, resources);
 
  // se valida si existe datos de tecnificación
  if(technificationData.federalSupport != 0){
    
    // reemplazamos el tag "<<FACTURA_DOS>>" en el body
    body.replaceText("<<FACTURA_DOS>>", calculatePercentage(technificationData.federalSupport, technificationData.totalInversion));
    
    // reemplazamos el tag "<<FACTURA_UNA>>" en el body
    body.replaceText("<<FACTURA_UNA>>", calculatePercentage(technificationData.producerContribution, technificationData.totalInversion));
    
    // reemplazamos el tag "<<TEC_APOYO_FEDERAL>>" en el body
    body.replaceText("<<TEC_APOYO_FEDERAL>>", formatValue(technificationData.federalSupport));
    
    // reemplazamos el tag "<<TEC_INVERSIÓN>>" en el body
    body.replaceText("<<TEC_INVERSIÓN>>", formatValue(technificationData.totalInversion));
    
    // reemplazamos el tag "<<TEC_APORTACIÓN_PRODUCTOR>>" en el body
    body.replaceText("<<TEC_APORTACIÓN_PRODUCTOR>>", formatValue(technificationData.producerContribution));
    
  }
  
  // obtenemos el correo del usuario
  var userMail = getUserMail();
  
  // referenciamos la hoja de cálculo general donde quedan todos los datos de las cotizaciones generadas
  var spreadsheetRequest = SpreadsheetApp.openById(resources.requestSheetId);
  
  // obtenemos la hoja de "Cotizaciones"
  var sheetCot = spreadsheetRequest.getSheetByName("Cotizaciones");
  
  // validamos si existe la hoja
  if(sheetCot){
    
    // guardamos el id del documento de federal
    sheetCot.getRange("O" + requestData.row).setNumberFormat("@STRING@").setValue(federalFile.getId());
    
  }
  
  // se retorna la respuesta
  return {
    message: "Cotización estatal creada exitosamente.",
    requestData: filterRequestByMail(userMail, resources.requestSheetId, userRole, year)
  };  
}

function sssLine(){
  var doc = DocumentApp.openById("1Vn0CZSzsC6TxrWPxXEYO8fLEOlGrdYmd07vrkOKt9RI");
  var body = doc.getBody();
  var recordsArray_ = [
    ['E.1.1. Cabezal de Riego', '8,721.16'],
    ['E.1.6. Linea Principal', '40,502.71'],
    ['E.1.8. Línea Laterales o Regantes', '25,520.94'],
    ['E.1.9. Emisores', '156,173.02'],
    ['E.1.10. Lineas Colectoras y Valvulas de Lavado', '1,164.72']
  ];
  insertProductLineTableData_(body, "<<Líneas de Producto>>", recordsArray_);
}
/**
 * Permite obtener una tabla de acuerdo a un tag
 */
function insertProductLineTableData_(body_, searchText_, recordsArray_) {
  let searchResult_ = body_.findText(searchText_);
  
  if (searchResult_) {
    let element_ = searchResult_.getElement();
    
    // Subir por la jerarquía de elementos hasta encontrar un elemento de tipo TABLE
    while (element_.getType() != DocumentApp.ElementType.TABLE && element_.getParent()) {
      element_ = element_.getParent();
    }
    
    // Verificar si el elemento es de tipo TABLE
    if (element_.getType() == DocumentApp.ElementType.TABLE) {
      let table_ = element_.asTable();
      let rowIndex_;
      
      // Subir por la jerarquía de elementos hasta encontrar un elemento de tipo TABLE_ROW
      element_ = searchResult_.getElement();
      while (element_.getType() != DocumentApp.ElementType.TABLE_ROW && element_.getParent()) {
        element_ = element_.getParent();
      }
      
      // Verificar si el elemento es de tipo TABLE_ROW
      if (element_.getType() == DocumentApp.ElementType.TABLE_ROW) {
        let tableRow_ = element_.asTableRow();
        rowIndex_ = table_.getChildIndex(tableRow_);
        const totalRows_ = table_.getNumRows() - 1;

        // Obtenemos los estilos
        let rowStyles_ = [];
        for (var i = 0; i < tableRow_.getNumCells(); i++) {
          var cell_ = tableRow_.getCell(i);
          // Agregamos el estilo
          rowStyles_.push(cell_.getAttributes());
        }
        
        // Duplicar la fila el número especificado de veces
        duplicateRow_(table_, rowIndex_, recordsArray_, rowStyles_);

        // Eliminamos la fila del tag
        table_.removeRow(rowIndex_);

        // Obtenemos la cantidad de filas vacias
        const startIndex_ = (rowIndex_ + recordsArray_.length);
        const emptyCols_ = totalRows_ - startIndex_;

        // verificamos si la cantidad de filas de la tabla aun existen vacios
        if(emptyCols_ > 0){
          for(let k = 0; k < emptyCols_; k++){
            table_.removeRow(startIndex_);
          }
        }
      }
      
      return table_;
    }
  }
  
  return null; // No se encontró una tabla que contenga el texto
}

/**
 * Permite duplicar una fila
 */
function duplicateRow_(tableRef_, rowIndex, recordsArray_, rowStyles_) {
  let currRow_, currCell_;
  for (let i = 0; i < recordsArray_.length; i++) {
    
    // Ajusta los datos de la nueva fila según recordsArray_
    var tempArray_ = [
      "",
      recordsArray_[i][0],
      "",
      "",
      "",
      recordsArray_[i][1]
    ];

    // Se valida si existe mas de 37 registros
    if(i > 37){
      currRow_ = tableRef_.appendTableRow();
      for (let j = 0; j < tempArray_.length; j++) {
        currCell_ = currRow_.appendTableCell(tempArray_[j]);
        currCell_.setAttributes(rowStyles_[j]);
      }
    } else {
      currRow_ = tableRef_.getRow(rowIndex + i + 1);
      for (let j = 0; j < tempArray_.length; j++) {
        currCell_ = currRow_.getCell(j);
        currCell_.setText(tempArray_[j]);
        currCell_.setAttributes(rowStyles_[j]);
      }
    }
  }
}

/**
* Función que permite validar si existe el tag y agregar las respectivas imagens
* @author Rubén Sánchez
*/
function initCreationFederalFile(imageList, requestData, fileUrl, estimateData, userRole, year) {

  // referenciamos el archivo de la cotización
  var projectSheet = DriveApp.getFileById(requestData.documentId);
  
  // se obtiene los recursos de hojas de cálculo
  var resources = getResources();
  
  // obtenemos las ubicaciones
  var parents = projectSheet.getParents();
  
  // variable para referenciar el padre
  var mainFolder = null;
  
  // se valida si existe algun padre
  if(parents.hasNext()){
    
    // referenciamos la carpeta donde esta ubicado el archivo
    mainFolder = parents.next();
    
  } else {
  
    // se referencia la carpeta de mi unidad
    mainFolder = DriveApp.getRootFolder();  
  }
  
  // se obtiene la fecha actual
  var currentDate = new Date();
  
  // fecha en formato DD de MES de YYYY
  var stringDate = formatStringDate(currentDate);
  
  // referenciamos la estructura actual
  var currentStructure = JSON.parse(requestData.structures[requestData.mainSheetName].structure);
  
  // se define el nombre del archivo
  var fileName = "Cotización: " + currentStructure.client + " - " + stringDate;
  
  // se realiza una copia del documento plantilla del federal
  var federalFile = DriveApp.getFileById(resources.federalTemplate).makeCopy(fileName, mainFolder);
  
  // se referencia el documento de Federal
  var document = DocumentApp.openById(federalFile.getId());
  
  // se referencia el cuerpo
  var body = document.getBody();
  
  // obtenemos el childIndex del parrafo donde se debe agregar las imagenes
  var conceptChildIndexArray = getChildIndexByTagImage(body, "<<PRESUPUESTO_DEL_SISTEMA>>");
  
  // validamos si es diferente a -1
  if(conceptChildIndexArray != null){
  
    // recorremos cada una de las imagenes
    for(var i = 0; i < imageList.length; i++){
      
      // generamos el blob a partir del base 64
      imageList[i] = base64toBlob(imageList[i], "conceptImage_" + i);
      
    }
    
    // recorremos la cantida de tags encontrados
    for(var j = 0; j < conceptChildIndexArray.length; j++){
      
      // agregamos las imagensen la ubicación respectiva
      insertImageToChildIndex(body, imageList, conceptChildIndexArray[j]);
      
    }
    
    // remplazamos el tag de "<<PRESUPUESTO_DEL_SISTEMA>>"
    body.replaceText("<<PRESUPUESTO_DEL_SISTEMA>>", "");
  }
  
  // agregamos los datos del cotizador al documneto
  insertFederalDataInDocument(document, body, requestData, currentStructure, estimateData);
  
  // insertamos los datos que vienen de los datos técnicos 
  insertTechnicalDataInFederal(fileUrl, body);
  
  // otenemos los datos de tecnificación
  var technificationData = getDataTechnification(currentStructure.federalRecord, resources);
 
  // se valida si existe datos de tecnificación
  if(technificationData.federalSupport != 0){
    
    // reemplazamos el tag "<<FACTURA_DOS>>" en el body
    body.replaceText("<<FACTURA_DOS>>", calculatePercentage(technificationData.federalSupport, technificationData.totalInversion));
    
    // reemplazamos el tag "<<FACTURA_UNA>>" en el body
    body.replaceText("<<FACTURA_UNA>>", calculatePercentage(technificationData.producerContribution, technificationData.totalInversion));
    
    // reemplazamos el tag "<<TEC_APOYO_FEDERAL>>" en el body
    body.replaceText("<<TEC_APOYO_FEDERAL>>", formatValue(technificationData.federalSupport));
    
    // reemplazamos el tag "<<TEC_INVERSIÓN>>" en el body
    body.replaceText("<<TEC_INVERSIÓN>>", formatValue(technificationData.totalInversion));
    
    // reemplazamos el tag "<<TEC_APORTACIÓN_PRODUCTOR>>" en el body
    body.replaceText("<<TEC_APORTACIÓN_PRODUCTOR>>", formatValue(technificationData.producerContribution));
    
  }
  
  // obtenemos el correo del usuario
  var userMail = getUserMail();
  
  // referenciamos la hoja de cálculo general donde quedan todos los datos de las cotizaciones generadas
  var spreadsheetRequest = SpreadsheetApp.openById(resources.requestSheetId);
  
  // obtenemos la hoja de "Cotizaciones"
  var sheetCot = spreadsheetRequest.getSheetByName("Cotizaciones");
  
  // validamos si existe la hoja
  if(sheetCot){
    
    // guardamos el id del documento de federal
    sheetCot.getRange("O" + requestData.row).setNumberFormat("@STRING@").setValue(federalFile.getId());    
  }
  
  // se retorna la respuesta
  return {
    message: "Cotización de federal creada exitosamente",
    requestData: filterRequestByMail(userMail, resources.requestSheetId, userRole, year)
  };  
}

/**
 * Función que permite insertar los datos tecnicos en la cotización federal
**/
function insertTechnicalDataInFederal(fileUrl, body){

  // Obtenemos el id de la url
  var urlInfo = fileUrl.match(/\/d\/(.*?)\//);
  
  // se valida si existe el id
  if(urlInfo && urlInfo.length > 1){
    
    // Obtenemos el id de la url
    var fileId = urlInfo[1];
    
    // variable para guardar la referencia del libro
    var spreadsheet = null;
    
    // encerramos en un try-catch por si la url proporcionada no pertenece a una hoja de cálculo
    try{
    
      // Referenciamos la hoja de datos tecnicos
      spreadsheet = SpreadsheetApp.openById(fileId);
      
    } catch(e){
      
      // dejamos en null la variable del libro
      spreadsheet = null;
    }
    
    // se valida si existe el libro
    if(spreadsheet != null){
      
      // se define un objeto con la relación de la hojas, tags y rangos respectivos
      var infoSheets = {
        "1.General": [
          {tag: "<FUENTE_DE_ABASTECIMIENTO>>", range: "C42"},
          {tag: "<<MARCO>>", range: "C66"},
          {tag: "<<PLANTAS>>", range: "C64"},
          {tag: "<<HILERAS>>", range: "C65"},
          {tag: "<<EVAPOTRANSPIRACIÓN>>", range: "C70"},
          {tag: "<<LÁMINA_HORARIA>>", range: "C77"},
          {tag: "<<TIEMPO>>", range: "C78"},
          {tag: "<<INTERVALO>>", range: "C79"},
          {tag: "<<PRESIÓN>>", range: "C82"},
          {tag: "<<CAUDAL>>", range: "C83"},
          {tag: "<<DESNIVEL_TOPOGRÁFICO>>", range: "C37"},
          {tag: "<<LATITUD_FUENTE>>", range: "C89"},
          {tag: "<<LONGITUD_FUENTE>>", range: "D89"},
          {tag: "<<LATITUD_PIVOTE>>", range: "C90"},
          {tag: "<<LONGITUD_PIVOTE>>", range: "D90"},
          {tag: "<<LATITUD_BOMBEO>>", range: "C91"},
          {tag: " <<LONGITUD_BOMBEO>>", range: "D91"}
        ],
        "2.Instalación": [
          {tag: "<<SEPARACIÓN_ENTRE_EMISORES>>", range: "D16"},
          {tag: "<<SEPARACIÓN_ENTRE_LATERALES>>", range: "D17"},
          {tag: "<<CAUDAL_DEL_EMISOR>>", range: "D19"},
          {tag: "<<PRESIÓN_DE_OPERACIÓN>>", range: "D18"},
          {tag: "<<PÉRDIDA_DE_CARGA_LATERAL>>", range: "D36"}
        ],
        "D.Localizado": [
          {tag: "<<NÚMERO_DE_SECCIONES>>", range: "C14"},
          {tag: "<<CAUDAL_POR_SECCIÓN>>", range: "C16"},
          {tag: "<<ÁREA_PROMEDIO_X_SECCIÓN>>", range: "C17"}
        ]
      };
      
      // variable para referenciar la hojar y tener el listado de tangos
      var sheet = null,
          ranges = []
      
      // recorremos cada una de las hojas 
      for(var sheetName in infoSheets){
        
        // se referencia la hoja 
        sheet = spreadsheet.getSheetByName(sheetName);
        
        // se valida que exista la hoja
        if(sheet){
          
          // se obtiene el lisatdo de rangos
          ranges = infoSheets[sheetName];
          
          // se recorre los registros
          for(var i = 0; i < ranges.length; i++){
            
            // reemplazamos el tag respectivo en el body por el valor parametrizado
            body.replaceText(ranges[i].tag, sheet.getRange(ranges[i].range).getValues());
            
          }
        }
      }
    }
  }
  
}

/**
 * Fúnción que permite calcular el porcentaje de 2 valores, siendo el segundo el parametro de refencia o mayor
**/
function calculatePercentage(minValue, totalValue){

  // se valida si los 2 parametros son números
  if(isNaN(minValue) || isNaN(totalValue) || minValue == 0 || totalValue == 0) return 0;
  
  // retornamos el valor 
  return roundDecimal((minValue * 100) /  totalValue, 2);
  
}

/**
 * Función que permite agregar los datos de federal y del cotizador en la cotiazción de federal
**/
function insertFederalDataInDocument(document, body, requestData, currentStructure, estimateData){
  
  // fecha de creación en array
  var dateArray = String(requestData.date).split("/");
  
  // obtenemos la fecha de creación
  var creationDate = new Date(dateArray[1] + "/" + dateArray[0] + "/" + dateArray[2]);
  
  // convertimos la fecha en formato DD de MES de YYYY
  var stringCreationDate = formatStringDate(creationDate);
  
  // se referencia el pie de página
  var footer = document.getFooter();
  
  // reemplazamos el tag "<<CLIENTE>>"
  footer.replaceText("<<CLIENTE>>", currentStructure.client);
  
  // se referencia el encabezado
  var head = document.getHeader();
  
  // reemplazamos el tag "<<STR_FECHA>>" de la primera página
  head.getParent().replaceText("<<STR_FECHA>>", stringCreationDate);

  // actualizamos los datos de ubicación
  head.replaceText("<<LATITUD>>", currentStructure.latitude);
  head.replaceText("<<LONGITUD>>", currentStructure.longitude);
  body.replaceText("<<LATITUD>>", currentStructure.latitude);
  body.replaceText("<<LONGITUD>>", currentStructure.longitude);
  footer.replaceText("<<LATITUD>>", currentStructure.latitude);
  footer.replaceText("<<LONGITUD>>", currentStructure.longitude);
  
  // reemplazamos el tag "<<STR_FECHA>>"
  head.replaceText("<<STR_FECHA>>", stringCreationDate);
  
  // reemplazamos el tag "<<CLIENTE>>"
  head.replaceText("<<CLIENTE>>", currentStructure.client);
  
  // reemplazamos el tag "<<CLIENTE>>" en el body
  body.replaceText("<<CLIENTE>>", currentStructure.client);
  
  // reemplazamos el tag "<<LOCALIDAD>>" en el body
  body.replaceText("<<LOCALIDAD>>", currentStructure.location);
  
  // reemplazamos el tag "<<MUNICIPIO>>" en el body
  body.replaceText("<<MUNICIPIO>>", currentStructure.municipality);
  
  // reemplazamos el tag "<<ESTADO>>" en el body
  body.replaceText("<<ESTADO>>", currentStructure.state);
  
  // reemplazamos el tag "<<CLASE>>" en el body
  body.replaceText("<<CLASE>>", currentStructure.systemClasification);
  
  // reemplazamos el tag "<<TIPO>>" en el body
  body.replaceText("<<TIPO>>", currentStructure.systemType);
  
  // reemplazamos el tag "<<FORMA>>" en el body
  body.replaceText("<<FORMA>>", currentStructure.formEmission);
  
  // reemplazamos el tag "<<FECHA>>" en el body
  body.replaceText("<<FECHA>>", stringCreationDate);
  
  // reemplazamos el tag "<<CULTIVO>>" en el body
  body.replaceText("<<CULTIVO>>", currentStructure.cultive);
  
  // reemplazamos el tag "<<SUPERFICIE>>" en el body
  body.replaceText("<<SUPERFICIE>>", currentStructure.surface);
  
  // reemplazamos el tag "<<SUBTOTAL_RIEGO>>" en el body
  body.replaceText("<<SUBTOTAL_RIEGO>>", formatValue(estimateData.prices["1. Sistema de riego"] || 0));
  
  // reemplazamos el tag "<<SUBTOTAL_OBRA>>" en el body
  body.replaceText("<<SUBTOTAL_OBRA>>", formatValue(estimateData.prices["2. Obra civil"] || 0));
  
  // reemplazamos el tag "<<SUBTOTAL_EQUIPO>>" en el body
  body.replaceText("<<SUBTOTAL_EQUIPO>>", formatValue(estimateData.prices["3. Equipo mecanico y electrico"] || 0));
  
  // reemplazamos el tag "<<SUBTOTAL_SERVICIO>>" en el body
  body.replaceText("<<SUBTOTAL_SERVICIO>>", formatValue(estimateData.prices["4. Servicios complementarios"] || 0));
  
  // reemplazamos el tag "<<OFERTA_FEDERAL>>" en el body
  body.replaceText("<<OFERTA_FEDERAL>>", estimateData.importe || 0);
  body.replaceText("<<OFERTA_ESTATAL>>", estimateData.importe || 0);
  
  // reemplazamos el tag "<<OFERTA_FEDERAL_LETRAS>>" en el body
  body.replaceText("<<OFERTA_FEDERAL_LETRAS>>", estimateData.importeLetter);
  body.replaceText("<<OFERTA_ESTATAL_LETRAS>>", estimateData.importeLetter);
  
}

/**
 * Función encargada de obtener la información del aplicativo de tecnificación de riego
 * @author Rubén Sánchez
 */
function getDataTechnification(rfc, resources){
  
  // referenciamos la hoja de usuarios
  var technificationSpreadsheet = SpreadsheetApp.openById(resources.technificationSheetId);
  
  // referenciamos la hoja de usuarios y roles
  var technificationSheet = technificationSpreadsheet.getSheetByName("Solicitudes");
  
  // definimos un objeto
  var resultDataTechnification = {
    federalSupport: 0,
    producerContribution: 0,
    totalInversion: 0
  };
  
  // Si existe la hoja
  if(technificationSheet){
    
    // obtenemos la ultima fila
    var lastRow = technificationSheet.getLastRow();
    
    // Si existe más de una fila
    if(lastRow > 1){
      
      // obtenemos los datos
      var arrayData = technificationSheet.getRange(2, 1, lastRow - 1, 65).getValues();
      
      // recorremos el listado de correo parametrizados
      for(var i = 0; i < arrayData.length; i++){
        
        // Si el RFC es igual al que se esta recorriendo
        if(rfc == arrayData[i][2]){
          
          // agregamos los datos de acuerdo al rfc
          resultDataTechnification.federalSupport = Number(arrayData[i][60]);
          resultDataTechnification.producerContribution = Number(arrayData[i][63]);
          resultDataTechnification.totalInversion = Number(arrayData[i][61]);
          
          // retornamos el objeto
          return resultDataTechnification;
          
        }
      }
    }
  }
  
  // retornamos el objeto vacio
  return resultDataTechnification;
}

/**
* Función encargada de dar formato de miles al valor ingresado en un campo de texto
* @param {event handler | Jquery object}
*/
function formatValue(value) {
  
  // eliminamos las copas existentes
  var stringValue = ("" + value).replace(/,/g,"");
  
  // generamos un array apartir de un etxto sepearandolo por un ponto(.)
  var arrayValue = stringValue.split(".");
  
  // se valida si solo existe un valor en el array
  if(arrayValue.length == 1){
    
    // retornamos el valor
    return arrayValue[0].replace(/,/g,"").replace(/\B(?=(\d{3})+(?!\d))/g,",");
    
  } else if(arrayValue.length == 2){
    
    // retornamos el valor y los decimales
    return arrayValue[0].replace(/,/g,"").replace(/\B(?=(\d{3})+(?!\d))/g,",") + "." + arrayValue[1];
    
  } else {
    
    // retornamos el valor por defecto
    return stringValue;
    
  }
}

/**
 * Función que permite convertir un new Date() a un formato "DD de MES de YYYY"
**/
function formatStringDate(date){
  
  // se define los meses
  var monthNames = [
    "Enero", "Febrero", "Marzo",
    "Abril", "Mayo", "Junio", "Julio",
    "Agosto", "Septiembre", "Octubre",
    "Noviembre", "Diciembre"
  ];
  
  // obtenemos datos de la fecha
  var day = date.getDate();
  var monthIndex = date.getMonth();
  var year = date.getFullYear();
  
  // retornamos la fecha
  return day + " de " + monthNames[monthIndex] + " del " + year;
  
}

/**
 * Función que permite agregar las imagenes a partir de la ubicación proporcionada
 * @author Rubén Sánchez
**/
function insertImageToChildIndex(body, imageList, childIndex) {
  
  // recorremos cada una de las imagenes
  for(var i = 0; i < imageList.length; i++){
    
    // se valida que exista una imagen
    if(imageList[i]){
    
      // insertamos la imagen 
      body.insertImage(childIndex, imageList[i]).setWidth(640);
      
      // aumentamos el childIndex
      childIndex++;
    }
  }
  
}

/**
 * Función que permite convertir de base64 a Blob
 * @author Rubén Sánchez
**/
function base64toBlob(base64Data, blobName) {
  
  // convertirmos el texto en un array de 2 posiciones de acuerdo al caracter(,)
  var splitBase = base64Data.split(",");
  
  // obtenemos el tipo de archivo eliminando el texto "data:" de la primera posicion del array anterior, eje: data:image/png;base64
  var mimeType = splitBase[0].split(";")[0].replace("data:","");
  
  // Utilizando la clase Utilities decodificamos el texto en Base64 a bytes
  var byteCharacters = Utilities.base64Decode(splitBase[1]);
  
  // se retorna el blob
  return Utilities.newBlob(byteCharacters, mimeType, blobName);
}

/**
 * Función que permite obtener la ubicación del tag
 * @author Rubén Sánchez
**/
function getChildIndexByTagImage(body, tag){
  
  // se realiza la búsqueda 
  var foundTag = body.findText(tag);
  
  // se valida si el elemento es diferente de null
  if(foundTag != null){
    
    // definimos el array organizados de abajo hacia arriba
    var arrayChildIndex = [];
    
    // se recorre tantas veces como sea posible
    while (foundTag != null) {
    
      // se obtiene el eleemnto donde esta el texto
      var tagElement = foundTag.getElement();
      
      // se encuentra el padre 
      var parent = tagElement.getParent();
      
      // agregamos la posición del elemento
      arrayChildIndex.unshift(parent.getParent().getChildIndex(parent));
      
      // Find the next match
      foundTag = body.findText(tag, foundTag);
    }
    
    // retornamos el array de ubicaciones del texto
    return arrayChildIndex;
    
  }
  
  //retornamos un valor null por defecto
  return null;
}
