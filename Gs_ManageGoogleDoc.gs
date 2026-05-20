/**
 * Función que permite agregar los datos en Google Doc
**/
function insertDataByGoogleDoc(document, conceptInfoObj, formData, consecutive, userData) {
  
  //document = DocumentApp.getActiveDocument();
  
  // referenciamos el cuerpo del archivo
  var body = document.getBody();
  
  // obtenemos el encabezado del documento
  var header = document.getHeader();
  
  // obtenemos el childIndex del parrafo donde el tag de la tabla de especies
  var conceptChildIndex = getChildIndexByTag(body, "<<TABLE_CONCEPTS>>");
  
  // validamos si es diferente a -1
  if(conceptChildIndex != -1){
  
    // agregamos la tabla de refacciones
    addTableInDocument(body, conceptChildIndex, conceptInfoObj);
  }
  
  // remplazamos el tag de "<<TABLE_CONCEPTS>>"
  body.replaceText("<<TABLE_CONCEPTS>>", "");
  
  // obtenemos el childIndex del parrafo donde el tag de la tabla de especies
  var deliveryChildIndex = getChildIndexByTag(body, "<<TABLE_INVENTARY>>");
  
  // validamos si es diferente a -1
  if(deliveryChildIndex != -1){
    
    // validamos si existen datos para agregar a la tabla de productos pendientes
    if(conceptInfoObj.deliveryTimeArray.length > 0){
    
      // agregamos la tabla de refacciones por entregar
      addTableInventary(body, deliveryChildIndex, conceptInfoObj.deliveryTimeArray);
      
    } else {
      
      // obtenemos la información de donde esta ubicado el siguiente texto
      var paragraphTimeIndex = getChildIndexByTag(body, "Los productos fuera de inventario, estarán disponibles en nuestra bodega de la siguiente manera.");
      
      // se valida si el index existe
      if(paragraphTimeIndex != -1){
      
        // eliminamos el texto "Los productos fuera de inventario, estarán disponibles en nuestra bodega de la siguiente manera."
        body.removeChild(body.getChild(paragraphTimeIndex));
      }
      
    }
  }
  
  // remplazamos el tag de "<<TABLE_INVENTARY>>"
  body.replaceText("<<TABLE_INVENTARY>>", "");
  
  // remplazamos el tag de "<<IMPORTE>>" por su respectivo valor
  body.replaceText("<<IMPORTE>>", conceptInfoObj.importe.number);
  
  // remplazamos el tag de "<<IMPORTE_LETTER>>" por su respectivo valor
  body.replaceText("<<IMPORTE_LETTER>>", conceptInfoObj.importe.letter);
  
  // remplazamos el tag de "<<COUNTER>>" por su respectivo valor
  body.replaceText("<<COUNTER>>", conceptInfoObj.counter);
  
  // array con los tags de titulo en la tabla de cambio
  var tagsTitleFromTable = ["<<TITLE_EU>>", "<<TITLE_USD>>", "<<TITLE_MN>>"];
  
  // variable para determinar en donde debe ir el titulo por defecto es el 2
  var titleIndex = 2;
  
  // se valida si existe especies con el tipo de moneda de euros
  if(conceptInfoObj.totalArray[0][1] != "0"){
  
    // remplazamos los tag de datos de EUROS
    body.replaceText("<<T_EU>>", conceptInfoObj.totalArray[0][0]);
    body.replaceText("<<S_EU>>", conceptInfoObj.totalArray[0][1]);
    body.replaceText("<<T_EUM>>", conceptInfoObj.totalArray[0][2]);
    
    // definimos el index de la posición 0 del array
    titleIndex = 0;
    
  } else {
    
    //removemos fila donde se muestra lo de euros
    deleteRowByTag("<<T_EU>>", body);
    
  }
  
  // se valida si existe especies con el tipo de moneda de euros
  if(conceptInfoObj.totalArray[1][1] != "0"){
  
    // remplazamos los tag de datos de DOLARES
    body.replaceText("<<T_USD>>", conceptInfoObj.totalArray[1][0]);
    body.replaceText("<<S_USD>>", conceptInfoObj.totalArray[1][1]);
    body.replaceText("<<T_USDM>>", conceptInfoObj.totalArray[1][2]);
    
    // validamos si el index del titulo es diferente de 0
    if(titleIndex != 0){
      
      // definimos el index de la posición 1 del array
      titleIndex = 1;
      
    }
    
  } else {
    
    //removemos fila donde se muestra lo de dolares
    deleteRowByTag("<<T_USD>>", body);
    
  }
  
  // se recorre los tags de titulos en la tabla de tipo de cambio
  for(var k = 0; k < tagsTitleFromTable.length; k++){
    
    // se valida si el consecutivo es igual al index donde debe ir el titulo
    if(k == titleIndex){
      
      // agregamos el titulo 
      body.replaceText(tagsTitleFromTable[k], "SUBTOTAL DE LA OFERTA");
      
    } else {
      
      // reemplazamos por un valor vacio
      body.replaceText(tagsTitleFromTable[k], "");
    }
    
  }
  
  // remplazamos los tag de datos de PESOS MEXICANOS
  body.replaceText("<<S_MN>>", conceptInfoObj.totalArray[2][1]);
  body.replaceText("<<T_MN>>", conceptInfoObj.totalArray[2][2]);
  
  // reemplazamos el campo de cliente
  body.replaceText("<<CLIENT>>", formData.client);

  // Obtenemos la fecha emisión
  let issueDate_ = formatIssueDate_(formData.issue);

  // reemplazamos los datos referentes al cliente en el CUERPO del documento por si los encuentra
  body.replaceText("<<ISSUE>>", issueDate_);
  body.replaceText("<<LOCATION>>", formData.state + ", " + formData.municipality + ", " + formData.location);
  body.replaceText("<<CONSECUTIVE>>", consecutive);
  body.replaceText("<<VALIDATE_DATE>>", formData.validity);
  
  // reemplazamos los datos referentes al cliente en el ENCABEZADO del documento por si los encuentra
  header.replaceText("<<ISSUE>>", issueDate_);
  header.replaceText("<<LOCATION>>", formData.state + ", " + formData.municipality + ", " + formData.location);
  header.replaceText("<<CONSECUTIVE>>", consecutive);
  header.replaceText("<<VALIDATE_DATE>>", formData.validity);
  
  // reemplazamos el nombre del creador
  body.replaceText("<<USER_CREATOR>>", userData.fullName);
  
  // obtenemos el importe en valor numerico y sin comas(,)
  var onlyImporte = Number(String(conceptInfoObj.importe.number).replace(/\,/g, ""));
  
  // calculamos el 70% del importe y el 30 %
  var inporte70 = roundDecimal((onlyImporte * 70) / 100, 2);
  var importe30 = roundDecimal(onlyImporte - inporte70, 2);
  
  // remplazamos el tag de "<<ADVANCE>>" por su respectivo valor
  body.replaceText("<<ADVANCE>>", formatValue(inporte70));
  
  // remplazamos el tag de "<<REMAINING>>" por su respectivo valor
  body.replaceText("<<REMAINING>>", formatValue(importe30));
  
  // remplazamos el tag de "<<ADVANCE_LETTER>>" por su respectivo valor
  body.replaceText("<<ADVANCE_LETTER>>", CifrasEnLetras.convertirNumeroEnLetras(inporte70));
  
  // remplazamos el tag de "<<REMAINING_LETTER>>" por su respectivo valor
  body.replaceText("<<REMAINING_LETTER>>", CifrasEnLetras.convertirNumeroEnLetras(importe30));
  
  // obtenemos la posición donde esta el tag de Precios
  var pricesChildIndex = getChildIndexByTag(body, "<<PRICES>>");
  
  // validamos si es diferente a -1
  if(pricesChildIndex != -1){
    
    //Definimos los estilos de una lista
    var listStyle = {};
    listStyle[DocumentApp.Attribute.BOLD] = false;
    listStyle[DocumentApp.Attribute.ITALIC] = false;
    listStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = "#000000";
    listStyle[DocumentApp.Attribute.GLYPH_TYPE] = DocumentApp.GlyphType.BULLET;
    
    // recorremos cada una de las opciones de precios seleccionadas
    for(var i = 0; i < formData.priceCategory.length; i++){
      
      // agregamos un item de lista
      body.insertListItem(pricesChildIndex + i, formData.priceCategory[i]).setAttributes(listStyle);
      
    }
    
  }
  
  // remplazamos el tag de "<<PRICES>>"
  body.replaceText("<<PRICES>>", "");
  
}

/*
* Función encargada de generar el formato de la fecha para el certificado laboral
*/
function formatIssueDate_(dateString_){

  // Obtenemos la fecha emisión
  let issueDate_ = (dateString_ || "").toString().trim();
  if(issueDate_){
  
    // Obtenemos la fecha en array
    let issueDateArray_ = issueDate_.split("/");

    // Se validar si existen 3 datos
    if(issueDateArray_ && issueDateArray_.length == 3){

      // Obtenemos el mes y restamos uno
      let month_ = parseInt(issueDateArray_[1]) - 1;
    
      // Array con los meses en español
      var monthList_ = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
      
      // Retornamos la fecha en el formato dd/MES/Año, eje: 20/Abril/2026
      return issueDateArray_[0] + "/" + monthList_[month_] + "/" + issueDateArray_[2];
    }
  }

  // Retornamos por defecto lo que nos envio
  return dateString_;
}

/**
 * Función que permite eliminar una fila de acuerdo a un tag enviado por parametro
**/
function deleteRowByTag(tagName, body){
  
  // se realiza la búsqueda 
  var foundTag = body.findText(tagName);
  
  // validamos si encontro algun elemento con ese texto
  if (foundTag != null) {
    
    // variable para referenciar la fila
    var tableRow = null,
        tagElement = foundTag.getElement();
    
    // referenciamos el elemento a validar
    var elemReference = DocumentApp.ElementType.TABLE_ROW;
    
    // se valida si el elemento es un Table_Row
    if(tagElement.getType() != elemReference){
      
      // realizamos un ciclo hasta encontralo
      while(tagElement.getParent() && !tableRow){
       
        // obtenemos el padre del elemento
        tagElement = tagElement.getParent();
        
        // se valida si el elemento es un Table_Row
        if(tagElement.getType() == elemReference){
          
          // cambiamos el estado
          tableRow = tagElement;
          
        }
      }
    }
    
    // validamos si existe un elemento
    if(tableRow){
      
      // eliminamos la respectiva fila
      tableRow.removeFromParent();
    }
  
  }
  
}

/** 
  Función que aproxima un valor decimal a cierto número de decimales. 
  @param {Number} number Número que se desea truncar. 
  @param {Number} [decimals=0] Número de decimales a truncar. 
  @return {Number} Retorna el número truncado. 
*/
function roundDecimal(number, decimals) {
  
  // convertimos en un numero el valor si el valor es diferente a un tipo numero
  if(typeof number !== "number"){
  
    // convertimos el valor en numero
    number = Number(String(number).replace(/\$|\,/g, "")) || 0;
  } else if(number === Infinity){
  
    // se supera la cantidad de decimales se establece cero
    number = 0;
  }
  
  if (decimals == undefined) {
    decimals = 0;
  }
  if (decimals >= 0) {
    /*var multipleValue = Math.pow(10, decimals);
    return Math.floor(number * multipleValue) / multipleValue;*/
    var newDecmal = +(Math.round(number + "e+" + decimals)  + "e-" + decimals);
      
    // validamos si no es número
    if(isNaN(newDecmal)){
      
      // retornamos el valor de cero
      return 0;
      
    } else {
      
      // retornamos el valor redondeado
      return newDecmal;
      
    }
  }
}

/**
 * Función que permiteagregar la tabla de las refacciones
**/
function addTableInDocument(body, childIndex, conceptInfoObj) {
  
  //Definimos los estilos del encabezado
  var headerStyle = {};
  headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = "#e0f7fa";
  headerStyle[DocumentApp.Attribute.BOLD] = true;
  headerStyle[DocumentApp.Attribute.ITALIC] = true;
  headerStyle[DocumentApp.Attribute.FONT_FAMILY] = "Arial Narrow";
  headerStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
  headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = "#000000";
  
  //Style for the cells other than header row
  var cellStyle = {};
  cellStyle[DocumentApp.Attribute.BOLD] = false;
  cellStyle[DocumentApp.Attribute.FONT_FAMILY] = "Arial Narrow";
  cellStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
  cellStyle[DocumentApp.Attribute.ITALIC] = false;
  cellStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  cellStyle[DocumentApp.Attribute.WIDTH] = 46;
  
  // establecemos un ancho de cada una de las celdas
  var widthArray = [20,46,46,0,0,197,46,46,46,46];
  
  //agregamos la tabla en la ubicación 
  var table = body.insertTable(childIndex);
  
  // creamos la fila del encabezado
  var trHeader = table.appendTableRow();
  
  // creamos cada una de los encabezados
  trHeader.appendTableCell("ID").setAttributes(headerStyle);
  trHeader.appendTableCell("MONEDA").setAttributes(headerStyle);
  trHeader.appendTableCell("INVENTARIO").setAttributes(headerStyle);
  trHeader.appendTableCell("DESCRIPCIÓN").setAttributes(headerStyle);
  trHeader.appendTableCell("UNIDAD").setAttributes(headerStyle);
  trHeader.appendTableCell("CANTIDAD").setAttributes(headerStyle);
  trHeader.appendTableCell("C. UNITARIO").setAttributes(headerStyle);
  trHeader.appendTableCell("IMPORTE").setAttributes(headerStyle);
  
  //recorremos cada uno de los datos
  for(var i = 0; i < conceptInfoObj.conceptsArray.length; i++){
    
    // creamos la respectiva tabla
    var tr = table.appendTableRow();
    
    //recorremos las columnas respectiva de la refacción
    for(var j = 0; j < conceptInfoObj.conceptsArray[i].length; j++){
      
      // validamos si la columna recorrida es diferente a la cuarta o quinta
      if(j != 3 && j != 4){
        
        // actualizamos el tamaño de la columna
        cellStyle[DocumentApp.Attribute.WIDTH] = widthArray[j];
        
        // creamos la celda y asignamos el respectivo estilo
        tr.appendTableCell(conceptInfoObj.conceptsArray[i][j]).setAttributes(cellStyle);
      }
    }
  }
  
  // se define los estilos generales de la tabla
  var tableStyle = {};
  tableStyle[DocumentApp.Attribute.BORDER_WIDTH] = 1; 
  tableStyle[DocumentApp.Attribute.BORDER_COLOR] = '#4a86e8';
  tableStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER; 
  
  // establecemos los estilos a la tabla
  table.setAttributes(tableStyle);
}

/**
 * Función que permite obtener la ubicación del tag
**/
function getChildIndexByTag(body, tag){
  
  // se realiza la búsqueda 
  var foundTag = body.findText(tag);
  
  // validamos si encontro algun elemento con ese texto
  if (foundTag != null) {
    
    // se obtiene el eleemnto donde esta el texto
    var tagElement = foundTag.getElement();
    
    // se encuentra el padre 
    var parent = tagElement.getParent();
    
    // se retorna la posición donde se va a generar un nuevo elemento
    return parent.getParent().getChildIndex(parent);
  
  }
  
  //retornamos un valor -1 por defecto
  return -1;
}

/**
 * Función que permiteagregar la tabla de las refacciones
**/
function addTableInventary(body, childIndex, arrayData) {
  
  //Definimos los estilos del encabezado
  var headerStyle = {};
  headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = "#4bacc6";
  headerStyle[DocumentApp.Attribute.BOLD] = false;
  headerStyle[DocumentApp.Attribute.ITALIC] = false;
  headerStyle[DocumentApp.Attribute.FONT_FAMILY] = "Arial";
  headerStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
  headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = "#FFFFFF";
  
  //Style for the cells other than header row
  var cellStyle = {};
  cellStyle[DocumentApp.Attribute.BOLD] = false;
  cellStyle[DocumentApp.Attribute.FONT_FAMILY] = "Arial";
  cellStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
  cellStyle[DocumentApp.Attribute.ITALIC] = false;
  cellStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  cellStyle[DocumentApp.Attribute.WIDTH] = 110;
  
  //agregamos la tabla en la ubicación 
  var table = body.insertTable(childIndex);
  
  // creamos la fila del encabezado
  var trHeader = table.appendTableRow();
  
  // creamos cada una de los encabezados
  trHeader.appendTableCell("Id producto").setAttributes(headerStyle);
  trHeader.appendTableCell("Tiempo de Entrega").setAttributes(headerStyle);
  
  //recorremos cada uno de los datos
  for(var i = 0; i < arrayData.length; i++){
    
    // creamos la respectiva tabla
    var tr = table.appendTableRow();
    
    // creamos la celda de Id del producto
    tr.appendTableCell(arrayData[i][0]).setAttributes(cellStyle);
    
    // creamos la celda de tiempo de entrega
    tr.appendTableCell(arrayData[i][1]).setAttributes(cellStyle);
  }
  
  // se define los estilos generales de la tabla
  var tableStyle = {};
  tableStyle[DocumentApp.Attribute.BORDER_WIDTH] = 1; 
  tableStyle[DocumentApp.Attribute.BORDER_COLOR] = "#78c0d4";
  tableStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER; 
  
  // establecemos los estilos a la tabla
  table.setAttributes(tableStyle);
}
