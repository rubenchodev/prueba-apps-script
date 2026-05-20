/**
 * Función que permite crear una cotización a partir de otra
**/
function createDuplicateQuotation(requestData, folderId, userRole, year) {

  // obtenemos los datos de la estructura de "Proyectos"
  var structure = requestData.structures.Proyecto.structure

  // obtenemos los datos de la estructura de la cotización selecciondad
  var requestDataSelected = JSON.parse(structure);

  // cambiamos la ruta de la carpeta
  requestDataSelected.folderId = folderId;
  
  // creamos la nueva cotización y recibimos la respuesta
  var response = createQuotation(requestDataSelected, null, {sheetName: "Proyecto"}, null, -1, null, userRole, year);
  
  // retornamos el mensaje y la lista actualizada de las solicitudes
  return response;
  
}

/**
 * Función que permite crear una cotización de refacción a partir de otra
**/
function createDuplicateRenovation(requestData, folderId, templateId, userRole) {

  // obtenemos los datos de la estructura de "Proyectos"
  var structure = requestData.structure;

  // obtenemos los datos de la estructura de la cotización selecciondad
  var requestDataSelected = JSON.parse(structure);

  // cambiamos la ruta de la carpeta
  requestDataSelected.folderId = folderId;
  
  // Agregamos el id de la plantilla
  requestDataSelected.templateId = templateId;

  // creamos la nueva cotización y recibimos la respuesta
  var response = generateCotRenovation(requestDataSelected, null, -1, null, null, userRole);
  
  // retornamos el mensaje y la lista actualizada de las solicitudes
  return response;
  
}