/**
 * Permite generar la lista de materiales a partir de los cruceros
 */
function csGenerateMaterials(projectList_){
  projectListOLD_ = [
    {
      "valveText": "3/4\"",
      "outletValue": "25",
      "valveType": "Angular",
      "valveDiameter": "0.75",
      "outletDiameter": "25 mm",
      "vaeConfig": "S/VAE",
      "crossingHeight": 80,
      "inletValue": "32",
      "inletDiameter": "32 mm",
      "trenchDepth": 90,
      "crossingType": "Sencillo",
      "count": 4
    }
  ]
  // Variables globales
  const resources_ = getResources();

  // refenciamos la clase del sistema
  const cruceroSystem_ = new CruceroSystem(resources_);

  // Procesamos los proyectos
  return cruceroSystem_.generateMaterials(projectList_);
}

/**
 * Permite obtener los datos de configuración
 */
function getCroisingParams() {
  
  // Variable a retornar
  let paramsObject_ = {
    diameters: {},
    valves: {},
  };

  // Obtenems los parametros globales
  const resources_ = getResources();

  // Referenciamos a la hoja de calculo 
  const spreadsheet_ = SpreadsheetApp.openById(resources_.crossingSheetId);
  
  // Refereciamos las pestañas de diametros
  var diameterSheet_ = spreadsheet_.getSheetByName(resources_.diameterSheetName);
  var valveSheet_ = spreadsheet_.getSheetByName(resources_.diameterValveSheetName);

  // Validamos que realmente existan hojas
  if(diameterSheet_ && valveSheet_){
  
    // Obtenemos los datos de diametros
    paramsObject_.diameters = consultDiameters_(diameterSheet_, '"');

    // Obtenemos los datos de válvula
    paramsObject_.valves = consultDiameters_(valveSheet_, ' mm');
  }
  
  // retornamos los parametros
  return paramsObject_;
}

/**
 * Permite obtener los datos de diametros
 */
function consultDiameters_(sheet_, suffix_){
  let diametersObject_ = {};

  // obtenemos la ultima fila
  const lastRow_ = sheet_.getLastRow();
  
  // validamos que exista mas de 1 fila
  if(lastRow_ > 1){
    
    // Obtenmos los datos
    const arrayData_ = sheet_.getRange(2, 1, lastRow_ - 1, 3).getValues();
    
    // Recorremos los datos 
    for(let i = 0; i < arrayData_.length; i++){
      let recordObject_ = arrayData_[i];
      
      // validamos si existe valores para continuar
      if(recordObject_[0] && recordObject_[1]){

        // verificamos que exista la propiedad
        if(!diametersObject_[recordObject_[1]]){

          // Iniciamos el objeto
          diametersObject_[recordObject_[1]] = {
            value: recordObject_[0],
            reference: getChipsValues_(recordObject_[2], suffix_)
          };
        }
      }
    }
  }

  // Retornamoslos datos
  return diametersObject_;
}

/**
 * Permite obtener valores relacionados en un chip de Google Sheet
 */
function getChipsValues_(chipString_, suffix_ = ""){
  if(!chipString_) return "";

  // Obtenemos el valor y eliminamos ls espacios
  chipString_ = String(chipString_).trim();

  // Obtenemos concidencias
  const matchs_ = chipString_.match(/\d+(?:\s\d+\/\d+|\/\d+)?/g);

  // Validamos si existen
  if(matchs_){
    return matchs_.map(s => (s.trim() + suffix_));
  }
  return "";
}
