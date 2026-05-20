function testGetFilderData() {
  // obtenemos el Id de la hoja de validaciones 
  var validationSheetId =  getResources().validationSheetId;
  
  // Referenciamos la hoja de validaciones
  var spreadSheet = SpreadsheetApp.openById(validationSheetId);
  
  return {
    department: getDepartmentData(spreadSheet),
    keyWord: getKeywordArrayData(spreadSheet)
  }
}

function sss(){
  console.log("oksss")
}
