    const ss =SpreadsheetApp.getActiveSpreadsheet();
    const formWs = ss.getSheetByName("Form");
    const settingWs = ss.getSheetByName("Settings");
    const dataWs = ss.getSheetByName("Data");
    const idCell = formWs.getRange("C5");
    const searchCell= formWs.getRange("C7");
    const fieldRange = ["C9","F9","C11","F11"];

    function saveRecord(){
      const id =idCell.getValue();
      if (id == ""){
        createNewRecord();
        return;
      }
      // Mencari Id didalam sheet Data
      const cellFound = dataWs.getRange("A:A")
                        .createTextFinder(id)
                        .matchCase(true)
                        .matchEntireCell(true)
                        .findNext();

      if (!cellFound) return;
      const row = cellFound.getRow();
      const fieldValues = fieldRange.map(f => formWs.getRange(f).getValue());
      fieldValues.unshift(id);
      dataWs.getRange(row,1,1,5).setValues([fieldValues]);
      searchCell.clearContent();
      ss.toast("id:"+ id,"Edited Succesed!");
    }

  function createNewRecord(){
 
    const fieldValues = fieldRange.map(f => formWs.getRange(f).getValue());
    const nextIdCell = settingWs.getRange("A2");
    const nextId = nextIdCell.getValue();
   
    fieldValues.unshift(nextId);
    dataWs.appendRow(fieldValues);
    idCell.setValue(nextId);
    nextIdCell.setValue(nextId+1);
    ss.toast("id:"+ nextId,"New Record created!");
  }

  function newRecord(){
    fieldValues = fieldRange.forEach(f => formWs.getRange(f).clearContent());
    idCell.clearContent();
    searchCell.clearContent();
    ss.toast("Clear Field !");
  }

function search(){
  
  const searchValue=searchCell.getValue();
  const data = dataWs.getRange("A2:F").getValues();
  //console.log("Horray!!");
  // r[6] --> search Column ada di colom 6 jika column berubah harus diganti.
  const recordFounds = data.filter(r => r[5] == searchValue);
  console.log(recordFounds.length);
  if (recordFounds.length === 0) return;
  idCell.setValue(recordFounds[0][0]);
  fieldRange.forEach((f,i) => formWs.getRange(f).setValue(recordFounds[0][i+1]));
  ss.toast("recordFound:"+ recordFounds.length,"SEARCH SUCCESED !");
}

function deletedRecord(){
  const id =idCell.getValue();
    if (id == ""){
     return;
    }
    
    // Mencari Id didalam sheet Data
    const cellFound = dataWs.getRange("A:A")
                      .createTextFinder(id)
                      .matchCase(true)
                      .matchEntireCell(true)
                      .findNext();

    if (!cellFound) return;
    const row = cellFound.getRow();
    dataWs.deleteRow(row);
    newRecord();
    ss.toast("id:"+ id,"Record Deleted!");
}

/// PDF ---------------------------------------------------
  const tempSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Temp"); 
  const nextIdCell = settingWs.getRange("A2");
  const nextId = nextIdCell.getValue();
  
function createTempRow(){
   
 // Create Temp row on new sheet
    const fieldValues = fieldRange.map(f => formWs.getRange(f).getDisplayValue());  
    fieldValues.unshift(nextId);
    //console.log(fieldValues)
    tempSheet.appendRow(fieldValues);
    ss.toast("id:"+ nextId,"Temp Created!");

}
function clearRowTemp(){

//Clear Row
  const idCellTemp = tempSheet.getRange("A2");
  const id =idCellTemp.getValue();
    if (id == ""){
     return;
    }
    // Mencari Id didalam sheet Data
    const cellFound = tempSheet.getRange("A:A")
                      .createTextFinder(id)
                      .matchCase(true)
                      .matchEntireCell(true)
                      .findNext();

    if (!cellFound) return;
    const row = cellFound.getRow();   
    tempSheet.deleteRow(row);
     ss.toast("id:"+ id,"Temp Clear!");
}

//--------Create PDF FILE----------------------------------------
function createPDF(pdfName,docFile,tempFolder,pdfFolder) 
{
  // CODE BUAT WORD FILE DARI ROW KE WORD------
  const tempFile = docFile.makeCopy(tempFolder).setName(pdfName);//buat copy ke word.
  const tempDocFile = DocumentApp.openById(tempFile.getId());
  const body = tempDocFile.getBody();
  const dataBody = tempSheet.getRange(2,1,tempSheet.getLastRow()-1,6).getDisplayValues();
  
  // Isi data dari column
  /* 
   // data pada row diakses dengan array data[0][0]= id
   */
    body.replaceText("{first}",     dataBody[0][1]);
    body.replaceText("{last}",      dataBody[0][2]);
    body.replaceText("{age}",       dataBody[0][3]);
    body.replaceText("{location}",  dataBody[0][4]);
 
  
  tempDocFile.saveAndClose();

  // CODE BUAT PDF FILE
  const pdfContentBlob = tempFile.getAs(MimeType.PDF);
  pdfFolder.createFile(pdfContentBlob).setName(pdfName);
  // Word File deleted-----------------------
  tempFile.setTrashed(true);
}

//PDF FOLDER. https://drive.google.com/drive/folders/14ZfGmBVxf8PkrsAI9Co3gGcR0cSJruxI
// Temp Folder https://drive.google.com/drive/folders/1ZEm-U8XwLMs7ZgS90Iqw8bqL6poxw6uw
// Doc File https://docs.google.com/document/d/18laFRaTRt7UfcfKAmXBYk98TO1_V6UWslqcY9PUoXXM/edit


function createPDFFile(){
    const tempFolder = DriveApp.getFolderById("1ZEm-U8XwLMs7ZgS90Iqw8bqL6poxw6uw");
    const pdfFolder = DriveApp.getFolderById("14ZfGmBVxf8PkrsAI9Co3gGcR0cSJruxI");
    const docFile = DriveApp.getFileById("18laFRaTRt7UfcfKAmXBYk98TO1_V6UWslqcY9PUoXXM");
    const id =idCell.getValue();
    if (id == ""){
      return;
    }
   createTempRow();
    const dataBody = tempSheet.getRange(2,1,tempSheet.getLastRow()-1,6).getDisplayValues();
    createPDF(dataBody[0][1]+" "+dataBody[0][0],docFile,tempFolder,pdfFolder);
    ss.toast("id:"+ nextId,"PDF LPD created!");
    clearRowTemp();
}






