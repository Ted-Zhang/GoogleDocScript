// This just demo for how to import thrid party library and modify google doc.
// The logic is only works for two people, but you can easily extend it by know the row number's in the table. just multiple by seven.
function onOpen() {
  loadMomentLibrary();
  
  var headerSection = DocumentApp.getActiveDocument().getHeader();

  if(!headerSection) {
    headerSection = DocumentApp.getActiveDocument().addHeader();
  }
  
  var rosterTable = headerSection.findElement(DocumentApp.ElementType.TABLE);
  if(!rosterTable){
    createRosterTable(headerSection);
  }
  else {
    rosterTable = rosterTable.getElement().asTable();
    updateRosterTable(rosterTable);
  }
}

function createRosterTable(container) {
  var people = [
   ['Tim', '04/06/2018'],
   ['Jared', '11/06/2018']
    
 ];
  
  container.appendTable(people);
}

function updateRosterTable(rosterTable) {
  var rowCount = rosterTable.getNumRows();
  
  for(var i= 0; i < rowCount; i++) {
    var speechDateCell = rosterTable.getCell(i, 1)
    
    var speechDate = speechDateCell.getText();
    
    var momentSpeechData = moment(speechDate, 'DD/MM/YYYY');
    if(momentSpeechData < moment()){
      var nextSpeechData = momentSpeechData.add(14, 'd');
      speechDateCell.setText(nextSpeechData.format("DD/MM/YYYY"));
      speechDateCell.setBackgroundColor('#ffffff');
    } 
    
    if (momentSpeechData <= moment().add(7, 'd') ){
      speechDateCell.setBackgroundColor('#1759c4');
    }
  }
}


function loadMomentLibrary() {
  var url = "https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.22.1/moment.min.js";
  var javascript = UrlFetchApp.fetch(url).getContentText();
  eval(javascript);
}
