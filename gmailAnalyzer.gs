function getMessagesWithLabel() {     
      
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  var threads = GmailApp.getUserLabelByName('VPN').getThreads(1,10);  

  for(var n in threads){
    var msg = threads[n].getMessages();
    
    for(var m in msg){
                       
      //inserta el registro
      var firstEmptyRow = sh.getLastRow() + 1;
      var cell = sh.getRange(firstEmptyRow,1)
      cell.setValue(m.getSubject());
      //cell = sheet.getRange(firstEmptyRow,2)
      //cell.setValue(params.user);
      //cell = sheet.getRange(firstEmptyRow,3)
      //cell.setValue(params.date);      
    }      
  }
}
