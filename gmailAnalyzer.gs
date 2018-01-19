function getMessagesWithLabel() {     
      
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  var threads = GmailApp.getUserLabelByName('VPN').getThreads(1,10);  
  var firstEmptyRow = sh.getLastRow();
  var cell = sh.getRange(firstEmptyRow,1)

  Logger.log("Threads: " + threads.length);
  
  for (var i = 0; i < threads.length; i++) {
    Logger.log("Thread: " + threads[i].getFirstMessageSubject());
    
    var messages = threads[i].getMessages();
    Logger.log("mensajes : " + messages.length);
    
    for (var j = 0; j < messages.length; j++) {
      Logger.log(messages[j].getSubject());
      //inserta el registro
      firstEmptyRow = firstEmptyRow + 1;
      cell = sh.getRange(firstEmptyRow,1)
      cell.setValue(messages[j].getDate());
      cell = sh.getRange(firstEmptyRow,2)
      cell.setValue(messages[j].getSubject());
      cell = sh.getRange(firstEmptyRow,3)
      cell.setValue(messages[j].getFrom());
    }    
  }
}
