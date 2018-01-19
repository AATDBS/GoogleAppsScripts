function getMessagesWithLabel() {     
      
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  var threads = GmailApp.getUserLabelByName('VPN').getThreads(0,100);  //threads from 0 to 100
  var firstEmptyRow = sh.getLastRow();
  var cell = sh.getRange(firstEmptyRow,1)
  var procesados = 0;

  Logger.log("Threads: " + threads.length);  
  
  for (var i = 0; i < threads.length; i++) {
    Logger.log("Thread: " + threads[i].getFirstMessageSubject());    
    
    var messages = threads[i].getMessages();
    Logger.log("mensajes : " + messages.length);
    
    for (var j = 0; j < messages.length; j++) {
      Logger.log("Message: " + messages[j].getSubject());
      //inserta el registro
      procesados++;
      firstEmptyRow = firstEmptyRow + 1;
      cell = sh.getRange(firstEmptyRow,1)
      cell.setValue(messages[j].getDate());
     
      cell = sh.getRange(firstEmptyRow,2)
      cell.setValue(messages[j].getSubject());
            
      cell = sh.getRange(firstEmptyRow,3)
      var remitente = messages[j].getFrom();      
      cell.setValue(remitente);
            
      cell = sh.getRange(firstEmptyRow,4)
      //extrae la parte de dominio del email
      dominio = remitente.substring(remitente.indexOf("@")+1, remitente.length-1);
      cell.setValue(dominio);
      
    } //end for j
    
  } //end for i
  
  Logger.log("Processed: " + procesados);  
  
}
