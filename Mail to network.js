/*function mailNetwork() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  libro.setActiveSheet(libro.getSheetByName("Mail to network"));
  const hoja = SpreadsheetApp.getActiveSheet();
  const filas = hoja.getRange("A2:B").getValues();
  
  
   for (indiceFila in filas) {
   var contacto = crearContacto(filas[indiceFila]);
       sendmail(contacto);   
   }
}

function crearContacto(datosFila) {
  const contacto = {
    mail:datosFila[0],
    sendEmail: datosFila[1],
  };
  return contacto;
}

function sendmail(contacto) {
  if (contacto.email=="") return;
      const plantilla = HtmlService.createTemplateFromFile("Mail content to network");
      plantilla.contacto = contacto;
      const mensaje = plantilla.evaluate().getContent();
      MailApp.sendEmail({to: estandar.email, replyTo: "claudio.sepulveda@aiesec.net", subject: "Resultados Account Management Minimums", htmlBody: mensaje});
}


function onOpen2() {
  const spreadsheet = SpreadsheetApp.getActive();
  const menuItems = [{name: 'Mail to network', functionName: 'mailNetwork'}];
  spreadsheet.addMenu('mailNetwork', menuItems);
} */



/*function onOpen2() {
  const spreadsheet2 = SpreadsheetApp.getActive();
  const menuItems2 = [{name: 'Mail to network', functionName: 'mailNetwork'}];
  spreadsheet.addMenu('mailNetwork', menuItems2);
}*/
