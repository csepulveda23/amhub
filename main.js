function sendemails() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  libro.setActiveSheet(libro.getSheetByName("Scoreboard form"));
  const hoja = SpreadsheetApp.getActiveSheet();
  const filas = hoja.getRange("A2:V").getValues();
   
   for (indiceFila in filas) {
   var estandar = crearEstandar(filas[indiceFila]);
     if(estandar.sendEmail == "Yes"){
       sendmail(estandar);}   
   }
}

function crearEstandar(datosFila) {
  var random = Math.floor(Math.random()*3);
  const estandar = {
    hora:datosFila[0],
    comite: datosFila[1],
    nombre: datosFila[2],
    estandarType: datosFila[3],
    comprobante: datosFila[4],
    cuenta1: datosFila[5],
    cuenta2: datosFila[6],
    cuenta3: datosFila[7],
    email: datosFila[8],
    emailSent: datosFila[9],
    sendEmail: datosFila[10],
    fotoam: datosFila[11],
    void4: datosFila[12],
    aprobacion: datosFila[13],
    prioridad: datosFila[14],
    vigente: datosFila[15],
    venceen: datosFila[16],
    porcentaje: datosFila[17],
    puntos: datosFila[18],
    index: datosFila[19],
    aliado: datosFila[20],
    comments: datosFila[21],
    randomnumber: random
  };
  return estandar;
}

//function numeroAleatorio(){
//  /*Numero aleatorio*/ const randomnumber = Math.floor(Math.random()*3);
//return randomnumber;
//}

function sendmail(estandar,randomnumber) {   
  if (estandar.email=="") return;
      const plantilla = HtmlService.createTemplateFromFile("AMM Message");
      plantilla.estandar = estandar;
      const mensaje = plantilla.evaluate().getContent();
      MailApp.sendEmail({to: estandar.email, replyTo: "claudio.sepulveda@aiesec.net", subject: "Account Management Minimums", htmlBody: mensaje});
}

/********************************************************************************************************************************************************/
function mailNetwork() {
  const libro2 = SpreadsheetApp.getActiveSpreadsheet();
  libro2.setActiveSheet(libro2.getSheetByName("Mail to network"));
  const hoja2 = SpreadsheetApp.getActiveSheet();
  const filas2 = hoja2.getRange("A2:B").getValues();
  
  
   for (indiceFila2 in filas2) {
   var contacto = crearContacto(filas2[indiceFila2]);
     if(contacto.sendEmail2 == "Yes"){
       mailtonetwork(contacto);}       
   }
}

function crearContacto(datosFila2) {
  const contacto = {
    mail:datosFila2[0],
    sendEmail2: datosFila2[1],
  };
  return contacto;
}

function mailtonetwork(contacto) {
  if (contacto.email=="") return;
      const plantilla2 = HtmlService.createTemplateFromFile("Mail content to network");
      plantilla2.contacto = contacto;
      const mensaje2 = plantilla2.evaluate().getContent();
      MailApp.sendEmail({to: contacto.mail, replyTo: "claudio.sepulveda@aiesec.net", subject: "Resultados Account Management Minimums", htmlBody: mensaje2});
}
/********************************************************************************************************************************************************/
function onOpen() {
  const spreadsheet = SpreadsheetApp.getActive();
  const menuItems = [{name: 'Enviar', functionName: 'sendemails'}];
  spreadsheet.addMenu('sendemails', menuItems);
  const spreadsheet2 = SpreadsheetApp.getActive();
  const menuItems2 = [{name: 'Mail to network', functionName: 'mailNetwork'}];
  spreadsheet2.addMenu('mailNetwork', menuItems2);
}

