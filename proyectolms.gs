function formSubmitReply(e){
  var emailsTM = []
  var emailsTT = []
  var emailsTN = []

  var msg = "Desde el usuario de " + e.values[1] + " se solicita el cambio de estado a no sistémico/tiempo no disponible. Motivo: " +  e.values[2] + "<br>" + "<a href='https://envios-lms.mercadolibre.com.ar/time-assign' class='button'>Asignar</a>"


  var fecha = new Date()
  var horaUTCmenos5 = fecha.getHours()
  var horaActual = horaUTCmenos5+2
  var horasTurnoManiana = [6,7,8,9,10,11,12,13]
  var horasTurnoTarde = [14,15,16,17,18,19,20,21]
  var horasTurnoNoche = [22,23,24,25,2,3,4,5,6]
  var diaDeSemana=fecha.getDay()

  // if para cuando es Sábado
if(diaDeSemana == 6){
    var horasTurnoManiana = [8,9,10,11,12,13,14,15]
    var horasTurnoNoche = [22,23,24,25,2,3,4,5,6]
  }
  
  // if para cuando es Domingo
  if(diaDeSemana == 0){
    var horasTurnoTarde = [8,9,10,11,12,13,14,15,16]
    var horasTurnoNoche = [21,22,23,24,25,2,3,4,5]
  }


  if(horasTurnoManiana.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios'), true);
    var usuarios = spreadsheet.getRange('A2:A').getDisplayValues()
    for(usuario in usuarios){
      let usuarioLimpio = usuarios[usuario]
      let usuarioSinLlaves = usuarioLimpio[0]
      if(usuarioSinLlaves != ""){
      emailsTM.push(usuarioSinLlaves)
      }
    } 
    
    for(email in emailsTM){
      MailApp.sendEmail({
      to: emailsTM[email],
      subject: "Nueva solicitud de cambio de actividad LMS | Retiros",
      htmlBody: msg
      })
      
    }
  }

if(horasTurnoTarde.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios'), true);
    var usuarios = spreadsheet.getRange('B2:B').getDisplayValues()
    for(usuario in usuarios){
      let usuarioLimpio = usuarios[usuario]
      let usuarioSinLlaves = usuarioLimpio[0]
      if(usuarioSinLlaves != ""){
      emailsTT.push(usuarioSinLlaves)
      }
    } 
    
    for(email in emailsTT){
      MailApp.sendEmail({
      to: emailsTT[email],
      subject: "Nueva solicitud de cambio de actividad LMS | Retiros",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoNoche.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios'), true);
    var usuarios = spreadsheet.getRange('C2:C').getDisplayValues()
    for(usuario in usuarios){
      let usuarioLimpio = usuarios[usuario]
      let usuarioSinLlaves = usuarioLimpio[0]
      if(usuarioSinLlaves != ""){
      emailsTN.push(usuarioSinLlaves)
      }
    } 
    
    for(email in emailsTN){
      MailApp.sendEmail({
      to: emailsTN[email],
      subject: "Nueva solicitud de cambio de actividad LMS | Retiros",
      htmlBody: msg
      })
      
    }
  }
}

function prueba(){

}
