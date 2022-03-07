function formSubmitReply(e){
  var emailsTM = []
  var emailsTT = []
  var emailsTN = []

  var tiempoTranscurrido = Date.now()+(1000*60*60)*2
  var fecha = new Date(tiempoTranscurrido)
  var horaActual = fecha.getHours()
  var minutos = fecha.getMinutes()
  var segundos = fecha.getSeconds() 
  var horaCompleta = horaActual + ':' + minutos + ':' + segundos
  var horasTurnoManiana = [6,7,8,9,10,11,12,13]
  var horasTurnoTarde = [14,15,16,17,18,19,20,21]
  var horasTurnoNoche = [22,23,0,1,2,3,4,5,6]
  var diaDeSemana=fecha.getDay()

  // if para cuando es Sábado
  if(diaDeSemana == 6){
    var horasTurnoManiana = [8,9,10,11,12,13,14,15]
    var horasTurnoNoche = [22,23,0,1,2,3,4,5,6]
  }
  
  // if para cuando es Domingo
  if(diaDeSemana == 0){
    var horasTurnoTarde = [8,9,10,11,12,13,14,15,16]
    var horasTurnoNoche = [21,22,23,0,1,2,3,4,5]
  }

  var msg = "Desde el usuario de " + e.values[3] + " a las " + horaCompleta + " se solicita el cambio de estado a no sistémico/tiempo no disponible. Motivo: " +  e.values[2] + "<br>" + "<a href='https://envios-lms.mercadolibre.com.ar/time-assign' class='button'>Asignar</a>" + "<br>" + "<img src='https://lh3.googleusercontent.com/OF72sU4NQldz4EvpO00Nl2NQM8UtvTyvr-HPBgHehPWRH0NaQUYUKmmwcXjjTOGJ7Xj8n8Bhg2eY4xQ47NUUZNSUReP1uwe61WZHgU6WzXqteEsYUe9fUut_ayNy-gNM_RbvhXxhgpGAXvP3l_Lv9wly0eRqNejqOxoYgZjLeeRtxeP0SW3Mdmy4DOgM171fHQHJX08i6qo8H2KQwrImgjESyEIQpqsQRpbA_e_AJtZimAP6qyZyDNpbsP_UJEvwFXApn6ZZjnXBEGC8IvPoz7HbKXIU84lIYAmiUglskQY7Q_hBQxxH7Z870HwpsvqVTe6Q7wAlMNBkl5oFynGFLUX_TCOlqaFkkrFTeD61BRB3YCOhv_7DkV4cE3MLiXUHH4CM-_irJ69xLDbMwP6wWCLBKMnglsT-BQggU4-h8Q2aPObvC5UsH3S_llcPhzWHWUNKE4gvQ4hS31OBiQDiaUkoXuxcBAf8BpoQTohLOWlhtIoJgTULJkUXXsWovpEnjaoYoG1wwVP1sUUDV8WdYosL3qKsJbkDv1EjM3d4kOyS3t5UOo8oa8-OokNc2paoxz0agpDQX0avjaFhrTQx0g4-MuPLi58paL49azADlSowJXzYZHTRgQbngZLFxFloELj5gn3bti4bh9PWFhxkXt9LMJ5jXtUPORR2VYYacLNAs-s9h0kYuCd4YbL8pGqU1wY_DFbz7RMcGZyY7snXi-Q=w1366-h414-no?authuser=2' alt='' width='150' height='50' /> &nbsp; &nbsp; &nbsp; &nbsp; <img src='https://lh3.googleusercontent.com/QgQTFg7gaVyBu6XySBE3K4pHBO-Q20kIDyjL14ycL8KcvSdnHCMnVpHWg-JQNVX2qdtapqz1mN6ff-oddvWod9jcWhsXc7pSy38BeHIIgoQLZgfuoP5xth2Wk8Bd8t42Ic3vPQGEH3UjsiGYYuNIzB2_oY9pTY2AwtOmL3zRPdnDEF1-SJ_EAZqRLtMPl_FKgEj8Hba_Ebs9F5msoEVOmdffWEGwPmXFyF006MnALcM0AOAcMGAm2oBU13QHzZIqkToE9wPY44Ju4gsPQY2TiOPo7AA5TnwJEqfpCQ4Um9g4XCpg3rxO_e2ATy9Vm_FEW54CvACzW6YBbAOrQWOOq5Ipx8zgYHn4cTyAPwiIkucyartxI81mx-uINaUs01mVhH_6lVJrYFQTvnxfQa1AWIySbIUECn47PmkIGlFfm248Rri88SiDKDY2xAH4TOof_VMOgPLOzpqAKZg-TEWZWCjAtSQoKjG2kXOlxI8LAL6ph7bVPwzJcWOxk2IZepwbrbJrCpcZfFtQ6CBPMPshL9ea_ai8s4OCPSjf5zYqksI8geg6Eu5lJ1vfMSkuLgFmob9MLRj_0xKJw27xdYSH3NgfSH6N_siVWG9KiOthkPMdm9r90wzy1_4XpooSg-jFzbSYH9S426A5Cp_dgRpnDq8iymFs4r7eOCKvaEC5492SW6cfsPo1R9ED0i3LtkG_UHNOq4mhBmKawMz4lMd3ecE=w636-h313-no?authuser=2' alt='' width='122' height='60' />"

var sector = e.values[1]


if (sector == "Retiros"){
  if(horasTurnoManiana.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('A38:A45').getDisplayValues()
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
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('B38:B45').getDisplayValues()
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
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('C38:C45').getDisplayValues()
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

else if (sector == "INBOUND | Receiving"){
  if(horasTurnoManiana.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('J4:J11').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Receiving",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoTarde.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('K4:K11').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Receiving",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoNoche.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('L4:L11').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Receiving",
      htmlBody: msg
      })
      
    }
  }
}

else if (sector == "INBOUND | Check In MZ"){
  if(horasTurnoManiana.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('A4:A11').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Check In MZ",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoTarde.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('B4:B11').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Check In MZ",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoNoche.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('C4:C11').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Check In MZ",
      htmlBody: msg
      })
      
    }
  }
}

else if (sector == "INBOUND | Put Away MZ"){
  if(horasTurnoManiana.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('D4:D11').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Put Away MZ",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoTarde.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('E4:E11').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Put Away MZ",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoNoche.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('F4:F11').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Put Away MZ",
      htmlBody: msg
      })
      
    }
  }
}

else if (sector == "INBOUND | RK"){
  if(horasTurnoManiana.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('G4:G11').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Inbound RK",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoTarde.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('H4:H11').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Inbound RK",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoNoche.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('I4:I11').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Inbound RK",
      htmlBody: msg
      })
      
    }
  }
}

else if (sector == "OUTBOUND | Picking MZ"){
  if(horasTurnoManiana.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('A16:A23').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Picking MZ",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoTarde.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('B16:B23').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Picking MZ",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoNoche.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('C16:C23').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Picking MZ",
      htmlBody: msg
      })
      
    }
  }
}

else if (sector == "OUTBOUND | Packing MZ"){
  if(horasTurnoManiana.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('D16:D23').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Packing MZ",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoTarde.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('E16:E23').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Packing MZ",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoNoche.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('F16:F23').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Packing MZ",
      htmlBody: msg
      })
      
    }
  }
}

else if (sector == "OUTBOUND | Wall"){
  if(horasTurnoManiana.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('G16:G23').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Wall",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoTarde.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('H16:H23').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Wall",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoNoche.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('I16:I23').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Wall",
      htmlBody: msg
      })
      
    }
  }
}

else if (sector == "OUTBOUND | RK"){
  if(horasTurnoManiana.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('J16:J23').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Outbound RK",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoTarde.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('K16:K23').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Outbound RK",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoNoche.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('L16:L23').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Outbound RK",
      htmlBody: msg
      })
      
    }
  }
}

else if (sector == "Inventario"){
  if(horasTurnoManiana.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('A27:A34').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Inventario",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoTarde.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('B27:B34').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Inventario",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoNoche.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('C27:C34').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Inventario",
      htmlBody: msg
      })
      
    }
  }
}

else if (sector == "Bienes en Rezago"){
  if(horasTurnoManiana.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('D27:D34').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Bienes en Rezago",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoTarde.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('E27:E34').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Bienes en Rezago",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoNoche.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('F27:F34').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Bienes en Rezago",
      htmlBody: msg
      })
      
    }
  }
}

else if (sector == "Calidad"){
  if(horasTurnoManiana.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('G27:G34').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Calidad",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoTarde.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('H27:H34').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Calidad",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoNoche.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('I27:I34').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Calidad",
      htmlBody: msg
      })
      
    }
  }
}

else if (sector == "Devoluciones"){
  if(horasTurnoManiana.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('D38:D45').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Devoluciones",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoTarde.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('E38:E45').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Devoluciones",
      htmlBody: msg
      })
      
    }
  }

  if(horasTurnoNoche.includes(horaActual)){
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Emails Usuarios 2'), true);
    var usuarios = spreadsheet.getRange('F38:F45').getDisplayValues()
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
      subject: "Nueva solicitud de cambio de actividad LMS | Devoluciones",
      htmlBody: msg
      })
      
    }
  }
}
}