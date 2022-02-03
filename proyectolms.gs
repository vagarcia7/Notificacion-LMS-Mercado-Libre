function formSubmitReply(e){
  var emailsTM = []
  var emailsTT = []
  var emailsTN = []

  var fecha = new Date()
  var horaUTCmenos5 = fecha.getHours()
  var horaActual = horaUTCmenos5+2
  var minutos = fecha.getMinutes()
  var segundos = fecha.getSeconds() 
  if(horaActual == 24){
    horaActual = 00
  }
  if(horaActual == 25){
    horaActual == 02
  }
  var horaCompleta = horaActual + ':' + minutos + ':' + segundos
  var horasTurnoManiana = [6,7,8,9,10,11,12,13]
  var horasTurnoTarde = [14,15,16,17,18,19,20,21]
  var horasTurnoNoche = [22,23,00,01,2,3,4,5,6]
  var diaDeSemana=fecha.getDay()


  var msg = "Desde el usuario de " + e.values[1] + " a las " + horaCompleta + " se solicita el cambio de estado a no sistémico/tiempo no disponible. Motivo: " +  e.values[2] + "<br>" + "<a href='https://envios-lms.mercadolibre.com.ar/time-assign' class='button'>Asignar</a>" + "<br>" + "<img src='https://lh3.googleusercontent.com/HTJvGNLS4KG5F0yKCrWHbQ2WbMLFaO1QLp7WgP6q3WuSA9nW6Izc9JDlpIAD7nUJMWkpY7-VDevxVpZ6GZGa1mk3i3e7lDZya3dL--DbbU-nzEMBNtnlD-_muyyAo4vhMnL-aZn6mA2II_El9Ty53nOOdnF2gdZfZ7krWztldmtY6Eni1Gz_gjjBY9OrV38mw7rgR_je5_bQBd1nA2T0yzlsTg1aH-iVZmzrAEPCNa1u1_FZ4NUBg1pZ3oh8Tx9CPTNuvoaiGgaS_eSDNJBOAW2vb2q1pAsaF_6bUiVpwVcWw8kmF5gQUNCUetOjvVcKkZ5hQiounJdVT2Y7w_W1h4MePvdlRy8gaTz0f7175SgU-77Blv7j5LaMZ-fDH-P-4jEpRZqCHNFff7FgAq90m-9cFg_N-eW0yJcYoGYa6UmdfxLJFss3dGg0jm-Lhsqnrf4HlXRtYsYub3T58t1H4n5UtZNurvN13VXsyQZUgqkH6J1rF2ls6ZBFNnFMyuzca-UlOENlm2Z5TwGhf4ieNmaQspWI1wNdY2pID7nw2Se__HOeFQb0F8kWWklHbPg7ZiR4yV7OjzDpSahri-JorKmyWTeRRCDNUxYXaVMcNCCXjyCexF2Bj8Tyhi6H-SiQBkmOTaPLCkEK0sU0W1Rast2-eoq1mNl3urbr6KWaSOD3SgguutsDp-YMZeFFc8ukp9-CNp-iEXcJirskwtADubi7=w1920-h582-no?authuser=0' alt='' width='150' height='50' /> &nbsp; &nbsp; &nbsp; &nbsp; <img src='https://lh3.googleusercontent.com/CVatTPuOztxhmZO9AVa8T9xIm7RPSfReaTe9g1HRVCryiRRjndsLDqKymPSMu7ZlMXyZmjsDBi33tdBz7yOWR92EMKZNsvjJeeoZj-BkbyiEGKvzrvssl0RuNox4Fj7HMXNxpBm0YwyX_YtUahLfWz2liX37rN7ybJ31v5sbCUWGx31TU9LkTqg6bksokoRHYLjMlp0xvbhkUFncUdTxuqvODo8rFxoD-hvlV1iO6_o6xahVBEzya_5OQQj1d-FDYoBNdxyGzsr9ncyGBaj5tNZpmhdDpEIwtN3lXX_NSjQx9TBqLm4pH0cdk8XWFvcxKFynLIwhyMKg6-7WQYdeBQ99vL7IAOZclb3HWOCqnnEh_SqslkfxGTUHhcIGlocMBxeVcxhXFrBs-aK6pN4osAUSGlLb-lyrfiK5Sh3dEnG3VxgIom-VZgwPXwY49l4mV9JnzJg8Ss_R51jZw21gOsjKkM0w94sVF8v_h30uOTZP0C12jfb0bUmLZ28bophW4Z9izIRXU4yQQ37okTPlIkw9N5CDIdJhXkoSPXGj2j_DH4zJpFobXdVFgiaO8t9l-8VY3cWT8gIVgsrMTp6kHh7jkcTOi4oiL8ZK0pMA_nPUjd1NGwdGhKYDgU6yBGvDoIb_oXnT4ZxcqHTXpiG_2IkrazxLtbUt8CruFjL8EHXGbju6uaELGI2XB28XTSugaPN-6L2p13-U-sUZRp5xiN1c=w636-h313-no?authuser=0' alt='' width='122' height='60' />"


  // if para cuando es Sábado
if(diaDeSemana == 6){
    var horasTurnoManiana = [8,9,10,11,12,13,14,15]
    var horasTurnoNoche = [22,23,00,01,2,3,4,5,6]
  }
  
  // if para cuando es Domingo
  if(diaDeSemana == 0){
    var horasTurnoTarde = [8,9,10,11,12,13,14,15,16]
    var horasTurnoNoche = [21,22,23,00,01,2,3,4,5]
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
