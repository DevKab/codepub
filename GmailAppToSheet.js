function onOpen() {
  var menu = SpreadsheetApp.getUi().createMenu("📑 Reporte Correo");
  menu.addSeparator();
  menu.addItem("🔄| Actualizar datos ", "Abril");
  menu.addSeparator();
  menu.addToUi();
}
function Abril() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]
  const buscafecha = Utilities.formatDate(new Date(), "GMT-6", "yyyy-M-dd");
  const before = Utilities.formatDate(new Date(), "GMT-6", "yyyy-12-31")
  var query = "after:" + buscafecha + " before:" + before + "label:inbox";
  const cadenas = GmailApp.search(query);
  hoja.getRange("A1").setValue("REMITENTE").setBackgroundColor('#46bdc6').setFontColor('#ffffff');
  hoja.getRange("B1").setValue("ASUNTO").setBackgroundColor('#46bdc6').setFontColor('#ffffff');
  hoja.getRange("C1").setValue("MENSAJE").setBackgroundColor('#46bdc6').setFontColor('#ffffff');
  hoja.getRange("D1").setValue("FECHA").setBackgroundColor('#46bdc6').setFontColor('#ffffff');
  hoja.getRange("E1").setValue("HORA").setBackgroundColor('#46bdc6').setFontColor('#ffffff');
  hoja.getRange("F1").setValue("NOMBRE").setBackgroundColor('#46bdc6').setFontColor('#ffffff');
  hoja.getRange("G1").setValue("MESA").setBackgroundColor('#46bdc6').setFontColor('#ffffff');
  var bucle = 0;
  cadenas.forEach(cadena => {
    bucle ++;
    const asunto = cadena.getFirstMessageSubject();
    const correo = cadena.getMessages()[0];
    const cuerpo = correo.getPlainBody();
    const remitente = correo.getFrom().split("<")[1].split(">")[0];
    const corto = remitente;
    const fechahoy = Utilities.formatDate(new Date(), "GMT-6", "yyyy-M-dd");
    const fechaor = Utilities.formatDate(correo.getDate(), "GMT-6", "yyyy-M-dd");
    if (fechaor == fechahoy) {
      var date = Utilities.formatDate(correo.getDate(), "GMT-6", "yyyy-M-dd");
      var horaoriginal = Utilities.formatDate(correo.getDate(), "GMT-6", "h a");
    } else {
      var date = Utilities.formatDate(cadena.getLastMessageDate(), "GMT-6", "yyyy-M-dd");
      var horaoriginal = Utilities.formatDate(cadena.getLastMessageDate(), "GMT-6", "h a");
    }
    //GmailApp.markMessagesRead(cadena.getMessages());
    //const lectura = cadena.isUnread();
    if (corto == "ejemplo@ejemplo.com" || corto == "ejemplo2@ejemplo.com") {
      switch (corto) {
        case "ejemplo@ejemplo.com":
          var nombre = "EJEMPLO";
          break;
        case "ejemplo2@ejemplo.com":
          var nombre = "EJEMPLO 2";
          break;
        default:
          var nombre = "DESCONOCIDO";
          break;
      }
      switch (horaoriginal) {
        case "12 AM":
          var tiempo = "MAÑANA";
          break;
        case "1 AM":
          var tiempo = "MAÑANA";
          break;
        case "2 AM":
          var tiempo = "MAÑANA";
          break;
        case "3 AM":
          var tiempo = "MAÑANA";
          break;
        case "4 AM":
          var tiempo = "MAÑANA";
          break;
        case "5 AM":
          var tiempo = "MAÑANA";
          break;
        case "6 AM":
          var tiempo = "MAÑANA";
          break;
        case "7 AM":
          var tiempo = "MAÑANA";
          break;
        case "8 AM":
          var tiempo = "MAÑANA";
          break;
        case "9 AM":
          var tiempo = "MAÑANA";
          break;
        case "10 AM":
          var tiempo = "MAÑANA";
          break;
        case "11 AM":
          var tiempo = "MAÑANA";
          break;
        case "12 PM":
          var tiempo = "MEDIO DIA";
          break;
        case "1 PM":
          var tiempo = "MEDIO DIA";
          break;
        case "2 PM":
          var tiempo = "MEDIO DIA";
          break;
        case "3 PM":
          var tiempo = "MEDIO DIA";
          break;
        case "4 PM":
          var tiempo = "TARDE";
          break;
        case "5 PM":
          var tiempo = "TARDE";
          break;
        case "6 PM":
          var tiempo = "TARDE";
          break;
        case "7 PM":
          var tiempo = "TARDE";
          break;
        default:
          var tiempo = "FUERA DE HORARIO";
          break;
      }
      if (date == buscafecha) {
        //if (lectura == true) {
          Logger.log(bucle + " .-" + corto) //solo para consola
        //hoja.appendRow([corto, asunto, cuerpo, date, horaoriginal, nombre, 'GREYZ', tiempo]) //descomentar para imprimir en hoja de sheet
        //}
      }
    }
  })
}
function DelTable() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]
  const del = hoja.getRange("A2:J").clearContent();
}
