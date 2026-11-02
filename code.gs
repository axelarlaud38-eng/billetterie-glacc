// Service appelé par la page web : ?code=XXXX -> renvoie du JSON
function doGet(e) {
  var code = (e && e.parameter && e.parameter.code) ? String(e.parameter.code) : '';
  var resultat = chercherCode(code);
  return ContentService
    .createTextOutput(JSON.stringify(resultat))
    .setMimeType(ContentService.MimeType.JSON);
}

function chercherCode(codeSaisi) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var codeNet = codeSaisi.trim().toUpperCase();
  if (!codeNet) return { success: false };

  var data = ss.getSheetByName('Codes').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var categorie = data[i][0];
    var codePlace = String(data[i][1]).trim().toUpperCase();
    var lien = data[i][2];
    if (codePlace && codePlace === codeNet) {
      logUtilisation(ss, codeSaisi, categorie, 'Redirigé');
      return { success: true, url: lien };
    }
  }
  logUtilisation(ss, codeSaisi, '', 'Code invalide');
  return { success: false };
}

function logUtilisation(ss, code, categorie, resultat) {
  ss.getSheetByName('Logs').appendRow([new Date(), code, categorie, resultat]);
}
