function enviarEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Clientes");
  var dados = sheet.getDataRange().getValues();

  // HTML base (arquivo separado: emailTemplate.html)
  var htmlTemplate = HtmlService.createTemplateFromFile("emailTemplate");

  var hoje = new Date(); // pega a data de hoje
  hoje.setHours(0,0,0,0); // zera hora/min/seg para evitar erro de comparação

  for (var i = 1; i < dados.length; i++) {
    var nome = dados[i][0];       // Coluna A
    var email = dados[i][1];      // Coluna B
    var valor = dados[i][4];      // Coluna C
    var vencimento = new Date(dados[i][3]); // Coluna D (formato de data no Sheets)

    vencimento.setHours(0,0,0,0);

    //Só envia se HOJE == vencimento
    if (vencimento.getTime() <= hoje.getTime()) {
      
      htmlTemplate.nome = nome;
      htmlTemplate.valor = valor;
      htmlTemplate.vencimento = Utilities.formatDate(vencimento, "GMT-3", "dd/MM/yyyy");

      var corpoFinal = htmlTemplate.evaluate().getContent();

      GmailApp.sendEmail(email, "Decisão Final: Cobrança Imimente", "", {
        htmlBody: corpoFinal
      });
    }
    
}
}

function extenso(numero) {
  let resultado = writtenNumber(numero, {lang:'pt' });
  return(resultado)
}

function onOpen(e)
{ui.createMenu("Teste").addItem("Open")}
