function enviarPrimeiroTeste() {

  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  const aba = planilha.getSheetByName("EMAILS_ORGANIZADOS");

  // Pega apenas primeira linha
  const dados = aba
    .getRange(2, 1, 1, 2)
    .getValues();

  const emailPrincipal = "ubametra.admexterno@gmail.com";

  const html = HtmlService
    .createHtmlOutputFromFile('email_nr1')
    .getContent();

  MailApp.sendEmail({
    to: emailPrincipal,
    subject: "[TESTE] Nova NR-1",
    htmlBody: html
  });

  Logger.log(`Teste enviado para: ${emailPrincipal}`);
}