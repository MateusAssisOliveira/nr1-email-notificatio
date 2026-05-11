function enviarEmail() {
  const html = HtmlService.createHtmlOutputFromFile('email_nr1').getContent();


  MailApp.sendEmail({
    to: "ubametra.admexterno@gmail.com",
    subject: "Nova NR-1",
    htmlBody: html
  });

}
