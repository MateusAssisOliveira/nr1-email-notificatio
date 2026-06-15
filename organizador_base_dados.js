function organizarEmails() {

  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  const abaBase = planilha.getSheetByName("BASE_BRUTA");
  const abaDestino = planilha.getSheetByName("EMAILS_ORGANIZADOS");

  const ultimaLinha = abaBase.getLastRow();

  const dados = abaBase
    .getRange(2, 1, ultimaLinha - 1, 1)
    .getValues();

  abaDestino.clearContents();

  const dominiosPessoais = [
    "gmail.com",
    "hotmail.com",
    "outlook.com",
    "yahoo.com"
  ];

  const prioridades = [
    "rh@",
    "dp@",
    "contato@",
    "financeiro@"
  ];

  const corporativos = {};
  const pessoais = new Set();

  // PROCESSAR BASE
  for (const linha of dados) {

    const email = String(linha[0])
      .trim()
      .toLowerCase();

    if (!email || !email.includes("@")) continue;

    const dominio = email.split("@")[1];

    // EMAIL PESSOAL
    if (dominiosPessoais.includes(dominio)) {

      pessoais.add(email);

      continue;
    }

    // CORPORATIVOS
    if (!corporativos[dominio]) {
      corporativos[dominio] = [];
    }

    // Evita duplicado exato
    if (!corporativos[dominio].includes(email)) {
      corporativos[dominio].push(email);
    }
  }

  // RESULTADO FINAL
  const resultado = [];

  resultado.push([
    "EMAIL_PRINCIPAL",
    "CC",
    "DOMINIO",
    "TIPO"
  ]);

  // ORGANIZAR CORPORATIVOS
  for (const dominio in corporativos) {

    const emails = corporativos[dominio];

    let principal = emails[0];

    // Escolher principal por prioridade
    for (const prioridade of prioridades) {

      const encontrado = emails.find(email =>
        email.startsWith(prioridade)
      );

      if (encontrado) {
        principal = encontrado;
        break;
      }
    }

    // CC = todos menos principal
    const cc = emails
      .filter(email => email !== principal)
      .join(",");

    resultado.push([
      principal,
      cc,
      dominio,
      "CORPORATIVO",
    ]);
  }

  // ORGANIZAR PESSOAIS
  pessoais.forEach(email => {

    resultado.push([
      email,
      "",
      email.split("@")[1],
      "PESSOAL"
    ]);
  });

  // ESCREVER TUDO
  abaDestino
    .getRange(1, 1, resultado.length, resultado[0].length)
    .setValues(resultado);

  Logger.log("Organização concluída!");
}