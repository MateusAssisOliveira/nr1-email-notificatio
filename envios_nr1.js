function enviarEmails() {

  // ==============================
  // CONFIGURAÇÕES
  // ==============================

  // Limite de destinatários
  // (TO + CC)
  const LIMITE_ENVIO = 100;

  // Diretora fixa em cópia
  const CC_FIXO = "";

  // Delay entre envios
  const DELAY_MS = 1500;

  // ==============================
  // CONTADORES
  // ==============================

  let totalDestinatarios = 0;
  let erros = 0;
  let ignorados = 0;
  let jaEnviados = 0;

  const inicioExecucao = new Date();

  Logger.log("======================================");
  Logger.log("INICIANDO DISPARO DE E-MAILS");
  Logger.log(`HORÁRIO: ${inicioExecucao}`);
  Logger.log(`LIMITE DESTINATÁRIOS: ${LIMITE_ENVIO}`);
  Logger.log("======================================");

  // ==============================
  // PLANILHA
  // ==============================

  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  const aba = planilha.getSheetByName(
    "EMAILS_ORGANIZADOS"
  );

  if (!aba) {

    Logger.log(
      "ERRO: Aba EMAILS_ORGANIZADOS não encontrada."
    );

    return;
  }

  const ultimaLinha = aba.getLastRow();

  Logger.log(`ÚLTIMA LINHA: ${ultimaLinha}`);

  if (ultimaLinha < 2) {

    Logger.log("NENHUM DADO ENCONTRADO");

    return;
  }

  const dados = aba
    .getRange(2, 1, ultimaLinha - 1, 5)
    .getValues();

  Logger.log(
    `TOTAL REGISTROS CARREGADOS: ${dados.length}`
  );

  // ==============================
  // HTML E ASSUNTOS
  // ==============================

  const html = HtmlService
    .createHtmlOutputFromFile('email_nr1')
    .getContent();

  const assuntos = [
    "Nova NR-1: informações importantes para sua empresa",
    "Sua empresa está preparada para a Nova NR-1?",
    "Adequação à Nova NR-1 – veja o que mudou",
    "Informações importantes sobre a atualização da NR-1"
  ];

  // ==============================
  // STATUS PARA ATUALIZAÇÃO
  // ==============================

  const atualizacoes = [];

  // ==============================
  // LOOP PRINCIPAL
  // ==============================

  for (let i = 0; i < dados.length; i++) {

    const linha = i + 2;

    const emailPrincipal = dados[i][0];
    const ccPlanilha = dados[i][1];
    const status = dados[i][4];

    Logger.log("--------------------------------------------------");
    Logger.log(`PROCESSANDO LINHA: ${linha}`);
    Logger.log(`EMAIL PRINCIPAL: ${emailPrincipal}`);
    Logger.log(`STATUS ATUAL: ${status}`);

    // ==============================
    // E-MAIL VAZIO
    // ==============================

    if (!emailPrincipal) {

      Logger.log("IGNORADO: E-mail vazio.");

      atualizacoes.push([
        "ERRO: E-mail vazio"
      ]);

      ignorados++;

      continue;
    }

    // ==============================
    // VALIDAÇÃO DE E-MAIL
    // ==============================

    const emailValido =
      /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    if (!emailValido.test(emailPrincipal)) {

      Logger.log(
        `ERRO: E-mail inválido -> ${emailPrincipal}`
      );

      atualizacoes.push([
        "ERRO: E-mail inválido"
      ]);

      erros++;

      continue;
    }

    // ==============================
    // MONTA CC
    // ==============================

    const ccFinal = ccPlanilha
      ? `${ccPlanilha},${CC_FIXO}`
      : CC_FIXO;

    Logger.log(`CC FINAL: ${ccFinal}`);

    // ==============================
    // CONTA DESTINATÁRIOS
    // ==============================

    const quantidadeCC = ccFinal
      ? ccFinal
          .split(",")
          .map(e => e.trim())
          .filter(e => e !== "")
          .length
      : 0;

    // TO + CC
    const destinatariosAtual =
      1 + quantidadeCC;

    Logger.log(
      `DESTINATÁRIOS DESTE ENVIO: ${destinatariosAtual}`
    );

    Logger.log(
      `TOTAL DESTINATÁRIOS ACUMULADOS: ${totalDestinatarios}`
    );

    // ==============================
    // JÁ ENVIADO
    // ==============================

    if (
      status &&
      status.toString().startsWith("ENVIADO")
    ) {

      Logger.log(
        "IGNORADO: Já enviado anteriormente."
      );

      atualizacoes.push([status]);

      jaEnviados++;

      continue;
    }

    // ==============================
    // LIMITE DESTINATÁRIOS
    // ==============================

    if (
      totalDestinatarios +
      destinatariosAtual >
      LIMITE_ENVIO
    ) {

      Logger.log("======================================");
      Logger.log(
        "LIMITE DE DESTINATÁRIOS ATINGIDO"
      );
      Logger.log(
        `LIMITE CONFIGURADO: ${LIMITE_ENVIO}`
      );
      Logger.log(
        `TOTAL ACUMULADO: ${totalDestinatarios}`
      );
      Logger.log("SCRIPT ENCERRADO");
      Logger.log("======================================");

      atualizacoes.push([
        "LIMITE DE DESTINATÁRIOS"
      ]);

      break;
    }

    // ==============================
    // ASSUNTO ALEATÓRIO
    // ==============================

    const assunto = assuntos[
      Math.floor(
        Math.random() * assuntos.length
      )
    ];

    Logger.log(
      `ASSUNTO ESCOLHIDO: ${assunto}`
    );

    // ==============================
    // ENVIO
    // ==============================

    try {

      MailApp.sendEmail({
        to: emailPrincipal,
        cc: ccFinal,
        subject: assunto,
        htmlBody: html
      });

      const dataEnvio =
        new Date().toLocaleString();

      atualizacoes.push([
        `ENVIADO ${dataEnvio}`
      ]);

      totalDestinatarios +=
        destinatariosAtual;

      Logger.log(
        "STATUS: ENVIADO COM SUCESSO"
      );

      Logger.log(
        `TOTAL DESTINATÁRIOS: ${totalDestinatarios}`
      );

    } catch (erro) {

      erros++;

      let mensagemErro = erro.message;

      Logger.log("ERRO AO ENVIAR");
      Logger.log(
        `MENSAGEM ORIGINAL: ${mensagemErro}`
      );

      // ==============================
      // LIMITE GOOGLE
      // ==============================

      if (
        mensagemErro.includes(
          "Service invoked too many times"
        ) ||
        mensagemErro.includes(
          "Limit exceeded"
        )
      ) {

        mensagemErro =
          "Quota de envio excedida";

        Logger.log("======================================");
        Logger.log(
          "IDENTIFICADO: LIMITE DE ENVIO GOOGLE"
        );
        Logger.log(
          "GOOGLE BLOQUEOU NOVOS ENVIOS"
        );
        Logger.log(
          "SCRIPT INTERROMPIDO"
        );
        Logger.log(
          `TOTAL DESTINATÁRIOS: ${totalDestinatarios}`
        );
        Logger.log("======================================");

        atualizacoes.push([
          `ERRO: ${mensagemErro}`
        ]);

        // SALVA IMEDIATAMENTE
        aba
          .getRange(
            2,
            5,
            atualizacoes.length,
            1
          )
          .setValues(atualizacoes);

        return;
      }

      // ==============================
      // E-MAIL INVÁLIDO
      // ==============================

      if (
        mensagemErro.includes(
          "Invalid email"
        )
      ) {

        mensagemErro =
          "E-mail inválido";

        Logger.log(
          "IDENTIFICADO: E-MAIL INVÁLIDO"
        );
      }

      atualizacoes.push([
        `ERRO: ${mensagemErro}`
      ]);

      Logger.log(
        `FALHA AO ENVIAR PARA: ${emailPrincipal}`
      );
    }

    // ==============================
    // DELAY ANTI-SPAM
    // ==============================

    Utilities.sleep(DELAY_MS);
  }

  // ==============================
  // ATUALIZA PLANILHA
  // ==============================

  aba
    .getRange(
      2,
      5,
      atualizacoes.length,
      1
    )
    .setValues(atualizacoes);

  // ==============================
  // FINALIZAÇÃO
  // ==============================

  const fimExecucao = new Date();

  Logger.log("======================================");
  Logger.log("FINALIZADO");
  Logger.log(`INÍCIO: ${inicioExecucao}`);
  Logger.log(`FIM: ${fimExecucao}`);
  Logger.log(
    `TOTAL DESTINATÁRIOS: ${totalDestinatarios}`
  );
  Logger.log(`ERROS: ${erros}`);
  Logger.log(`IGNORADOS: ${ignorados}`);
  Logger.log(`JÁ ENVIADOS: ${jaEnviados}`);
  Logger.log("======================================");
}