// ============================================================================
// GOOGLE APPS SCRIPT — CADASTRO DAS PARTICIPANTES (HTMLSERVICE + PLANILHA + DRIVE)
// ============================================================================
// Como instalar no Google Apps Script:
// 1. Crie uma pasta privada no Google Drive para guardar as fotos dos documentos.
// 2. Copie o ID da pasta (o código no final da URL da pasta no navegador).
// 3. Cole esse ID na constante FOLDER_ID logo abaixo.
// 4. Na sua Planilha Google, abra Extensões > Apps Script.
// 5. Cole este conteúdo no arquivo Codigo.gs.
// 6. Crie um arquivo HTML chamado "index" (menu "+" > HTML > digite: index)
//    e cole o conteúdo completo do arquivo index.html.
// 7. Clique em Implantar > Nova implantação > App da Web
//    (Executar como: "Eu" | Quem pode acessar: "Qualquer pessoa").
// ============================================================================

const FOLDER_ID = "19A7eppBBg4mzeHco0drx-I3MFjA3sS0I";

const PAGE_TITLE = "Cadastro de Participantes — PROJETO SAAD";
const SHEET_CADASTROS = "Cadastros";
const SHEET_RESUMO = "Resumo Camisas";

const CABECALHO_CADASTROS = [
  "Protocolo",
  "Data/Hora",
  "Nome Completo",
  "Telefone",
  "Data de Nascimento",
  "Idade",
  "Tamanho da Camisa",
  "CEP",
  "Rua",
  "Número",
  "Complemento",
  "Bairro",
  "Cidade",
  "UF",
  "Link Foto Frente",
  "Link Foto Verso",
  "Consentimento LGPD"
];

/**
 * Serve o arquivo index.html autocontido pelo HtmlService do Google Apps Script.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle(PAGE_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * Função chamada diretamente pelo index.html via google.script.run.salvarCadastro(dados)
 * e também reutilizada pelo doPost(e) caso seja feito envio HTTP externo.
 */
function salvarCadastro(dados) {
  const lock = LockService.getScriptLock();

  try {
    // Evita conflitos quando duas pessoas enviam o cadastro ao mesmo tempo
    lock.waitLock(30000);

    if (!dados || typeof dados !== "object") {
      return {
        ok: false,
        error: "Nenhum dado foi recebido no envio."
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaCadastros = garantirAbaCadastros(ss);
    garantirAbaResumoCamisas(ss);

    // Verifica duplicidade pelo telefone (comparando apenas os números)
    const telefoneNovoLimpo = String(dados.telefone || "").replace(/\D/g, "");
    if (!telefoneNovoLimpo) {
      return {
        ok: false,
        error: "Telefone não informado."
      };
    }

    if (telefoneJaCadastrado(abaCadastros, telefoneNovoLimpo)) {
      return {
        ok: false,
        error: "Já existe um cadastro com este telefone."
      };
    }

    if (!FOLDER_ID || FOLDER_ID === "COLE_AQUI_O_ID_DA_PASTA_DO_DRIVE") {
      return {
        ok: false,
        error: "O ID da pasta do Google Drive ainda não foi configurado no arquivo Codigo.gs."
      };
    }

    // Gera protocolo sequencial: ex. CAD-2026-0001
    const protocoloGerado = gerarProximoProtocolo(abaCadastros);
    const nomeParaArquivo = sanitizarNomeArquivo(dados.nome || "PARTICIPANTE");
    const pastaDrive = DriveApp.getFolderById(FOLDER_ID);

    // Salva as fotos na pasta privada do Google Drive (sem compartilhamento público)
    let linkFotoFrente = "";
    if (dados.fotoFrenteBase64) {
      const nomeArquivoFrente = protocoloGerado + "_" + nomeParaArquivo + "_FRENTE.jpg";
      linkFotoFrente = salvarFotoNoDrive(pastaDrive, dados.fotoFrenteBase64, nomeArquivoFrente);
    }

    let linkFotoVerso = "";
    if (dados.fotoVersoBase64) {
      const nomeArquivoVerso = protocoloGerado + "_" + nomeParaArquivo + "_VERSO.jpg";
      linkFotoVerso = salvarFotoNoDrive(pastaDrive, dados.fotoVersoBase64, nomeArquivoVerso);
    }

    const dataHoraFormatada = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone() || "America/Sao_Paulo",
      "dd/MM/yyyy HH:mm:ss"
    );

    const novaLinha = [
      protocoloGerado,
      dataHoraFormatada,
      String(dados.nome || "").trim(),
      String(dados.telefone || "").trim(),
      String(dados.dataNascimento || "").trim(),
      Number(dados.idade || 0),
      String(dados.tamanhoCamisa || "").trim(),
      String(dados.cep || "").trim(),
      String(dados.rua || "").trim(),
      String(dados.numero || "").trim(),
      String(dados.complemento || "").trim(),
      String(dados.bairro || "").trim(),
      String(dados.cidade || "").trim(),
      String(dados.uf || "").trim().toUpperCase(),
      linkFotoFrente,
      linkFotoVerso,
      dados.consentimentoLgpd ? "Sim (Autorizado)" : "Não"
    ];

    abaCadastros.appendRow(novaLinha);

    return {
      ok: true,
      protocolo: protocoloGerado
    };
  } catch (erro) {
    return {
      ok: false,
      error: "Erro ao salvar cadastro: " + (erro && erro.message ? erro.message : String(erro))
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Endpoint HTTP POST opcional para integrações externas via fetch.
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return criarRespostaJson({
        ok: false,
        error: "Nenhum dado foi recebido no envio."
      });
    }

    const dados = JSON.parse(e.postData.contents);
    const resultado = salvarCadastro(dados);
    return criarRespostaJson(resultado);
  } catch (erro) {
    return criarRespostaJson({
      ok: false,
      error: "Erro ao processar requisição: " + (erro && erro.message ? erro.message : String(erro))
    });
  }
}

function garantirAbaCadastros(ss) {
  let aba = ss.getSheetByName(SHEET_CADASTROS);
  if (!aba) {
    aba = ss.insertSheet(SHEET_CADASTROS);
    aba.appendRow(CABECALHO_CADASTROS);

    const faixaCabecalho = aba.getRange(1, 1, 1, CABECALHO_CADASTROS.length);
    faixaCabecalho
      .setFontWeight("bold")
      .setBackground("#702459")
      .setFontColor("#FFFFFF")
      .setHorizontalAlignment("center");

    aba.setFrozenRows(1);
    aba.autoResizeColumns(1, CABECALHO_CADASTROS.length);
  }
  return aba;
}

function garantirAbaResumoCamisas(ss) {
  let aba = ss.getSheetByName(SHEET_RESUMO);
  if (!aba) {
    aba = ss.insertSheet(SHEET_RESUMO);
    aba.getRange("A1:B1").setValues([["Tamanho da Camisa", "Quantidade"]]);
    aba.getRange("A1:B1")
      .setFontWeight("bold")
      .setBackground("#702459")
      .setFontColor("#FFFFFF");

    const tamanhos = ["PP", "P", "M", "G", "GG", "XG", "XGG"];
    for (let i = 0; i < tamanhos.length; i++) {
      const linha = i + 2;
      aba.getRange(linha, 1).setValue(tamanhos[i]);
      // COUNTIF é exibido automaticamente como CONT.SE nas planilhas em Português
      aba.getRange(linha, 2).setFormula("=COUNTIF(" + SHEET_CADASTROS + "!G2:G, A" + linha + ")");
    }

    const linhaTotal = tamanhos.length + 2;
    aba.getRange(linhaTotal, 1).setValue("Total Geral").setFontWeight("bold");
    aba.getRange(linhaTotal, 2).setFormula("=SUM(B2:B" + (linhaTotal - 1) + ")").setFontWeight("bold");

    aba.setFrozenRows(1);
    aba.autoResizeColumns(1, 2);
  }
  return aba;
}

function telefoneJaCadastrado(abaCadastros, telefoneNovoLimpo) {
  const ultimaLinha = abaCadastros.getLastRow();
  if (ultimaLinha <= 1) return false;

  // Coluna D (índice 4) armazena o Telefone
  const valoresTelefone = abaCadastros.getRange(2, 4, ultimaLinha - 1, 1).getValues();
  for (let i = 0; i < valoresTelefone.length; i++) {
    const existenteLimpo = String(valoresTelefone[i][0] || "").replace(/\D/g, "");
    if (existenteLimpo && existenteLimpo === telefoneNovoLimpo) {
      return true;
    }
  }
  return false;
}

function gerarProximoProtocolo(abaCadastros) {
  const anoAtual = new Date().getFullYear();
  const ultimaLinha = abaCadastros.getLastRow();
  const proximoNumero = Math.max(1, ultimaLinha); // Linha 1 é o cabeçalho
  const numeroFormatado = ("0000" + proximoNumero).slice(-4);
  return "CAD-" + anoAtual + "-" + numeroFormatado;
}

function salvarFotoNoDrive(pastaDrive, base64String, nomeArquivo) {
  const base64Limpo = String(base64String).replace(/^data:image\/[a-zA-Z0-9.+-]+;base64,/, "");
  const bytes = Utilities.base64Decode(base64Limpo);
  const blob = Utilities.newBlob(bytes, "image/jpeg", nomeArquivo);
  const arquivo = pastaDrive.createFile(blob);

  // LGPD: Não alteramos a permissão para público.
  // Apenas quem já tem acesso à pasta do projeto no Google Drive consegue abrir o link.
  return arquivo.getUrl();
}

function sanitizarNomeArquivo(nome) {
  return String(nome)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[^a-zA-Z0-9_]/g, "")
    .toUpperCase();
}

function criarRespostaJson(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}