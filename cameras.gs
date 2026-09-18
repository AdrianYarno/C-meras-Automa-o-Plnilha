/** ============================================================================
 *  QUEO NETICS — MONITORAMENTO PREVENTIVO DE CÂMERAS (multi-cliente)
 *  ----------------------------------------------------------------------
 *  Como usar (resumo — veja o GUIA_DE_IMPLANTACAO.md para o passo a passo):
 *   1) Cole este arquivo inteiro em Extensões > Apps Script (substitua tudo).
 *   2) Volte à planilha, recarregue a página.
 *   3) Menu "⚙️ Queo Netics" > "Configuração inicial (rodar 1x)".
 *   4) No dia a dia: "Nova aba de importação" > colar CSV > "Formatar aba atual".
 *   5) Quando quiser atualizar o painel: "Atualizar Base Unificada" e depois
 *      "Recriar Dashboard" (ou só mexer nos filtros do Dashboard, que já
 *      recalcula sozinho a partir da Base_Unificada).
 *
 *  ----------------------------------------------------------------------
 *  CORREÇÃO (2026-09-11): a remoção de "linhas de lixo" no topo da aba
 *  colada usava posições fixas (linhas 2 e 3), assumindo que o export
 *  sempre trazia 2 linhas de lixo ali. Quando o export não tem essas
 *  linhas, isso apagava placas reais. Agora a remoção verifica o
 *  conteúdo da célula de Placa (via regex) antes de decidir apagar.
 * ========================================================================= */

// ---------------------------------------------------------------------------
// CONFIGURAÇÃO GERAL — mude aqui se algo estrutural mudar no futuro
// ---------------------------------------------------------------------------
const NOME_ABA_CONFIG    = "Config";
const NOME_ABA_RESUMO    = "Resumo_Diario";
const NOME_ABA_BASE      = "Base_Unificada";
const NOME_ABA_DASHBOARD = "Dashboard";

const LOGO_URL = "https://www.queonetics.com/wp-content/themes/space-odyssey/assets/images/brand-white.png";

// Paleta institucional
const COR_PRIMARIA   = "#2e4a62";
const COR_SECUNDARIA = "#3a5d7c";
const COR_BORDA      = "#d1d5db";
const COR_TEXTO_MUTED = "#5f6a7a";

// Clientes padrão usados apenas na primeira criação da aba Config.
// Depois disso, edite direto na aba Config — o script sempre lê de lá.
const CLIENTES_PADRAO = ["Andina", "Uberlândia", "Urbam", "Solar"];

// Posição das colunas no CSV BRUTO exportado pelo sistema de origem
// (0 = coluna A). Se o layout do export mudar de novo, ajuste só aqui —
// não precisa mexer no resto do script.
const RAW_COL = {
  PLACA: 0,
  CLIENTE: 1,
  ORGANIZACAO: 2,
  STATUS: 3,          // genérico ("Problema"), não é usado no resultado final
  PROBLEMAS: 4,
  CHAMADOS_60D: 5,     // não é usado no resultado final
  NUMERO_CHAMADO: 6    // não é mais usado — "Chamado Aberto" agora é sempre preenchido manualmente
};

// Regex de validação de placa (usada para detectar linhas de lixo no topo
// da aba colada, em vez de posições fixas). Cobre placa antiga (ABC1234)
// e Mercosul (ABC1D23).
const PADRAO_PLACA = /^[A-Z]{3}\d[A-Z0-9]\d{2}$/i;

// Estados possíveis da coluna "Estado" + cores do dropdown
const MAPEAMENTO_CORES = [
  { valor: "Sem transmissão (oficina)",          bg: "#1a73e8", fg: "#ffffff" },
  { valor: "Sem transmissão (em rota)",          bg: "#673ab7", fg: "#ffffff" },
  { valor: "Sem transmissão (CD)",               bg: "#137333", fg: "#ffffff" },
  { valor: "Salto de Posição (GPS)",             bg: "#c2e7ff", fg: "#000000" },
  { valor: "Falso Positivo",                     bg: "#c5221f", fg: "#ffffff" },
  { valor: "Obstruída",                          bg: "#00ff00", fg: "#000000" },
  { valor: "Desfixada",                          bg: "#795548", fg: "#ffffff" },
  { valor: "Alto consumo de dados",              bg: "#fef7e0", fg: "#000000" },
  { valor: "Erro SMARTSPEAKER",                  bg: "#e6f4ea", fg: "#000000" },
  { valor: "Intervenção Física",                 bg: "#ff6d01", fg: "#ffffff" },
  { valor: "Câmera mal posicionada",             bg: "#3d3d3d", fg: "#ffffff" },
  { valor: "Captura de dados CAN incorreto",     bg: "#0040a0", fg: "#ffffff" },
  { valor: "SD removido",                        bg: "#11734b", fg: "#ffffff" }
];

// ---------------------------------------------------------------------------
// MENU
// ---------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi().createMenu("⚙️ Queo Netics")
    .addItem("Configuração inicial (rodar 1x)", "iniciarQueoNetics")
    .addSeparator()
    .addItem("Nova aba de importação (cliente/dia)", "criarAbaImportacao")
    .addItem("Formatar aba atual (aplicar padrão)", "formatarAbaAtual")
    .addSeparator()
    .addItem("Atualizar Base Unificada", "atualizarBaseUnificada")
    .addItem("Recriar Dashboard", "criarDashboard")
    .addToUi();
}

// ---------------------------------------------------------------------------
// CONFIGURAÇÃO INICIAL (rodar uma vez, ou de novo se quiser resetar o layout)
// ---------------------------------------------------------------------------
function iniciarQueoNetics() {
  criarOuResetarConfig();
  criarResumoDiarioSeNaoExistir();
  var ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(NOME_ABA_BASE)) ss.insertSheet(NOME_ABA_BASE);
  atualizarBaseUnificada();
  criarDashboard();
  SpreadsheetApp.getUi().alert(
    "Estrutura criada!\n\nAbas: Config, Resumo_Diario, Base_Unificada, Dashboard.\n" +
    "Suas abas de clientes (ex: Solar 27/08) continuam funcionando normalmente."
  );
}

function criarOuResetarConfig() {
  var ss = SpreadsheetApp.getActive();
  var aba = ss.getSheetByName(NOME_ABA_CONFIG);
  if (!aba) aba = ss.insertSheet(NOME_ABA_CONFIG);
  else return; // já existe: não sobrescreve a lista de clientes que você já editou
  aba.getRange("A1").setValue("Clientes")
     .setFontWeight("bold").setBackground(COR_SECUNDARIA).setFontColor("#ffffff");
  aba.getRange(2, 1, CLIENTES_PADRAO.length, 1)
     .setValues(CLIENTES_PADRAO.map(function (c) { return [c]; }));
  aba.setColumnWidth(1, 200);
  aba.getRange("C2").setValue(
    "← Adicione uma linha aqui sempre que entrar um cliente novo. " +
    "Não é preciso mexer em nenhum script."
  ).setFontColor(COR_TEXTO_MUTED).setFontStyle("italic");
}

function criarResumoDiarioSeNaoExistir() {
  var ss = SpreadsheetApp.getActive();
  var aba = ss.getSheetByName(NOME_ABA_RESUMO);
  if (!aba) aba = ss.insertSheet(NOME_ABA_RESUMO);
  if (aba.getLastRow() > 0) return; // já tem dados, não mexe
  var cab = ["Data", "Cliente", "Quant. de Eventos", "À Analisar", "Problemas Identificados"];
  aba.getRange(1, 1, 1, cab.length).setValues([cab])
     .setFontWeight("bold").setBackground(COR_SECUNDARIA).setFontColor("#ffffff");
  aba.setFrozenRows(1);
  aba.setColumnWidth(1, 100);
  aba.setColumnWidth(2, 140);
  aba.getRange("G1").setValue(
    "Preencha uma linha por dia/cliente aqui (esses 3 números vêm do seu sistema de câmeras, " +
    "não são calculados pela planilha). Isso alimenta os KPIs e o gráfico de eventos do Dashboard."
  ).setFontColor(COR_TEXTO_MUTED).setFontStyle("italic");
}

// ---------------------------------------------------------------------------
// HELPERS DE CLIENTE / NOME DE ABA
// ---------------------------------------------------------------------------
function getListaClientes() {
  var ss = SpreadsheetApp.getActive();
  var aba = ss.getSheetByName(NOME_ABA_CONFIG);
  if (!aba) return CLIENTES_PADRAO.slice();
  var n = Math.max(aba.getLastRow() - 1, 0);
  if (n === 0) return CLIENTES_PADRAO.slice();
  var valores = aba.getRange(2, 1, n, 1).getValues();
  var lista = valores.map(function (r) { return (r[0] || "").toString().trim(); })
                      .filter(function (v) { return v !== ""; });
  return lista.length ? lista : CLIENTES_PADRAO.slice();
}

// Reconhece se um nome de aba é do tipo "{Cliente} dd/MM"
function parseAbaCliente(nomeAba, listaClientes) {
  var ordenada = listaClientes.slice().sort(function (a, b) { return b.length - a.length; });
  for (var i = 0; i < ordenada.length; i++) {
    var cli = ordenada[i];
    if (nomeAba.indexOf(cli + " ") === 0) {
      var resto = nomeAba.substring(cli.length + 1).trim();
      if (/^\d{2}\/\d{2}/.test(resto)) {
        return { cliente: cli, ddMM: resto.substring(0, 5) };
      }
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// 1) CRIAR NOVA ABA DE IMPORTAÇÃO
// ---------------------------------------------------------------------------
function criarAbaImportacao() {
  var ui = SpreadsheetApp.getUi();
  var listaClientes = getListaClientes();

  var respCliente = ui.prompt(
    "Nova aba de importação",
    "Digite o cliente exatamente como aparece na aba Config:\n" + listaClientes.join(", "),
    ui.ButtonSet.OK_CANCEL
  );
  if (respCliente.getSelectedButton() !== ui.Button.OK) return;
  var cliente = respCliente.getResponseText().trim();
  if (listaClientes.indexOf(cliente) === -1) {
    ui.alert("Cliente não encontrado na aba Config. Adicione-o lá primeiro, ou confira a grafia.");
    return;
  }

  var respData = ui.prompt("Data da análise", "Formato dd/mm (ex: 27/08):", ui.ButtonSet.OK_CANCEL);
  if (respData.getSelectedButton() !== ui.Button.OK) return;
  var ddMM = respData.getResponseText().trim();
  if (!/^\d{2}\/\d{2}$/.test(ddMM)) { ui.alert("Data inválida. Use o formato dd/mm."); return; }

  var nomeAba = cliente + " " + ddMM;
  var ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(nomeAba)) { ui.alert("Já existe uma aba \"" + nomeAba + "\"."); return; }

  var novaAba = ss.insertSheet(nomeAba);
  ss.setActiveSheet(novaAba);
  ui.alert(
    "Aba \"" + nomeAba + "\" criada.\n\n" +
    "Agora cole o CSV bruto a partir da célula A1 e rode \"Formatar aba atual\"."
  );
}

// ---------------------------------------------------------------------------
// 2) FORMATAR ABA ATUAL
// ---------------------------------------------------------------------------
function formatarAbaAtual() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var listaClientes = getListaClientes();
  var info = parseAbaCliente(sheet.getName(), listaClientes);
  var cliente = info ? info.cliente : perguntarCliente(listaClientes);
  if (!cliente) return;

  // 1. Remove linhas de lixo logo após o cabeçalho do CSV colado
  //    (verifica o conteúdo da célula de Placa em vez de posições fixas,
  //    para não apagar placas reais quando o export não tiver linhas de lixo)
  removerLinhasLixoTopo(sheet);

  var dadosBrutos = sheet.getDataRange().getValues();
  if (dadosBrutos.length < 2) { ui.alert("Não encontrei dados colados nesta aba."); return; }
  var linhasDados = dadosBrutos.slice(1);

  var linhasLimpas = linhasDados
    .filter(function (l) { return l[RAW_COL.PLACA] && l[RAW_COL.PLACA].toString().trim() !== ""; })
    .map(function (l) {
      return [
        l[RAW_COL.PLACA] || "",
        cliente,
        l[RAW_COL.ORGANIZACAO] || "",
        l[RAW_COL.PROBLEMAS] || "",
        "",                                     // Estado — preenchido manualmente
        "",                                     // Observação — preenchido manualmente
        "",                                     // Chamado Aberto — preenchido manualmente (nunca vem do CSV)
        ""                                      // Data - Abertura — auto-preenchida no onEdit
      ];
    });

  sheet.clearContents();
  sheet.clearFormats();

  // Faixa de marca (linhas 1-2)
  sheet.setRowHeight(1, 45);
  sheet.setRowHeight(2, 45);
  sheet.getRange("A1:H2").setBackground(COR_PRIMARIA);
  sheet.getRange("A1").setValue('=IMAGE("' + LOGO_URL + '"; 2)');
  sheet.getRange("D1").setValue("Monitoramento preventivo — " + cliente)
       .setFontSize(20).setFontWeight("bold").setFontColor("#ffffff")
       .setVerticalAlignment("bottom").setHorizontalAlignment("center");

  // Cabeçalho de dados (linha 3)
  var colunas = [["Placa", "Cliente", "Organização", "Problemas", "Estado", "Observação", "Chamado Aberto", "Data - Abertura"]];
  sheet.getRange(3, 1, 1, 8).setValues(colunas)
       .setBackground(COR_SECUNDARIA).setFontColor("#ffffff").setFontWeight("bold")
       .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(3, 32);

  // Corpo
  if (linhasLimpas.length > 0) {
    sheet.getRange(4, 1, linhasLimpas.length, 8).setValues(linhasLimpas);
  }
  var totalLinhas = 3 + linhasLimpas.length;

  sheet.getRange(3, 1, Math.max(linhasLimpas.length + 1, 1), 8)
       .setBorder(true, true, true, true, true, true, COR_BORDA, SpreadsheetApp.BorderStyle.SOLID);

  aplicarDropdownEstado(sheet, 4, linhasLimpas.length);
  aplicarLargurasColunas(sheet);
  sheet.setFrozenRows(3);

  // Remove sobras de linhas/colunas
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 8) sheet.deleteColumns(9, maxCols - 8);
  var maxRows = sheet.getMaxRows();
  if (maxRows > totalLinhas) sheet.deleteRows(totalLinhas + 1, maxRows - totalLinhas);

  sheet.getRange(4, 1, Math.max(linhasLimpas.length, 1), 1).setHorizontalAlignment("center");
  sheet.getRange(4, 7, Math.max(linhasLimpas.length, 1), 2).setHorizontalAlignment("center");

  SpreadsheetApp.getActive().toast(
    "Aba \"" + sheet.getName() + "\" formatada — cliente: " + cliente, "Queo Netics", 5
  );
}

// Remove linhas de lixo logo abaixo do cabeçalho do CSV colado (linhas 2 e 3),
// mas só quando a célula de Placa daquela linha NÃO parecer uma placa real.
// Isso evita apagar dados de verdade quando o export não traz linhas de lixo.
function removerLinhasLixoTopo(sheet) {
  var limiteChecagem = Math.min(sheet.getMaxRows(), 4); // só olha as linhas 2 e 3
  for (var r = limiteChecagem; r >= 2; r--) {
    var valor = sheet.getRange(r, RAW_COL.PLACA + 1).getValue();
    var texto = (valor || "").toString().trim();
    var pareceLixo = texto === "" || !PADRAO_PLACA.test(texto);
    if (pareceLixo) {
      sheet.deleteRow(r);
    }
  }
}

function perguntarCliente(listaClientes) {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    "Não consegui identificar o cliente pelo nome da aba",
    "Digite o cliente exatamente como na aba Config:\n" + listaClientes.join(", "),
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return null;
  var cliente = resp.getResponseText().trim();
  if (listaClientes.indexOf(cliente) === -1) { ui.alert("Cliente não encontrado na Config."); return null; }
  return cliente;
}

function aplicarDropdownEstado(sheet, startRow, count) {
  if (count <= 0) return;
  var listaOpcoes = MAPEAMENTO_CORES.map(function (item) { return item.valor; });
  var regraValidacao = SpreadsheetApp.newDataValidation()
    .requireValueInList(listaOpcoes)
    .setAllowInvalid(false);
  if (typeof regraValidacao.setColors === "function") {
    regraValidacao.setColors(MAPEAMENTO_CORES.map(function (item) {
      return { value: item.valor, background: item.bg, foreground: item.fg };
    }));
  }
  var intervaloDropdown = sheet.getRange(startRow, 5, count, 1);
  intervaloDropdown.setDataValidation(regraValidacao.build());

  var regrasCondicionais = MAPEAMENTO_CORES.map(function (item) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(item.valor)
      .setBackground(item.bg)
      .setFontColor(item.fg)
      .setRanges([intervaloDropdown])
      .build();
  });
  sheet.setConditionalFormatRules(regrasCondicionais);
}

function aplicarLargurasColunas(sheet) {
  sheet.setColumnWidth(1, 160); // Placa
  sheet.setColumnWidth(2, 140); // Cliente
  sheet.setColumnWidth(3, 240); // Organização
  sheet.setColumnWidth(4, 420); // Problemas
  sheet.setColumnWidth(5, 230); // Estado
  sheet.setColumnWidth(6, 200); // Observação
  sheet.setColumnWidth(7, 150); // Chamado Aberto
  sheet.setColumnWidth(8, 130); // Data - Abertura
}

// ---------------------------------------------------------------------------
// AUTO-PREENCHER "Data - Abertura" QUANDO UM CHAMADO É MARCADO
// ---------------------------------------------------------------------------
function onEdit(e) {
  try {
    var sheet = e.range.getSheet();
    var listaClientes = getListaClientes();
    if (!parseAbaCliente(sheet.getName(), listaClientes)) return;
    if (e.range.getRow() < 4 || e.range.getColumn() !== 7) return; // coluna G = Chamado Aberto
    var valor = e.range.getValue();
    var celData = sheet.getRange(e.range.getRow(), 8); // H = Data - Abertura
    if (valor && valor.toString().trim() !== "" && celData.getValue() === "") {
      celData.setValue(new Date()).setNumberFormat("dd/mm/yyyy");
    }
  } catch (err) {
    // ignorar erros não impeditivos
  }
}

// ---------------------------------------------------------------------------
// 3) ATUALIZAR BASE UNIFICADA
// ---------------------------------------------------------------------------
function atualizarBaseUnificada() {
  var ss = SpreadsheetApp.getActive();
  var listaClientes = getListaClientes();
  var baseSheet = ss.getSheetByName(NOME_ABA_BASE) || ss.insertSheet(NOME_ABA_BASE);
  baseSheet.clear();

  var cab = ["Data", "Cliente", "Placa", "Organização", "Problemas", "Estado", "Observação", "Chamado Aberto"];
  baseSheet.getRange(1, 1, 1, cab.length).setValues([cab])
           .setFontWeight("bold").setBackground(COR_SECUNDARIA).setFontColor("#ffffff");
  baseSheet.setFrozenRows(1);

  var linhasFinal = [];
  var anoAtual = new Date().getFullYear();

  ss.getSheets().forEach(function (aba) {
    var info = parseAbaCliente(aba.getName(), listaClientes);
    if (!info) return;
    var partes = info.ddMM.split("/");
    var dataAba = new Date(anoAtual, parseInt(partes[1], 10) - 1, parseInt(partes[0], 10));
    var totalLinhasAba = Math.max(aba.getLastRow() - 3, 0);
    if (totalLinhasAba === 0) return;
    var dados = aba.getRange(4, 1, totalLinhasAba, 8).getValues();
    dados.forEach(function (l) {
      if (!l[0]) return; // sem placa
      linhasFinal.push([dataAba, info.cliente, l[0], l[2], l[3], l[4], l[5], l[6]]);
    });
  });

  if (linhasFinal.length > 0) {
    baseSheet.getRange(2, 1, linhasFinal.length, cab.length).setValues(linhasFinal);
    baseSheet.getRange(2, 1, linhasFinal.length, 1).setNumberFormat("dd/mm/yyyy");
  }
  baseSheet.autoResizeColumns(1, cab.length);

  SpreadsheetApp.getActive().toast(linhasFinal.length + " linhas consolidadas.", "Base Unificada", 5);
}

// ---------------------------------------------------------------------------
// FUNÇÕES DE CONSULTA (usadas pelo Dashboard e disponíveis como fórmulas)
// ---------------------------------------------------------------------------
function normalizarDatas(intervaloDeDatas) {
  if (!Array.isArray(intervaloDeDatas)) intervaloDeDatas = [[intervaloDeDatas]];
  var datasAlvo = {};
  intervaloDeDatas.forEach(function (linha) {
    linha.forEach(function (cel) {
      if (cel instanceof Date && !isNaN(cel.getTime())) {
        datasAlvo[Utilities.formatDate(cel, "UTC", "dd/MM/yyyy")] = true;
      }
    });
  });
  return datasAlvo;
}

function CONSOLIDAR_CHAMADOS_INTERVALO(intervaloDeDatas, clienteFiltro) {
  var datas = normalizarDatas(intervaloDeDatas);
  if (Object.keys(datas).length === 0) return [["Nenhuma data válida", 0]];
  var base = SpreadsheetApp.getActive().getSheetByName(NOME_ABA_BASE);
  if (!base) return [["Rode 'Atualizar Base Unificada' primeiro", 0]];
  var dados = base.getDataRange().getValues();
  var consolidado = {};
  for (var i = 1; i < dados.length; i++) {
    var l = dados[i];
    if (!(l[0] instanceof Date)) continue;
    if (!datas[Utilities.formatDate(l[0], "UTC", "dd/MM/yyyy")]) continue;
    if (clienteFiltro && clienteFiltro !== "" && clienteFiltro !== "(Todos)" && l[1] !== clienteFiltro) continue;
    var chamado = (l[7] || "").toString().trim();
    if (!chamado || chamado === "-" || chamado === "0" ||
        chamado.toLowerCase() === "null" || chamado.toLowerCase() === "false") continue;
    var org = l[3] || "(sem organização)";
    consolidado[org] = (consolidado[org] || 0) + 1;
  }
  var resultado = Object.keys(consolidado).map(function (k) { return [k, consolidado[k]]; });
  resultado.sort(function (a, b) { return b[1] - a[1]; });
  return resultado.length ? resultado : [["Nenhum chamado aberto no período", 0]];
}

function DETALHAR_CHAMADOS_INTERVALO(intervaloDeDatas, clienteFiltro) {
  var datas = normalizarDatas(intervaloDeDatas);
  if (Object.keys(datas).length === 0) return [["Nenhuma data válida", "", "", "", ""]];
  var base = SpreadsheetApp.getActive().getSheetByName(NOME_ABA_BASE);
  if (!base) return [["Rode 'Atualizar Base Unificada' primeiro", "", "", "", ""]];
  var dados = base.getDataRange().getValues();
  var linhas = [];
  for (var i = 1; i < dados.length; i++) {
    var l = dados[i];
    if (!(l[0] instanceof Date)) continue;
    var dataFmt = Utilities.formatDate(l[0], "UTC", "dd/MM/yyyy");
    if (!datas[dataFmt]) continue;
    if (clienteFiltro && clienteFiltro !== "" && clienteFiltro !== "(Todos)" && l[1] !== clienteFiltro) continue;
    var chamado = (l[7] || "").toString().trim();
    if (!chamado || chamado === "-" || chamado === "0" ||
        chamado.toLowerCase() === "null" || chamado.toLowerCase() === "false") continue;
    linhas.push([dataFmt, l[1], l[3], l[5], chamado]);
  }
  if (linhas.length === 0) return [["Nenhum chamado encontrado no período", "", "", "", ""]];
  linhas.sort(function (a, b) { return a[0] === b[0] ? (a[2] < b[2] ? -1 : 1) : (a[0] < b[0] ? -1 : 1); });
  linhas.unshift(["Data", "Cliente", "Organização", "Estado", "Chamado"]);
  return linhas;
}

function RESUMO_FILTRADO(cliente, dataIni, dataFim) {
  var aba = SpreadsheetApp.getActive().getSheetByName(NOME_ABA_RESUMO);
  if (!aba || aba.getLastRow() < 2) return [["Sem dados em Resumo_Diario", "", "", ""]];
  var dados = aba.getDataRange().getValues();
  var linhas = [];
  
  var dIni = (dataIni instanceof Date) ? dataIni : new Date(dataIni);
  var dFim = (dataFim instanceof Date) ? dataFim : new Date(dataFim);

  for (var i = 1; i < dados.length; i++) {
    var l = dados[i];
    if (!(l[0] instanceof Date)) continue;
    if (cliente && cliente !== "(Todos)" && l[1] !== cliente) continue;
    if (!isNaN(dIni.getTime()) && l[0] < dIni) continue;
    if (!isNaN(dFim.getTime()) && l[0] > dFim) continue;
    linhas.push([Utilities.formatDate(l[0], "America/Sao_Paulo", "dd/MM"), Number(l[2]) || 0, Number(l[3]) || 0, Number(l[4]) || 0]);
  }
  linhas.sort(function (a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0; });
  linhas.unshift(["Data", "Quant. de Eventos", "À Analisar", "Problemas Identificados"]);
  return linhas.length > 1 ? linhas : [["Sem dados no período selecionado", 0, 0, 0]];
}

function KPI_RESUMO(campo, cliente, dataIni, dataFim) {
  var mapa = { eventos: 2, analisar: 3, problemas: 4 };
  var idx = mapa[campo];
  if (!idx) return 0;
  var aba = SpreadsheetApp.getActive().getSheetByName(NOME_ABA_RESUMO);
  if (!aba || aba.getLastRow() < 2) return 0;
  var dados = aba.getDataRange().getValues();
  var total = 0;

  var dIni = (dataIni instanceof Date) ? dataIni : new Date(dataIni);
  var dFim = (dataFim instanceof Date) ? dataFim : new Date(dataFim);

  for (var i = 1; i < dados.length; i++) {
    var l = dados[i];
    if (!(l[0] instanceof Date)) continue;
    if (cliente && cliente !== "(Todos)" && l[1] !== cliente) continue;
    if (!isNaN(dIni.getTime()) && l[0] < dIni) continue;
    if (!isNaN(dFim.getTime()) && l[0] > dFim) continue;
    total += Number(l[idx]) || 0;
  }
  return total;
}

function CHAMADOS_POR_ORGANIZACAO_PERIODO(cliente, dataIni, dataFim) {
  var base = SpreadsheetApp.getActive().getSheetByName(NOME_ABA_BASE);
  if (!base) return [["Rode 'Atualizar Base Unificada' primeiro", ""]];
  var dados = base.getDataRange().getValues();
  var consolidado = {};

  var dIni = (dataIni instanceof Date) ? dataIni : new Date(dataIni);
  var dFim = (dataFim instanceof Date) ? dataFim : new Date(dataFim);

  for (var i = 1; i < dados.length; i++) {
    var l = dados[i];
    if (!(l[0] instanceof Date)) continue;
    if (cliente && cliente !== "(Todos)" && l[1] !== cliente) continue;
    if (!isNaN(dIni.getTime()) && l[0] < dIni) continue;
    if (!isNaN(dFim.getTime()) && l[0] > dFim) continue;
    var chamado = (l[7] || "").toString().trim();
    if (!chamado || chamado === "-" || chamado === "0" ||
        chamado.toLowerCase() === "null" || chamado.toLowerCase() === "false") continue;
    var org = l[3] || "(sem organização)";
    consolidado[org] = (consolidado[org] || 0) + 1;
  }
  var resultado = Object.keys(consolidado).map(function (k) { return [k, consolidado[k]]; });
  resultado.sort(function (a, b) { return b[1] - a[1]; });
  resultado.unshift(["Organização", "Chamados"]);
  return resultado.length > 1 ? resultado : [["Nenhum chamado no período selecionado", 0]];
}

function TOTAL_CHAMADOS_PERIODO(cliente, dataIni, dataFim) {
  var tabela = CHAMADOS_POR_ORGANIZACAO_PERIODO(cliente, dataIni, dataFim);
  var total = 0;
  for (var i = 1; i < tabela.length; i++) total += Number(tabela[i][1]) || 0;
  return total;
}

// ---------------------------------------------------------------------------
// 4) DASHBOARD CLEAN — TOTALMENTE EM GRÁFICOS E TABELAS ESTRUTURADAS (CORRIGIDO)
// ---------------------------------------------------------------------------
function criarDashboard() {
  var ss = SpreadsheetApp.getActive();
  
  // 1. Preparar aba oculta de suporte para os gráficos (evita poluir o fundo)
  var NOME_ABA_DADOS_DASH = "_DadosDash";
  var abaDados = ss.getSheetByName(NOME_ABA_DADOS_DASH);
  if (abaDados) ss.deleteSheet(abaDados);
  abaDados = ss.insertSheet(NOME_ABA_DADOS_DASH);
  abaDados.hideSheet(); // oculta a aba técnica

  // 2. Reiniciar aba Dashboard
  var antiga = ss.getSheetByName(NOME_ABA_DASHBOARD);
  if (antiga) ss.deleteSheet(antiga);
  var dash = ss.insertSheet(NOME_ABA_DASHBOARD, 0);

  // --- CORREÇÃO DO ERRO AQUI ---
  dash.setHiddenGridlines(true);

  var listaClientes = getListaClientes();

  // --- CORES INSTITUCIONAIS ---
  var COR_HEADER = "#1E293B";       // Azul Slate
  var COR_CARD_BG = "#F8FAFC";      // Fundo cinza ultraleve
  var COR_TEXTO_MUTED = "#64748B"; // Cinza rótulo
  var COR_DESTAQUE = "#0F172A";    // Texto escuro
  var COR_BORDA_TAB = "#E2E8F0";   // Borda leve de tabelas

  // --- A. CABEÇALHO ---
  dash.setRowHeight(1, 55);
  dash.getRange("A1:M1").setBackground(COR_HEADER);
  dash.getRange("A1").setValue('=IMAGE("' + LOGO_URL + '"; 2)');
  dash.getRange("C1").setValue("MONITORAMENTO PREVENTIVO — DASHBOARD EXECUTIVO")
      .setFontSize(15).setFontWeight("bold").setFontColor("#FFFFFF").setVerticalAlignment("middle");

  // --- B. BARRA DE FILTROS ---
  dash.setRowHeight(3, 35);
  dash.getRange("A3:M3").setBackground("#F1F5F9").setVerticalAlignment("middle");
  
  dash.getRange("A3").setValue("  Cliente:").setFontWeight("bold").setFontColor(COR_TEXTO_MUTED);
  var opcoesCliente = ["(Todos)"].concat(listaClientes);
  var regraCliente = SpreadsheetApp.newDataValidation().requireValueInList(opcoesCliente).setAllowInvalid(false).build();
  dash.getRange("B3").setDataValidation(regraCliente).setValue("(Todos)")
      .setBackground("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");

  dash.getRange("D3").setValue("Período de:").setFontWeight("bold").setFontColor(COR_TEXTO_MUTED);
  dash.getRange("E3").setValue(new Date(new Date().getFullYear(), new Date().getMonth(), 1))
      .setNumberFormat("dd/mm/yyyy").setBackground("#FFFFFF").setHorizontalAlignment("center");
  
  dash.getRange("F3").setValue("até:").setFontWeight("bold").setFontColor(COR_TEXTO_MUTED);
  dash.getRange("G3").setValue(new Date()).setNumberFormat("dd/mm/yyyy").setBackground("#FFFFFF").setHorizontalAlignment("center");

  dash.setFrozenRows(3);

  // --- C. CARDS DE KPIS ---
  var kpis = [
    { range: "A5:C6", labelCell: "A5", valCell: "A6", label: "TOTAL DE EVENTOS", formula: '=KPI_RESUMO("eventos"; $B$3; $E$3; $G$3)' },
    { range: "D5:F6", labelCell: "D5", valCell: "D6", label: "À ANALISAR", formula: '=KPI_RESUMO("analisar"; $B$3; $E$3; $G$3)' },
    { range: "G5:I6", labelCell: "G5", valCell: "G6", label: "PROBLEMAS IDENTIFICADOS", formula: '=KPI_RESUMO("problemas"; $B$3; $E$3; $G$3)' },
    { range: "J5:L6", labelCell: "J5", valCell: "J6", label: "CHAMADOS ABERTOS", formula: '=TOTAL_CHAMADOS_PERIODO($B$3; $E$3; $G$3)' }
  ];

  kpis.forEach(function (k) {
    dash.getRange(k.range).setBackground(COR_CARD_BG)
        .setBorder(true, true, true, true, false, false, COR_BORDA, SpreadsheetApp.BorderStyle.SOLID);
    
    dash.getRange(k.labelCell).setValue(k.label)
        .setFontWeight("bold").setFontColor(COR_TEXTO_MUTED).setFontSize(9).setHorizontalAlignment("center");
    
    dash.getRange(k.valCell).setFormula(k.formula)
        .setFontSize(22).setFontWeight("bold").setFontColor(COR_DESTAQUE).setHorizontalAlignment("center");
  });

  // --- D. ALIMENTAR ABA SUPORTE TÉCNICO (DADOS DOS GRÁFICOS) ---
  abaDados.getRange("A1").setFormula('=' + NOME_ABA_DASHBOARD + '!B3');
  abaDados.getRange("B1").setFormula('=' + NOME_ABA_DASHBOARD + '!E3');
  abaDados.getRange("C1").setFormula('=' + NOME_ABA_DASHBOARD + '!G3');
  
  abaDados.getRange("A3").setFormula('=RESUMO_FILTRADO(A1; B1; C1)');
  abaDados.getRange("G3").setFormula('=CHAMADOS_POR_ORGANIZACAO_PERIODO(A1; B1; C1)');

  SpreadsheetApp.flush();

  // --- E. GRÁFICOS DO DASHBOARD ---
  
  // Gráfico 1: Evolução Diária (Coluna)
  var chart1 = dash.newChart()
    .asColumnChart()
    .addRange(abaDados.getRange("A3:D300"))
    .setNumHeaders(1)
    .setPosition(8, 1, 0, 0)
    .setOption("title", "EVOLUÇÃO DIÁRIA DE EVENTOS E PROBLEMAS")
    .setOption("titleTextStyle", { color: COR_HEADER, fontSize: 11, bold: true })
    .setOption("legend", { position: "top" })
    .setOption("colors", [COR_SECUNDARIA, "#E11D48", "#F59E0B"])
    .setOption("width", 540).setOption("height", 290)
    .build();
  dash.insertChart(chart1);

  // Gráfico 2: Chamados por Organização (Barra)
  var chart2 = dash.newChart()
    .asBarChart()
    .addRange(abaDados.getRange("G3:H50"))
    .setNumHeaders(1)
    .setPosition(8, 7, 0, 0)
    .setOption("title", "VERIFICAÇÃO DE ABERTURA DE CHAMADOS POR ORGANIZAÇÃO")
    .setOption("titleTextStyle", { color: COR_HEADER, fontSize: 11, bold: true })
    .setOption("colors", [COR_PRIMARIA])
    .setOption("legend", { position: "none" })
    .setOption("width", 520).setOption("height", 290)
    .build();
  dash.insertChart(chart2);

  // --- F. TABELAS DE DETALHAMENTO NO DASHBOARD (A PARTIR DA LINHA 25) ---

  // TITULO TABELA 1
  dash.getRange("A24").setValue("DETALHAMENTO DE CHAMADOS POR UNIDADE")
      .setFontWeight("bold").setFontSize(11).setFontColor(COR_HEADER);
  dash.getRange("A25").setFormula('=DETALHAR_CHAMADOS_INTERVALO(E3:G3; B3)');

  // TITULO TABELA 2
  dash.getRange("G24").setValue("CONTROLE SEMANAL DE ORGANIZAÇÕES")
      .setFontWeight("bold").setFontSize(11).setFontColor(COR_HEADER);
  dash.getRange("G25").setFormula('=CHAMADOS_POR_ORGANIZACAO_PERIODO($B$3; $E$3; $G$3)');

  // Formatação estética das colunas do Dashboard
  dash.setColumnWidth(1, 100); // Data
  dash.setColumnWidth(2, 110); // Cliente
  dash.setColumnWidth(3, 220); // Organização
  dash.setColumnWidth(4, 210); // Estado
  dash.setColumnWidth(5, 170); // Chamado

  dash.setColumnWidth(6, 20);  // Espaçador entre tabelas

  dash.setColumnWidth(7, 240); // Organização (Controle Semanal)
  dash.setColumnWidth(8, 90);  // Chamados

  SpreadsheetApp.getActive().toast("Dashboard reconstruído com sucesso!", "Queo Netics", 4);
}
