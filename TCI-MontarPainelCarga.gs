/**
 * PAINEL CARGA TCI — Miscelânia (ontem × hoje) + Modems + Forms instalados
 *
 * Setup:
 *   1. Nova planilha Google em branco
 *   2. Extensões → Apps Script → cole este arquivo → Salvar
 *   3. Volte na planilha → F5 → Menu TCI Carga
 *   4. TCI Carga → Montar estrutura inicial
 *   5. Após colar BASE MICELANEAS (10h): Registrar snapshot de hoje
 *   6. TCI Carga → Atualizar painel miscelânia
 *
 * Planilhas fonte (produção):
 *   MATERIAIS — BASE MICELANEAS / ALMOX SERIALIZADA
 *   EQUIPES   — lista TT dos técnicos TCI (Página4)
 *   FORMS     — instalações do dia (Respostas ao formulário 1)
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TCI Carga')
    .addItem('Montar estrutura inicial', 'montarEstruturaInicial')
    .addItem('Corrigir CONFIG vazio', 'corrigirConfigVazio')
    .addItem('TESTE: gravar B2 (= OK script)', 'testeGravarB2')
    .addSeparator()
    .addItem('Registrar snapshot de hoje (misc)', 'registrarSnapshotMisc')
    .addItem('Atualizar painel miscelânia', 'montarPainelMiscelania')
    .addItem('Atualizar painel modems', 'montarPainelModems')
    .addSeparator()
    .addItem('Tudo (snapshot + painéis)', 'atualizarTudo')
    .addSeparator()
    .addItem('Mostrar link gerencial (Web App)', 'mostrarLinkWebApp')
    .addSeparator()
    .addItem('Ativar atualização automática diária (10h30)', 'ativarGatilhoDiario')
    .addItem('Ativar atualização frequente (a cada 15 min)', 'ativarAtualizacaoFrequente')
    .addItem('Desativar atualização automática', 'desativarGatilhoDiario')
    .addToUi();
}

/**
 * WEB APP — visão gerencial em HTML, com URL fixa.
 * Publicar: Implantar → Nova implantação → Tipo "App da Web" → Executar como "eu" → Acesso conforme necessidade.
 * A URL /exec gerada é permanente e mostra os painéis sempre atualizados ao abrir.
 */
function doGet() {
  const html = construirHtmlGerencial_();
  return HtmlService.createHtmlOutput(html)
    .setTitle('Painel Carga TCI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Mostra a URL do Web App (se já publicado) com instruções. */
function mostrarLinkWebApp() {
  const ui = SpreadsheetApp.getUi();
  let url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { url = ''; }
  if (url) {
    ui.alert('Link gerencial (Web App)',
      'Abra ou compartilhe este link — mostra os painéis sempre atualizados:\n\n' + url +
      '\n\nDica: salve nos favoritos ou envie no grupo da equipe.',
      ui.ButtonSet.OK);
  } else {
    ui.alert('Web App ainda não publicado',
      'Para gerar o link gerencial:\n\n' +
      '1. Menu Extensões → Apps Script\n' +
      '2. Botão "Implantar" → "Nova implantação"\n' +
      '3. Tipo: "App da Web"\n' +
      '4. Executar como: "Eu"\n' +
      '5. Quem tem acesso: escolha (ex.: qualquer pessoa com o link)\n' +
      '6. Implantar → copie a URL /exec\n\n' +
      'Depois volte aqui em "Mostrar link gerencial" que ele aparece pronto.',
      ui.ButtonSet.OK);
  }
}

function construirHtmlGerencial_() {
  const cfg    = lerConfig_();
  const equipe = carregarEquipeTci_(cfg);
  const tz     = Session.getScriptTimeZone();
  const hoje   = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const agora  = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');

  // ── MODEMS ──
  const modems = lerModemsAlmox_(cfg, equipe.porTt);
  const forms  = lerFormsInstalados_(cfg, equipe);
  const encerradosTotal = coletarSeriaisInstalados_(forms, null);

  const porTt = {};
  modems.forEach(function (m) {
    if (!porTt[m.tt]) porTt[m.tt] = { velocidades: {}, total: 0 };
    if (m.serial && encerradosTotal[m.serial]) return;
    porTt[m.tt].velocidades[m.velocidade] = (porTt[m.tt].velocidades[m.velocidade] || 0) + 1;
    porTt[m.tt].total += 1;
  });

  const modemRows = equipe.lista.map(function (tec) {
    const almox = porTt[tec.tt];
    const velocs = almox ? Object.keys(almox.velocidades).sort() : [];
    const r = resumirFormsTt_(forms[tec.tt], hoje, tz);
    const farol = farolEncerramento_(r.diasDesde);
    return {
      tt: tec.tt, nome: tec.nome, area: tec.area || '',
      velocHtml: velocs.map(function (v) { return esc_(v || '(sem descrição)'); }).join('<br>') || '—',
      qtdHtml:   velocs.map(function (v) { return almox.velocidades[v]; }).join('<br>') || '0',
      total: almox ? almox.total : 0,
      encHoje: r.encerradosHoje, encTotal: r.totalEncerrados,
      ultimo: farol.txt, status: farol.status, cor: farol.cor, ordem: ordemStatus_(farol.status)
    };
  }).sort(function (a, b) { return b.ordem - a.ordem || b.total - a.total; });

  let modemTbody = '';
  modemRows.forEach(function (r) {
    modemTbody +=
      '<tr><td class="tt">' + esc_(r.tt) + '</td><td>' + esc_(r.nome) + '</td><td>' + esc_(r.area) + '</td>' +
      '<td>' + r.velocHtml + '</td><td class="ctr">' + r.qtdHtml + '</td>' +
      '<td class="num"><b>' + r.total + '</b></td>' +
      '<td class="num">' + (r.encHoje || '') + '</td><td class="num">' + (r.encTotal || '') + '</td>' +
      '<td class="ctr" style="background:' + r.cor + '">' + esc_(r.ultimo) + '</td>' +
      '<td class="ctr" style="background:' + r.cor + '">' + esc_(r.status) + '</td></tr>';
  });

  // ── MISCELÂNIA ──
  let miscResumo = '';
  const ss = obterPlanilha_();
  const shHist = ss.getSheetByName('HISTORICO MISC');
  if (shHist && shHist.getLastRow() >= 2) {
    const snapshots = lerSnapshots_(shHist);
    const dataHoje  = snapshots.datas[snapshots.datas.length - 1];
    const dataOntem = snapshots.datas.length >= 2 ? snapshots.datas[snapshots.datas.length - 2] : null;

    let ths = '<th>TT</th><th>Técnico</th><th>Área</th>';
    CATS_MISC_.forEach(function (c) { ths += '<th>' + esc_(c.label) + '</th>'; });

    let rows = '';
    Object.keys(equipe.porTt).sort().forEach(function (tt) {
      const tec = equipe.porTt[tt];
      const h = somarCatsMiscTecnico_(snapshots, dataHoje,  tt);
      const o = somarCatsMiscTecnico_(snapshots, dataOntem, tt);
      if (!CATS_MISC_.some(function (c) { return h[c.key] > 0 || o[c.key] > 0; })) return;
      let tds = '<td class="tt">' + esc_(tt) + '</td><td>' + esc_(tec.nome) + '</td><td>' + esc_(tec.area || '') + '</td>';
      CATS_MISC_.forEach(function (c) {
        const vh = h[c.key] || 0, vo = o[c.key] || 0;
        const cor = !dataOntem ? '' : (vh !== vo ? '#e8f5e9' : '#ffcdd2');
        tds += '<td class="num" style="background:' + cor + '">' + vh.toLocaleString('pt-BR') + '</td>';
      });
      rows += '<tr>' + tds + '</tr>';
    });
    miscResumo =
      '<div class="card"><div class="titulo">PAINEL MISCELÂNIA TCI — CARGA ATUAL POR TÉCNICO</div>' +
      '<div class="sub">Saldo col ' + esc_(cfg.COL_SALDO_MISC) + ' · ' +
      (dataOntem ? ('comparação ' + dataOntem + ' → ' + dataHoje) : ('1º dia: ' + dataHoje)) +
      ' · 🟢 movimentou · 🔴 parado</div>' +
      '<table><thead><tr>' + ths + '</tr></thead><tbody>' + rows + '</tbody></table></div>';
  } else {
    miscResumo = '<div class="card"><div class="titulo">PAINEL MISCELÂNIA</div>' +
      '<div class="sub">Sem histórico ainda — rode "Atualizar painel miscelânia" na planilha.</div></div>';
  }

  return '<!DOCTYPE html><html lang="pt-br"><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,Helvetica,sans-serif;background:#f1f3f4;margin:0;padding:16px;}' +
    '.card{background:#fff;border:1px solid #c0c0c0;box-shadow:0 2px 8px rgba(0,0,0,.12);margin-bottom:24px;border-radius:6px;overflow:hidden;}' +
    '.titulo{background:#1565c0;color:#fff;font-weight:bold;font-size:15px;text-align:center;padding:11px;}' +
    '.sub{font-style:italic;text-align:center;color:#555;font-size:11px;padding:6px;background:#fafafa;}' +
    'table{border-collapse:collapse;width:100%;font-size:12px;}' +
    'td,th{border:1px solid #d8d8d8;padding:6px 8px;}' +
    'thead th{background:#e3f2fd;font-weight:bold;text-align:center;position:sticky;top:0;}' +
    '.tt{color:#1565c0;font-weight:bold;}.num{text-align:right;}.ctr{text-align:center;}' +
    'h1{font-size:18px;color:#333;}' +
    '</style></head><body>' +
    '<h1>📊 Painel Carga TCI <small style="font-weight:normal;color:#888;font-size:12px;">— atualizado ' + esc_(agora) + '</small></h1>' +
    '<div class="card"><div class="titulo">PAINEL MODEMS TCI — CARGA LÍQUIDA POR TT (ONT · MESH · MODEM)</div>' +
    '<div class="sub">Carga = ALMOX − todos os seriais já encerrados no Forms · 🟢 encerrou hoje · 🟡 1 dia · 🟠 2 dias · 🔴 3+ dias</div>' +
    '<table><thead><tr><th>TT</th><th>Técnico</th><th>Área</th><th>Velocidade / Descrição</th>' +
    '<th>Qtd<br>(linha)</th><th>Total<br>carga</th><th>Enc.<br>hoje</th><th>Total<br>enc.</th>' +
    '<th>Último<br>encerram.</th><th>Status</th></tr></thead><tbody>' + modemTbody + '</tbody></table></div>' +
    miscResumo +
    '<p style="font-size:11px;color:#777;">Visão gerencial · dados batem com os painéis da planilha · recarregue a página para atualizar.</p>' +
    '</body></html>';
}

/** Escapa texto para HTML. */
function esc_(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/**
 * Atualiza só os painéis (sem registrar novo snapshot).
 * Usado pelo gatilho frequente — lê FORMS e MATERIAIS atualizados e redesenha.
 */
function atualizarPaineis() {
  montarPainelMiscelania();
  montarPainelModems();
}

/**
 * Gatilho a cada 15 min: reflete mudanças do FORMS e da BASE MATERIAIS no painel.
 * Como FORMS e MATERIAIS são planilhas externas, o Apps Script não recebe onEdit
 * delas — por isso usamos tempo (a cada 15 min) para reler e atualizar sozinho.
 */
function instalarGatilhosFrequentes_() {
  desativarGatilhoDiario();
  ScriptApp.newTrigger('atualizarPaineis').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('registrarSnapshotMisc').timeBased().everyDays(1).atHour(10).nearMinute(30).create();
}

function ativarAtualizacaoFrequente() {
  instalarGatilhosFrequentes_();
  SpreadsheetApp.getUi().alert(
    'Atualização frequente ativada',
    'Os painéis MODEMS e MISCELÂNIA serão atualizados sozinhos a cada 15 minutos,\n' +
    'refletindo automaticamente o que mudar no FORMS e na BASE MATERIAIS.\n\n' +
    'O snapshot de miscelânia (comparativo ontem×hoje) continua 1x/dia às 10h30.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/** Cria gatilho diário que roda atualizarTudo automaticamente (~10h30). */
function ativarGatilhoDiario() {
  desativarGatilhoDiario();
  ScriptApp.newTrigger('atualizarTudo')
    .timeBased().everyDays(1).atHour(10).nearMinute(30).create();
  SpreadsheetApp.getUi().alert(
    'Automático ativado',
    'Os painéis serão atualizados sozinhos todo dia por volta das 10h30.\n' +
    'Garanta que a BASE MICELANEAS já esteja colada antes desse horário.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function desativarGatilhoDiario() {
  const alvos = { atualizarTudo: 1, atualizarPaineis: 1, registrarSnapshotMisc: 1 };
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (alvos[t.getHandlerFunction()]) ScriptApp.deleteTrigger(t);
  });
}

function atualizarTudo() {
  registrarSnapshotMisc();
  montarPainelMiscelania();
  montarPainelModems();
  obterPlanilha_().toast('Snapshot e painéis atualizados.', 'TCI Carga', 8);
}

/** Planilha painel TCI Carga */
var PAINEL_ID_ = '1YAYbViLyvQoIEkeBizwOHSKMc9sxkE0Rl8L8bUGro9M';

/** IDs de produção — editáveis em CONFIG */
var CFG_PADRAO_ = {
  MATERIAIS_ID: '1W1K7OzbZejmSbEUnQxJ9u43phV0sHYNWkeW3KIbv0vM',
  EQUIPES_ID:   '1qOimWoX06HPnfErX---a1Tnn6wDTK9Bgxeogvomox90',
  FORMS_ID:     '1h_u41TfElB651eKIhDgT0aQdnMx6_DASLEsBNThA5R0',
  ABA_MISC:     'BASE MICELANEAS',
  ABA_ALMOX:    'ALMOX SERIALIZADA',
  ABA_EQUIPES:  'FERNANDA TT TR TELEFONE',
  GID_EQUIPES:  '1514437459',
  ABA_FORMS:    'Respostas ao formulário 1',
  COL_SALDO_MISC: 'Y'   // coluna do saldo atual na BASE MICELANEAS (após atualização diária)
};

function lerConfig_() {
  const ss = obterPlanilha_();
  const sh = ss.getSheetByName('CONFIG');
  if (!sh) return Object.assign({}, CFG_PADRAO_);

  const vals = sh.getRange('A2:B20').getValues();
  const map = {};
  vals.forEach(function (row) {
    if (row[0]) map[String(row[0]).trim()] = String(row[1] || '').trim();
  });
  // Se o ID em CONFIG estiver inválido (ex.: OK-TESTE-... de um teste antigo),
  // ignora e usa o ID de produção. Assim o painel se conserta sozinho.
  return {
    MATERIAIS_ID: idValido_(map.MATERIAIS_ID) ? map.MATERIAIS_ID : CFG_PADRAO_.MATERIAIS_ID,
    EQUIPES_ID:   idValido_(map.EQUIPES_ID)   ? map.EQUIPES_ID   : CFG_PADRAO_.EQUIPES_ID,
    FORMS_ID:     idValido_(map.FORMS_ID)     ? map.FORMS_ID     : CFG_PADRAO_.FORMS_ID,
    ABA_MISC:     map.ABA_MISC     || CFG_PADRAO_.ABA_MISC,
    ABA_ALMOX:    map.ABA_ALMOX    || CFG_PADRAO_.ABA_ALMOX,
    ABA_EQUIPES:  map.ABA_EQUIPES  || CFG_PADRAO_.ABA_EQUIPES,
    GID_EQUIPES:  map.GID_EQUIPES  || CFG_PADRAO_.GID_EQUIPES,
    ABA_FORMS:    map.ABA_FORMS    || CFG_PADRAO_.ABA_FORMS,
    COL_SALDO_MISC: map.COL_SALDO_MISC || CFG_PADRAO_.COL_SALDO_MISC
  };
}

/** Converte letra de coluna (ex.: 'Y') em índice 0-based (Y → 24). Vazio → -1. */
function letraColParaIndice_(letra) {
  const s = String(letra || '').trim().toUpperCase();
  if (!/^[A-Z]+$/.test(s)) return -1;
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n - 1;
}

/** ID de planilha Google: só letras, dígitos, _ e -, e tamanho >= 30. */
function idValido_(v) {
  const s = String(v || '').trim();
  return /^[A-Za-z0-9_-]{30,}$/.test(s);
}

function obterPlanilha_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  return SpreadsheetApp.openById(PAINEL_ID_);
}

function testeGravarB2() {
  // NÃO escreve em B2 (= MATERIAIS_ID). Usa célula isolada E1 para não corromper CONFIG.
  var ss = obterPlanilha_();
  var sh = ss.getSheetByName('CONFIG') || ss.insertSheet('CONFIG');
  var marca = 'OK-TESTE-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');
  sh.getRange('E1').setValue('Teste gravação:');
  sh.getRange('F1').setValue(marca);
  SpreadsheetApp.flush();
  ss.setActiveSheet(sh);
  SpreadsheetApp.getUi().alert(
    'Teste',
    'Planilha: ' + ss.getName() + '\nAba CONFIG, célula F1 deve mostrar:\n' + marca +
      '\n\nSe F1 continuar vazio, o Apps Script NÃO está ligado a esta planilha.' +
      '\n\nObs: este teste NÃO altera os IDs (B2 etc.).',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function corrigirConfigVazio() {
  try {
    const ss = obterPlanilha_();
    const sh = garantirAbaConfig_(ss);
    preencherConfig_(sh);
    SpreadsheetApp.flush();
    const teste = sh.getRange('B2').getValue();
    ss.setActiveSheet(sh);
    SpreadsheetApp.getUi().alert(
      'CONFIG OK',
      'Planilha: ' + ss.getName() + '\nB2 (MATERIAIS_ID): ' + teste +
        '\n\nPróximo: TCI Carga → Tudo (snapshot + painéis)',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erro CONFIG', String(e.message || e), SpreadsheetApp.getUi().ButtonSet.OK);
    throw e;
  }
}

function montarEstruturaInicial() {
  try {
    const ss = obterPlanilha_();
    criarAbaInicio_(ss);
    preencherConfig_(garantirAbaConfig_(ss));
    criarAbaHistorico_(ss);
    criarAbaResumoDrop_(ss);
    removerFolhasVaziasPadrao_(ss);
    const shIni = ss.getSheetByName('INÍCIO');
    if (shIni) ss.setActiveSheet(shIni);
    // Atualização automática a cada 15 min já fica ativa por padrão
    try { instalarGatilhosFrequentes_(); } catch (e) { /* autorização pode faltar; ativar pelo menu */ }
    SpreadsheetApp.flush();
    ss.toast('Estrutura OK + atualização automática ativada (15 min). Próximo: TCI Carga → Tudo.', 'TCI Carga', 12);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erro', String(e.message || e), SpreadsheetApp.getUi().ButtonSet.OK);
    throw e;
  }
}

function criarAbaInicio_(ss) {
  const nome = 'INÍCIO';
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome, 0);
  sh.clear();
  sh.getRange('A1').setValue('PAINEL CARGA TCI v2').setFontSize(16).setFontWeight('bold');
  const passos = [
    ['1', 'Após colar BASE MICELANEAS (~10h) → menu TCI Carga → Tudo'],
    ['2', 'PAINEL MISCELANIA → ver RESUMO DROP (ontem × hoje)'],
    ['3', 'PAINEL MODEMS → seriais por TT (Almox) + instalados hoje (Forms)'],
    ['4', 'Amanhã repita o passo 1 para ver quem BAIXOU drop'],
    ['', 'Aba CONFIG = IDs das planilhas fonte (produção)']
  ];
  sh.getRange('A2:B6').setValues(passos);
  sh.setColumnWidth(1, 40);
  sh.setColumnWidth(2, 520);
}

function removerFolhasVaziasPadrao_(ss) {
  ['Página1', 'Pagina1', 'Folha1', 'Sheet1'].forEach(function (nome) {
    const sh = ss.getSheetByName(nome);
    if (!sh || ss.getSheets().length <= 1) return;
    if (sh.getLastRow() <= 1 && sh.getLastColumn() <= 1) {
      try { ss.deleteSheet(sh); } catch (e) { /* ok */ }
    }
  });
}

function garantirAbaConfig_(ss) {
  let sh = ss.getSheetByName('CONFIG');
  if (!sh) sh = ss.insertSheet('CONFIG');
  return sh;
}

function preencherConfig_(sh) {
  const dados = [
    ['Chave',        'Valor'],
    ['MATERIAIS_ID', CFG_PADRAO_.MATERIAIS_ID],
    ['EQUIPES_ID',   CFG_PADRAO_.EQUIPES_ID],
    ['FORMS_ID',     CFG_PADRAO_.FORMS_ID],
    ['ABA_MISC',     CFG_PADRAO_.ABA_MISC],
    ['ABA_ALMOX',    CFG_PADRAO_.ABA_ALMOX],
    ['ABA_EQUIPES',  CFG_PADRAO_.ABA_EQUIPES],
    ['GID_EQUIPES',  CFG_PADRAO_.GID_EQUIPES],
    ['ABA_FORMS',    CFG_PADRAO_.ABA_FORMS],
    ['COL_SALDO_MISC', CFG_PADRAO_.COL_SALDO_MISC],
    ['', ''],
    ['Regra ONT',    'Excluir Grupo=ONT ou subsegmento FIBRA ONT (SERIAL)'],
    ['Técnicos',     'Somente TT listados na planilha EQUIPES (FERNANDA TT TR TELEFONE)'],
    ['Snapshot',     'Registrar após colar BASE MICELANEAS (~10h)'],
    ['Painel misc',  'Compara último dia × penúltimo dia do HISTÓRICO'],
    ['Forms',        'Cruza TR (col A do Forms) com TR da EQUIPES p/ achar TT']
  ];
  sh.clear();
  for (var i = 0; i < dados.length; i++) {
    sh.getRange(i + 1, 1).setValue(dados[i][0]);
    sh.getRange(i + 1, 2).setValue(dados[i][1]);
  }
  sh.getRange('A1:B1').setFontWeight('bold');
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 420);
}

function criarAbaHistorico_(ss) {
  const nome = 'HISTORICO MISC';
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  if (sh.getLastRow() < 1) {
    sh.getRange('A1:J1').setValues([[
      'Data snapshot', 'TT', 'Nome técnico', 'Cód. material', 'Material',
      'Grupo material', 'Agregador', 'Saldo', 'Segmento', 'Observação'
    ]]).setFontWeight('bold').setBackground('#e8eaf6');
    sh.setFrozenRows(1);
  }
}

function criarAbaResumoDrop_(ss) {
  const nome = 'RESUMO DROP';
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
}

// ─── Equipe TCI (EQUIPES Página4, índice por TT e por TR) ────────────────────

function carregarEquipeTci_(cfg) {
  const ssExt = SpreadsheetApp.openById(cfg.EQUIPES_ID);
  let sh = ssExt.getSheetByName(cfg.ABA_EQUIPES);
  if (!sh) {
    const sheets = ssExt.getSheets();
    sh = sheets.find(function (s) {
      return String(s.getSheetId()) === String(cfg.GID_EQUIPES);
    }) || sheets[0];
  }
  const last = sh.getLastRow();
  if (last < 2) return { lista: [], porTt: {}, porTr: {} };

  const numCols = Math.max(sh.getLastColumn(), 5);
  const dados = sh.getRange(1, 1, last, numCols).getValues();
  const header = dados[0].map(normalizarTexto_);

  const ixNome   = indiceColuna_(header, ['nome do tecnico', 'nome']);
  const ixTt     = indiceColuna_(header, ['tt']);
  const ixTr     = indiceColuna_(header, ['tr']);
  const ixArea   = indiceColuna_(header, ['area atual', 'area']);
  // STATUS: tenta pelo header; Página4 tem header vazio na col 5 → fallback posição 4
  let ixStatus   = indiceColuna_(header, ['status', 'situacao']);
  if (ixStatus < 0 && numCols >= 5) ixStatus = 4;

  const porTt = {};
  const porTr = {};
  const porNome = {};
  const lista = [];

  for (let i = 1; i < dados.length; i++) {
    const tt = normalizarTt_(dados[i][ixTt]);
    if (!tt) continue;
    const nome = String(dados[i][ixNome] || '').trim();
    if (!nome) continue;
    const tr = normalizarTr_(ixTr >= 0 ? dados[i][ixTr] : '');
    const item = {
      tt:     tt,
      tr:     tr,
      nome:   nome,
      area:   ixArea  >= 0 ? String(dados[i][ixArea]   || '').trim() : '',
      status: ixStatus >= 0 ? String(dados[i][ixStatus] || '').trim() : ''
    };
    porTt[tt] = item;
    if (tr) porTr[tr] = item;
    const chaveNome = chaveNome_(nome);
    if (chaveNome) porNome[chaveNome] = item;
    lista.push(item);
  }
  return { lista: lista, porTt: porTt, porTr: porTr, porNome: porNome };
}

/** Chave de nome normalizada (sem acento, maiúsculo, espaços colapsados) para fallback de cruzamento. */
function chaveNome_(v) {
  return normalizarTexto_(v).replace(/\s+/g, ' ').trim();
}

// ─── Snapshot miscelânea ─────────────────────────────────────────────────────

function registrarSnapshotMisc() {
  const ss = obterPlanilha_();
  const cfg = lerConfig_();
  const equipe = carregarEquipeTci_(cfg);
  if (!equipe.lista.length) {
    throw new Error('Nenhum técnico na planilha EQUIPES. Confira ABA_EQUIPES em CONFIG.');
  }

  const linhasMisc = lerMiscFiltrada_(cfg, equipe.porTt);
  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  let shHist = ss.getSheetByName('HISTORICO MISC');
  if (!shHist) {
    criarAbaHistorico_(ss);
    shHist = ss.getSheetByName('HISTORICO MISC');
  }

  removerSnapshotData_(shHist, hoje);

  const saida = linhasMisc.map(function (r) {
    return [hoje, r.tt, r.nome, r.codmaterial, r.material, r.grupo, r.agregador, r.saldo, r.segmento, r.observacao || ''];
  });

  if (saida.length) {
    const start = shHist.getLastRow() + 1;
    rangeLinhas_(shHist, start, 1, saida.length, 10).setValues(saida);
  }

  ss.toast(
    'Snapshot ' + hoje + ': ' + saida.length + ' itens (TT da equipe, sem ONT).',
    'TCI Carga', 10
  );
}

function removerSnapshotData_(sh, dataIso) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const datas = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = datas.length - 1; i >= 0; i--) {
    const d = formatarDataIso_(datas[i][0]);
    if (d === dataIso) sh.deleteRow(i + 2);
  }
}

function lerMiscFiltrada_(cfg, porTt) {
  const ssExt = SpreadsheetApp.openById(cfg.MATERIAIS_ID);
  const sh = ssExt.getSheetByName(cfg.ABA_MISC);
  if (!sh) throw new Error('Aba não encontrada: ' + cfg.ABA_MISC);

  const last = sh.getLastRow();
  if (last < 2) return [];

  const dados = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
  const header = dados[0].map(normalizarTexto_);
  const ix = {
    tt:       indiceColuna_(header, ['codarmazem', 'cod armazem', 'tt']),
    nome:     indiceColuna_(header, ['armazem']),
    cod:      indiceColuna_(header, ['codmaterial', 'cod material']),
    mat:      indiceColuna_(header, ['material']),
    grupo:    indiceColuna_(header, ['grupo material']),
    agreg:    indiceColuna_(header, ['agregador']),
    saldo:    indiceColuna_(header, ['saldo']),
    sub:      indiceColuna_(header, ['subsegmento', 'sub grupo agregador']),
    segmento: indiceColuna_(header, ['segmento']),
    obs:      indiceColuna_(header, ['observacao', 'observacoes', 'obs'])
  };

  // Saldo: usa a coluna fixa Y (saldo atual após atualização diária da base).
  // Se a coluna Y não existir/estiver fora do range, cai no detector por header.
  const ixSaldoY = letraColParaIndice_(cfg.COL_SALDO_MISC);
  if (ixSaldoY >= 0 && ixSaldoY < dados[0].length) ix.saldo = ixSaldoY;

  const mapa = {};

  for (let i = 1; i < dados.length; i++) {
    const row = dados[i];
    const tt = normalizarTt_(row[ix.tt]);
    if (!tt || !porTt[tt]) continue;

    const grupo    = String(row[ix.grupo] || '');
    const material = String(row[ix.mat]   || '');
    const sub      = ix.sub >= 0 ? String(row[ix.sub] || '') : '';
    // Miscelânia = tudo que NÃO é equipamento (modem, ONT, roteador). Só material complementar.
    if (ehOnt_(grupo, sub, material) || ehEquipamentoNaoMisc_(grupo, material)) continue;

    const saldo = parseNumero_(row[ix.saldo]);
    if (saldo <= 0) continue;

    const cod  = String(row[ix.cod] || '').trim();
    const chave = tt + '\t' + cod;
    const nome  = porTt[tt].nome;

    if (!mapa[chave]) {
      mapa[chave] = {
        tt:          tt,
        nome:        nome,
        codmaterial: cod,
        material:    material,
        grupo:       grupo,
        agregador:   ix.agreg    >= 0 ? String(row[ix.agreg]    || '') : '',
        segmento:    ix.segmento >= 0 ? String(row[ix.segmento] || '') : '',
        observacao:  ix.obs      >= 0 ? String(row[ix.obs]      || '').trim() : '',
        saldo:       0
      };
    }
    mapa[chave].saldo += saldo;
  }

  return Object.keys(mapa).map(function (k) { return mapa[k]; });
}

function ehOnt_(grupo, subsegmento, material) {
  const g = String(grupo      || '').trim().toUpperCase();
  const s = String(subsegmento|| '').toUpperCase();
  const m = String(material   || '').trim().toUpperCase();
  if (s.indexOf('FIBRA ONT (SERIAL)') >= 0) return true;
  if (g === 'ONT') return true;
  if (m.indexOf('ONT ') === 0) return true;
  return false;
}

function ehDrop_(material, agregador, grupo) {
  const t = [material, agregador, grupo].join(' ').toUpperCase();
  return t.indexOf('DROP') >= 0;
}

/** Equipamento que NÃO é miscelânia: apenas modem. Roteador, ONU e todo o resto são miscelânia. */
function ehEquipamentoNaoMisc_(grupo, material) {
  const t = [grupo, material].join(' ').toUpperCase();
  return t.indexOf('MODEM') >= 0;
}

/** Rolo de 500m (caixa) — medido em metros. Demais drops/cabos contados por unidade (peça). */
function ehRolo500_(material, agregador) {
  const t = [material, agregador].join(' ').toUpperCase();
  return t.indexOf('ROLO') >= 0 && /500/.test(t);
}

/** Unidade de medida: só rolo 500m em metros; todo o resto (inclusive drop em peças) em unidades. */
function medidaMaterial_(material, agregador, grupo) {
  return ehRolo500_(material, agregador) ? 'metros' : 'un';
}

/**
 * Categoria-chave do material (por palavra-chave, pois a descrição varia na base).
 * Retorna null se não for um dos materiais-chave acompanhados.
 */
function categoriaChave_(material, agregador) {
  const t = [material, agregador].join(' ').toUpperCase();
  if (t.indexOf('CONECTOR') >= 0 && t.indexOf('INTERN') >= 0) return 'CONECTOR INTERNO';
  if (t.indexOf('CONECTOR') >= 0 && t.indexOf('EXTERN') >= 0) return 'CONECTOR EXTERNO';
  if (t.indexOf('CONECTOR') >= 0)                              return 'CONECTOR';
  if (t.indexOf('PLAQUETA') >= 0)                             return 'PLAQUETA';
  if (t.indexOf('ESTICADOR') >= 0 || t.indexOf('ANEL') >= 0)  return 'ESTICADOR / ANEL';
  if (t.indexOf('CUNHA') >= 0)                                return 'CUNHA';
  if (ehRolo500_(material, agregador))                        return 'CABO ROLO 500M';
  if (t.indexOf('DROP') >= 0 || t.indexOf('CABO') >= 0)       return 'CABO / DROP (peças)';
  return null;
}

/**
 * Categoria simplificada da miscelânia (painel novo).
 * DROP = só descrição com DROP (peças) + rolo 500m. CABO genérico sem DROP é ignorado.
 * Retorna null se não for categoria acompanhada.
 */
function categoriaMisc_(material, agregador, grupo) {
  const t = [material, agregador, grupo].join(' ').toUpperCase();
  if (t.indexOf('CONECTOR') >= 0) return 'CONECTOR';
  if (t.indexOf('PLAQUETA')  >= 0) return 'PLAQUETA';
  if (t.indexOf('CUNHA')     >= 0) return 'CUNHA';
  if (ehRolo500_(material, agregador)) return 'DROP ROLO';
  if (t.indexOf('DROP')      >= 0) return 'DROP PECAS';
  return null;
}

/** Colunas do resumo de miscelânia por técnico, na ordem de exibição. */
var CATS_MISC_ = [
  { key: 'DROP PECAS', label: 'DROP peças (un)', medida: 'un' },
  { key: 'DROP ROLO',  label: 'DROP rolo 500m (m)', medida: 'metros' },
  { key: 'CONECTOR',   label: 'CONECTOR (un)', medida: 'un' },
  { key: 'PLAQUETA',   label: 'PLAQUETA (un)', medida: 'un' },
  { key: 'CUNHA',      label: 'CUNHA (un)', medida: 'un' }
];

/** Ordem de exibição da seção materiais-chave. */
var ORDEM_CHAVE_ = [
  'CABO / DROP (peças)', 'CABO ROLO 500M', 'CONECTOR INTERNO', 'CONECTOR EXTERNO',
  'CONECTOR', 'PLAQUETA', 'ESTICADOR / ANEL', 'CUNHA'
];

/** Medida de cada categoria-chave. */
function medidaCategoria_(cat) {
  return cat === 'CABO ROLO 500M' ? 'metros' : 'un';
}

/**
 * Dias consecutivos (mais recentes) em que o saldo NÃO baixou para uma chave TT+material.
 * 0 = baixou no último dia (ou só há 1 snapshot). 4+ = alerta de estoque parado.
 */
function diasSemBaixar_(snapshots, chave) {
  const datas = snapshots.datas;
  if (datas.length < 2) return 0;
  let dias = 0;
  for (let i = datas.length - 1; i >= 1; i--) {
    const cur  = snapshots.porData[datas[i]]     && snapshots.porData[datas[i]][chave];
    const prev = snapshots.porData[datas[i - 1]] && snapshots.porData[datas[i - 1]][chave];
    const sCur  = cur  ? cur.saldo  : 0;
    const sPrev = prev ? prev.saldo : 0;
    if (sCur < sPrev) break;   // baixou nesta transição → para de contar
    dias++;
  }
  return dias;
}

// ─── Forms instalados do dia ─────────────────────────────────────────────────

/**
 * Lê FORMS e retorna instalações agrupadas por TT.
 * Col A formato: "NOME - TR######"  →  extrai TR  →  busca TT em equipe.porTr
 */
function lerFormsInstalados_(cfg, equipe) {
  const ssExt = SpreadsheetApp.openById(cfg.FORMS_ID);
  const sh = ssExt.getSheetByName(cfg.ABA_FORMS);
  if (!sh) return {};

  const last = sh.getLastRow();
  if (last < 2) return {};

  const dados = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
  const header = dados[0].map(normalizarTexto_);

  const ixData     = indiceColuna_(header, ['carimbo de data/hora', 'data/hora']);
  const ixSerial1  = indiceColuna_(header, ['serial do equipamento instalado 1', 'serial 1']);
  const ixSerial2  = indiceColuna_(header, ['serial do equipamento instalado 2', 'serial 2']);
  const ixProtocolo= indiceColuna_(header, ['numero do protocolo', 'protocolo']);
  const ixTipo     = indiceColuna_(header, ['tipo de servico', 'tipo de servi']);
  const ixErro     = indiceColuna_(header, ['erro de baixa', 'erro']);

  const porTt = {};

  for (let i = 1; i < dados.length; i++) {
    const row   = dados[i];
    const colA  = String(row[0] || '').trim();
    const tr    = extrairTrDaColA_(colA);

    // 1º tenta cruzar pelo TR; se não achar, cai para o nome (col A traz "NOME - TR######")
    let tec = tr ? equipe.porTr[tr] : null;
    if (!tec) tec = casarPorNome_(colA, equipe.porNome);
    if (!tec) continue;

    const tt = tec.tt;
    if (!porTt[tt]) {
      porTt[tt] = { tt: tt, nome: tec.nome, area: tec.area, instalacoes: [] };
    }

    const s1 = String(ixSerial1  >= 0 ? row[ixSerial1]   || '' : '').trim().toUpperCase();
    const s2 = String(ixSerial2  >= 0 ? row[ixSerial2]   || '' : '').trim().toUpperCase();
    const seriais = [s1, s2].filter(Boolean);
    if (!seriais.length) continue;

    const dataRaw = ixData >= 0 ? row[ixData] : '';
    const dataIso = formatarDataIso_(dataRaw);

    porTt[tt].instalacoes.push({
      data:      dataIso,
      protocolo: ixProtocolo >= 0 ? String(row[ixProtocolo] || '') : '',
      tipo:      ixTipo      >= 0 ? String(row[ixTipo]      || '') : '',
      erro:      ixErro      >= 0 ? String(row[ixErro]      || '') : '',
      seriais:   seriais
    });
  }
  return porTt;
}

function extrairTrDaColA_(v) {
  const m = String(v || '').match(/\bTR\d+\b/i);
  return m ? m[0].toUpperCase() : '';
}

/**
 * Fallback de cruzamento por nome quando o TR não bate.
 * Pega a parte de nome da col A (antes do " - TR..."), normaliza e procura em porNome.
 * Tenta nome completo e depois primeiro+último token.
 */
function casarPorNome_(colA, porNome) {
  if (!porNome) return null;
  let s = String(colA || '');
  s = s.replace(/\bT[RT]\d+\b/ig, ' ');          // remove códigos TR/TT
  s = s.replace(/[-–|].*$/, ' ');                 // corta a partir de traço/barra
  const chave = chaveNome_(s);
  if (!chave) return null;
  if (porNome[chave]) return porNome[chave];
  // tenta primeiro + último nome
  const toks = chave.split(' ').filter(Boolean);
  if (toks.length >= 2) {
    const reduzida = toks[0] + ' ' + toks[toks.length - 1];
    if (porNome[reduzida]) return porNome[reduzida];
  }
  // tenta achar uma chave que comece pelo mesmo primeiro+segundo nome
  if (toks.length >= 2) {
    const prefixo = toks[0] + ' ' + toks[1];
    const achou = Object.keys(porNome).find(function (k) { return k.indexOf(prefixo) === 0; });
    if (achou) return porNome[achou];
  }
  return null;
}

// ─── Painel miscelânia (ontem × hoje) ────────────────────────────────────────

function montarPainelMiscelania() {
  const ss = obterPlanilha_();
  const cfg = lerConfig_();
  const equipe = carregarEquipeTci_(cfg);

  // Garante que o snapshot de hoje está salvo (a base é trocada todo dia — guardamos o histórico).
  try { registrarSnapshotMisc(); } catch (e) { /* segue com o que houver */ }

  const shHist = ss.getSheetByName('HISTORICO MISC');
  if (!shHist || shHist.getLastRow() < 2) {
    throw new Error('Sem histórico. Rode primeiro: Registrar snapshot de hoje.');
  }

  const snapshots = lerSnapshots_(shHist);
  if (snapshots.datas.length < 1) throw new Error('Histórico vazio.');

  const dataHoje  = snapshots.datas[snapshots.datas.length - 1];
  const dataOntem = snapshots.datas.length >= 2
    ? snapshots.datas[snapshots.datas.length - 2]
    : null;

  const nome = 'PAINEL MISCELANIA';
  let old = ss.getSheetByName(nome);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(nome, 0);

  const NC = CATS_MISC_.length + 3; // TT, Técnico, Área + categorias
  const colFim = colLetra_(Math.max(NC, 9));

  titulo_(sh, 'A1:' + colFim + '1', 'PAINEL MISCELANIA TCI — CARGA ATUAL POR TÉCNICO (saldo col ' + cfg.COL_SALDO_MISC + ')');
  const txtComp = dataOntem
    ? ('Comparação com ontem: ' + dataOntem + ' → ' + dataHoje)
    : ('1º snapshot: ' + dataHoje + ' — farol de movimentação aparece a partir de amanhã');
  sh.getRange('A2:' + colFim + '2').merge()
    .setValue(txtComp + '  |  Atualizado: ' + Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'
    ))
    .setFontStyle('italic').setHorizontalAlignment('center').setFontSize(10);

  escreverDescricaoRegras_(sh, 'A3:' + colFim + '3',
    'Saldo atual = coluna ' + cfg.COL_SALDO_MISC + ' da BASE MICELANEAS. A base é trocada diariamente até 11h e o histórico de ontem fica salvo. ' +
    'Farol por célula: VERDE = saldo baixou de ontem→hoje (aplicou) · VERMELHO = saldo igual (sem movimentação). ' +
    'DROP = só descrição DROP (peças) + rolo 500m. CABO genérico sem DROP não entra.');

  let row = 5;

  // ── SEÇÃO 1: RESUMO POR TÉCNICO — quantidade por categoria, farol por célula ──
  sh.getRange(row, 1).setValue('RESUMO — quantidade em carga por técnico e categoria').setFontWeight('bold').setFontSize(11);
  row++;
  const cabResumo = ['TT', 'Técnico', 'Área'].concat(CATS_MISC_.map(function (c) { return c.label; }));
  cabecalhoTabela_(sh, 'A' + row + ':' + colLetra_(NC) + row, cabResumo);
  const linCabResumo = row;
  row++;

  const linhasResumo = [];
  const corResumo = []; // matriz de cores por célula de categoria
  Object.keys(equipe.porTt).sort().forEach(function (tt) {
    const tec = equipe.porTt[tt];
    const h = somarCatsMiscTecnico_(snapshots, dataHoje,  tt);
    const o = somarCatsMiscTecnico_(snapshots, dataOntem, tt);
    const temAlgo = CATS_MISC_.some(function (c) { return h[c.key] > 0 || o[c.key] > 0; });
    if (!temAlgo) return;

    const lin = [tt, tec.nome, tec.area || ''];
    const cores = [];
    CATS_MISC_.forEach(function (c) {
      const vh = h[c.key] || 0;
      const vo = o[c.key] || 0;
      lin.push(vh);
      // VERDE se houve movimentação (saldo mudou); VERMELHO se igual; sem ontem → sem cor
      if (!dataOntem)      cores.push(null);
      else if (vh !== vo)  cores.push('#e8f5e9');
      else                 cores.push('#ffcdd2');
    });
    linhasResumo.push(lin);
    corResumo.push(cores);
  });

  if (linhasResumo.length) {
    const fim = row + linhasResumo.length - 1;
    rangeLinhas_(sh, row, 1, linhasResumo.length, NC).setValues(linhasResumo);
    rangeLinhas_(sh, row, 4, linhasResumo.length, CATS_MISC_.length).setNumberFormat('#,##0');
    for (let i = 0; i < corResumo.length; i++) {
      for (let j = 0; j < corResumo[i].length; j++) {
        if (corResumo[i][j]) sh.getRange(row + i, 4 + j).setBackground(corResumo[i][j]);
      }
    }
    sh.getRange('A' + linCabResumo + ':' + colLetra_(NC) + fim).setBorder(
      true, true, true, true, true, true, '#9e9e9e', SpreadsheetApp.BorderStyle.SOLID);
    row = fim + 3;
  } else {
    row += 2;
  }

  // ── SEÇÃO 2: DETALHE POR TÉCNICO E MATERIAL — farol por linha (sem ontem/hoje) ──
  sh.getRange(row, 1).setValue('DETALHE POR TÉCNICO E MATERIAL').setFontWeight('bold').setFontSize(11);
  row++;
  cabecalhoTabela_(sh, 'A' + row + ':I' + row, [
    'TT', 'Técnico', 'Área', 'Categoria', 'Material', 'Medida', 'Saldo atual', 'Status', 'Observação'
  ]);
  const linhaCab = row;
  row++;

  const chaves = {};
  equipe.lista.forEach(function (tec) {
    [dataHoje, dataOntem].forEach(function (d) {
      if (!d || !snapshots.porData[d]) return;
      Object.keys(snapshots.porData[d]).forEach(function (k) {
        if (k.indexOf(tec.tt + '\t') === 0) chaves[k] = true;
      });
    });
  });

  const linhas = [];
  const coresStatus = [];
  const notasObs = [];
  Object.keys(chaves).sort().forEach(function (chave) {
    const tt  = chave.split('\t')[0];
    const tec = equipe.porTt[tt];
    if (!tec) return;

    const snapHoje  = (snapshots.porData[dataHoje]  && snapshots.porData[dataHoje][chave])  || null;
    const snapOntem = dataOntem && snapshots.porData[dataOntem]
      ? snapshots.porData[dataOntem][chave] : null;
    const ref = snapHoje || snapOntem;

    const cat = categoriaMisc_(ref.material, ref.agregador, ref.grupo);
    if (!cat) return; // só categorias acompanhadas (DROP, CONECTOR, PLAQUETA, CUNHA)

    const saldoHoje  = snapHoje  ? snapHoje.saldo  : 0;
    const saldoOntem = snapOntem ? snapOntem.saldo : 0;

    let status, cor, obsMov;
    if (!dataOntem) { status = 'AGUARDANDO 2º DIA'; cor = null; obsMov = ''; }
    else if (saldoHoje !== saldoOntem) { status = 'MOVIMENTOU'; cor = '#e8f5e9'; obsMov = ''; }
    else { status = 'PARADO'; cor = '#ffcdd2'; obsMov = 'sem movimentação'; }

    const obsBase = (ref.observacao || '').trim();
    if (obsBase) notasObs.push({ offset: linhas.length, texto: 'Observação almox: ' + obsBase });

    linhas.push([
      tt, tec.nome, tec.area || '', catMiscLabel_(cat),
      ref.material, medidaMaterial_(ref.material, ref.agregador, ref.grupo),
      saldoHoje, status, obsMov
    ]);
    coresStatus.push(cor);
  });

  if (linhas.length) {
    const fim = row + linhas.length - 1;
    rangeLinhas_(sh, row, 1, linhas.length, 9).setValues(linhas);
    rangeLinhas_(sh, row, 7, linhas.length, 1).setNumberFormat('#,##0');
    for (let i = 0; i < coresStatus.length; i++) {
      if (coresStatus[i]) {
        sh.getRange(row + i, 8).setBackground(coresStatus[i]); // Status
        sh.getRange(row + i, 9).setBackground(coresStatus[i]); // Observação
      }
    }
    notasObs.forEach(function (n) { sh.getRange(row + n.offset, 5).setNote(n.texto); });
    sh.getRange('A' + linhaCab + ':I' + fim).setBorder(
      true, true, true, true, true, true, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);
  }

  // Larguras: cabeçalho do detalhe tem 9 colunas
  [90, 200, 70, 110, 280, 70, 90, 120, 150].forEach(function (w, i) {
    sh.setColumnWidth(i + 1, w);
  });
  sh.setFrozenRows(4);   // congela só título/subtítulo/regras — não congela tabelas
  SpreadsheetApp.flush();
  ss.toast('Painel miscelânia atualizado.', 'TCI Carga', 8);
}

/** Soma saldo por categoria simplificada (CATS_MISC_) para um técnico numa data. */
function somarCatsMiscTecnico_(snapshots, data, tt) {
  const out = {};
  CATS_MISC_.forEach(function (c) { out[c.key] = 0; });
  if (!data || !snapshots.porData[data]) return out;
  Object.keys(snapshots.porData[data]).forEach(function (chave) {
    if (chave.indexOf(tt + '\t') !== 0) return;
    const item = snapshots.porData[data][chave];
    const cat = categoriaMisc_(item.material, item.agregador, item.grupo);
    if (cat && out[cat] !== undefined) out[cat] += item.saldo;
  });
  return out;
}

/** Rótulo amigável da categoria no detalhe. */
function catMiscLabel_(cat) {
  if (cat === 'DROP PECAS') return 'DROP (peças)';
  if (cat === 'DROP ROLO')  return 'DROP (rolo 500m)';
  return cat;
}

/** Texto de regras/descrição numa faixa mesclada, fonte pequena, sem poluir. */
function escreverDescricaoRegras_(sh, range, texto) {
  sh.getRange(range).merge().setValue(texto)
    .setFontSize(9).setFontColor('#555').setWrap(true)
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBackground('#f5f5f5');
}

/** SEÇÃO 1 — Materiais-chave: total TCI (ontem × hoje) somando todos os técnicos. */
function escreverMateriaisChave_(sh, row, snapshots, dataOntem, dataHoje, porTt) {
  sh.getRange(row, 1).setValue('MATERIAIS-CHAVE — TOTAL TCI (ontem × hoje)').setFontWeight('bold').setFontSize(11);
  row++;
  cabecalhoTabela_(sh, 'A' + row + ':G' + row, [
    'Material-chave', 'Medida', 'Téc. c/ estoque', 'Total ontem', 'Total hoje', 'Variação', 'Consumiu?'
  ]);
  const linCab = row;
  row++;

  const totHoje  = somarPorCategoria_(snapshots, dataHoje, porTt);
  const totOntem = somarPorCategoria_(snapshots, dataOntem, porTt);

  const linhas = [];
  ORDEM_CHAVE_.forEach(function (cat) {
    const h = totHoje[cat]  || { soma: 0, tecs: {} };
    const o = totOntem[cat] || { soma: 0, tecs: {} };
    if (h.soma <= 0 && o.soma <= 0) return;
    const variacao = h.soma - o.soma;
    const consumiu = dataOntem ? (variacao < 0 ? 'SIM' : 'NÃO') : '';
    linhas.push([
      cat, medidaCategoria_(cat), Object.keys(h.tecs).length,
      dataOntem ? o.soma : '', h.soma,
      dataOntem ? variacao : '', consumiu
    ]);
  });

  if (linhas.length) {
    const fim = row + linhas.length - 1;
    rangeLinhas_(sh, row, 1, linhas.length, 7).setValues(linhas);
    rangeLinhas_(sh, row, 4, linhas.length, 3).setNumberFormat('#,##0');
    // colore a coluna "Consumiu?" (col G = 7): SIM verde, NÃO amarelo
    for (let i = 0; i < linhas.length; i++) {
      const c = String(linhas[i][6] || '').toUpperCase();
      if (c === 'SIM')      sh.getRange(row + i, 7).setBackground('#e8f5e9');
      else if (c === 'NÃO') sh.getRange(row + i, 7).setBackground('#fff8e1');
    }
    sh.getRange('A' + linCab + ':G' + fim).setBorder(
      true, true, true, true, true, true, '#9e9e9e', SpreadsheetApp.BorderStyle.SOLID
    );
    row = fim + 2;
  } else {
    row += 2;
  }
  return row;
}

/** Soma saldo por categoria-chave (todos os TT da equipe) numa data. */
function somarPorCategoria_(snapshots, data, porTt) {
  const out = {};
  if (!data || !snapshots.porData[data]) return out;
  Object.keys(snapshots.porData[data]).forEach(function (chave) {
    const tt = chave.split('\t')[0];
    if (!porTt[tt]) return;
    const item = snapshots.porData[data][chave];
    const cat = categoriaChave_(item.material, item.agregador);
    if (!cat) return;
    if (!out[cat]) out[cat] = { soma: 0, tecs: {} };
    out[cat].soma += item.saldo;
    if (item.saldo > 0) out[cat].tecs[tt] = true;
  });
  return out;
}

/**
 * RESUMO POR TÉCNICO — todas as categorias de miscelânia (ontem × hoje).
 * Colunas: TT | Técnico | Área | CABO/DROP (un) | Rolo 500m (m) | CONECTOR | PLAQUETA | ESTICADOR/ANEL | Consumiu hoje?
 */
function escreverResumoDrop_(sh, row, snapshots, dataOntem, dataHoje, porTt) {
  sh.getRange(row, 1).setValue('RESUMO MISCELÂNIA POR TÉCNICO — todas as categorias (ontem × hoje)').setFontWeight('bold').setFontSize(11);
  row++;

  // Cabeçalho duplo: categoria acima, ontem/hoje abaixo
  const CATS = [
    { label: 'CABO/DROP (un)',     key: 'CABO / DROP (peças)' },
    { label: 'Rolo 500m (m)',      key: 'CABO ROLO 500M'      },
    { label: 'CONECTOR (un)',      key: 'CONECTOR'             },
    { label: 'PLAQUETA (un)',      key: 'PLAQUETA'             },
    { label: 'ESTICADOR/ANEL (un)',key: 'ESTICADOR / ANEL'     }
  ];

  // Linha de cabeçalho simples
  const cabRow = ['TT', 'Técnico', 'Área'];
  CATS.forEach(function (c) {
    cabRow.push(c.label + ' ontem');
    cabRow.push(c.label + ' hoje');
  });
  cabRow.push('Consumiu hoje?');

  cabecalhoTabela_(sh, 'A' + row + ':' + colLetra_(cabRow.length) + row, cabRow);
  const linCab = row;
  row++;

  const resumo = [];
  Object.keys(porTt).sort().forEach(function (tt) {
    const somaH = somarCatsTecnico_(snapshots, dataHoje,  tt, CATS);
    const somaO = somarCatsTecnico_(snapshots, dataOntem, tt, CATS);

    // Pula técnico sem nenhum material nas categorias
    const temAlgo = CATS.some(function (c) { return somaH[c.key] > 0 || somaO[c.key] > 0; });
    if (!temAlgo) return;

    const lin = [tt, porTt[tt].nome, porTt[tt].area || ''];
    let consumiu = false;
    CATS.forEach(function (c) {
      const h = somaH[c.key] || 0;
      const o = somaO[c.key] || 0;
      lin.push(dataOntem ? o : '');
      lin.push(h);
      if (dataOntem && h < o) consumiu = true;
    });
    lin.push(dataOntem ? (consumiu ? 'SIM' : 'NÃO') : '');
    resumo.push(lin);
  });

  if (resumo.length) {
    const ncols = cabRow.length;
    const fim = row + resumo.length - 1;
    rangeLinhas_(sh, row, 1, resumo.length, ncols).setValues(resumo);
    // Colorir "Consumiu hoje?" (última coluna)
    const colCons = ncols;
    for (let i = 0; i < resumo.length; i++) {
      const v = String(resumo[i][colCons - 1] || '').toUpperCase();
      let bg = null;
      if (v === 'SIM') bg = '#e8f5e9';
      else if (v === 'NÃO') bg = '#fff8e1';
      if (bg) sh.getRange(row + i, colCons).setBackground(bg);
    }
    sh.getRange('A' + linCab + ':' + colLetra_(ncols) + fim).setBorder(
      true, true, true, true, true, true, '#9e9e9e', SpreadsheetApp.BorderStyle.SOLID
    );
    row = fim + 3;
  } else {
    row += 2;
  }
  return row;
}

/** Soma saldo por categoria-chave para um técnico (TT) numa data. */
function somarCatsTecnico_(snapshots, data, tt, cats) {
  const out = {};
  cats.forEach(function (c) { out[c.key] = 0; });
  if (!data || !snapshots.porData[data]) return out;
  Object.keys(snapshots.porData[data]).forEach(function (chave) {
    if (chave.indexOf(tt + '\t') !== 0) return;
    const item = snapshots.porData[data][chave];
    // Para CONECTOR agrupamos qualquer variante
    const cat = categoriaChaveResumo_(item.material, item.agregador);
    if (cat && out[cat] !== undefined) out[cat] += item.saldo;
  });
  return out;
}

/** Mesmo que categoriaChave_ mas mapeia variantes de CONECTOR para a chave genérica. */
function categoriaChaveResumo_(material, agregador) {
  const t = [material, agregador].join(' ').toUpperCase();
  if (t.indexOf('CONECTOR') >= 0)                              return 'CONECTOR';
  if (t.indexOf('PLAQUETA') >= 0)                             return 'PLAQUETA';
  if (t.indexOf('ESTICADOR') >= 0 || t.indexOf('ANEL') >= 0)  return 'ESTICADOR / ANEL';
  if (ehRolo500_(material, agregador))                        return 'CABO ROLO 500M';
  if (t.indexOf('DROP') >= 0 || t.indexOf('CABO') >= 0)       return 'CABO / DROP (peças)';
  return null;
}

/** Converte número de coluna (1-based) para letra(s) de coluna. */
function colLetra_(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Colore coluna "Dias s/ baixar": 4+ vermelho forte, 2-3 amarelo, 0-1 verde. */
function colorirDiasSemBaixar_(sh, rowIni, rowFim, col) {
  const range = sh.getRange(rowIni, col, rowFim - rowIni + 1, 1);
  const vals  = range.getValues();
  for (let i = 0; i < vals.length; i++) {
    const v = vals[i][0];
    if (v === '' || v === null) continue;
    const n = Number(v);
    let c = null;
    if (n >= 4)      c = '#ef9a9a';   // vermelho forte — estoque parado (alerta)
    else if (n >= 2) c = '#fff8e1';   // amarelo
    else             c = '#e8f5e9';   // verde — está girando
    if (c) sh.getRange(rowIni + i, col).setBackground(c);
  }
}

function classificarVariacao_(ontem, hoje) {
  if (ontem <= 0 && hoje > 0) return 'ENTROU';
  if (ontem > 0  && hoje <= 0) return 'ZEROU';
  if (hoje < ontem) return 'BAIXOU';
  if (hoje > ontem) return 'SUBIU';
  return 'MANTEVE';
}

function lerSnapshots_(shHist) {
  const last = shHist.getLastRow();
  const numCols = Math.min(shHist.getLastColumn(), 10);
  const dados = rangeLinhas_(shHist, 2, 1, last - 1, numCols).getValues();
  const porData = {};
  const setDatas = {};

  dados.forEach(function (row) {
    const data = formatarDataIso_(row[0]);
    if (!data) return;
    setDatas[data] = true;
    const tt  = normalizarTt_(row[1]);
    const cod = String(row[3] || '').trim();
    const chave = tt + '\t' + cod;
    if (!porData[data]) porData[data] = {};
    porData[data][chave] = {
      tt:          tt,
      nome:        String(row[2] || ''),
      codmaterial: cod,
      material:    String(row[4] || ''),
      grupo:       String(row[5] || ''),
      agregador:   String(row[6] || ''),
      saldo:       parseNumero_(row[7]),
      segmento:    String(row[8] || ''),
      observacao:  String(row[9] || '')
    };
  });

  const datas = Object.keys(setDatas).sort();
  return { datas: datas, porData: porData };
}

function colorirSituacao_(sh, rowIni, rowFim, colSit) {
  const range = sh.getRange(rowIni, colSit, rowFim - rowIni + 1, 1);
  const vals  = range.getValues();
  for (let i = 0; i < vals.length; i++) {
    const s = String(vals[i][0] || '').toUpperCase();
    let c = null;
    if (s === 'BAIXOU' || s === 'ZEROU')    c = '#ffebee';
    else if (s === 'MANTEVE')               c = '#fff8e1';
    else if (s === 'SUBIU' || s === 'ENTROU') c = '#e8f5e9';
    if (c) sh.getRange(rowIni + i, colSit).setBackground(c);
  }
}

// ─── Painel modems (Almox + Forms instalados hoje) ───────────────────────────

function montarPainelModems() {
  const ss     = obterPlanilha_();
  const cfg    = lerConfig_();
  const equipe = carregarEquipeTci_(cfg);
  const modems = lerModemsAlmox_(cfg, equipe.porTt);
  const forms  = lerFormsInstalados_(cfg, equipe);

  const tz = Session.getScriptTimeZone();
  const hoje = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  // TODOS os serials já encerrados no Forms (histórico completo) — saem da carga na hora.
  const encerradosTotal = coletarSeriaisInstalados_(forms, null);

  const nome = 'PAINEL MODEMS';
  let old = ss.getSheetByName(nome);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(nome, 0);

  // 9 colunas: TT | Técnico | Área | Velocidade | Qtd em carga (líquida) | Encerrados hoje | Total encerrados (histórico) | Último encerramento | Status
  titulo_(sh, 'A1:I1', 'PAINEL MODEMS TCI — CARGA LÍQUIDA POR TT (só ONT · MESH · MODEM)');
  sh.getRange('A2:I2').merge()
    .setValue('Carga = ALMOX SERIALIZADA − TODOS os seriais já encerrados no Forms (histórico) · Somente TT da EQUIPES · Atualizado: ' +
      Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm'))
    .setFontStyle('italic').setHorizontalAlignment('center').setFontSize(10);

  escreverDescricaoRegras_(sh, 'A3:I3',
    'Serial encerrado no Forms = equipamento já instalado no cliente → sai da carga imediatamente (independe da observação BASE/TÉCNICO/PENDENTE da col I). ' +
    'Status = tempo desde o último encerramento: hoje (verde) · 1 dia (amarelo) · 2 dias (laranja) · 3+ dias (vermelho). ' +
    'Fontes: EQUIPES (online) · FORMS de encerradas · ALMOX SERIALIZADA.');

  cabecalhoTabela_(sh, 'A4:I4', [
    'TT', 'Técnico', 'Área', 'Velocidade / Descrição', 'Qtd em carga (líquida)',
    'Encerrados hoje', 'Total encerrados (histórico)', 'Último encerramento', 'Status'
  ]);

  // Agrupa por TT e por velocidade — abatendo serial JÁ ENCERRADO (qualquer dia) da carga
  const porTt = {};
  modems.forEach(function (m) {
    if (!porTt[m.tt]) porTt[m.tt] = { velocidades: {}, total: 0 };
    if (m.serial && encerradosTotal[m.serial]) return; // já no cliente — fora da carga
    porTt[m.tt].velocidades[m.velocidade] = (porTt[m.tt].velocidades[m.velocidade] || 0) + 1;
    porTt[m.tt].total += 1;
  });

  const farolCores = [];
  const linhas = equipe.lista.map(function (tec) {
    const almox = porTt[tec.tt];
    const form  = forms[tec.tt];

    const velocOrdenadas = almox ? Object.keys(almox.velocidades).sort() : [];
    const colVeloc = velocOrdenadas.map(function (v) { return v || '(sem descrição)'; }).join('\n');
    const colQtd   = velocOrdenadas.map(function (v) { return almox.velocidades[v]; }).join('\n');

    const r = resumirFormsTt_(form, hoje, tz);
    const farol = farolEncerramento_(r.diasDesde);
    farolCores.push(farol.cor);

    return [
      tec.tt, tec.nome, tec.area, colVeloc, colQtd,
      r.encerradosHoje || '', r.totalEncerrados || '', farol.txt, farol.status
    ];
  });

  linhas.sort(function (a, b) {
    // ordena por status mais crítico primeiro (3+ dias no topo), depois por carga
    const oa = ordemStatus_(a[8]), ob = ordemStatus_(b[8]);
    if (oa !== ob) return ob - oa;
    return Number(b[4].toString().split('\n')[0] || 0) - Number(a[4].toString().split('\n')[0] || 0);
  });
  // recoloca cores na ordem ordenada
  const corPorTt = {};
  equipe.lista.forEach(function (tec, i) { corPorTt[tec.tt] = farolCores[i]; });

  if (linhas.length) {
    rangeLinhas_(sh, 5, 1, linhas.length, 9).setValues(linhas);
    rangeLinhas_(sh, 5, 4, linhas.length, 1).setWrap(true);
    rangeLinhas_(sh, 5, 5, linhas.length, 1).setWrap(true).setHorizontalAlignment('center');
    for (let i = 0; i < linhas.length; i++) {
      const cor = corPorTt[linhas[i][0]];
      if (cor) {
        sh.getRange(5 + i, 8).setBackground(cor); // Último encerramento
        sh.getRange(5 + i, 9).setBackground(cor); // Status
      }
    }
  }

  [90, 220, 80, 280, 120, 110, 150, 130, 110].forEach(function (w, i) {
    sh.setColumnWidth(i + 1, w);
  });
  sh.setFrozenRows(4);

  // ── Totalizador: carga TCI por equipamento/ONT (soma de todos, líquida) ──
  const totalPorEquip = {};
  modems.forEach(function (m) {
    if (m.serial && encerradosTotal[m.serial]) return; // já encerrado, fora da carga
    const v = m.velocidade || '(sem descrição)';
    totalPorEquip[v] = (totalPorEquip[v] || 0) + 1;
  });

  let rTot = 5 + (linhas.length ? linhas.length : 0) + 2;
  sh.getRange(rTot, 1).setValue('TOTAL TCI EM CARGA POR EQUIPAMENTO / ONT (toda a equipe)')
    .setFontWeight('bold').setFontSize(11);
  rTot++;
  cabecalhoTabela_(sh, 'A' + rTot + ':B' + rTot, ['Equipamento / Velocidade', 'Total TCI em carga']);
  const rTotCab = rTot;
  rTot++;

  const totEquipLinhas = Object.keys(totalPorEquip).sort(function (a, b) {
    return totalPorEquip[b] - totalPorEquip[a];
  }).map(function (v) { return [v, totalPorEquip[v]]; });

  let totalGeral = 0;
  totEquipLinhas.forEach(function (l) { totalGeral += Number(l[1]); });

  if (totEquipLinhas.length) {
    rangeLinhas_(sh, rTot, 1, totEquipLinhas.length, 2).setValues(totEquipLinhas);
    rangeLinhas_(sh, rTot, 2, totEquipLinhas.length, 1).setNumberFormat('#,##0');
    const fimTot = rTot + totEquipLinhas.length - 1;
    sh.getRange(fimTot + 1, 1).setValue('TOTAL GERAL TCI').setFontWeight('bold');
    sh.getRange(fimTot + 1, 2).setValue(totalGeral).setFontWeight('bold').setNumberFormat('#,##0');
    sh.getRange('A' + rTotCab + ':B' + (fimTot + 1)).setBorder(
      true, true, true, true, true, true, '#9e9e9e', SpreadsheetApp.BorderStyle.SOLID);
  }

  SpreadsheetApp.flush();
  ss.toast('Painel modems atualizado.', 'TCI Carga', 8);
}

/**
 * Resume os encerramentos do Forms de um técnico:
 *  - totalEncerrados: nº de seriais encerrados em todo o histórico
 *  - encerradosHoje:  nº de seriais encerrados hoje
 *  - diasDesde:       dias desde o último encerramento (0 = hoje); null = nunca encerrou
 */
function resumirFormsTt_(form, hojeIso, tz) {
  const out = { totalEncerrados: 0, encerradosHoje: 0, diasDesde: null };
  if (!form || !form.instalacoes || !form.instalacoes.length) return out;
  let ultimaData = null;
  form.instalacoes.forEach(function (ins) {
    const n = (ins.seriais && ins.seriais.length) ? ins.seriais.length : 0;
    out.totalEncerrados += n;
    if (ins.data === hojeIso) out.encerradosHoje += n;
    if (ins.data && (!ultimaData || ins.data > ultimaData)) ultimaData = ins.data;
  });
  if (ultimaData) out.diasDesde = diasEntreIso_(ultimaData, hojeIso);
  return out;
}

/** Dias entre duas datas ISO (yyyy-MM-dd). hoje − data. */
function diasEntreIso_(dataIso, hojeIso) {
  const a = new Date(dataIso + 'T00:00:00');
  const b = new Date(hojeIso + 'T00:00:00');
  if (isNaN(a.getTime()) || isNaN(b.getTime())) return null;
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}

/** Farol pelo nº de dias desde o último encerramento no Forms. */
function farolEncerramento_(diasDesde) {
  if (diasDesde === null)   return { txt: 'nunca encerrou', status: '🔴 3+ dias', cor: '#ef9a9a' };
  if (diasDesde <= 0)       return { txt: 'hoje',           status: '🟢 OK',      cor: '#e8f5e9' };
  if (diasDesde === 1)      return { txt: 'há 1 dia',       status: '🟡 1 dia',   cor: '#fff8e1' };
  if (diasDesde === 2)      return { txt: 'há 2 dias',      status: '🟠 2 dias',  cor: '#ffe0b2' };
  return { txt: 'há ' + diasDesde + ' dias', status: '🔴 3+ dias', cor: '#ef9a9a' };
}

/** Peso de ordenação do status (maior = mais crítico, vai pro topo). */
function ordemStatus_(status) {
  const s = String(status || '');
  if (s.indexOf('3+') >= 0)  return 4;
  if (s.indexOf('2 dias') >= 0) return 3;
  if (s.indexOf('1 dia') >= 0)  return 2;
  return 1; // OK hoje
}

/** Serials baixados no Forms na data indicada (qualquer técnico). serial → true */
function coletarSeriaisInstalados_(formsPorTt, dataIso) {
  const set = {};
  Object.keys(formsPorTt).forEach(function (tt) {
    formsPorTt[tt].instalacoes.forEach(function (ins) {
      if (dataIso && ins.data && ins.data !== dataIso) return;
      ins.seriais.forEach(function (s) {
        const v = String(s || '').trim().toUpperCase();
        if (v) set[v] = true;
      });
    });
  });
  return set;
}

function lerModemsAlmox_(cfg, porTt) {
  const ssExt = SpreadsheetApp.openById(cfg.MATERIAIS_ID);
  const sh = ssExt.getSheetByName(cfg.ABA_ALMOX);
  if (!sh) throw new Error('Aba não encontrada: ' + cfg.ABA_ALMOX);

  const last = sh.getLastRow();
  if (last < 2) return [];

  const dados = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
  const header = dados[0].map(normalizarTexto_);

  const ixTt     = indiceColuna_(header, ['f/tt', 'gestech', 'tt']);
  const ixVeloc  = indiceColuna_(header, ['velocidade', 'plano', 'texto breve material', 'descricao material', 'material']);
  const ixSerial = indiceColuna_(header, ['serial', 'numero de serie', 'serie']);

  const out = [];
  for (let i = 1; i < dados.length; i++) {
    const tt = normalizarTt_(dados[i][ixTt]);
    if (!tt || !porTt[tt]) continue;
    const serial = ixSerial >= 0 ? String(dados[i][ixSerial] || '').trim().toUpperCase() : '';
    if (!serial) continue;
    const velocidade = ixVeloc >= 0 ? String(dados[i][ixVeloc] || '').trim() : '';
    // CABO/DROP têm serial mas pertencem à miscelânia — excluir do painel modems
    if (ehCaboDropAlmox_(velocidade)) continue;
    out.push({ tt: tt, nome: porTt[tt].nome, velocidade: velocidade, serial: serial });
  }
  return out;
}

/**
 * Retorna true se a descrição da ALMOX SERIALIZADA for cabo/drop/fibra —
 * esses itens têm serial mas pertencem à miscelânia, não ao painel modems.
 */
function ehCaboDropAlmox_(descricao) {
  const t = String(descricao || '').toUpperCase();
  return t.indexOf('CABO') >= 0 || t.indexOf('DROP') >= 0 ||
         (t.indexOf('FIBRA') >= 0 && t.indexOf('ONT') < 0);
}

// ─── Utilitários ─────────────────────────────────────────────────────────────

function normalizarTt_(v) {
  const s = String(v || '').trim().toUpperCase();
  if (!s || s === 'PENDENTE') return '';
  if (s.indexOf('TR') === 0) return '';      // rejeita TR (TT/TR trocados na planilha)
  if (s.indexOf('TT') === 0) return s;
  if (/^\d+$/.test(s)) return 'TT' + s;
  return s;
}

function normalizarTr_(v) {
  const s = String(v || '').trim().toUpperCase();
  if (!s || s === 'PENDENTE') return '';
  if (s.indexOf('TR') === 0) return s;
  if (s.indexOf('TT') === 0) return '';      // rejeita TT no campo TR
  if (/^\d+$/.test(s)) return 'TR' + s;
  return s;
}

function normalizarTexto_(v) {
  return String(v || '').trim().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '');
}

/**
 * Busca coluna pelo header (candidatos e headers normalizados).
 * Primeiro passa: match exato em todos os candidatos.
 * Segunda passa: substring — evita que 'codmaterial' seja retornado ao buscar 'material'.
 */
function indiceColuna_(header, candidatos) {
  for (let c = 0; c < candidatos.length; c++) {
    const alvo = normalizarTexto_(candidatos[c]);
    for (let i = 0; i < header.length; i++) {
      if (header[i] === alvo) return i;
    }
  }
  for (let c = 0; c < candidatos.length; c++) {
    const alvo = normalizarTexto_(candidatos[c]);
    for (let i = 0; i < header.length; i++) {
      if (header[i].indexOf(alvo) >= 0) return i;
    }
  }
  return -1;
}

function parseNumero_(v) {
  if (typeof v === 'number' && !isNaN(v)) return v;
  const s = String(v || '').replace(/\./g, '').replace(',', '.').replace(/[^\d.-]/g, '');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function formatarDataIso_(v) {
  if (!v) return '';
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return '';
}

function rangeLinhas_(sh, linha, col, numLinhas, numCols) {
  return sh.getRange(linha, col, numLinhas, numCols);
}

function titulo_(sh, range, text) {
  sh.getRange(range).merge().setValue(text)
    .setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center')
    .setBackground('#1565c0').setFontColor('#ffffff');
}

function cabecalhoTabela_(sh, range, values) {
  sh.getRange(range).setValues([values])
    .setFontWeight('bold').setBackground('#e3f2fd').setHorizontalAlignment('center');
}
