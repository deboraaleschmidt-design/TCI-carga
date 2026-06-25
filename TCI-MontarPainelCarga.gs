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
    .addItem('Ativar atualização automática diária (10h30)', 'ativarGatilhoDiario')
    .addItem('Ativar atualização frequente (a cada 15 min)', 'ativarAtualizacaoFrequente')
    .addItem('Desativar atualização automática', 'desativarGatilhoDiario')
    .addToUi();
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
  ABA_EQUIPES:  'Página4',
  GID_EQUIPES:  '1514437459',
  ABA_FORMS:    'Respostas ao formulário 1'
};

function lerConfig_() {
  const ss = obterPlanilha_();
  const sh = ss.getSheetByName('CONFIG');
  if (!sh) return Object.assign({}, CFG_PADRAO_);

  const vals = sh.getRange('A2:B14').getValues();
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
    ABA_FORMS:    map.ABA_FORMS    || CFG_PADRAO_.ABA_FORMS
  };
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
    ['', ''],
    ['Regra ONT',    'Excluir Grupo=ONT ou subsegmento FIBRA ONT (SERIAL)'],
    ['Técnicos',     'Somente TT listados na planilha EQUIPES (Página4)'],
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
    sh.getRange('A1:I1').setValues([[
      'Data snapshot', 'TT', 'Nome técnico', 'Cód. material', 'Material',
      'Grupo material', 'Agregador', 'Saldo', 'Segmento'
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
    lista.push(item);
  }
  return { lista: lista, porTt: porTt, porTr: porTr };
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
    return [hoje, r.tt, r.nome, r.codmaterial, r.material, r.grupo, r.agregador, r.saldo, r.segmento];
  });

  if (saida.length) {
    const start = shHist.getLastRow() + 1;
    rangeLinhas_(shHist, start, 1, saida.length, 9).setValues(saida);
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
    segmento: indiceColuna_(header, ['segmento'])
  };

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

/** Unidade de medida para o painel: DROP em metros, demais em unidades. */
function medidaMaterial_(material, agregador, grupo) {
  return ehDrop_(material, agregador, grupo) ? 'metros' : 'un';
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
    if (!tr) continue;

    const tec = equipe.porTr[tr];
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

// ─── Painel miscelânia (ontem × hoje) ────────────────────────────────────────

function montarPainelMiscelania() {
  const ss = obterPlanilha_();
  const cfg = lerConfig_();
  const equipe = carregarEquipeTci_(cfg);
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

  titulo_(sh, 'A1:K1', 'PAINEL MISCELANIA TCI — ONTEM × HOJE (por TT)');
  const txtComp = dataOntem
    ? ('Comparação: ' + dataOntem + ' → ' + dataHoje)
    : ('1º snapshot: ' + dataHoje + ' — rode amanhã para ver ontem × hoje');
  sh.getRange('A2:K2').merge()
    .setValue(txtComp + '  |  Atualizado: ' + Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'
    ))
    .setFontStyle('italic').setHorizontalAlignment('center').setFontSize(10);

  let row = 4;
  row = escreverResumoDrop_(sh, row, snapshots, dataOntem, dataHoje, equipe.porTt);

  sh.getRange(row, 1).setValue('DETALHE POR TÉCNICO E MATERIAL').setFontWeight('bold').setFontSize(11);
  row++;
  const cab = [
    'TT', 'Técnico', 'Área', 'Segmento', 'Material', 'Grupo', 'Medida',
    'Saldo ontem', 'Saldo hoje', 'Variação', 'Situação'
  ];
  cabecalhoTabela_(sh, 'A' + row + ':K' + row, cab);
  const linhaCab = row;
  row++;

  const chaves = {};
  equipe.lista.forEach(function (tec) {
    Object.keys(snapshots.porData[dataHoje] || {}).forEach(function (k) {
      if (k.indexOf(tec.tt + '\t') === 0) chaves[k] = true;
    });
    if (dataOntem) {
      Object.keys(snapshots.porData[dataOntem] || {}).forEach(function (k) {
        if (k.indexOf(tec.tt + '\t') === 0) chaves[k] = true;
      });
    }
  });

  const linhas = [];
  Object.keys(chaves).sort().forEach(function (chave) {
    const partes = chave.split('\t');
    const tt  = partes[0];
    const tec = equipe.porTt[tt];
    if (!tec) return;

    const snapHoje  = (snapshots.porData[dataHoje]  && snapshots.porData[dataHoje][chave])  || null;
    const snapOntem = dataOntem && snapshots.porData[dataOntem]
      ? snapshots.porData[dataOntem][chave]
      : null;

    const saldoHoje  = snapHoje  ? snapHoje.saldo  : 0;
    const saldoOntem = snapOntem ? snapOntem.saldo : 0;
    const ref      = snapHoje || snapOntem;
    const variacao = saldoHoje - saldoOntem;
    const situacao = classificarVariacao_(saldoOntem, saldoHoje);

    linhas.push([
      tt,
      tec.nome,
      tec.area,
      ref.segmento || '',
      ref.material,
      ref.grupo,
      medidaMaterial_(ref.material, ref.agregador, ref.grupo),
      dataOntem ? saldoOntem : '',
      saldoHoje,
      dataOntem ? variacao   : '',
      dataOntem ? situacao   : 'AGUARDANDO 2º DIA'
    ]);
  });

  if (linhas.length) {
    const fim = row + linhas.length - 1;
    rangeLinhas_(sh, row, 1, linhas.length, 11).setValues(linhas);
    rangeLinhas_(sh, row, 8, linhas.length, 3).setNumberFormat('#,##0');
    colorirSituacao_(sh, row, fim, 11);
    sh.getRange('A' + linhaCab + ':K' + fim).setBorder(
      true, true, true, true, true, true, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID
    );
  }

  [90, 220, 80, 100, 260, 110, 50, 90, 90, 80, 110].forEach(function (w, i) {
    sh.setColumnWidth(i + 1, w);
  });
  sh.setFrozenRows(linhaCab);
  SpreadsheetApp.flush();
  ss.toast('Painel miscelânia atualizado.', 'TCI Carga', 8);
}

function escreverResumoDrop_(sh, row, snapshots, dataOntem, dataHoje, porTt) {
  sh.getRange(row, 1).setValue('RESUMO DROP (metros / unidades em estoque)').setFontWeight('bold').setFontSize(11);
  row++;
  cabecalhoTabela_(sh, 'A' + row + ':H' + row, [
    'TT', 'Técnico', 'DROP ontem', 'DROP hoje', 'Variação', 'Situação', 'Consumiu?', 'Obs'
  ]);
  const linCab = row;
  row++;

  const resumo = [];
  Object.keys(porTt).sort().forEach(function (tt) {
    const dropOntem = somarDropTecnico_(snapshots, dataOntem, tt);
    const dropHoje  = somarDropTecnico_(snapshots, dataHoje,  tt);
    if (dropOntem <= 0 && dropHoje <= 0) return;

    const variacao = dropHoje - dropOntem;
    const sit = classificarVariacao_(dropOntem, dropHoje);
    resumo.push([
      tt,
      porTt[tt].nome,
      dataOntem ? dropOntem : '',
      dropHoje,
      dataOntem ? variacao : '',
      dataOntem ? sit : 'AGUARDANDO 2º DIA',
      dataOntem && variacao < 0 ? 'SIM' : (dataOntem && variacao === 0 ? 'NÃO' : ''),
      variacao < 0 ? 'Baixou ' + Math.abs(variacao) + ' vs ontem' : ''
    ]);
  });

  if (resumo.length) {
    const fim = row + resumo.length - 1;
    rangeLinhas_(sh, row, 1, resumo.length, 8).setValues(resumo);
    rangeLinhas_(sh, row, 3, resumo.length, 3).setNumberFormat('#,##0');
    colorirSituacao_(sh, row, fim, 6);
    sh.getRange('A' + linCab + ':H' + fim).setBorder(
      true, true, true, true, true, true, '#9e9e9e', SpreadsheetApp.BorderStyle.SOLID
    );
    row = fim + 3;
  } else {
    row += 2;
  }
  return row;
}

function somarDropTecnico_(snapshots, data, tt) {
  if (!data || !snapshots.porData[data]) return 0;
  let total = 0;
  Object.keys(snapshots.porData[data]).forEach(function (chave) {
    if (chave.indexOf(tt + '\t') !== 0) return;
    const item = snapshots.porData[data][chave];
    if (ehDrop_(item.material, item.agregador, item.grupo)) total += item.saldo;
  });
  return total;
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
  const numCols = Math.min(shHist.getLastColumn(), 9);
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
      segmento:    String(row[8] || '')
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
  const hoje   = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const ontem  = Utilities.formatDate(new Date(Date.now() - 86400000), tz, 'yyyy-MM-dd');

  // Serials baixados HOJE no Forms — usados para abater da carga (serial = equipamento único)
  const instaladosHoje = coletarSeriaisInstalados_(forms, hoje);

  const nome = 'PAINEL MODEMS';
  let old = ss.getSheetByName(nome);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(nome, 0);

  // 10 colunas: TT | Técnico | Área | Velocidade | Qtd em carga (líquida) | Total em carga | Baixados hoje | Total Forms (histórico) | Preencheu hoje/ontem? | Status
  titulo_(sh, 'A1:J1', 'PAINEL MODEMS TCI — CARGA POR TT · controle de preenchimento Forms');
  sh.getRange('A2:J2').merge()
    .setValue('Carga líquida = ALMOX menos serials já baixados no Forms hoje · Somente TT da EQUIPES · Atualizado: ' +
      Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm'))
    .setFontStyle('italic').setHorizontalAlignment('center').setFontSize(10);

  cabecalhoTabela_(sh, 'A4:J4', [
    'TT', 'Técnico', 'Área', 'Velocidade / Descrição', 'Qtd em carga (líquida)',
    'Total em carga', 'Baixados hoje', 'Total Forms (histórico)', 'Preencheu hoje/ontem?', 'Status'
  ]);

  // Agrupa por TT e por velocidade — abatendo serial já instalado no Forms hoje
  const porTt = {};
  modems.forEach(function (m) {
    if (!porTt[m.tt]) porTt[m.tt] = { velocidades: {}, total: 0, baixados: 0 };
    if (m.serial && instaladosHoje[m.serial]) {
      porTt[m.tt].baixados += 1;
      return;
    }
    porTt[m.tt].velocidades[m.velocidade] = (porTt[m.tt].velocidades[m.velocidade] || 0) + 1;
    porTt[m.tt].total += 1;
  });

  const linhas = equipe.lista.map(function (tec) {
    const almox = porTt[tec.tt];
    const form  = forms[tec.tt];

    const totalCarga  = almox ? almox.total   : 0;
    const baixados    = almox ? almox.baixados : 0;
    const velocOrdenadas = almox ? Object.keys(almox.velocidades).sort() : [];
    const colVeloc = velocOrdenadas.map(function (v) { return v || '(sem descrição)'; }).join('\n');
    const colQtd   = velocOrdenadas.map(function (v) { return almox.velocidades[v]; }).join('\n');

    // Total de preenchimentos históricos no Forms (todos os dias)
    const totalForms = form ? form.instalacoes.length : 0;

    // Preencheu hoje OU ontem?
    let preencheuStatus = 'NÃO PREENCHEU';
    if (form && form.instalacoes.length > 0) {
      const preencheuHoje  = form.instalacoes.some(function (ins) { return ins.data === hoje; });
      const preencheuOntem = form.instalacoes.some(function (ins) { return ins.data === ontem; });
      if (preencheuHoje)       preencheuStatus = 'SIM — hoje';
      else if (preencheuOntem) preencheuStatus = 'SIM — ontem';
      else                     preencheuStatus = 'NÃO (últimos 2 dias)';
    }

    return [
      tec.tt, tec.nome, tec.area, colVeloc, colQtd,
      totalCarga, baixados || '', totalForms || '', preencheuStatus, tec.status
    ];
  });

  linhas.sort(function (a, b) { return Number(b[5]) - Number(a[5]); });

  if (linhas.length) {
    rangeLinhas_(sh, 5, 1, linhas.length, 10).setValues(linhas);
    rangeLinhas_(sh, 5, 6, linhas.length, 1).setNumberFormat('#,##0');
    rangeLinhas_(sh, 5, 4, linhas.length, 1).setWrap(true);
    rangeLinhas_(sh, 5, 5, linhas.length, 1).setWrap(true).setHorizontalAlignment('center');
    // Colorir coluna "Preencheu hoje/ontem?" (col I = índice 9)
    colorirPreenchimento_(sh, 5, 5 + linhas.length - 1, 9);
  }

  [90, 220, 80, 260, 110, 100, 100, 110, 140, 90].forEach(function (w, i) {
    sh.setColumnWidth(i + 1, w);
  });
  sh.setFrozenRows(4);

  // ── Totalizador: carga TCI por equipamento/ONT (soma de todos os técnicos) ──
  const totalPorEquip = {};
  modems.forEach(function (m) {
    if (m.serial && instaladosHoje[m.serial]) return; // já baixado hoje, fora da carga
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

function colorirPreenchimento_(sh, rowIni, rowFim, col) {
  const range = sh.getRange(rowIni, col, rowFim - rowIni + 1, 1);
  const vals  = range.getValues();
  for (let i = 0; i < vals.length; i++) {
    const s = String(vals[i][0] || '').toUpperCase();
    let bg = null;
    if (s.indexOf('SIM — HOJE') >= 0)   bg = '#e8f5e9';  // verde
    else if (s.indexOf('SIM — ONTEM') >= 0) bg = '#fff8e1'; // amarelo
    else if (s.indexOf('NÃO') >= 0)     bg = '#ffebee';  // vermelho
    if (bg) sh.getRange(rowIni + i, col).setBackground(bg);
  }
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
    // ALMOX SERIALIZADA = equipamento com serial. Sem serial (ex.: cabo) não entra no painel modems.
    if (!serial) continue;
    const velocidade = ixVeloc >= 0 ? String(dados[i][ixVeloc] || '').trim() : '';
    out.push({ tt: tt, nome: porTt[tt].nome, velocidade: velocidade, serial: serial });
  }
  return out;
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
