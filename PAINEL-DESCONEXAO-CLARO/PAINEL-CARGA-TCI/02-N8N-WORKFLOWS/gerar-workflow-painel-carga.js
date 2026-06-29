/**
 * Gera IMPORTAR-PAINEL-CARGA-GESTECH.json para n8n.
 * Uso: node gerar-workflow-painel-carga.js
 *
 * Correcao BASE TCI (append):
 *   misc/modems  -> limpa a aba inteira e regrava (snapshot do dia)
 *   basetci      -> APPEND (insere a linha sem apagar as ~1900 linhas manuais)
 */
const fs = require('fs');
const path = require('path');

const MATERIAIS_ID = '1W1K7OzbZejmSbEUnQxJ9u43phV0sHYNWkeW3KIbv0vM';
const WAHA_KEY = 'minha_chave_123';
const WAHA_API = 'http://waha:3000';
const CONFIRM_CHAT = '555185670380@c.us';
const SESSION = 'default';

const codeFiltrar = `const API = '${WAHA_API}';
const KEY = '${WAHA_KEY}';
const CONFIRM_CHATS = ['555185670380@c.us', '5551985670380@c.us'];
const SESSION = '${SESSION}';

const item = $input.first().json;
const body = item.body || item;
const payload = body.payload || body;
const session = body.session || payload.session || SESSION;
const from = String(payload.from || '');
const fromMe = payload.fromMe === true;
const event = String(body.event || '');
const hoje = $now.setZone('America/Sao_Paulo').toFormat('dd/MM/yyyy');

function acharGrupo(p) {
  const campos = [
    p.from, p.chatId, p.to,
    p._data && p._data.key && p._data.key.remoteJid,
    p._data && p._data.id && p._data.id.remote,
  ];
  for (let i = 0; i < campos.length; i++) {
    const v = String(campos[i] || '');
    if (v.includes('@g.us')) return v;
  }
  return '';
}

function tipoObrigatorio(nome) {
  const n = String(nome || '').toLowerCase();
  if (/follow.?up|saldo.?gestech|gestech.*follow/i.test(n)) return 'misc';
  if (/serializados|controle.*hc|controle\\s*serializados/i.test(n)) return 'modems';
  return '';
}

function ehPlanilha(ext, mime, nome) {
  if (['xlsx','xlsm','xls'].includes(ext)) return true;
  if (mime.includes('spreadsheet') || mime.includes('excel')) return true;
  if (/follow.?up|saldo.?gestech|serializados|controle.*hc|\\.xls/i.test(nome)) return true;
  return false;
}

const grupoId = acharGrupo(payload);

if (event && !/message/i.test(event)) {
  return [{ json: { skip: true, motivo: 'evento: ' + event, from: from, grupoId: grupoId } }];
}

const type = String(payload.type || payload._data?.type || '').toLowerCase();
const media = payload.media || payload._data?.media || {};
const filename = String(
  media.filename || payload._data?.filename || payload._data?.message?.documentMessage?.fileName || ''
).trim();
const mimetype = String(
  media.mimetype || media.mimeType || payload._data?.message?.documentMessage?.mimetype || ''
).toLowerCase();
const msgId = String(payload.id || payload._data?.id?.id || payload._data?.key?.id || '');
const hasMedia = payload.hasMedia === true || type.includes('document') || !!media.url || (!!msgId && type.includes('document'));
const obrigatorio = tipoObrigatorio(filename) || tipoObrigatorio(mimetype);

if (obrigatorio && !grupoId) {
  throw new Error('OBRIGATORIO bloqueado: planilha ' + obrigatorio + ' sem grupo @g.us — confira WAHA webhook');
}
if (!grupoId && !obrigatorio) {
  return [{ json: { skip: true, motivo: 'nao e grupo (@g.us)', from: from, grupoId: '' } }];
}
if (fromMe && !obrigatorio) {
  return [{ json: { skip: true, motivo: 'fromMe', from: from, grupoId: grupoId } }];
}
if (fromMe && obrigatorio) {
  return [{ json: { skip: true, motivo: 'fromMe (obrigatorio so de outrem no grupo)', from: from, grupoId: grupoId, obrigatorio: obrigatorio } }];
}

if (msgId) {
  const sd = $getWorkflowStaticData('global');
  if (!sd.seenMsg) sd.seenMsg = {};
  if (sd.seenMsg[msgId]) {
    return [{ json: { skip: true, motivo: 'duplicado webhook', from: from, grupoId: grupoId, obrigatorio: obrigatorio } }];
  }
  sd.seenMsg[msgId] = Date.now();
  const keys = Object.keys(sd.seenMsg);
  if (keys.length > 400) {
    keys.sort(function (a, b) { return sd.seenMsg[a] - sd.seenMsg[b]; });
    keys.slice(0, keys.length - 250).forEach(function (k) { delete sd.seenMsg[k]; });
  }
}

if (hasMedia && (media.url || msgId)) {
  let url = media.url ? String(media.url) : '';
  if (url && !url.startsWith('http')) url = API + url;
  if (!url && msgId) {
    url = API + '/api/' + session + '/files/' + encodeURIComponent(msgId);
  }
  const ext = (filename.split('.').pop() || '').toLowerCase();
  if (!ehPlanilha(ext, mimetype, filename)) {
    if (obrigatorio) {
      throw new Error('OBRIGATORIO ' + obrigatorio + ' rejeitado: anexo nao e planilha (' + (filename || type) + ')');
    }
    return [{ json: { skip: true, motivo: 'anexo nao e planilha: ' + (filename || type), from: from, grupoId: grupoId } }];
  }
  let tipo = obrigatorio || 'modems';
  if (!obrigatorio) {
    if (/follow.?up|saldo.?gestech|gestech.*follow/i.test(filename)) tipo = 'misc';
    else if (/serializados|controle.*hc/i.test(filename)) tipo = 'modems';
    else if (/follow.?up|saldo.?gestech/i.test(mimetype + ' ' + filename)) tipo = 'misc';
  }
  return [{ json: {
    skip: false,
    obrigatorio: obrigatorio || '',
    tipo,
    session,
    from,
    grupoId,
    filename,
    downloadUrl: url,
    msgId: msgId,
    hoje: hoje,
    materiaisId: '${MATERIAIS_ID}',
    confirmChat: CONFIRM_CHATS[0],
    confirmChats: CONFIRM_CHATS,
    apiKey: KEY,
  }}];
}

const texto = String(payload.body || payload._data?.body || '').trim();
const pareceBaseTci = /serial\\s*[:\\s]|(?:de|origem)\\s*[:\\s].*para|\\bpara\\s*[:\\s]/i.test(texto);
if (texto.length > 8 && pareceBaseTci) {
  return [{ json: {
    skip: false, tipo: 'basetci', session, from, texto, obrigatorio: '',
    materiaisId: '${MATERIAIS_ID}', confirmChat: CONFIRM_CHATS[0], confirmChats: CONFIRM_CHATS, apiKey: KEY,
  }}];
}

return [{ json: { skip: true, motivo: 'sem planilha nem texto BASE TCI', from: from, grupoId: grupoId, type: type, filename: filename } }];`;

const codeParseMisc = `const XLSX = require('xlsx');
const meta = $('Filtrar Grupo WAHA').first().json;
const bin = $input.first().binary?.data;
if (!bin) throw new Error('Download vazio — confira WAHA e URL do anexo');

const buf = Buffer.from(bin.data, 'base64');
const wb = XLSX.read(buf, {
  type: 'buffer', cellFormula: false, cellHTML: false, cellStyles: false, sheetStubs: false,
});

let sheetName = wb.SheetNames.find(function (n) {
  return String(n).toUpperCase().indexOf('SALDO GESTECH') >= 0;
});
if (!sheetName) {
  throw new Error('Aba SALDO GESTECH nao encontrada. Abas: ' + wb.SheetNames.join(' | '));
}

const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '', raw: false });

function pick(row, aliases) {
  const norm = {};
  Object.keys(row).forEach(function (k) {
    norm[String(k).toLowerCase().replace(/\\s+/g, ' ').trim()] = row[k];
  });
  for (let i = 0; i < aliases.length; i++) {
    const v = norm[aliases[i].toLowerCase()];
    if (v !== undefined && v !== '') return v;
  }
  return '';
}

const header = [
  'CodArmazem','Armazem','CodMaterial','Material',
  'Grupo Material','Agregador','Sub Grupo Agregador','Segmento','Subsegmento','Saldo'
];
const dataRows = [];

rows.forEach(function (row) {
  let tt = String(pick(row, ['codarmazem','cod armazem','tt'])).trim().toUpperCase();
  if (!tt.startsWith('TT')) return;
  const saldoRaw = pick(row, ['saldo']);
  const saldo = Number(String(saldoRaw).replace(/\\./g, '').replace(',', '.')) || Number(saldoRaw) || 0;
  if (saldo <= 0) return;
  dataRows.push([
    tt,
    pick(row, ['armazem']),
    pick(row, ['codmaterial','cod material']),
    pick(row, ['material']),
    pick(row, ['grupo material']),
    pick(row, ['agregador']),
    pick(row, ['sub grupo agregador','subsegmento']),
    pick(row, ['segmento','segmento2','_segmento']),
    pick(row, ['subsegmento','_subsegmento']),
    saldo,
  ]);
});

if (!dataRows.length) {
  if (meta.obrigatorio === 'misc') throw new Error('OBRIGATORIO misc: SALDO GESTECH sem linhas TT');
  throw new Error('SALDO GESTECH: nenhuma linha TT com saldo > 0');
}

return [{
  json: {
    sheetName: 'BASE MICELANEAS',
    documentId: meta.materiaisId,
    filename: meta.filename,
    tipo: 'misc',
    total: dataRows.length,
    header: header,
    rows: dataRows,
    confirmChat: meta.confirmChat,
    apiKey: meta.apiKey,
    session: meta.session,
  }
}];`;

const codeParseModems = `const XLSX = require('xlsx');
const meta = $('Filtrar Grupo WAHA').first().json;
const bin = $input.first().binary?.data;
if (!bin) throw new Error('Download vazio');

const buf = Buffer.from(bin.data, 'base64');
const wb = XLSX.read(buf, {
  type: 'buffer', cellFormula: false, cellHTML: false, cellStyles: false, sheetStubs: false,
});

let sheetName = wb.SheetNames.find(function (n) {
  return String(n).toUpperCase() === 'GESTECH';
});
if (!sheetName) {
  sheetName = wb.SheetNames.find(function (n) {
    return String(n).toUpperCase().indexOf('BASE SERIALIZADOS') >= 0;
  });
}
if (!sheetName) throw new Error('Aba GESTECH / BASE SERIALIZADOS nao encontrada');

const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '', raw: false });
const isBase = String(sheetName).toUpperCase().indexOf('BASE SERIALIZADOS') >= 0;

function pick(row, aliases) {
  const norm = {};
  Object.keys(row).forEach(function (k) {
    norm[String(k).toLowerCase().replace(/\\s+/g, ' ').trim()] = row[k];
  });
  for (let i = 0; i < aliases.length; i++) {
    const v = norm[aliases[i].toLowerCase()];
    if (v !== undefined && v !== '') return v;
  }
  return '';
}

const header = ['Serial','Gestech','Material','Velocidade'];
const dataRows = [];

rows.forEach(function (row) {
  const serial = String(
    pick(row, isBase ? ['n de serie','numero de serie','serial'] : ['serial'])
  ).trim().toUpperCase();
  let tt = String(pick(row, isBase ? ['gestech'] : ['codarm','armazem','gestech','matricula'])).trim().toUpperCase();
  if (!serial || !tt.startsWith('TT')) return;
  const material = pick(row, isBase ? ['texto breve material','material'] : ['material']);
  const vel = pick(row, ['velocidade','texto breve material','material']);
  dataRows.push([serial, tt, material, vel]);
});

if (!dataRows.length) {
  if (meta.obrigatorio === 'modems') throw new Error('OBRIGATORIO modems: GESTECH sem serial TT');
  throw new Error('Modems: nenhum serial TT valido');
}

return [{
  json: {
    sheetName: 'ALMOX SERIALIZADA',
    documentId: meta.materiaisId,
    filename: meta.filename,
    tipo: 'modems',
    total: dataRows.length,
    header: header,
    rows: dataRows,
    confirmChat: meta.confirmChat,
    apiKey: meta.apiKey,
    session: meta.session,
  }
}];`;

const codeParseBaseTci = `const meta = $('Filtrar Grupo WAHA').first().json;
const t = String(meta.texto || '').replace(/\\r/g, '');

function extrair(re, idx) {
  const m = t.match(re);
  return m && m[idx] ? String(m[idx]).trim() : '';
}

const serial = extrair(/serial[:\\s]+([A-Z0-9]+)/i, 1) || extrair(/\\b([A-Z0-9]{8,20})\\b/, 1);
const de = extrair(/(?:de|origem)[:\\s]+(.+?)(?:\\n|para|$)/i, 1);
const para = extrair(/para[:\\s]+(.+?)(?:\\n|$)/i, 1);
const obs = extrair(/obs[:\\s]+(.+)/i, 1) || (t.toUpperCase().indexOf('PENDENTE') >= 0 ? 'PENDENCIA TRANSFERENCIA' : t.substring(0, 120));

if (!serial && !de && !para) {
  throw new Error('Texto BASE TCI nao reconhecido (precisa DE/PARA ou serial)');
}

const hoje = $now.setZone('America/Sao_Paulo').toFormat('dd/MM/yyyy');

return [{
  json: {
    sheetName: 'BASE TCI',
    documentId: meta.materiaisId,
    tipo: 'basetci',
    header: ['Serial','Colaborador E','Receptor F','OK G','Data H','Obs I'],
    rows: [[serial, de, para, '', hoje, obs]],
    confirmChat: meta.confirmChat,
    apiKey: meta.apiKey,
    session: meta.session,
    total: 1,
  }
}];`;

const codePrepararGravar = `const item = $input.first().json;
const meta = $('Filtrar Grupo WAHA').first().json;
const header = item.header || [];
const rows = item.rows || [];
if (!rows.length) {
  if (meta.obrigatorio) {
    throw new Error('OBRIGATORIO ' + meta.obrigatorio + ' falhou: zero linhas para gravar');
  }
  throw new Error('Sem linhas para gravar');
}
// BASE TCI = append (NAO apaga a base manual). misc/modems = substituicao do snapshot.
const append = item.tipo === 'basetci';
const values = append ? rows : [header].concat(rows);
const hoje = $now.setZone('America/Sao_Paulo').toFormat('dd/MM/yyyy HH:mm');
const obr = meta.obrigatorio || '';
const rotulo = obr === 'misc' ? 'FOLLOW UP → BASE MICELANEAS' : (obr === 'modems' ? 'SERIALIZADOS → ALMOX' : item.sheetName);
const prefix = obr ? '🚨 OBRIGATÓRIO DO DIA — ' : '✅ ';
const acao = append ? 'append' : 'substituicao';
const msg = prefix + rotulo + '\\n' + item.sheetName + ' (' + acao + '): ' + rows.length + ' linha(s)\\nArquivo: ' + (item.filename || meta.filename || '') + '\\n' + hoje;
return [{
  json: {
    documentId: item.documentId,
    sheetName: item.sheetName,
    values: values,
    rowCount: rows.length,
    append: append,
    msg: msg,
    obrigatorio: obr,
    confirmChat: meta.confirmChat,
    confirmChats: meta.confirmChats || [meta.confirmChat],
    apiKey: item.apiKey || meta.apiKey,
    session: item.session || meta.session || '${SESSION}',
  }
}];`;

const codeRegistroObrigatorio = `const prep = $('Preparar MATERIAIS').first().json;
const meta = $('Filtrar Grupo WAHA').first().json;
const hojeKey = $now.setZone('America/Sao_Paulo').toFormat('dd/MM/yyyy');
const sd = $getWorkflowStaticData('global');
if (!sd.obrigatorioDia) sd.obrigatorioDia = {};
if (!sd.obrigatorioDia[hojeKey]) sd.obrigatorioDia[hojeKey] = {};
const chave = prep.obrigatorio || meta.obrigatorio;
if (chave === 'misc' || chave === 'modems') {
  sd.obrigatorioDia[hojeKey][chave] = {
    ok: true,
    linhas: prep.rowCount,
    aba: prep.sheetName,
    arquivo: meta.filename || '',
    hora: $now.setZone('America/Sao_Paulo').toFormat('HH:mm'),
  };
}
const keys = Object.keys(sd.obrigatorioDia);
if (keys.length > 14) {
  keys.sort();
  keys.slice(0, keys.length - 10).forEach(function (k) { delete sd.obrigatorioDia[k]; });
}
return [{ json: { ok: true, dia: hojeKey, registro: sd.obrigatorioDia[hojeKey] } }];`;

const CREDS_SHEETS = {
  googleSheetsOAuth2Api: {
    id: 'VINCULAR_NO_N8N',
    name: 'Conta do Google Sheets 2',
  },
};

const workflow = {
  name: 'PAINEL CARGA TCI — Gestech WhatsApp',
  nodes: [
    {
      parameters: {
        httpMethod: 'POST',
        path: 'painel-carga-waha',
        responseMode: 'onReceived',
        options: {},
      },
      type: 'n8n-nodes-base.webhook',
      typeVersion: 2.1,
      position: [-800, 300],
      id: 'pc001-webhook',
      name: 'Webhook WAHA',
      webhookId: 'painel-carga-waha-001',
    },
    {
      parameters: {
        mode: 'runOnceForAllItems',
        jsCode: codeFiltrar,
      },
      type: 'n8n-nodes-base.code',
      typeVersion: 2,
      position: [-560, 300],
      id: 'pc002-filtrar',
      name: 'Filtrar Grupo WAHA',
    },
    {
      parameters: {
        conditions: {
          options: { caseSensitive: true, leftValue: '', typeValidation: 'strict', version: 2 },
          conditions: [{
            id: 'c-skip',
            leftValue: '={{ $json.skip }}',
            rightValue: true,
            operator: { type: 'boolean', operation: 'notEquals' },
          }],
          combinator: 'and',
        },
        options: {},
      },
      type: 'n8n-nodes-base.if',
      typeVersion: 2.2,
      position: [-320, 300],
      id: 'pc003-if',
      name: 'Processar?',
    },
    {
      parameters: {
        rules: {
          values: [
            {
              conditions: {
                options: { caseSensitive: false, typeValidation: 'loose', version: 2 },
                conditions: [{
                  leftValue: '={{ $json.tipo }}',
                  rightValue: 'misc',
                  operator: { type: 'string', operation: 'equals' },
                }],
                combinator: 'and',
              },
              renameOutput: true,
              outputKey: 'misc',
            },
            {
              conditions: {
                options: { caseSensitive: false, typeValidation: 'loose', version: 2 },
                conditions: [{
                  leftValue: '={{ $json.tipo }}',
                  rightValue: 'modems',
                  operator: { type: 'string', operation: 'equals' },
                }],
                combinator: 'and',
              },
              renameOutput: true,
              outputKey: 'modems',
            },
            {
              conditions: {
                options: { caseSensitive: false, typeValidation: 'loose', version: 2 },
                conditions: [{
                  leftValue: '={{ $json.tipo }}',
                  rightValue: 'basetci',
                  operator: { type: 'string', operation: 'equals' },
                }],
                combinator: 'and',
              },
              renameOutput: true,
              outputKey: 'basetci',
            },
          ],
        },
        options: {},
      },
      type: 'n8n-nodes-base.switch',
      typeVersion: 3.2,
      position: [-80, 280],
      id: 'pc004-switch',
      name: 'Tipo entrada',
    },
    {
      parameters: {
        url: '={{ $json.downloadUrl }}',
        sendHeaders: true,
        headerParameters: {
          parameters: [{ name: 'X-Api-Key', value: '={{ $json.apiKey }}' }],
        },
        options: { response: { response: { responseFormat: 'file' } }, timeout: 120000 },
      },
      type: 'n8n-nodes-base.httpRequest',
      typeVersion: 4.2,
      position: [160, 80],
      id: 'pc005-dl-misc',
      name: 'Download planilha misc',
    },
    {
      parameters: {
        url: '={{ $json.downloadUrl }}',
        sendHeaders: true,
        headerParameters: {
          parameters: [{ name: 'X-Api-Key', value: '={{ $json.apiKey }}' }],
        },
        options: { response: { response: { responseFormat: 'file' } }, timeout: 120000 },
      },
      type: 'n8n-nodes-base.httpRequest',
      typeVersion: 4.2,
      position: [160, 280],
      id: 'pc006-dl-modems',
      name: 'Download planilha modems',
    },
    {
      parameters: { mode: 'runOnceForAllItems', jsCode: codeParseMisc },
      type: 'n8n-nodes-base.code',
      typeVersion: 2,
      position: [400, 80],
      id: 'pc007-parse-misc',
      name: 'Valores SALDO GESTECH',
    },
    {
      parameters: { mode: 'runOnceForAllItems', jsCode: codeParseModems },
      type: 'n8n-nodes-base.code',
      typeVersion: 2,
      position: [400, 280],
      id: 'pc008-parse-modems',
      name: 'Valores GESTECH modems',
    },
    {
      parameters: { mode: 'runOnceForAllItems', jsCode: codeParseBaseTci },
      type: 'n8n-nodes-base.code',
      typeVersion: 2,
      position: [160, 480],
      id: 'pc009-parse-btci',
      name: 'Parse texto BASE TCI',
    },
    {
      parameters: { mode: 'runOnceForAllItems', jsCode: codePrepararGravar },
      type: 'n8n-nodes-base.code',
      typeVersion: 2,
      position: [640, 280],
      id: 'pc010-prep',
      name: 'Preparar MATERIAIS',
    },
    {
      parameters: {
        conditions: {
          options: { caseSensitive: true, leftValue: '', typeValidation: 'strict', version: 2 },
          conditions: [{
            id: 'c-append',
            leftValue: '={{ $json.append }}',
            rightValue: true,
            operator: { type: 'boolean', operation: 'equals', singleValue: true },
          }],
          combinator: 'and',
        },
        options: {},
      },
      type: 'n8n-nodes-base.if',
      typeVersion: 2.2,
      position: [820, 280],
      id: 'pc016-if-append',
      name: 'É append (BASE TCI)?',
    },
    {
      parameters: {
        method: 'POST',
        url: "=https://sheets.googleapis.com/v4/spreadsheets/{{ $('Preparar MATERIAIS').item.json.documentId }}/values/{{ encodeURIComponent($('Preparar MATERIAIS').item.json.sheetName + '!A1') }}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS",
        authentication: 'predefinedCredentialType',
        nodeCredentialType: 'googleSheetsOAuth2Api',
        sendBody: true,
        specifyBody: 'json',
        jsonBody: "={{ JSON.stringify({ range: $('Preparar MATERIAIS').item.json.sheetName + '!A1', majorDimension: 'ROWS', values: $('Preparar MATERIAIS').item.json.values }) }}",
        options: {},
      },
      type: 'n8n-nodes-base.httpRequest',
      typeVersion: 4.2,
      position: [1040, 160],
      id: 'pc017-append',
      name: 'Append BASE TCI',
      credentials: CREDS_SHEETS,
    },
    {
      parameters: {
        operation: 'clear',
        documentId: {
          __rl: true,
          value: '={{ $json.documentId }}',
          mode: 'id',
        },
        sheetName: {
          __rl: true,
          value: '={{ $json.sheetName }}',
          mode: 'name',
        },
        clear: 'wholeSheet',
        options: {},
      },
      type: 'n8n-nodes-base.googleSheets',
      typeVersion: 4.7,
      position: [1040, 360],
      id: 'pc011-clear',
      name: 'Limpar aba MATERIAIS',
      credentials: CREDS_SHEETS,
    },
    {
      parameters: {
        method: 'POST',
        url: "=https://sheets.googleapis.com/v4/spreadsheets/{{ $('Preparar MATERIAIS').item.json.documentId }}/values/{{ encodeURIComponent($('Preparar MATERIAIS').item.json.sheetName + '!A1') }}?valueInputOption=USER_ENTERED",
        authentication: 'predefinedCredentialType',
        nodeCredentialType: 'googleSheetsOAuth2Api',
        sendBody: true,
        specifyBody: 'json',
        jsonBody: "={{ JSON.stringify({ range: $('Preparar MATERIAIS').item.json.sheetName + '!A1', majorDimension: 'ROWS', values: $('Preparar MATERIAIS').item.json.values }) }}",
        options: {},
      },
      type: 'n8n-nodes-base.httpRequest',
      typeVersion: 4.2,
      position: [1260, 360],
      id: 'pc012-gravar',
      name: 'Gravar valores MATERIAIS',
      credentials: CREDS_SHEETS,
    },
    {
      parameters: {
        method: 'POST',
        url: `${WAHA_API}/api/sendText`,
        sendHeaders: true,
        headerParameters: {
          parameters: [
            { name: 'X-Api-Key', value: "={{ $('Preparar MATERIAIS').item.json.apiKey }}" },
            { name: 'Content-Type', value: 'application/json' },
          ],
        },
        sendBody: true,
        specifyBody: 'json',
        jsonBody: "={{ JSON.stringify({ session: $('Preparar MATERIAIS').item.json.session, chatId: $('Preparar MATERIAIS').item.json.confirmChat, text: $('Preparar MATERIAIS').item.json.msg }) }}",
        options: {},
      },
      type: 'n8n-nodes-base.httpRequest',
      typeVersion: 4.2,
      position: [1480, 280],
      id: 'pc013-aviso',
      name: 'Aviso WhatsApp',
    },
    {
      parameters: { mode: 'runOnceForAllItems', jsCode: codeRegistroObrigatorio },
      type: 'n8n-nodes-base.code',
      typeVersion: 2,
      position: [1700, 280],
      id: 'pc015-registro',
      name: 'Registro OBRIGATORIO dia',
    },
    {
      parameters: {},
      type: 'n8n-nodes-base.noOp',
      typeVersion: 1,
      position: [-320, 480],
      id: 'pc014-skip',
      name: 'Ignorado',
    },
  ],
  connections: {
    'Webhook WAHA': { main: [[{ node: 'Filtrar Grupo WAHA', type: 'main', index: 0 }]] },
    'Filtrar Grupo WAHA': { main: [[{ node: 'Processar?', type: 'main', index: 0 }]] },
    'Processar?': {
      main: [
        [{ node: 'Tipo entrada', type: 'main', index: 0 }],
        [{ node: 'Ignorado', type: 'main', index: 0 }],
      ],
    },
    'Tipo entrada': {
      main: [
        [{ node: 'Download planilha misc', type: 'main', index: 0 }],
        [{ node: 'Download planilha modems', type: 'main', index: 0 }],
        [{ node: 'Parse texto BASE TCI', type: 'main', index: 0 }],
      ],
    },
    'Download planilha misc': { main: [[{ node: 'Valores SALDO GESTECH', type: 'main', index: 0 }]] },
    'Download planilha modems': { main: [[{ node: 'Valores GESTECH modems', type: 'main', index: 0 }]] },
    'Valores SALDO GESTECH': { main: [[{ node: 'Preparar MATERIAIS', type: 'main', index: 0 }]] },
    'Valores GESTECH modems': { main: [[{ node: 'Preparar MATERIAIS', type: 'main', index: 0 }]] },
    'Parse texto BASE TCI': { main: [[{ node: 'Preparar MATERIAIS', type: 'main', index: 0 }]] },
    'Preparar MATERIAIS': { main: [[{ node: 'É append (BASE TCI)?', type: 'main', index: 0 }]] },
    'É append (BASE TCI)?': {
      main: [
        [{ node: 'Append BASE TCI', type: 'main', index: 0 }],
        [{ node: 'Limpar aba MATERIAIS', type: 'main', index: 0 }],
      ],
    },
    'Append BASE TCI': { main: [[{ node: 'Aviso WhatsApp', type: 'main', index: 0 }]] },
    'Limpar aba MATERIAIS': { main: [[{ node: 'Gravar valores MATERIAIS', type: 'main', index: 0 }]] },
    'Gravar valores MATERIAIS': { main: [[{ node: 'Aviso WhatsApp', type: 'main', index: 0 }]] },
    'Aviso WhatsApp': { main: [[{ node: 'Registro OBRIGATORIO dia', type: 'main', index: 0 }]] },
  },
  pinData: {},
  settings: { executionOrder: 'v1', timezone: 'America/Sao_Paulo' },
  staticData: null,
  tags: [],
  meta: { templateCredsSetupCompleted: false },
  versionId: 'painel-carga-gestech-v5-append-basetci',
};

const out = path.join(__dirname, 'IMPORTAR-PAINEL-CARGA-GESTECH.json');
fs.writeFileSync(out, JSON.stringify(workflow, null, 2), 'utf8');
console.log('OK:', out);
