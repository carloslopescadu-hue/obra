/**
 * OBRA DA CASA — backend (Google Apps Script)
 * A página fica no GitHub Pages e conversa com este script.
 * Dados: Planilha Google · Fotos e áudios: Google Drive · Avisos: Gmail
 * Transcrição do áudio: Groq (Whisper) · Sugestão de execução e materiais: API do Claude
 *
 * Configuração (uma vez):
 *   1. Rode `setup` (menu Executar) e autorize. Ela cria a planilha, a pasta no Drive,
 *      o PIN de aprovação e o aviso automático por e-mail. O PIN aparece no Registro de execução.
 *   2. Configurações do projeto > Propriedades do script:
 *        ANTHROPIC_API_KEY = chave da API do Claude (console.anthropic.com)
 *        GROQ_API_KEY      = chave da Groq (console.groq.com)
 *      Opcionais: EMAIL_DONO, EMAIL_PRESTADOR, NOME_DONO, NOME_PRESTADOR, PAGE_URL,
 *                 OWNER_PIN, CLAUDE_MODEL (padrão claude-sonnet-5-5), AUTO_SUGGEST=false.
 *   3. Implantar > Nova implantação > App da Web — Executar como: Eu · Acesso: Qualquer pessoa.
 *      Copie a URL que termina em /exec para a página (CONFIG.API_URL).
 */

const SHEET_NAME = 'Servicos';
const HEADERS = [
  'id', 'numero', 'criadoEm', 'atualizadoEm', 'criadoPor', 'titulo', 'descricao',
  'maoDeObra', 'materialManual', 'materialTotal', 'total', 'inicio', 'termino',
  'status', 'observacoes', 'fotos', 'audios', 'materiais', 'sugestao', 'sugestaoStatus',
  'tituloAuto', 'avisos', 'valorCobrado', 'diferenca'
];
const JSON_FIELDS = ['fotos', 'audios', 'materiais', 'sugestao', 'avisos'];
const EDITABLE = ['titulo', 'descricao', 'maoDeObra', 'materialManual', 'valorCobrado', 'inicio', 'termino', 'observacoes', 'materiais'];
const VALUE_FIELDS = ['maoDeObra', 'materialManual', 'valorCobrado', 'materiais'];
const STATUSES = ['aguardando', 'aprovado', 'nao_aprovado'];
const STATUS_LABEL = { aguardando: 'Aguardando', aprovado: 'Aprovado', nao_aprovado: 'Não aprovado' };
const TITULO_PROVISORIO = 'Novo serviço';

/* ---------- Web app ---------- */

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, app: 'Obra da Casa' })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let out;
  try {
    const body = JSON.parse(e.postData.contents);
    out = { ok: true, data: api(body.action, body.args || []) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function api(action, args) {
  const fn = ROUTES[action];
  if (!fn) throw new Error('Ação desconhecida: ' + action);
  return fn.apply(null, args || []);
}

const ROUTES = {
  list: function () {
    return {
      items: readAll_(),
      hasClaude: !!prop_('ANTHROPIC_API_KEY'),
      hasTranscription: !!prop_('GROQ_API_KEY'),
      autoSuggest: prop_('AUTO_SUGGEST') !== 'false',
      dono: nomeDono_(), prestador: nomePrestador_()
    };
  },
  create: function (fields) {
    return withLock_(function () {
      const items = readAll_();
      const now = new Date().toISOString();
      const titulo = clean_(fields.titulo, 140);
      const it = {
        id: Utilities.getUuid().slice(0, 8),
        numero: items.reduce(function (m, x) { return Math.max(m, Number(x.numero) || 0); }, 0) + 1,
        criadoEm: now, atualizadoEm: now,
        criadoPor: clean_(fields.criadoPor, 40),
        titulo: titulo || TITULO_PROVISORIO, tituloAuto: !titulo,
        descricao: clean_(fields.descricao, 4000),
        maoDeObra: 0, materialManual: 0, materialTotal: 0, total: 0, valorCobrado: '', diferenca: 0,
        inicio: '', termino: '', status: 'aguardando', observacoes: '',
        fotos: [], audios: [], materiais: [], sugestao: null, sugestaoStatus: '',
        avisos: { novo: now }
      };
      sheet_().appendRow(itemToRow_(it));
      return it;
    });
  },
  update: function (id, patch, actor) {
    return mutate_(id, function (it) {
      let valores = false;
      Object.keys(patch || {}).forEach(function (k) {
        if (EDITABLE.indexOf(k) === -1) return;
        let v = patch[k];
        if (k === 'maoDeObra' || k === 'materialManual') v = num_(v);
        else if (k === 'valorCobrado') v = String(v == null ? '' : v).trim() === '' ? '' : num_(v);
        else if (k === 'inicio' || k === 'termino') v = /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : '';
        else if (k === 'materiais') v = sanitizeMaterials_(v);
        else v = clean_(v, k === 'titulo' ? 140 : 4000);
        if (k === 'titulo') it.tituloAuto = false;
        if (VALUE_FIELDS.indexOf(k) !== -1) valores = true;
        it[k] = v;
      });
      if (valores && clean_(actor, 40) !== nomeDono_()) {
        it.avisos = it.avisos || {};
        it.avisos.orcamento = new Date().toISOString();
      }
    });
  },
  addPhoto: function (id, dataUrl, nome) {
    const file = saveDataUrl_(dataUrl, nome || 'foto.jpg', 8 * 1024 * 1024);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return mutate_(id, function (it) { it.fotos.push({ id: file.getId(), nome: file.getName() }); });
  },
  removePhoto: function (id, fileId) {
    return mutate_(id, function (it) {
      it.fotos = it.fotos.filter(function (f) { return f.id !== fileId; });
      trash_(fileId);
    });
  },
  addAudio: function (id, dataUrl, duracao, nome) {
    if (num_(duracao) > 65) throw new Error('O áudio pode ter no máximo 1 minuto.');
    const file = saveDataUrl_(dataUrl, nome || 'audio.webm', 8 * 1024 * 1024);
    let transcricao = '', erro = '';
    try { transcricao = transcribe_(file.getBlob()); } catch (e) { erro = e.message; }
    return mutate_(id, function (it) {
      it.audios.push({
        id: file.getId(), duracao: Math.round(num_(duracao)), mime: file.getMimeType(),
        criadoEm: new Date().toISOString(), transcricao: transcricao, erroTranscricao: erro
      });
    });
  },
  removeAudio: function (id, fileId) {
    return mutate_(id, function (it) {
      it.audios = it.audios.filter(function (a) { return a.id !== fileId; });
      trash_(fileId);
    });
  },
  getAudio: function (fileId) {
    const blob = DriveApp.getFileById(fileId).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  },
  checkPin: function (pin) { return checkPin_(pin); },
  setStatus: function (id, status, pin) {
    if (!checkPin_(pin)) throw new Error('PIN incorreto.');
    if (STATUSES.indexOf(status) === -1) throw new Error('Status inválido.');
    const it = mutate_(id, function (x) { x.status = status; });
    try { avisoStatus_(it); } catch (e) { Logger.log('Aviso de status falhou: ' + e.message); }
    return it;
  },
  remove: function (id, pin) {
    if (!checkPin_(pin)) throw new Error('PIN incorreto.');
    return withLock_(function () {
      const r = findRow_(id);
      if (!r) return true;
      r.item.fotos.concat(r.item.audios).forEach(function (f) { trash_(f.id); });
      sheet_().deleteRow(r.row);
      return true;
    });
  },
  suggest: function (id) { return suggest_(id); }
};

/* ---------- Transcrição (Groq · Whisper) ---------- */

function transcribe_(blob) {
  const key = prop_('GROQ_API_KEY');
  if (!key) return '';
  const res = UrlFetchApp.fetch('https://api.groq.com/openai/v1/audio/transcriptions', {
    method: 'post', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + key },
    payload: { file: blob, model: 'whisper-large-v3-turbo', language: 'pt', response_format: 'json', temperature: '0' }
  });
  const body = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200) throw new Error('Transcrição falhou: ' + ((body.error && body.error.message) || res.getResponseCode()));
  return clean_(body.text, 4000);
}

/* ---------- Sugestão do Claude ---------- */

function suggest_(id) {
  const key = prop_('ANTHROPIC_API_KEY');
  if (!key) throw new Error('A chave da API do Claude ainda não foi configurada.');
  const r = findRow_(id);
  if (!r) throw new Error('Serviço não encontrado.');
  const it = r.item;
  mutate_(id, function (x) { x.sugestaoStatus = 'gerando'; });

  const content = [];
  it.fotos.slice(0, 3).forEach(function (f) {
    try {
      const blob = DriveApp.getFileById(f.id).getBlob();
      content.push({ type: 'image', source: { type: 'base64', media_type: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) } });
    } catch (e) { /* foto apagada */ }
  });
  content.push({ type: 'text', text: buildPrompt_(it, content.length) });

  let s;
  try {
    const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: prop_('CLAUDE_MODEL') || 'claude-sonnet-5-5', max_tokens: 3500, messages: [{ role: 'user', content: content }] })
    });
    const body = JSON.parse(res.getContentText());
    if (res.getResponseCode() !== 200) throw new Error((body.error && body.error.message) || ('HTTP ' + res.getResponseCode()));
    s = parseJson_(body.content.filter(function (c) { return c.type === 'text'; }).map(function (c) { return c.text; }).join('\n'));
  } catch (err) {
    mutate_(id, function (x) { x.sugestaoStatus = 'erro'; });
    throw new Error('Não consegui gerar a sugestão: ' + err.message);
  }

  const done = mutate_(id, function (x) {
    if (x.tituloAuto && clean_(s.titulo, 140)) { x.titulo = clean_(s.titulo, 140); x.tituloAuto = false; }
    if (!x.descricao && clean_(s.descricao, 4000)) x.descricao = clean_(s.descricao, 4000);
    x.sugestao = {
      diagnostico: clean_(s.diagnostico, 2000),
      passos: (s.passos || []).slice(0, 14).map(function (p) { return { titulo: clean_(p.titulo, 120), detalhe: clean_(p.detalhe, 800) }; }),
      cuidados: (s.cuidados || []).slice(0, 8).map(function (c) { return clean_(c, 400); }),
      prazoDias: num_(s.prazoDias),
      geradoEm: new Date().toISOString()
    };
    x.sugestaoStatus = 'ok';
    const keep = x.materiais.filter(function (m) { return !m.sugerido || m.usar || num_(m.valor) > 0; });
    const names = keep.map(function (m) { return String(m.nome).toLowerCase(); });
    (s.materiais || []).slice(0, 25).forEach(function (m) {
      const nome = clean_(m.nome, 120);
      if (!nome || names.indexOf(nome.toLowerCase()) !== -1) return;
      keep.push({ id: Utilities.getUuid().slice(0, 6), nome: nome, qtd: clean_(m.quantidade, 60), obs: clean_(m.obs, 200), usar: false, valor: 0, sugerido: true });
    });
    x.materiais = keep;
  });
  try { enviarAvisos(done.id); } catch (e) { Logger.log('Aviso falhou: ' + e.message); }
  return done;
}

function buildPrompt_(it, nFotos) {
  const falas = it.audios.filter(function (a) { return a.transcricao; }).map(function (a, i) { return 'ÁUDIO ' + (i + 1) + ' (transcrição automática): ' + a.transcricao; });
  return [
    'Você é um mestre de obras brasileiro experiente (reformas, alvenaria, elétrica, hidráulica, pintura, impermeabilização).',
    'O morador de uma casa relatou um problema. O prestador de serviços vai ler sua sugestão para orçar e executar.',
    '',
    'TÍTULO: ' + (it.tituloAuto ? '(ainda não definido)' : it.titulo),
    'DESCRIÇÃO ESCRITA: ' + (it.descricao || '(nenhuma)'),
    falas.length ? falas.join('\n') : 'Sem áudio.',
    nFotos ? ('Há ' + nFotos + ' foto(s) anexada(s); use-as no diagnóstico.') : 'Sem foto.',
    '',
    'Responda SOMENTE com um objeto JSON, sem texto antes ou depois, neste formato:',
    '{"titulo": "título curto do serviço, até 60 caracteres, ex.: Infiltração na parede do quarto",',
    ' "descricao": "o problema relatado, em 1 a 3 frases claras, na voz do morador",',
    ' "diagnostico": "2 a 4 frases: o que provavelmente é, causas prováveis e o que verificar antes de começar",',
    ' "passos": [{"titulo": "verbo no infinitivo, curto", "detalhe": "como fazer, em 1 a 3 frases práticas"}],',
    ' "cuidados": ["segurança, normas (ex.: NBR 5410 em elétrica), erros comuns"],',
    ' "prazoDias": número estimado de dias de execução, incluindo secagem ou cura,',
    ' "materiais": [{"nome": "material com especificação usual no Brasil", "quantidade": "estimativa com unidade, ou \\"a definir\\"", "obs": "curta, opcional"}]}',
    '',
    'Regras: 5 a 10 passos na ordem de execução; 2 a 5 cuidados; 4 a 15 materiais (inclua consumíveis);',
    'não informe preços; linguagem simples de obra; português do Brasil.',
    'A transcrição pode ter erros de reconhecimento de fala: interprete pelo contexto.'
  ].join('\n');
}

function parseJson_(text) {
  const fence = text.match(/```(?:json)?\s*([\s\S]*?)```/);
  return JSON.parse(fence ? fence[1] : text.slice(text.indexOf('{'), text.lastIndexOf('}') + 1));
}

/* ---------- Avisos por e-mail ---------- */

/** Roda a cada 5 minutos (gatilho criado no setup). Também é chamada logo após cada sugestão. */
function enviarAvisos(soId) {
  const agora = Date.now();
  const items = readAll_();
  items.forEach(function (it) {
    if (soId && typeof soId === 'string' && it.id !== soId) return;
    const av = it.avisos || {};
    const enviar = [];
    const pronto = it.sugestaoStatus === 'ok' || it.sugestaoStatus === 'erro';
    if (av.novo && (pronto || agora - Date.parse(av.novo) > 3 * 60000)) enviar.push('novo');
    if (av.orcamento && agora - Date.parse(av.orcamento) > 5 * 60000) enviar.push('orcamento');
    if (!enviar.length) return;
    enviar.forEach(function (tipo) {
      try {
        if (tipo === 'novo') avisoNovo_(it); else avisoOrcamento_(it);
      } catch (e) { Logger.log('Falha no aviso ' + tipo + ' de ' + it.id + ': ' + e.message); }
    });
    mutate_(it.id, function (x) {
      x.avisos = x.avisos || {};
      enviar.forEach(function (tipo) { if (x.avisos[tipo] === av[tipo]) delete x.avisos[tipo]; });
    }, true);
  });
}

function avisoNovo_(it) {
  const doDono = it.criadoPor === nomeDono_();
  const para = doDono ? emailPrestador_() : emailDono_();
  const quem = it.criadoPor || 'Alguém';
  sendMail_(para, 'Novo serviço: ' + osNum_(it) + ' · ' + it.titulo,
    quem + ' cadastrou um serviço novo na obra.',
    it, doDono ? 'Abra para ver a sugestão do Claude e preencher mão de obra, materiais e datas.' : 'Abra para ver e aprovar quando o orçamento estiver pronto.');
}
function avisoOrcamento_(it) {
  sendMail_(emailDono_(), 'Orçamento atualizado: ' + osNum_(it) + ' · ' + it.titulo,
    nomePrestador_() + ' preencheu ou alterou o orçamento.', it, 'Abra para conferir e aprovar.');
}
function avisoStatus_(it) {
  sendMail_(emailPrestador_(), STATUS_LABEL[it.status] + ': ' + osNum_(it) + ' · ' + it.titulo,
    nomeDono_() + ' marcou este serviço como "' + STATUS_LABEL[it.status] + '".', it,
    it.status === 'aprovado' ? 'Pode programar a execução nas datas combinadas.' : 'Abra para ver os detalhes.');
}
function sendMail_(to, subject, intro, it, cta) {
  if (!to) return;
  const brl = function (n) { return 'R$ ' + Number(n || 0).toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.'); };
  const datas = it.inicio ? (fmtData_(it.inicio) + (it.termino ? ' a ' + fmtData_(it.termino) : '')) : 'a definir';
  const link = (prop_('PAGE_URL') || '') + '#os-' + it.id;
  const linha = function (k, v) { return '<tr><td style="padding:6px 12px 6px 0;color:#59645F">' + k + '</td><td style="padding:6px 0;font-weight:600">' + esc_(v) + '</td></tr>'; };
  const html =
    '<div style="font-family:Arial,Helvetica,sans-serif;color:#17201C;max-width:560px">' +
    '<div style="height:6px;background:repeating-linear-gradient(-45deg,#F2B705 0 12px,#17201C 12px 24px)"></div>' +
    '<p style="font-size:15px">' + esc_(intro) + '</p>' +
    '<p style="font-size:12px;color:#59645F;margin:0">' + osNum_(it) + '</p>' +
    '<h2 style="margin:2px 0 8px;font-size:20px">' + esc_(it.titulo) + '</h2>' +
    (it.descricao ? '<p style="color:#59645F;margin:0 0 12px">' + esc_(it.descricao) + '</p>' : '') +
    '<table style="border-collapse:collapse;font-size:14px">' +
    linha('Status', STATUS_LABEL[it.status]) + linha('Mão de obra', brl(it.maoDeObra)) + linha('Material', brl(it.materialTotal)) +
    linha('Total', brl(it.total)) +
    (it.valorCobrado !== '' ? linha('Valor cobrado', brl(it.valorCobrado)) + linha('Diferença', brl(it.diferenca)) : '') +
    linha('Execução', datas) + '</table>' +
    '<p style="font-size:14px">' + esc_(cta) + '</p>' +
    (prop_('PAGE_URL') ? '<p><a href="' + link + '" style="display:inline-block;background:#F2B705;color:#17201C;padding:10px 18px;border-radius:8px;font-weight:700;text-decoration:none">Abrir ' + osNum_(it) + '</a></p>' : '') +
    '</div>';
  MailApp.sendEmail({ to: to, subject: subject, htmlBody: html, name: 'Obra da Casa' });
}

/* ---------- Setup ---------- */

function setup() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SHEET_ID');
  if (!ssId) {
    ssId = SpreadsheetApp.create('Obra da Casa — base de dados').getId();
    props.setProperty('SHEET_ID', ssId);
  }
  const ss = SpreadsheetApp.openById(ssId);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.getSheets()[0]; sh.setName(SHEET_NAME); }
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(1, 1, sh.getMaxRows(), HEADERS.length).setNumberFormat('@');

  const defaults = {
    FOLDER_ID: function () { return DriveApp.createFolder('Obra da Casa — fotos e áudios').getId(); },
    OWNER_PIN: function () { return String(Math.floor(100000 + Math.random() * 900000)); },
    EMAIL_DONO: function () { return Session.getEffectiveUser().getEmail(); },
    EMAIL_PRESTADOR: function () { return Session.getEffectiveUser().getEmail(); },
    NOME_DONO: function () { return 'Carlos'; },
    NOME_PRESTADOR: function () { return 'Júlio'; },
    PAGE_URL: function () { return 'https://carloslopescadu-hue.github.io/obra/'; }
  };
  Object.keys(defaults).forEach(function (k) { if (!props.getProperty(k)) props.setProperty(k, defaults[k]()); });

  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'enviarAvisos') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('enviarAvisos').timeBased().everyMinutes(5).create();
  MailApp.getRemainingDailyQuota(); // pede a autorização de e-mail já no setup

  Logger.log('Planilha: ' + ss.getUrl());
  Logger.log('SEU PIN DE APROVAÇÃO: ' + props.getProperty('OWNER_PIN'));
  Logger.log('Avisos para: dono ' + props.getProperty('EMAIL_DONO') + ' · prestador ' + props.getProperty('EMAIL_PRESTADOR'));
  ['ANTHROPIC_API_KEY', 'GROQ_API_KEY', 'PAGE_URL'].forEach(function (k) {
    Logger.log(k + ': ' + (props.getProperty(k) ? 'ok' : 'FALTA configurar'));
  });
}

/* ---------- Planilha ---------- */

let SHEET_CACHE_ = null;
function sheet_() {
  if (SHEET_CACHE_) return SHEET_CACHE_;
  const id = prop_('SHEET_ID');
  if (!id) throw new Error('Rode a função setup antes de usar.');
  const sh = SpreadsheetApp.openById(id).getSheetByName(SHEET_NAME);
  if (sh.getLastColumn() < HEADERS.length) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.getRange(1, 1, sh.getMaxRows(), HEADERS.length).setNumberFormat('@');
  }
  SHEET_CACHE_ = sh;
  return sh;
}

function readAll_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues()
    .map(rowToItem_).filter(function (x) { return x.id; })
    .sort(function (a, b) { return a.numero - b.numero; });
}

function findRow_(id) {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      const row = i + 2;
      return { row: row, item: rowToItem_(sh.getRange(row, 1, 1, HEADERS.length).getValues()[0]) };
    }
  }
  return null;
}

function mutate_(id, fn, keepTimestamp) {
  return withLock_(function () {
    const r = findRow_(id);
    if (!r) throw new Error('Este serviço foi apagado.');
    fn(r.item);
    computeTotals_(r.item);
    if (!keepTimestamp) r.item.atualizadoEm = new Date().toISOString();
    sheet_().getRange(r.row, 1, 1, HEADERS.length).setValues([itemToRow_(r.item)]);
    return r.item;
  });
}

function computeTotals_(it) {
  const priced = (it.materiais || []).filter(function (m) { return m.usar && num_(m.valor) > 0; });
  it.materialTotal = priced.length ? priced.reduce(function (s, m) { return s + num_(m.valor); }, 0) : num_(it.materialManual);
  it.total = num_(it.maoDeObra) + it.materialTotal;
  const temCobrado = it.valorCobrado !== '' && it.valorCobrado != null;
  it.diferenca = temCobrado ? Math.round((it.total - num_(it.valorCobrado)) * 100) / 100 : 0;
}

function rowToItem_(row) {
  const it = {};
  HEADERS.forEach(function (h, i) {
    let v = row[i];
    if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (JSON_FIELDS.indexOf(h) !== -1) {
      try { v = v ? JSON.parse(v) : null; } catch (e) { v = null; }
    }
    it[h] = v;
  });
  it.fotos = it.fotos || []; it.audios = it.audios || []; it.materiais = it.materiais || []; it.avisos = it.avisos || {};
  ['numero', 'maoDeObra', 'materialManual', 'materialTotal', 'total', 'diferenca'].forEach(function (k) { it[k] = num_(it[k]); });
  it.valorCobrado = (it.valorCobrado === '' || it.valorCobrado == null) ? '' : num_(it.valorCobrado);
  it.tituloAuto = it.tituloAuto === true || String(it.tituloAuto).toLowerCase() === 'true';
  it.id = String(it.id || '');
  return it;
}

function itemToRow_(it) {
  return HEADERS.map(function (h) {
    const v = it[h];
    if (JSON_FIELDS.indexOf(h) !== -1) return v == null ? '' : JSON.stringify(v);
    if (typeof v === 'boolean') return v ? 'true' : 'false';
    return v == null ? '' : v;
  });
}

/* ---------- Utilidades ---------- */

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function nomeDono_() { return prop_('NOME_DONO') || 'Carlos'; }
function nomePrestador_() { return prop_('NOME_PRESTADOR') || 'Júlio'; }
function emailDono_() { return prop_('EMAIL_DONO'); }
function emailPrestador_() { return prop_('EMAIL_PRESTADOR'); }
function osNum_(it) { return 'OS ' + ('00' + it.numero).slice(-3); }
function fmtData_(s) { const p = String(s).split('-'); return p.length === 3 ? p[2] + '/' + p[1] : s; }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function checkPin_(pin) { const p = prop_('OWNER_PIN'); return !!p && String(pin || '').trim() === p; }
function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let s = String(v == null ? '' : v).replace(/[^\d,.\-]/g, '');
  if (s.indexOf(',') !== -1) s = s.replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isFinite(n) ? Math.round(n * 100) / 100 : 0;
}
function clean_(v, max) { return String(v == null ? '' : v).trim().slice(0, max || 1000); }
function sanitizeMaterials_(list) {
  return (Array.isArray(list) ? list : []).slice(0, 60).map(function (m) {
    return {
      id: clean_(m.id, 12) || Utilities.getUuid().slice(0, 6), nome: clean_(m.nome, 120), qtd: clean_(m.qtd, 60),
      obs: clean_(m.obs, 200), usar: !!m.usar, valor: num_(m.valor), sugerido: !!m.sugerido
    };
  }).filter(function (m) { return m.nome || m.qtd || m.valor; });
}
function saveDataUrl_(dataUrl, nome, maxBytes) {
  const m = String(dataUrl).match(/^data:([^;,]+)[^,]*;base64,(.+)$/);
  if (!m) throw new Error('Arquivo inválido.');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > maxBytes) throw new Error('Arquivo grande demais.');
  return DriveApp.getFolderById(prop_('FOLDER_ID')).createFile(Utilities.newBlob(bytes, m[1], nome));
}
function trash_(fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* já removido */ } }
