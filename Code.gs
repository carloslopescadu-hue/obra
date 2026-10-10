/**
 * OBRA DA CASA — backend (Google Apps Script)
 * A página fica no GitHub Pages e conversa com este script.
 * Dados: Planilha Google · Fotos e áudios: Google Drive · Avisos: Gmail
 * Transcrição do áudio: Groq (Whisper) · Sugestão de execução e materiais: API do Claude
 *
 * Segurança: o script usa só a permissão "drive.file" (enxerga apenas a planilha, a pasta e os
 * arquivos que ele mesmo criou), mais envio de e-mail, chamadas externas e gatilhos.
 * Planilha e Drive são acessados pelas APIs REST do Google com essa permissão restrita.
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
  'tituloAuto', 'avisos', 'valorCobrado', 'diferenca', 'historico'
];
const JSON_FIELDS = ['fotos', 'audios', 'materiais', 'sugestao', 'avisos', 'historico'];
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
    limite_('create', 40, 3600, 'Muitos serviços criados em pouco tempo. Tente de novo mais tarde.');
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
        avisos: { novo: now }, historico: [{ tipo: 'criado', por: clean_(fields.criadoPor, 40), em: now }]
      };
      appendRow_(it);
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
    if (!findRow_(id)) throw new Error('Este serviço foi apagado.');
    limite_('upload', 80, 3600, 'Muitos arquivos enviados em pouco tempo. Tente de novo mais tarde.');
    const f = saveDataUrl_(dataUrl, nome || 'foto.jpg', 8 * 1024 * 1024, /^image\/(jpeg|png|webp|gif)$/);
    shareAnyone_(f.id);
    return mutate_(id, function (it) { it.fotos.push({ id: f.id, nome: f.name }); });
  },
  /** Só apaga o arquivo se ele pertencer a este serviço. */
  removePhoto: function (id, fileId) {
    return mutate_(id, function (it) {
      const tinha = it.fotos.some(function (f) { return f.id === fileId; });
      it.fotos = it.fotos.filter(function (f) { return f.id !== fileId; });
      if (tinha) trash_(fileId);
    });
  },
  addAudio: function (id, dataUrl, duracao, nome) {
    if (num_(duracao) > 65) throw new Error('O áudio pode ter no máximo 1 minuto.');
    if (!findRow_(id)) throw new Error('Este serviço foi apagado.');
    limite_('upload', 80, 3600, 'Muitos arquivos enviados em pouco tempo. Tente de novo mais tarde.');
    limite_('audio', 40, 3600, 'Muitos áudios em pouco tempo. Tente de novo mais tarde.');
    const f = saveDataUrl_(dataUrl, nome || 'audio.webm', 8 * 1024 * 1024, /^(audio|video)\/[\w.+-]+$/);
    let transcricao = '', erro = '';
    try { transcricao = transcribe_(f.blob); } catch (e) { erro = e.message; }
    return mutate_(id, function (it) {
      it.audios.push({
        id: f.id, duracao: Math.round(num_(duracao)), mime: f.mimeType,
        criadoEm: new Date().toISOString(), transcricao: transcricao, erroTranscricao: erro
      });
    });
  },
  /** Só apaga o arquivo se ele pertencer a este serviço. */
  removeAudio: function (id, fileId) {
    return mutate_(id, function (it) {
      const tinha = it.audios.some(function (a) { return a.id === fileId; });
      it.audios = it.audios.filter(function (a) { return a.id !== fileId; });
      if (tinha) trash_(fileId);
    });
  },
  /** Só entrega áudios que pertencem a algum serviço da obra. */
  getAudio: function (fileId) {
    const dono = readAll_().some(function (it) { return it.audios.some(function (a) { return a.id === fileId; }); });
    if (!dono) throw new Error('Áudio não encontrado.');
    const blob = getBlob_(fileId);
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  },
  checkPin: function (pin) { return guardPin_(pin); },
  /** Qualquer pessoa pode aprovar; cada mudança fica no histórico com o nome de quem fez. */
  setStatus: function (id, status, por) {
    if (STATUSES.indexOf(status) === -1) throw new Error('Status inválido.');
    limite_('status', 120, 3600, 'Muitas mudanças de aprovação em pouco tempo. Tente de novo mais tarde.');
    const nome = clean_(por, 40);
    if (!nome) throw new Error('Informe seu nome para registrar a aprovação.');
    let mudou = false;
    const it = mutate_(id, function (x) {
      if (x.status === status) return;
      x.historico.push({ tipo: 'status', de: x.status, para: status, por: nome, em: new Date().toISOString() });
      if (x.historico.length > 100) x.historico = x.historico.slice(-100);
      x.status = status; mudou = true;
    });
    if (mudou) { try { avisoStatus_(it, nome); } catch (e) { Logger.log('Aviso de status falhou: ' + e.message); } }
    return it;
  },
  /** Ajuste em lote pedido pelo dono: limpa datas e volta todos para "aguardando". Exige PIN. */
  resetAll: function (pin, por) {
    if (!guardPin_(pin)) throw new Error('PIN incorreto.');
    const nome = clean_(por, 40) || nomeDono_();
    readAll_().forEach(function (item) {
      mutate_(item.id, function (x) {
        x.inicio = ''; x.termino = '';
        if (x.status !== 'aguardando') {
          x.historico.push({ tipo: 'status', de: x.status, para: 'aguardando', por: nome + ' (ajuste geral)', em: new Date().toISOString() });
          x.status = 'aguardando';
        }
      });
    });
    return readAll_();
  },
  remove: function (id, pin) {
    if (!guardPin_(pin)) throw new Error('PIN incorreto.');
    return withLock_(function () {
      const r = findRow_(id);
      if (!r) return true;
      r.item.fotos.concat(r.item.audios).forEach(function (f) { trash_(f.id); });
      deleteRow_(r.row);
      return true;
    });
  },
  suggest: function (id) {
    limite_('sug:' + id, 1, 60, 'Aguarde um minuto antes de pedir outra sugestão para este serviço.');
    limite_('sug', 20, 3600, 'Limite de sugestões por hora atingido. Tente de novo mais tarde.');
    return suggest_(id);
  }
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
      const blob = getBlob_(f.id);
      content.push({ type: 'image', source: { type: 'base64', media_type: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) } });
    } catch (e) { /* foto apagada */ }
  });
  content.push({ type: 'text', text: buildPrompt_(it, content.length) });

  let s;
  try {
    const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: prop_('CLAUDE_MODEL') || 'claude-sonnet-5-5', max_tokens: 1500, messages: [{ role: 'user', content: content }] })
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
      diagnostico: clean_(s.diagnostico, 400),
      passos: (s.passos || []).slice(0, 7).map(function (p) { return { titulo: clean_(p.titulo, 60), detalhe: clean_(p.detalhe, 220) }; }),
      cuidados: (s.cuidados || []).slice(0, 3).map(function (c) { return clean_(c, 160); }),
      prazoDias: num_(s.prazoDias),
      geradoEm: new Date().toISOString()
    };
    x.sugestaoStatus = 'ok';
    const keep = x.materiais.filter(function (m) { return !m.sugerido || m.usar || num_(m.valor) > 0; });
    const names = keep.map(function (m) { return String(m.nome).toLowerCase(); });
    (s.materiais || []).slice(0, 12).forEach(function (m) {
      const nome = clean_(m.nome, 120);
      if (!nome || names.indexOf(nome.toLowerCase()) !== -1) return;
      keep.push({ id: Utilities.getUuid().slice(0, 6), nome: nome, qtd: clean_(m.quantidade, 40), obs: clean_(m.obs, 60), usar: false, valor: 0, sugerido: true });
    });
    x.materiais = keep;
  });
  try { enviarAvisos(done.id); } catch (e) { Logger.log('Aviso falhou: ' + e.message); }
  return done;
}

function buildPrompt_(it, nFotos) {
  const falas = it.audios.filter(function (a) { return a.transcricao; }).map(function (a, i) { return 'ÁUDIO ' + (i + 1) + ' (transcrição automática): ' + a.transcricao; });
  return [
    'Você é um mestre de obras brasileiro experiente. O morador relatou um problema na casa; o prestador vai ler sua sugestão no celular para orçar e executar.',
    '',
    'TÍTULO: ' + (it.tituloAuto ? '(ainda não definido)' : it.titulo),
    'DESCRIÇÃO ESCRITA: ' + (it.descricao || '(nenhuma)'),
    falas.length ? falas.join('\n') : 'Sem áudio.',
    nFotos ? ('Há ' + nFotos + ' foto(s); use-as no diagnóstico.') : 'Sem foto.',
    '',
    'Seja OBJETIVO e CURTO. Cada informação aparece uma única vez: não repita no passo o que já está no diagnóstico, nem nos cuidados o que já está nos passos.',
    '',
    'Responda SOMENTE com um objeto JSON, sem texto antes ou depois:',
    '{"titulo": "até 50 caracteres, ex.: Infiltração na parede do quarto",',
    ' "descricao": "o problema em 1 frase, na voz do morador",',
    ' "diagnostico": "1 ou 2 frases curtas (máx. 220 caracteres): causa provável e o que conferir antes",',
    ' "passos": [{"titulo": "verbo no infinitivo, até 4 palavras", "detalhe": "1 frase curta, máx. 110 caracteres"}],',
    ' "cuidados": ["máx. 90 caracteres cada; só segurança/normas que não estejam nos passos"],',
    ' "prazoDias": número de dias, incluindo secagem ou cura,',
    ' "materiais": [{"nome": "nome curto e usual, sem marca", "quantidade": "estimativa simples, ex.: 1 galão 3,6 L", "obs": "opcional, até 40 caracteres"}]}',
    '',
    'Limites: 4 a 6 passos; 1 a 3 cuidados; 3 a 10 materiais (só o necessário). Sem preços. Português do Brasil, linguagem simples de obra.',
    'A transcrição pode ter erros de fala: interprete pelo contexto.'
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
function avisoStatus_(it, por) {
  const para = por === nomePrestador_() ? emailDono_() : emailPrestador_();
  sendMail_(para, STATUS_LABEL[it.status] + ': ' + osNum_(it) + ' · ' + it.titulo,
    (por || nomeDono_()) + ' marcou este serviço como "' + STATUS_LABEL[it.status] + '".', it,
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
  if (!props.getProperty('SHEET_ID')) {
    const ss = gapi_('post', 'https://sheets.googleapis.com/v4/spreadsheets', { properties: { title: 'Obra da Casa — base de dados' }, sheets: [{ properties: { title: SHEET_NAME, gridProperties: { frozenRowCount: 1 } } }] });
    props.setProperty('SHEET_ID', ss.spreadsheetId);
  }
  if (!props.getProperty('FOLDER_ID')) props.setProperty('FOLDER_ID', createFolder_('Obra da Casa — fotos e áudios'));
  const defaults = {
    OWNER_PIN: function () { return String(Math.floor(100000 + Math.random() * 900000)); },
    NOME_DONO: function () { return 'Carlos'; },
    NOME_PRESTADOR: function () { return 'Júlio'; },
    PAGE_URL: function () { return 'https://carloslopescadu-hue.github.io/obra/'; }
  };
  Object.keys(defaults).forEach(function (k) { if (!props.getProperty(k)) props.setProperty(k, defaults[k]()); });
  ensureHeader_();

  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'enviarAvisos') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('enviarAvisos').timeBased().everyMinutes(5).create();
  MailApp.getRemainingDailyQuota();

  Logger.log('Planilha: https://docs.google.com/spreadsheets/d/' + props.getProperty('SHEET_ID') + '/edit');
  Logger.log('Avisos para: dono ' + (props.getProperty('EMAIL_DONO') || 'FALTA') + ' · prestador ' + (props.getProperty('EMAIL_PRESTADOR') || 'FALTA'));
  ['ANTHROPIC_API_KEY', 'GROQ_API_KEY', 'PAGE_URL'].forEach(function (k) { Logger.log(k + ': ' + (props.getProperty(k) ? 'ok' : 'FALTA configurar')); });
}

/**
 * Migração única para a permissão restrita (drive.file): copia a planilha, as fotos e os áudios
 * para uma planilha e uma pasta novas, criadas pelo próprio app. Os originais ficam intactos
 * (renomeados como "antiga") como cópia de segurança. Rode uma vez com o acesso amplo liberado.
 */
function migrarParaAcessoRestrito() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('MIGRADO_EM')) { Logger.log('Já migrado em ' + props.getProperty('MIGRADO_EM')); return; }
  const velhaPlanilha = props.getProperty('SHEET_ID'), velhaPasta = props.getProperty('FOLDER_ID');
  const antigos = readAll_();
  Logger.log('Serviços lidos da planilha antiga: ' + antigos.length);

  const nova = gapi_('post', 'https://sheets.googleapis.com/v4/spreadsheets', { properties: { title: 'Obra da Casa — base de dados' }, sheets: [{ properties: { title: SHEET_NAME, gridProperties: { frozenRowCount: 1 } } }] });
  const pasta = createFolder_('Obra da Casa — fotos e áudios');
  let copiados = 0, falhas = 0;
  const copiar = function (id, publico) {
    try {
      const f = gapi_('post', DRIVE_API + 'files/' + encodeURIComponent(id) + '/copy?fields=id', { parents: [pasta] });
      if (publico) shareAnyone_(f.id);
      copiados++; return f.id;
    } catch (e) { falhas++; Logger.log('Não copiei ' + id + ': ' + e.message); return id; }
  };
  const linhas = antigos.map(function (it) {
    it.fotos = it.fotos.map(function (f) { f.id = copiar(f.id, true); return f; });
    it.audios = it.audios.map(function (a) { a.id = copiar(a.id, false); return a; });
    return itemToRow_(it);
  });
  const total = linhas.length + 1;
  gapi_('put', SHEETS_API + nova.spreadsheetId + '/values/' + rng_('A1:' + LAST_COL_ + total) + '?valueInputOption=RAW', { values: [HEADERS].concat(linhas) });

  try { gapi_('patch', DRIVE_API + 'files/' + velhaPlanilha + '?fields=id', { name: 'Obra da Casa — base de dados (antiga, cópia de segurança)' }); } catch (e) {}
  try { gapi_('patch', DRIVE_API + 'files/' + velhaPasta + '?fields=id', { name: 'Obra da Casa — fotos e áudios (antiga, cópia de segurança)' }); } catch (e) {}
  props.setProperties({ SHEET_ID_ANTIGA: velhaPlanilha, FOLDER_ID_ANTIGA: velhaPasta, SHEET_ID: nova.spreadsheetId, FOLDER_ID: pasta, MIGRADO_EM: new Date().toISOString() });
  HEADER_OK_ = false;
  const conferido = readAll_().length;
  Logger.log('Nova planilha: https://docs.google.com/spreadsheets/d/' + nova.spreadsheetId + '/edit');
  Logger.log('Arquivos copiados: ' + copiados + ' · falhas: ' + falhas + ' · serviços na nova planilha: ' + conferido);
}

/** Rode pelo menu Executar para conferir se o script ainda alcança a planilha, a pasta e os arquivos. */
function verificarAcesso() {
  const itens = readAll_();
  Logger.log('Planilha OK: ' + itens.length + ' serviços.');
  const pasta = gapi_('get', DRIVE_API + 'files/' + prop_('FOLDER_ID') + '?fields=id,name');
  Logger.log('Pasta OK: ' + pasta.name);
  let testados = 0, falhas = 0;
  itens.forEach(function (it) {
    it.audios.concat(it.fotos).slice(0, 1).forEach(function (f) {
      testados++;
      try { gapi_('get', DRIVE_API + 'files/' + f.id + '?fields=id'); } catch (e) { falhas++; Logger.log('Sem acesso ao arquivo ' + f.id + ': ' + e.message); }
    });
  });
  Logger.log('Arquivos testados: ' + testados + ' · sem acesso: ' + falhas);
}

/* ---------- Planilha ---------- */

const SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets/';
const DRIVE_API = 'https://www.googleapis.com/drive/v3/';
function sid_() { const id = prop_('SHEET_ID'); if (!id) throw new Error('Rode a função setup antes de usar.'); return id; }
function colA1_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function rng_(a1) { return encodeURIComponent(SHEET_NAME + '!' + a1); }
const LAST_COL_ = colA1_(HEADERS.length);

let HEADER_OK_ = false;
function ensureHeader_() {
  if (HEADER_OK_) return;
  const r = gapi_('get', SHEETS_API + sid_() + '/values/' + rng_('A1:' + LAST_COL_ + '1'));
  const atual = (r.values && r.values[0]) || [];
  if (atual.join('|') !== HEADERS.join('|')) {
    gapi_('put', SHEETS_API + sid_() + '/values/' + rng_('A1:' + LAST_COL_ + '1') + '?valueInputOption=RAW', { values: [HEADERS] });
  }
  HEADER_OK_ = true;
}
function rows_() {
  ensureHeader_();
  const r = gapi_('get', SHEETS_API + sid_() + '/values/' + rng_('A2:' + LAST_COL_) + '?valueRenderOption=UNFORMATTED_VALUE');
  return (r.values || []).map(function (row) { while (row.length < HEADERS.length) row.push(''); return row; });
}

function readAll_() {
  return rows_().map(rowToItem_).filter(function (x) { return x.id; })
    .sort(function (a, b) { return a.numero - b.numero; });
}

function findRow_(id) {
  const rows = rows_();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) return { row: i + 2, item: rowToItem_(rows[i]) };
  }
  return null;
}

function writeRow_(row, it) {
  const a1 = 'A' + row + ':' + LAST_COL_ + row;
  gapi_('put', SHEETS_API + sid_() + '/values/' + rng_(a1) + '?valueInputOption=RAW', { values: [itemToRow_(it)] });
}
function appendRow_(it) {
  ensureHeader_();
  gapi_('post', SHEETS_API + sid_() + '/values/' + rng_('A1:' + LAST_COL_) + ':append?valueInputOption=RAW&insertDataOption=INSERT_ROWS', { values: [itemToRow_(it)] });
}
function deleteRow_(row) {
  const meta = gapi_('get', SHEETS_API + sid_() + '?fields=sheets.properties');
  const sh = (meta.sheets || []).map(function (x) { return x.properties; }).filter(function (p) { return p.title === SHEET_NAME; })[0];
  if (!sh) throw new Error('Aba da planilha não encontrada.');
  gapi_('post', SHEETS_API + sid_() + ':batchUpdate', { requests: [{ deleteDimension: { range: { sheetId: sh.sheetId, dimension: 'ROWS', startIndex: row - 1, endIndex: row } } }] });
}

function mutate_(id, fn, keepTimestamp) {
  return withLock_(function () {
    const r = findRow_(id);
    if (!r) throw new Error('Este serviço foi apagado.');
    fn(r.item);
    computeTotals_(r.item);
    if (!keepTimestamp) r.item.atualizadoEm = new Date().toISOString();
    writeRow_(r.row, r.item);
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
    if (JSON_FIELDS.indexOf(h) !== -1) {
      try { v = v ? JSON.parse(v) : null; } catch (e) { v = null; }
    }
    it[h] = v;
  });
  it.fotos = it.fotos || []; it.audios = it.audios || []; it.materiais = it.materiais || []; it.avisos = it.avisos || {}; it.historico = it.historico || [];
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
/** PIN com trava: 5 erros em 15 minutos bloqueiam novas tentativas por 15 minutos. */
function guardPin_(pin) {
  const c = CacheService.getScriptCache();
  if (c.get('pin:lock')) throw new Error('Muitas tentativas de PIN. Tente de novo em 15 minutos.');
  if (checkPin_(pin)) { c.remove('pin:fail'); return true; }
  const n = Number(c.get('pin:fail') || 0) + 1;
  c.put('pin:fail', String(n), 900);
  if (n >= 5) c.put('pin:lock', '1', 900);
  return false;
}
/** Limite simples por janela de tempo (contador no cache do script). */
function limite_(chave, max, segundos, msg) {
  const c = CacheService.getScriptCache();
  const k = 'lim:' + chave + ':' + Math.floor(Date.now() / 1000 / segundos);
  const n = Number(c.get(k) || 0);
  if (n >= max) throw new Error(msg);
  c.put(k, String(n + 1), segundos + 60);
}
/** Chamada às APIs do Google com o token restrito do próprio script. */
function gapi_(method, url, body) {
  const params = { method: method, muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } };
  if (body !== undefined) { params.contentType = 'application/json'; params.payload = JSON.stringify(body); }
  const res = UrlFetchApp.fetch(url, params);
  const code = res.getResponseCode(), txt = res.getContentText();
  if (code >= 300) throw new Error('Google ' + code + ': ' + txt.slice(0, 180));
  return txt ? JSON.parse(txt) : {};
}
function createFolder_(name) {
  return gapi_('post', DRIVE_API + 'files?fields=id', { name: name, mimeType: 'application/vnd.google-apps.folder' }).id;
}
function createFile_(blob, name, folderId) {
  const b = 'obra' + Utilities.getUuid().replace(/-/g, '');
  const type = blob.getContentType() || 'application/octet-stream';
  const meta = JSON.stringify({ name: name, parents: [folderId], mimeType: type });
  const head = Utilities.newBlob('--' + b + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta + '\r\n--' + b + '\r\nContent-Type: ' + type + '\r\n\r\n').getBytes();
  const tail = Utilities.newBlob('\r\n--' + b + '--').getBytes();
  const res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,mimeType', {
    method: 'post', muteHttpExceptions: true, contentType: 'multipart/related; boundary=' + b,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: head.concat(blob.getBytes(), tail)
  });
  if (res.getResponseCode() >= 300) throw new Error('Não consegui salvar o arquivo: ' + res.getContentText().slice(0, 180));
  return JSON.parse(res.getContentText());
}
function shareAnyone_(fileId) {
  gapi_('post', DRIVE_API + 'files/' + fileId + '/permissions?fields=id', { role: 'reader', type: 'anyone' });
}
function getBlob_(fileId) {
  const res = UrlFetchApp.fetch(DRIVE_API + 'files/' + encodeURIComponent(fileId) + '?alt=media', {
    muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  });
  if (res.getResponseCode() >= 300) throw new Error('Arquivo indisponível.');
  return res.getBlob();
}
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
function saveDataUrl_(dataUrl, nome, maxBytes, tipoOk) {
  const m = String(dataUrl).match(/^data:([^;,]+)[^,]*;base64,(.+)$/);
  if (!m) throw new Error('Arquivo inválido.');
  if (tipoOk && !tipoOk.test(m[1])) throw new Error('Tipo de arquivo não aceito.');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > maxBytes) throw new Error('Arquivo grande demais.');
  const blob = Utilities.newBlob(bytes, m[1], clean_(nome, 80).replace(/[^\w.\- ]/g, '_') || 'arquivo');
  const f = createFile_(blob, blob.getName(), prop_('FOLDER_ID'));
  return { id: f.id, name: f.name, mimeType: f.mimeType || m[1], blob: blob };
}
function trash_(fileId) { try { gapi_('patch', DRIVE_API + 'files/' + encodeURIComponent(fileId) + '?fields=id', { trashed: true }); } catch (e) { /* já removido ou fora do alcance */ } }
