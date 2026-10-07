/**
 * VIRTHVS pre-order intake
 * Container-bound script: create it from inside the Google Sheet
 * (Extensions > Apps Script). Receives orders from the order page,
 * validates them, writes one row per order and sends the emails.
 *
 * Run setup() once before the first deployment.
 */

/* ------------------------------------------------------------------
   CONFIG: edit only this block (keep it in sync with the order page)
------------------------------------------------------------------- */
const CONFIG = {
  BRAND_NAME: 'VIRTHVS',
  PIECE_NAME: '[Nome da peça]',
  PRICE: null,                       // número em EUR, por exemplo 35. null deixa o total vazio.
  SIZES: ['S', 'M', 'L', 'XL'],      // tem de coincidir com a página do formulário
  MAX_PIECES: 3,                     // peças por encomenda, cada uma com o seu tamanho
  DELIVERY_FROM_LABEL: '[data]',     // por exemplo '25 de outubro de 2026'
  SHEET_NAME: 'Encomendas',
  DAILY_ORDER_LIMIT: 40,             // limite diário de segurança contra abusos
  EMAIL_COOLDOWN_SECONDS: 120,       // o mesmo email não pode enviar duas vezes dentro deste tempo
  NOTIFY_OWNER: true                 // recebes um email quando chega uma encomenda nova
};

const STATUS_OPTIONS = ['Recebida', 'A aguardar pagamento', 'Paga', 'Em produção', 'Pronta para entrega', 'Entregue', 'Cancelada'];
const PAYMENT_OPTIONS = ['Pendente', 'Confirmado'];

// Uma linha por encomenda, uma coluna por tamanho (os tamanhos vêm de CONFIG.SIZES)
const H = {
  id: 'Nº encomenda', at: 'Data', status: 'Estado', payment: 'Pagamento',
  name: 'Nome', email: 'Email', instagram: 'Instagram', piece: 'Peça',
  pieces: 'Peças', total: 'Total (EUR)', terms: 'Condições aceites', notes: 'Notas'
};
const HEADERS = [H.id, H.at, H.status, H.payment, H.name, H.email, H.instagram, H.piece]
  .concat(CONFIG.SIZES)
  .concat([H.pieces, H.total, H.terms, H.notes]);

/* ------------------------------------------------------------------
   One-time setup
------------------------------------------------------------------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  } else if (sheet.getLastRow() > 1) {
    Logger.log('O separador "' + CONFIG.SHEET_NAME + '" já tem encomendas. ' +
      'O setup NÃO foi executado, para proteger os dados. Apagar ou renomear esse separador e executar setup() outra vez.');
    return;
  } else {
    // Only a header row exists: reset it cleanly
    sheet.clear();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  }

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  const statusCol = HEADERS.indexOf(H.status) + 1;
  const paymentCol = HEADERS.indexOf(H.payment) + 1;
  const rows = 1000;

  sheet.getRange(2, statusCol, rows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUS_OPTIONS, true).setAllowInvalid(false).build()
  );
  sheet.getRange(2, paymentCol, rows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(PAYMENT_OPTIONS, true).setAllowInvalid(false).build()
  );

  // Manter o número da encomenda como texto simples
  sheet.getRange(2, HEADERS.indexOf(H.id) + 1, rows, 1).setNumberFormat('@');

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ORDER_COUNTER')) props.setProperty('ORDER_COUNTER', '0');

  sheet.autoResizeColumns(1, HEADERS.length);
  Logger.log('Setup concluído. O separador "' + CONFIG.SHEET_NAME + '" está pronto.');
}

/**
 * Run this only after deleting your test orders, to restart numbering at VRT-001.
 */
function resetCounter() {
  PropertiesService.getScriptProperties().setProperty('ORDER_COUNTER', '0');
  Logger.log('Contador de encomendas reposto a 0.');
}

/* ------------------------------------------------------------------
   Web app entry point
------------------------------------------------------------------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    const data = parseBody_(e);
    const order = validate_(data);               // throws on invalid input

    checkLimits_(order);                         // throws on abuse

    const orderId = nextOrderId_();
    writeRow_(orderId, order);

    sendConfirmation_(orderId, order);
    if (CONFIG.NOTIFY_OWNER) notifyOwner_(orderId, order);

    return json_({ ok: true, id: orderId });
  } catch (err) {
    console.error(err && err.message ? err.message : err);
    return json_({ ok: false });
  } finally {
    try { lock.releaseLock(); } catch (x) { /* not held */ }
  }
}

function doGet() {
  return ContentService.createTextOutput('VIRTHVS').setMimeType(ContentService.MimeType.TEXT);
}

/* ------------------------------------------------------------------
   Parsing and validation (never trust the page)
------------------------------------------------------------------- */
function parseBody_(e) {
  if (!e || !e.postData || !e.postData.contents) throw new Error('Empty request');
  if (e.postData.contents.length > 5000) throw new Error('Request too large');
  return JSON.parse(e.postData.contents);
}

function clean_(value, max) {
  let s = String(value === undefined || value === null ? '' : value);
  s = s.replace(/[\u0000-\u001F\u007F]/g, ' ').replace(/\s+/g, ' ').trim();
  return s.substring(0, max);
}

/**
 * Stops values that start with = + - @ from being read as formulas in Sheets.
 */
function safeCell_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function validate_(d) {
  const o = {
    name: clean_(d.name, 120),
    email: clean_(d.email, 160).toLowerCase(),
    instagram: clean_(d.instagram, 60).replace(/^@+/, '')
  };

  if (!o.name) throw new Error('Missing field');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(o.email)) throw new Error('Invalid email');
  if (!/^[A-Za-z0-9._]{1,30}$/.test(o.instagram)) throw new Error('Invalid Instagram username');
  if (d.terms_accepted !== true) throw new Error('Terms not accepted');

  // Sizes: one entry per piece
  if (!Array.isArray(d.sizes) || d.sizes.length < 1 || d.sizes.length > CONFIG.MAX_PIECES) {
    throw new Error('Invalid pieces');
  }
  o.counts = {};
  d.sizes.forEach(function (raw) {
    const size = clean_(raw, 10);
    if (CONFIG.SIZES.indexOf(size) === -1) throw new Error('Invalid size');
    o.counts[size] = (o.counts[size] || 0) + 1;
  });
  o.pieces = d.sizes.length;
  o.total = CONFIG.PRICE ? CONFIG.PRICE * o.pieces : '';
  return o;
}

/** e.g. "M x 1, XL x 2" */
function sizesSummary_(o) {
  return CONFIG.SIZES.filter(function (sz) { return o.counts[sz]; })
    .map(function (sz) { return sz + ' x ' + o.counts[sz]; })
    .join(', ');
}

function checkLimits_(o) {
  // Cooldown per email
  const cache = CacheService.getScriptCache();
  const key = 'email:' + o.email;
  if (cache.get(key)) throw new Error('Duplicate submission');
  cache.put(key, '1', CONFIG.EMAIL_COOLDOWN_SECONDS);

  // Daily cap
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const dayKey = 'DAY_' + today;
  const count = Number(props.getProperty(dayKey) || 0);
  if (count >= CONFIG.DAILY_ORDER_LIMIT) throw new Error('Daily limit reached');
  props.setProperty(dayKey, String(count + 1));
}

/* ------------------------------------------------------------------
   Writing
------------------------------------------------------------------- */
function nextOrderId_() {
  const props = PropertiesService.getScriptProperties();
  const n = Number(props.getProperty('ORDER_COUNTER') || 0) + 1;
  props.setProperty('ORDER_COUNTER', String(n));
  return 'VRT-' + ('000' + n).slice(-3);
}

function writeRow_(orderId, o) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Separador não encontrado. Executar setup() primeiro.');

  const row = HEADERS.map(function (h) {
    switch (h) {
      case H.id: return orderId;
      case H.at: return new Date();
      case H.status: return 'Recebida';
      case H.payment: return 'Pendente';
      case H.name: return safeCell_(o.name);
      case H.email: return safeCell_(o.email);
      case H.instagram: return safeCell_(o.instagram);
      case H.piece: return CONFIG.PIECE_NAME;
      case H.pieces: return o.pieces;
      case H.total: return o.total;
      case H.terms: return 'Sim';
      default:
        // Colunas de tamanho: número de peças nesse tamanho
        return CONFIG.SIZES.indexOf(h) !== -1 ? (o.counts[h] || '') : '';
    }
  });

  // Write at the first empty row, with text-safe formats already set by setup()
  const target = sheet.getLastRow() + 1;
  sheet.getRange(target, 1, 1, row.length).setValues([row]);
}

/* ------------------------------------------------------------------
   Emails (plain text, English)
------------------------------------------------------------------- */
function sendConfirmation_(orderId, o) {
  const total = o.total !== '' ? '\nTotal: ' + Number(o.total).toFixed(2).replace('.', ',') + ' €' : '';
  const body =
    'Pré-encomenda registada.\n\n' +
    'Referência: ' + orderId + '\n' +
    'Peça: ' + CONFIG.PIECE_NAME + '\n' +
    'Tamanhos: ' + sizesSummary_(o) + '\n' +
    'Peças: ' + o.pieces +
    total + '\n\n' +
    'As instruções de pagamento seguem por mensagem direta no Instagram.\n' +
    'Entrega em mão, a partir de ' + CONFIG.DELIVERY_FROM_LABEL + '. O local e a hora combinam-se por mensagem.\n\n' +
    'Preserve. Protect. Remain.';

  MailApp.sendEmail({
    to: o.email,
    subject: orderId + ' | Recebida',
    body: body,
    name: CONFIG.BRAND_NAME
  });
}

function notifyOwner_(orderId, o) {
  const owner = Session.getEffectiveUser().getEmail();
  if (!owner) return;
  const body =
    orderId + '\n' +
    o.name + ' (@' + o.instagram + ')\n' +
    CONFIG.PIECE_NAME + ': ' + sizesSummary_(o) + ' (' + o.pieces + ' peças)\n\n' +
    'Próximo passo: enviar as instruções de pagamento por mensagem direta e passar o estado para "A aguardar pagamento".';
  MailApp.sendEmail({ to: owner, subject: 'Nova pré-encomenda ' + orderId, body: body, name: CONFIG.BRAND_NAME });
}

/* ------------------------------------------------------------------
   Helpers and test
------------------------------------------------------------------- */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Test without the page: change the email to your own, run it,
 * then check the sheet and your inbox. Delete the test row afterwards
 * and run resetCounter() if you want numbering to restart.
 */
function testOrder() {
  const fake = {
    postData: {
      contents: JSON.stringify({
        name: 'Teste Encomenda',
        email: 'O_TEU_EMAIL@exemplo.pt',
        instagram: 'test.account',
        sizes: [CONFIG.SIZES[0], CONFIG.SIZES[CONFIG.SIZES.length - 1]],
        terms_accepted: true
      })
    }
  };
  Logger.log(doPost(fake).getContent());
}
