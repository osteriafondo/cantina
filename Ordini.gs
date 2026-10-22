// =============================================================================
//  Osteria Fondo — Ordini vini (backend)
//  File SEPARATO dello stesso progetto Apps Script: in editor fai
//  "+" → Script, chiamalo "Ordini" e incolla questo file.
//  Usa API_KEY e ACCESS_TOKEN definiti nel file principale: non ripeterli qui.
//  Il file principale gli passa tutte le azioni che iniziano con "ordini".
//
//  Fogli usati (creati da soli al primo salvataggio, quelli dell'inventario
//  non vengono toccati):
//    FORNITORI · LISTINI · ORDINI · ORDINI_IMPOSTAZIONI
// =============================================================================

var ORD_MODELLO = 'claude-opus-5-5';

var ORD_FOGLI = {
  FORNITORI: ['Id', 'Nome', 'Tipo', 'Rappresentante', 'Telefono', 'Email', 'OrdineMinimo', 'PortoFranco', 'Pagamento', 'Consegna', 'Condizioni', 'ListinoAl', 'Aggiornato'],
  LISTINI: ['FornitoreId', 'Id', 'Codice', 'Vino', 'Produttore', 'Annata', 'Formato', 'Prezzo', 'Cartone'],
  ORDINI: ['Id', 'FornitoreId', 'Fornitore', 'Stato', 'Bottiglie', 'Imponibile', 'Nota', 'Creato', 'Aggiornato', 'Inviato', 'Righe'],
  ORDINI_IMPOSTAZIONI: ['Chiave', 'Valore']
};

function gestisciOrdini(ss, action, params) {
  if (params.token !== ACCESS_TOKEN) return { errore: 'non autorizzato' };
  var dati = params.dati ? JSON.parse(params.dati) : {};
  switch (action) {
    case 'ordiniLeggi':             return ordLeggi(ss);
    case 'ordiniSalvaFornitore':    return ordConLock(function() { return ordSalvaFornitore(ss, dati); });
    case 'ordiniEliminaFornitore':  return ordConLock(function() { return ordEliminaFornitore(ss, dati.id); });
    case 'ordiniSalvaOrdine':       return ordConLock(function() { return ordSalvaOrdine(ss, dati); });
    case 'ordiniEliminaOrdine':     return ordConLock(function() { return ordEliminaRiga(ordFoglio(ss, 'ORDINI'), dati.id) ? { ok: true } : { ok: true, giaEliminato: true }; });
    case 'ordiniSalvaImpostazioni': return ordConLock(function() { return ordSalvaImpostazioni(ss, dati); });
    case 'ordiniAnalizzaListino':   return ordAnalizzaListino(params.pdfBase64, params.pagine || '');
  }
  return { errore: 'Azione non riconosciuta' };
}

// ── Utility ──────────────────────────────────────────────────────────────────

function ordConLock(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { errore: 'Il foglio è occupato da un altro salvataggio, riprova tra un attimo.' };
  try { return fn(); } finally { lock.releaseLock(); }
}
function ordFoglio(ss, nome) { return getOrCreateSheet(ss, nome, ORD_FOGLI[nome]); }
function ordRighe(sh) {
  var v = sh.getDataRange().getValues();
  return v.slice(1).filter(function(r) { return r[0] !== '' && r[0] !== null; });
}
function ordTesto(v) { return (v === null || v === undefined) ? '' : (v instanceof Date ? dataISO(v) : v.toString()); }
function ordNum(v) { return arrotonda(parseFloat(v) || 0); }
function ordTempo(v) { return (v instanceof Date) ? v.getTime() : (parseFloat(v) || 0); }
function ordTrovaRiga(sh, id) {
  var col = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (var i = 1; i < col.length; i++) if (col[i][0].toString() === id.toString()) return i + 1;
  return 0;
}
function ordScriviRiga(sh, id, riga) {
  var n = ordTrovaRiga(sh, id);
  if (n) sh.getRange(n, 1, 1, riga.length).setValues([riga]);
  else sh.appendRow(riga);
}
function ordEliminaRiga(sh, id) {
  var n = ordTrovaRiga(sh, id);
  if (n) sh.deleteRow(n);
  return !!n;
}

// ── Lettura ──────────────────────────────────────────────────────────────────

function ordLeggi(ss) {
  var listini = {};
  ordRighe(ordFoglio(ss, 'LISTINI')).forEach(function(r) {
    var fid = r[0].toString();
    (listini[fid] = listini[fid] || []).push({
      id: ordTesto(r[1]), code: ordTesto(r[2]), name: ordTesto(r[3]), producer: ordTesto(r[4]),
      vintage: ordTesto(r[5]), format: ordTesto(r[6]), price: ordNum(r[7]), 'case': parseInt(r[8]) || 0
    });
  });
  var fornitori = ordRighe(ordFoglio(ss, 'FORNITORI')).map(function(r) {
    var id = r[0].toString();
    return {
      id: id, name: ordTesto(r[1]), kind: ordTesto(r[2]) || 'distributore', repName: ordTesto(r[3]), repPhone: ordTesto(r[4]),
      repEmail: ordTesto(r[5]), minOrder: ordNum(r[6]), freeShipping: ordNum(r[7]), payment: ordTesto(r[8]),
      delivery: ordTesto(r[9]), conditions: ordTesto(r[10]), priceListDate: formattaData(r[11]), updatedAt: ordTempo(r[12]),
      products: listini[id] || []
    };
  });
  var ordini = ordRighe(ordFoglio(ss, 'ORDINI')).map(function(r) {
    var righe = {};
    try { righe = JSON.parse(r[10] || '{}'); } catch (e) {}
    return {
      id: r[0].toString(), supplierId: ordTesto(r[1]), supplierName: ordTesto(r[2]), status: ordTesto(r[3]) || 'bozza',
      note: ordTesto(r[6]), createdAt: ordTempo(r[7]), updatedAt: ordTempo(r[8]), closedAt: ordTempo(r[9]), lines: righe
    };
  });
  ordini.sort(function(a, b) { return b.updatedAt - a.updatedAt; });
  var impostazioni = {};
  ordRighe(ordFoglio(ss, 'ORDINI_IMPOSTAZIONI')).forEach(function(r) {
    try { impostazioni[r[0]] = JSON.parse(r[1]); } catch (e) { impostazioni[r[0]] = r[1]; }
  });
  return { fornitori: fornitori, ordini: ordini, impostazioni: impostazioni };
}

// ── Scrittura ────────────────────────────────────────────────────────────────

function ordSalvaFornitore(ss, f) {
  if (!f.id || !(f.name || '').toString().trim()) return { errore: 'Manca il nome del fornitore.' };
  var adesso = new Date();
  ordScriviRiga(ordFoglio(ss, 'FORNITORI'), f.id, [
    f.id, f.name.toString().trim(), f.kind || 'distributore', f.repName || '', f.repPhone || '', f.repEmail || '',
    ordNum(f.minOrder), ordNum(f.freeShipping), f.payment || '', f.delivery || '', f.conditions || '',
    f.priceListDate || '', adesso
  ]);
  // Listino: riscrivo solo le righe di questo fornitore, le altre restano come sono
  var sh = ordFoglio(ss, 'LISTINI');
  var tutte = sh.getDataRange().getValues();
  var intest = tutte[0];
  var altre = tutte.slice(1).filter(function(r) { return r[0] !== '' && r[0].toString() !== f.id.toString(); });
  var nuove = (f.products || []).filter(function(p) { return (p.name || '').toString().trim(); }).map(function(p) {
    return [f.id, p.id || Utilities.getUuid().slice(0, 12), p.code || '', p.name, p.producer || '', p.vintage || '', p.format || '', ordNum(p.price), parseInt(p['case']) || 0];
  });
  var righe = altre.concat(nuove);
  if (tutte.length > 1) sh.getRange(2, 1, tutte.length - 1, intest.length).clearContent();
  if (righe.length) {
    // codici e annate come testo (non perdono gli zeri iniziali), prezzi e cartoni come numeri
    sh.getRange(2, 1, righe.length, 7).setNumberFormat('@');
    sh.getRange(2, 8, righe.length, 1).setNumberFormat('0.00');
    sh.getRange(2, 1, righe.length, intest.length).setValues(righe);
  }
  return { ok: true, updatedAt: adesso.getTime(), referenze: nuove.length };
}

function ordEliminaFornitore(ss, id) {
  if (!id) return { errore: 'Fornitore mancante.' };
  ordEliminaRiga(ordFoglio(ss, 'FORNITORI'), id);
  var sh = ordFoglio(ss, 'LISTINI');
  var tutte = sh.getDataRange().getValues();
  var resto = tutte.slice(1).filter(function(r) { return r[0] !== '' && r[0].toString() !== id.toString(); });
  if (tutte.length > 1) sh.getRange(2, 1, tutte.length - 1, tutte[0].length).clearContent();
  if (resto.length) sh.getRange(2, 1, resto.length, tutte[0].length).setValues(resto);
  return { ok: true };
}

function ordSalvaOrdine(ss, o) {
  if (!o.id || !o.supplierId) return { errore: 'Ordine incompleto.' };
  var bottiglie = 0, imponibile = 0;
  Object.keys(o.lines || {}).forEach(function(k) {
    var l = o.lines[k];
    bottiglie += parseInt(l.qty) || 0;
    imponibile += (parseInt(l.qty) || 0) * (parseFloat(l.price) || 0);
  });
  var adesso = Date.now();
  ordScriviRiga(ordFoglio(ss, 'ORDINI'), o.id, [
    o.id, o.supplierId, o.supplierName || '', o.status || 'bozza', bottiglie, arrotonda(imponibile), o.note || '',
    o.createdAt ? new Date(o.createdAt) : new Date(), new Date(adesso), o.closedAt ? new Date(o.closedAt) : '',
    JSON.stringify(o.lines || {})
  ]);
  return { ok: true, updatedAt: adesso };
}

function ordSalvaImpostazioni(ss, dati) {
  var sh = ordFoglio(ss, 'ORDINI_IMPOSTAZIONI');
  Object.keys(dati || {}).forEach(function(k) { ordScriviRiga(sh, k, [k, JSON.stringify(dati[k])]); });
  return { ok: true };
}

// ── Lettura dei listini PDF con Claude ───────────────────────────────────────
//  La pagina divide il PDF in pezzi di poche pagine e li manda uno per volta,
//  così ogni chiamata resta breve.

var ORD_PROMPT_LISTINO =
  'Questo PDF è (una parte di) un listino prezzi di vini di un distributore o di una cantina, letto da un ristorante italiano. ' +
  'Estrai TUTTE le referenze di vino acquistabili presenti in queste pagine.\n\n' +
  'Per ogni referenza:\n' +
  '- name: il vino o la denominazione (es. "Barolo Ravera", "Franciacorta Brut"), senza produttore e senza annata.\n' +
  '- producer: la cantina. Se il listino raggruppa i vini sotto il titolo di una cantina, usa quel titolo per ogni vino del gruppo. ' +
  'Se il gruppo è iniziato in una pagina precedente e qui non si vede il titolo, lascia vuoto.\n' +
  '- vintage: annata a 4 cifre, oppure "" (spumanti senza annata, "s.a.", "NV").\n' +
  '- format: es. "75 cl", "150 cl", "37,5 cl"; "" se non indicato.\n' +
  '- price: prezzo di UNA bottiglia, IVA esclusa, in euro. Se ci sono più colonne di prezzo (listino, sconto, netto) usa il prezzo netto finale per il cliente. ' +
  'Se il prezzo è per cartone, dividilo per le bottiglie del cartone. Se il listino dice che i prezzi sono IVA inclusa, scorpora il 22%. 0 se non c\'è un prezzo.\n' +
  '- case: bottiglie per cartone se indicate, altrimenti 0.\n' +
  '- code: codice articolo se presente, altrimenti "".\n\n' +
  'Salta intestazioni, note, condizioni di vendita e prodotti che chiaramente non sono vino. Non inventare valori: lascia "" o 0.';

var ORD_SCHEMA_LISTINO = {
  type: 'object',
  properties: {
    referenze: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          code: { type: 'string' }, name: { type: 'string' }, producer: { type: 'string' },
          vintage: { type: 'string' }, format: { type: 'string' }, price: { type: 'number' }, 'case': { type: 'integer' }
        },
        required: ['code', 'name', 'producer', 'vintage', 'format', 'price', 'case'],
        additionalProperties: false
      }
    }
  },
  required: ['referenze'],
  additionalProperties: false
};

function ordAnalizzaListino(pdfBase64, pagine) {
  if (!pdfBase64) return { errore: 'PDF mancante.' };
  if (!API_KEY || API_KEY === 'LA_MIA_CHIAVE') return { errore: 'Chiave API Anthropic non impostata nello script.' };
  var testo = ORD_PROMPT_LISTINO + (pagine ? '\n\nQueste sono le pagine ' + pagine + ' del listino.' : '');
  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    muteHttpExceptions: true,
    headers: {
      'x-api-key': API_KEY,
      'anthropic-version': '2023-06-01',
      'anthropic-beta': 'server-side-fallback-2026-07-01',
      'content-type': 'application/json'
    },
    payload: JSON.stringify({
      model: ORD_MODELLO,
      max_tokens: 16000,
      fallbacks: 'default',
      output_config: { effort: 'low', format: { type: 'json_schema', schema: ORD_SCHEMA_LISTINO } },
      messages: [{ role: 'user', content: [
        { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: pdfBase64 } },
        { type: 'text', text: testo }
      ] }]
    })
  });
  var result;
  try { result = JSON.parse(response.getContentText()); } catch (e) { return { errore: 'Risposta non valida da Claude (HTTP ' + response.getResponseCode() + ').' }; }
  if (result.error) return { errore: 'Claude: ' + (result.error.message || 'errore sconosciuto') };
  if (result.stop_reason === 'refusal') return { errore: 'Claude non ha voluto leggere queste pagine.' };
  if (result.stop_reason === 'max_tokens') return { errore: 'Troppe referenze in queste pagine per una sola lettura.' };
  var blocco = (result.content || []).filter(function(b) { return b.type === 'text'; })[0];
  if (!blocco) return { errore: 'Claude non ha restituito referenze.' };
  try {
    return { referenze: JSON.parse(blocco.text).referenze || [] };
  } catch (e) {
    return { errore: 'Non sono riuscito a leggere la risposta di Claude.' };
  }
}
