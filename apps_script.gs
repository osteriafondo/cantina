// =============================================================================
//  Osteria Fondo — Backend Google Apps Script (file completo)
//  Copia TUTTO questo file e incollalo nell'editor Apps Script,
//  dopo aver svuotato completamente l'editor (Ctrl/Cmd+A -> Canc).
//  Poi rimetti API_KEY e ACCESS_TOKEN qui sotto e fai il Deploy.
// =============================================================================

// ⬇️ LE TUE CREDENZIALI
var API_KEY = 'LA_MIA_CHIAVE';                    // chiave API Anthropic
var ACCESS_TOKEN = 'la-tua-password-gestionale';  // password del gestionale

var CATEGORIE = ['BIANCHI', 'ROSSI', 'BOLLICINE', 'MACERATI', 'ROSATI'];

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  var params = e.parameter;
  var action = params.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result;
  var protette = ['aggiungi','aggiorna','salvaAdj','generaAdj','analizzaBolla','statistiche','salvaSnapshot','nascondi'];

  try {
    if (action === 'verificaToken') {
      result = (params.token === ACCESS_TOKEN) ? { ok: true } : { errore: 'non autorizzato' };
    } else if (protette.indexOf(action) !== -1 && params.token !== ACCESS_TOKEN) {
      result = { errore: 'non autorizzato' };
    } else if (action === 'leggi') {
      result = leggiTutti(ss);
    } else if (action === 'aggiorna') {
      result = aggiornaBottiglie(ss, params.categoria, parseInt(params.id), parseInt(params.quantita), (params.acquisto !== undefined && params.acquisto !== '') ? parseFloat(params.acquisto) : null);
    } else if (action === 'aggiungi') {
      result = aggiungiVino(ss, params.categoria, params.nome, params.produttore, params.vitigni, params.regione, parseFloat(params.acquisto), parseFloat(params.vendita), parseInt(params.quantita), parseInt(params.soglia), params.adj || '', params.area || '');
    } else if (action === 'salvaAdj') {
      result = salvaAdj(ss, params.categoria, parseInt(params.id), params.adj);
    } else if (action === 'nascondi') {
      result = salvaNascosto(ss, params.categoria, parseInt(params.id), params.nascosto === '1');
    } else if (action === 'analizzaBolla') {
      result = analizzaBolla(ss, params.imageBase64, params.mimeType);
    } else if (action === 'generaAdj') {
      result = generaAdj(params.nome, params.produttore, params.vitigno);
    } else if (action === 'statistiche') {
      result = statistiche();
    } else if (action === 'salvaSnapshot') {
      result = salvaSnapshot();
    } else {
      result = { errore: 'Azione non riconosciuta' };
    }
  } catch(err) {
    result = { errore: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ── Lettura / scrittura vini ──────────────────────────────────────────────────
function leggiTutti(ss) {
  var dati = {};
  CATEGORIE.forEach(function(cat) {
    var sheet = ss.getSheetByName(cat);
    if (!sheet) { dati[cat.toLowerCase()] = []; return; }
    var rows = sheet.getDataRange().getValues();
    var vini = [];
    for (var i = 1; i < rows.length; i++) {
      var qty = rows[i][0], nome = rows[i][1], produttore = rows[i][2],
          vitigno = rows[i][3], regione = rows[i][4], vendita = rows[i][6] || '',
          soglia = rows[i][8] || 3, adj = rows[i][9] || '', area = rows[i][10] || '',
          nascosto = rows[i][11];   // colonna 12 (L) = nascosto
      if (!nome || nome.toString().trim() === '') continue;
      var nv = nascosto ? nascosto.toString().trim().toLowerCase() : '';
      vini.push({
        id: i,
        nome: nome.toString().trim(),
        produttore: produttore ? produttore.toString().trim() : '',
        vitigno: vitigno ? vitigno.toString().trim() : '',
        regione: regione ? regione.toString().trim() : '',
        vendita: vendita ? '€ ' + parseFloat(vendita).toFixed(2) : '',
        adj: adj ? adj.toString().trim() : '',
        area: area ? area.toString().trim() : '',
        nascosto: (nv === 'si' || nv === '1' || nv === 'true' || nv === 'vero' || nv === 'x'),
        qty: parseInt(qty) || 0,
        soglia: parseInt(soglia) || 3
      });
    }
    dati[cat.toLowerCase()] = vini;
  });
  return dati;
}

function aggiornaBottiglie(ss, categoria, id, nuovaQty, acquisto) {
  var sheet = ss.getSheetByName(categoria.toUpperCase());
  if (!sheet) return { errore: 'Categoria non trovata' };
  var riga = id + 1;
  var vecchia = parseInt(sheet.getRange(riga, 1).getValue()) || 0;
  var delta = (parseInt(nuovaQty) || 0) - vecchia;
  // se arriva un prezzo d'acquisto valido (es. da una bolla) aggiorna la colonna F
  if (acquisto && acquisto > 0) sheet.getRange(riga, 6).setValue(acquisto);
  if (delta > 0) {   // e' un carico (rifornimento): registralo
    var nome = sheet.getRange(riga, 2).getValue();
    var prezzoCarico = parseFloat(sheet.getRange(riga, 6).getValue()) || 0;
    logCarico(ss, categoria, nome, delta, prezzoCarico);
  }
  sheet.getRange(riga, 1).setValue(nuovaQty);
  return { ok: true, id: id, qty: nuovaQty };
}

function aggiungiVino(ss, categoria, nome, produttore, vitigni, regione, acquisto, vendita, qty, soglia, adj, area) {
  var sheet = ss.getSheetByName(categoria.toUpperCase());
  if (!sheet) return { errore: 'Categoria non trovata' };
  var lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1).setValue(qty || 0);
  sheet.getRange(lastRow, 2).setValue(nome);
  sheet.getRange(lastRow, 3).setValue(produttore || '');
  sheet.getRange(lastRow, 4).setValue(vitigni || '');
  sheet.getRange(lastRow, 5).setValue(regione || '');
  sheet.getRange(lastRow, 6).setValue(acquisto || 0);   // prezzo di acquisto (per il wine cost)
  sheet.getRange(lastRow, 7).setValue(vendita || '');
  sheet.getRange(lastRow, 9).setValue(soglia || 3);
  sheet.getRange(lastRow, 10).setValue(adj || '');
  sheet.getRange(lastRow, 11).setValue(area || '');
  if ((parseInt(qty) || 0) > 0) logCarico(ss, categoria, nome, qty, acquisto || 0);
  return { ok: true, riga: lastRow };
}

function salvaAdj(ss, categoria, id, adj) {
  var sheet = ss.getSheetByName(categoria.toUpperCase());
  if (!sheet) return { errore: 'Categoria non trovata' };
  sheet.getRange(id + 1, 10).setValue(adj);
  return { ok: true };
}

function salvaNascosto(ss, categoria, id, nascosto) {
  var sheet = ss.getSheetByName(categoria.toUpperCase());
  if (!sheet) return { errore: 'Categoria non trovata' };
  sheet.getRange(id + 1, 12).setValue(nascosto ? 'SI' : '');   // colonna 12 (L)
  return { ok: true };
}

// ── AI ────────────────────────────────────────────────────────────────────────
function generaAdj(nome, produttore, vitigno) {
  var prompt = 'Sei un sommelier esperto. Per il vino "' + (nome || '') +
    '" prodotto da "' + (produttore || 'produttore sconosciuto') +
    '", vitigno "' + (vitigno || 'non specificato') +
    '", rispondi SOLO con 2 aggettivi in italiano che descrivono le note ' +
    'aromatiche o la struttura, separati da " · ". Esempio: "Sapido · Minerale". ' +
    'Nientaltro, nessuna punteggiatura finale.';
  try {
    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      headers: { 'x-api-key': API_KEY, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
      payload: JSON.stringify({ model: 'claude-haiku-4-5', max_tokens: 20, messages: [{ role: 'user', content: prompt }] }),
      muteHttpExceptions: true
    });
    var result = JSON.parse(response.getContentText());
    var adj = (result.content && result.content[0] && result.content[0].text) ? result.content[0].text.trim() : '';
    return { adj: adj };
  } catch(e) { return { errore: e.toString() }; }
}

function analizzaBolla(ss, imageBase64, mimeType) {
  // Elenco dei vini gia in inventario (anche a giacenza 0), per far riconoscere
  // all'AI i duplicati ed evitare che lo stesso vino compaia due volte.
  var inventario = leggiTutti(ss);
  var elenco = [];
  CATEGORIE.forEach(function(cat) {
    var key = cat.toLowerCase();
    (inventario[key] || []).forEach(function(v) {
      elenco.push('- categoria "' + key + '" | nome "' + v.nome + '"' +
        (v.produttore ? ' | produttore "' + v.produttore + '"' : ''));
    });
  });
  var testoInventario = elenco.length
    ? elenco.join('\n')
    : '(inventario vuoto)';

  var prompt = 'Sei un assistente per la gestione di una cantina di un ristorante. Analizza questo documento (bolla di consegna di un fornitore di vini) ed estrai TUTTI i vini presenti.\n\n' +
    'Per ogni vino restituisci un JSON array con questi campi:\n' +
    '- nome: nome del vino (stringa)\n' +
    '- produttore: nome del produttore/cantina se presente (stringa, puo essere vuoto)\n' +
    '- quantita: numero di bottiglie (numero intero)\n' +
    '- prezzo: SOLO il prezzo unitario di acquisto per bottiglia, IVA ESCLUSA (numero decimale, 0 se non presente). NON usare il totale di riga, il totale documento, l imponibile complessivo o il prezzo IVA inclusa.\n' +
    '- categoria: "bianchi", "rossi", "bollicine", "macerati" o "rosati" (deduci dal tipo di vino)\n' +
    '- esistente: se il vino corrisponde a uno gia presente nell inventario qui sotto, un oggetto {"categoria": <categoria del vino in inventario>, "nome": <nome ESATTO come scritto in inventario>}; altrimenti null.\n\n' +
    'REGOLA IMPORTANTE per "esistente": e lo stesso vino se sono la stessa etichetta dello stesso produttore, ANCHE se cambia l annata (es. 2019 vs 2020), la scrittura, le abbreviazioni o le maiuscole. In quel caso copia esattamente il nome gia presente in inventario. Se il vino non c e (nemmeno a giacenza 0), metti esistente = null.\n\n' +
    'VINI GIA IN INVENTARIO:\n' + testoInventario + '\n\n' +
    'Ignora tutti gli altri numeri della bolla (valore IVA, aliquote, sconti, totali di riga o di documento, colli, codici): non vanno riportati.\n\n' +
    'Rispondi SOLO con il JSON array, nessun testo aggiuntivo.';
  var contentItem = mimeType === 'application/pdf'
    ? { type: 'document', source: { type: 'base64', media_type: mimeType, data: imageBase64 } }
    : { type: 'image', source: { type: 'base64', media_type: mimeType, data: imageBase64 } };
  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    headers: { 'x-api-key': API_KEY, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
    payload: JSON.stringify({ model: 'claude-sonnet-4-6', max_tokens: 1000, messages: [{ role: 'user', content: [contentItem, { type: 'text', text: prompt }] }] })
  });
  var result = JSON.parse(response.getContentText());
  var clean = result.content[0].text.trim().replace(/```json|```/g, '').trim();
  return JSON.parse(clean);
}

function generaDescrizioniTutti() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  CATEGORIE.forEach(function(cat) {
    var sheet = ss.getSheetByName(cat);
    if (!sheet) return;
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var nome = rows[i][1], produttore = rows[i][2] || '', vitigno = rows[i][3] || '', adjEsistente = rows[i][9];
      if (!nome || nome.toString().trim() === '') continue;
      if (adjEsistente && adjEsistente.toString().trim() !== '') continue;
      var res = generaAdj(nome, produttore, vitigno);
      if (res.adj) { sheet.getRange(i + 1, 10).setValue(res.adj); Utilities.sleep(300); }
    }
  });
}

function warmUp() { Logger.log('Ponte attivo: ' + new Date()); }

// ── STATISTICHE ─────────────────────────────────────────────────────────────────
function dataISO(d) {
  d = d || new Date();
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}
function formattaData(v) { return (v instanceof Date) ? dataISO(v) : (v || '').toString().slice(0, 10); }
function chiaveNome(n) { return (n || '').toString().trim().toLowerCase(); }
function arrotonda(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function nomeMese(yyyymm) {
  var nomi = ['Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno','Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre'];
  var p = yyyymm.split('-');
  return nomi[parseInt(p[1], 10) - 1] + ' ' + p[0];
}
function getOrCreateSheet(ss, nome, intestazioni) {
  var sh = ss.getSheetByName(nome);
  if (!sh) { sh = ss.insertSheet(nome); sh.appendRow(intestazioni); }
  return sh;
}
function logCarico(ss, categoria, nome, quantita, acquisto) {
  quantita = parseInt(quantita) || 0;
  if (quantita <= 0) return;
  var sh = getOrCreateSheet(ss, 'CARICHI', ['Data','Categoria','Nome','Quantita','Acquisto']);
  sh.appendRow([new Date(), (categoria || '').toString().toLowerCase(), (nome || '').toString().trim(), quantita, parseFloat(acquisto) || 0]);
}

function salvaSnapshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOrCreateSheet(ss, 'STORICO', ['Data','Categoria','Nome','Giacenza','Acquisto','Vendita']);
  var oggi = dataISO();
  var esistenti = sh.getDataRange().getValues();
  for (var r = 1; r < esistenti.length; r++) {
    if (formattaData(esistenti[r][0]) === oggi) return { ok: true, righe: 0, data: oggi, nota: 'Fotografia gia presente per oggi' };
  }
  var righe = [];
  CATEGORIE.forEach(function(cat) {
    var sheet = ss.getSheetByName(cat);
    if (!sheet) return;
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var nome = rows[i][1];
      if (!nome || nome.toString().trim() === '') continue;
      righe.push([oggi, cat.toLowerCase(), nome.toString().trim(), parseInt(rows[i][0]) || 0, parseFloat(rows[i][5]) || 0, parseFloat(rows[i][6]) || 0]);
    }
  });
  if (righe.length) sh.getRange(sh.getLastRow() + 1, 1, righe.length, 6).setValues(righe);
  return { ok: true, righe: righe.length, data: oggi };
}

function installaTriggerSettimanale() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'salvaSnapshot') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('salvaSnapshot').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(23).create();
  return 'Trigger installato: fotografia ogni domenica sera.';
}

function caricaStorico(ss) {
  var sh = ss.getSheetByName('STORICO');
  var snaps = {};
  if (!sh) return { snaps: snaps, date: [] };
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var data = formattaData(rows[i][0]);
    var nome = (rows[i][2] || '').toString().trim();
    if (!data || !nome) continue;
    if (!snaps[data]) snaps[data] = {};
    snaps[data][chiaveNome(nome)] = {
      cat: (rows[i][1] || '').toString().toLowerCase(), nome: nome,
      qty: parseInt(rows[i][3]) || 0, acquisto: parseFloat(rows[i][4]) || 0, vendita: parseFloat(rows[i][5]) || 0
    };
  }
  return { snaps: snaps, date: Object.keys(snaps).sort() };
}

function caricaCarichi(ss) {
  var sh = ss.getSheetByName('CARICHI');
  var lista = [];
  if (!sh) return lista;
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var giorno = formattaData(rows[i][0]);
    if (!giorno) continue;
    lista.push({ giorno: giorno, nome: chiaveNome(rows[i][2]), qty: parseInt(rows[i][3]) || 0 });
  }
  return lista;
}

function venditePeriodo(snaps, prevData, currData, carichi) {
  var prev = snaps[prevData] || {}, curr = snaps[currData] || {};
  var caricoPerNome = {};
  carichi.forEach(function(c) {
    if (c.giorno > prevData && c.giorno <= currData) caricoPerNome[c.nome] = (caricoPerNome[c.nome] || 0) + c.qty;
  });
  var chiavi = {};
  Object.keys(prev).forEach(function(k) { chiavi[k] = 1; });
  Object.keys(curr).forEach(function(k) { chiavi[k] = 1; });
  var out = [];
  Object.keys(chiavi).forEach(function(k) {
    var info = curr[k] || prev[k];
    var venduto = (prev[k] ? prev[k].qty : 0) + (caricoPerNome[k] || 0) - (curr[k] ? curr[k].qty : 0);
    if (venduto <= 0) return;
    out.push({ nome: info.nome, cat: info.cat, sold: venduto, costo: venduto * (info.acquisto || 0), incasso: venduto * (info.vendita || 0) });
  });
  return out;
}

function aggrega(vendite) {
  var perNome = {};
  vendite.forEach(function(v) {
    var k = chiaveNome(v.nome);
    if (!perNome[k]) perNome[k] = { nome: v.nome, cat: v.cat, sold: 0, costo: 0, incasso: 0 };
    perNome[k].sold += v.sold; perNome[k].costo += v.costo; perNome[k].incasso += v.incasso; perNome[k].cat = v.cat;
  });
  var lista = Object.keys(perNome).map(function(k) { return perNome[k]; });
  var bottiglie = 0, costo = 0, incasso = 0;
  lista.forEach(function(v) { bottiglie += v.sold; costo += v.costo; incasso += v.incasso; });
  var top = lista.slice().sort(function(a, b) { return b.sold - a.sold; }).slice(0, 3)
    .map(function(v) { return { nome: v.nome, cat: v.cat, sold: v.sold, costo: arrotonda(v.costo) }; });
  var topCat = {};
  ['bollicine','bianchi','macerati','rosati','rossi'].forEach(function(cat) {
    topCat[cat] = lista.filter(function(v) { return v.cat === cat; }).sort(function(a, b) { return b.sold - a.sold; }).slice(0, 3)
      .map(function(v) { return { nome: v.nome, sold: v.sold, costo: arrotonda(v.costo) }; });
  });
  return {
    disponibile: true, bottiglie: bottiglie, costo: arrotonda(costo), incasso: arrotonda(incasso),
    wineCostPct: incasso > 0 ? Math.round(costo / incasso * 100) : null, top: top, topCat: topCat
  };
}

function statistiche() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var st = caricaStorico(ss), carichi = caricaCarichi(ss), date = st.date;
  var out = { ok: true, aggiornato: new Date().toISOString() };
  if (date.length < 2) {
    var msg = date.length === 1
      ? 'Hai una fotografia. Il primo report arrivera dopo la prossima domenica.'
      : 'Nessuna fotografia ancora. Premi "Salva fotografia giacenze ora" per iniziare.';
    out.settimana = { disponibile: false, messaggio: msg };
    out.mese = { disponibile: false, messaggio: msg };
    return out;
  }
  var currData = date[date.length - 1], prevData = date[date.length - 2];
  out.settimana = aggrega(venditePeriodo(st.snaps, prevData, currData, carichi));
  out.settimana.da = prevData; out.settimana.a = currData;
  var meseCorr = currData.slice(0, 7), venditeMese = [];
  for (var i = 1; i < date.length; i++) {
    if (date[i].slice(0, 7) === meseCorr) venditeMese = venditeMese.concat(venditePeriodo(st.snaps, date[i - 1], date[i], carichi));
  }
  out.mese = venditeMese.length ? aggrega(venditeMese) : { disponibile: false, messaggio: 'Nessuna vendita registrata questo mese.' };
  if (out.mese.disponibile) out.mese.mese = nomeMese(meseCorr);
  return out;
}
