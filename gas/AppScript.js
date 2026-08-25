// ═══════════════════════════════════════════════════════════════════
//  CASSA AZIENDALE — Google Apps Script v2
//  Incolla questo codice in Google Apps Script e segui le istruzioni
// ═══════════════════════════════════════════════════════════════════
//
//  ISTRUZIONI SETUP (una volta sola):
//
//  1. Vai su https://sheets.google.com e crea un nuovo foglio
//     Rinominalo come vuoi (es. "Cassa Aziendale")
//
//  2. Vai su Estensioni → Apps Script
//
//  3. Cancella il codice di default e incolla TUTTO questo file
//
//  4. Clicca su "Esegui" → seleziona la funzione "setupFoglio"
//     (questo crea le intestazioni automaticamente)
//
//  5. Vai su Distribuisci → Nuova distribuzione
//     - Tipo: App Web
//     - Esegui come: Me (il tuo account)
//     - Chi può accedere: Chiunque
//     - Clicca "Distribuisci" e autorizza
//
//  6. Copia l'URL della distribuzione (inizia con https://script.google.com/...)
//
//  7. Nell'app HTML, vai su ⚙️ Impostazioni e incolla quell'URL
//
// ═══════════════════════════════════════════════════════════════════

const NOME_FOGLIO              = 'Movimenti';
const NOME_FOGLIO_PERSONALE    = 'Personale';
const NOME_FOGLIO_CATEGORIE    = 'Categorie';
const NOME_FOGLIO_BUDGET_FAM   = 'BudgetFamiliare';
const NOME_FOGLIO_SPESE_FAM    = 'SpeseFamiliari';
const NOME_FOGLIO_PRESTITO     = 'Prestito';

// Categorie del budget familiare (allineate ai gruppi Splitwise)
const CATEGORIE_FAM = ['Auto','Casa','Spritz','Riccardo','Regali','Ristoranti','Spesa Cibo','Viaggi'];

// Default budget annuali (dal foglio storico)
const BUDGET_FAM_DEFAULT = {
  2025: {'Auto':3000,'Casa':3000,'Spritz':600,'Riccardo':2000,'Regali':1000,'Ristoranti':3000,'Spesa Cibo':6000,'Viaggi':5000},
  2026: {'Auto':1500,'Casa':3000,'Spritz':800,'Riccardo':5000,'Regali':1000,'Ristoranti':3000,'Spesa Cibo':5000,'Viaggi':3000}
};

// Spese 2025 mensili (importi positivi, già verificati dal Google Sheet)
const SPESE_FAM_2025_DEFAULT = {
  'Auto':       [27,360,609,40,677,780,98,59,483,5,8,60],
  'Casa':       [729,71,1514,23,0,20,286,116,86,50,190,492],
  'Spritz':     [37,50,28,121,111,0,0,224,76,20,96,261],
  'Riccardo':   [303,408,287,231,743,270,306,170,251,958,879,962],
  'Regali':     [0,0,0,0,98,0,35,0,0,0,0,675],
  'Ristoranti': [244,185,299,248,354,221,147,364,173,209,244,203],
  'Spesa Cibo': [297,234,295,244,386,247,606,335,297,733,535,660],
  'Viaggi':     [0,0,0,81,230,1181,222,111,458,0,0,0]
};

// ── Setup iniziale: crea il foglio con le intestazioni ──
function setupFoglio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let foglio = ss.getSheetByName(NOME_FOGLIO);

  if (!foglio) {
    foglio = ss.insertSheet(NOME_FOGLIO);
  }

  // Intestazioni
  const intestazioni = ['ID', 'Data', 'Tipo', 'Importo (€)', 'Nota'];
  const primaRiga = foglio.getRange(1, 1, 1, intestazioni.length);
  primaRiga.setValues([intestazioni]);
  primaRiga.setFontWeight('bold');
  primaRiga.setBackground('#4CAF50');
  primaRiga.setFontColor('white');

  // Larghezza colonne
  foglio.setColumnWidth(1, 160); // ID
  foglio.setColumnWidth(2, 110); // Data
  foglio.setColumnWidth(3, 90);  // Tipo
  foglio.setColumnWidth(4, 110); // Importo
  foglio.setColumnWidth(5, 250); // Nota

  // Blocca intestazione
  foglio.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('Foglio "Movimenti" configurato correttamente!');
}

// ── Gestisce le richieste GET: serve l'HTML o risponde come API ──
function doGet(e) {
  const action   = e.parameter && e.parameter.action;
  const callback = e.parameter && e.parameter.callback;

  // Nessuna action → serve l'interfaccia HTML
  if (!action) {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Cassa Aziendale')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Modalità API (per chiamate dirette o JSONP)
  let risultato;
  if      (action === 'get')           risultato = leggiMovimenti();
  else if (action === 'add')           risultato = aggiungiMovimento(e.parameter);
  else if (action === 'modifica')      risultato = modificaMovimento(e.parameter);
  else if (action === 'elimina')       risultato = eliminaMovimento(e.parameter.id);
  else if (action === 'get_personal')  risultato = leggiPersonale();
  else if (action === 'add_personal')       risultato = aggiungiPersonale(e.parameter);
  else if (action === 'modifica_personal')  risultato = modificaPersonale(e.parameter);
  else if (action === 'elimina_personal')   risultato = eliminaPersonale(e.parameter.id);
  else if (action === 'get_categorie') risultato = leggiCategorie();
  else if (action === 'add_categoria') risultato = aggiungiCategoria(e.parameter.nome);
  else if (action === 'del_categoria') risultato = eliminaCategoria(e.parameter.nome);
  else if (action === 'get_fam')       risultato = leggiBudgetFam();
  else if (action === 'set_spesa_fam') risultato = salvaSpesaFam(e.parameter);
  else if (action === 'set_budget_fam')risultato = salvaBudgetFam(e.parameter);
  else if (action === 'del_spesa_fam') risultato = eliminaSpesaFam(e.parameter);
  else if (action === 'setup_fam')     risultato = setupFamCore();
  else if (action === 'set_token')     risultato = setSplitwiseTokenSafe(e.parameter);
  else if (action === 'sync_splitwise')risultato = aggiornaDaSplitwise();
  else if (action === 'install_trigger')risultato = installaTriggerSplitwise();
  else if (action === 'list_triggers') risultato = elencoTriggers();
  else if (action === 'get_prestito')  risultato = leggiPrestito();
  else if (action === 'add_prestito')  risultato = aggiungiPrestito(e.parameter);
  else if (action === 'del_prestito')  risultato = eliminaPrestito(e.parameter.id);
  else                                 risultato = { success: false, error: 'Azione sconosciuta: ' + action };

  const json = JSON.stringify(risultato);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Chiamata da google.script.run (dall'HTML interno) ──
function salvaMovimento(mov) {
  return aggiungiMovimento({
    id:      mov.id,
    data:    mov.data,
    tipo:    mov.tipo,
    importo: mov.importo.toString(),
    nota:    mov.nota || ''
  });
}

// ── Modifica un movimento esistente ──
function modificaMovimento(mov) {
  try {
    const foglio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_FOGLIO);
    const dati   = foglio.getDataRange().getValues();
    for (let i = 1; i < dati.length; i++) {
      if (dati[i][0].toString() === mov.id.toString()) {
        foglio.getRange(i + 1, 2).setValue(mov.data);
        foglio.getRange(i + 1, 3).setValue(mov.tipo);
        foglio.getRange(i + 1, 4).setValue(parseFloat(mov.importo));
        foglio.getRange(i + 1, 5).setValue(mov.nota || '');
        const colore = mov.tipo === 'incasso' ? '#E8F5E9' : '#FFEBEE';
        foglio.getRange(i + 1, 1, 1, 5).setBackground(colore);
        foglio.getRange(i + 1, 4).setNumberFormat('€#,##0.00');
        return { success: true };
      }
    }
    return { success: false, error: 'Record non trovato' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ── Elimina un movimento ──
function eliminaMovimento(id) {
  try {
    const foglio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_FOGLIO);
    const dati   = foglio.getDataRange().getValues();
    for (let i = 1; i < dati.length; i++) {
      if (dati[i][0].toString() === id.toString()) {
        foglio.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Record non trovato' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ── Cerca un id nella colonna A di un foglio. -1 se non c'è ──
// Serve per rendere gli "add" IDEMPOTENTI: la stessa chiamata ripetuta (retry del client,
// re-invio del sync, doppio tap) non deve creare una seconda riga. 25/08/2026.
function trovaRigaPerId(foglio, id) {
  const ultima = foglio.getLastRow();
  if (ultima < 2) return -1;
  const ids = foglio.getRange(2, 1, ultima - 1, 1).getValues();
  const cercato = id.toString();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0].toString() === cercato) return i + 2;
  }
  return -1;
}

// ── Aggiunge un nuovo movimento ──
function aggiungiMovimento(params) {
  // Il lock serializza le scritture concorrenti: senza, due chiamate in parallelo
  // leggono lo stesso getLastRow() e la formattazione finiva sulla riga sbagliata
  // (è il motivo delle righe con importo "38" invece di "€38,00").
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const id      = params.id      || Date.now().toString();
    const data    = params.data    || '';
    const tipo    = params.tipo    || '';
    const importo = parseFloat(params.importo) || 0;
    const nota    = params.nota    || '';

    if (!data || !tipo || importo <= 0) {
      return { success: false, error: 'Parametri mancanti o non validi' };
    }

    const foglio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_FOGLIO);

    // già registrato con questo id → non riscrivere, rispondi comunque OK
    if (trovaRigaPerId(foglio, id) !== -1) {
      return { success: true, id: id, duplicato: true };
    }

    foglio.appendRow([id, data, tipo, importo, nota]);

    const ultimaRiga = foglio.getLastRow();
    const coloreRiga = tipo === 'incasso' ? '#E8F5E9' : '#FFEBEE';
    foglio.getRange(ultimaRiga, 1, 1, 5).setBackground(coloreRiga);
    foglio.getRange(ultimaRiga, 4).setNumberFormat('€#,##0.00');

    return { success: true, id: id };

  } catch (err) {
    return { success: false, error: err.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ── Legge tutti i movimenti ──
function leggiMovimenti() {
  try {
    const foglio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_FOGLIO);
    const dati = foglio.getDataRange().getValues();

    if (dati.length <= 1) {
      return { success: true, data: [] };
    }

    const movimenti = dati.slice(1)
      .filter(riga => riga[0] !== '')
      .map(riga => ({
        id:      riga[0].toString(),
        data:    formattaData(riga[1]),
        tipo:    riga[2].toString(),
        importo: parseFloat(riga[3]) || 0,
        nota:    riga[4] ? riga[4].toString() : ''
      }))
      .sort((a, b) => {
        const da = new Date(a.data);
        const db = new Date(b.data);
        return db - da || b.id.localeCompare(a.id);
      });

    return { success: true, data: movimenti };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ── Modifica un movimento personale ──
function modificaPersonale(mov) {
  try {
    const foglio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_FOGLIO_PERSONALE);
    if (!foglio) return { success: false, error: 'Foglio "Personale" non trovato' };
    const dati = foglio.getDataRange().getValues();
    for (let i = 1; i < dati.length; i++) {
      if (dati[i][0].toString() === mov.id.toString()) {
        foglio.getRange(i + 1, 2).setValue(mov.data);
        foglio.getRange(i + 1, 3).setValue(mov.categoria);
        foglio.getRange(i + 1, 4).setValue(mov.tipo);
        foglio.getRange(i + 1, 5).setValue(parseFloat(mov.importo));
        foglio.getRange(i + 1, 6).setValue(mov.nota || '');
        foglio.getRange(i + 1, 1, 1, 6).setBackground(mov.tipo === 'entrata' ? '#E8F5E9' : '#FFEBEE');
        foglio.getRange(i + 1, 5).setNumberFormat('€#,##0.00');
        return { success: true };
      }
    }
    return { success: false, error: 'Record non trovato' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ── Elimina un movimento personale ──
function eliminaPersonale(id) {
  try {
    const foglio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_FOGLIO_PERSONALE);
    if (!foglio) return { success: false, error: 'Foglio "Personale" non trovato' };
    const dati = foglio.getDataRange().getValues();
    for (let i = 1; i < dati.length; i++) {
      if (dati[i][0].toString() === id.toString()) {
        foglio.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Record non trovato' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ── Setup foglio Personale ──
function setupFoglioPersonale() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let foglio = ss.getSheetByName(NOME_FOGLIO_PERSONALE);
  if (!foglio) foglio = ss.insertSheet(NOME_FOGLIO_PERSONALE);

  const intestazioni = ['ID', 'Data', 'Categoria', 'Tipo', 'Importo (€)', 'Nota'];
  const primaRiga = foglio.getRange(1, 1, 1, intestazioni.length);
  primaRiga.setValues([intestazioni]);
  primaRiga.setFontWeight('bold');
  primaRiga.setBackground('#1565C0');
  primaRiga.setFontColor('white');
  foglio.setColumnWidth(1, 160);
  foglio.setColumnWidth(2, 110);
  foglio.setColumnWidth(3, 180);
  foglio.setColumnWidth(4, 90);
  foglio.setColumnWidth(5, 110);
  foglio.setColumnWidth(6, 250);
  foglio.setFrozenRows(1);
  SpreadsheetApp.getUi().alert('Foglio "Personale" configurato correttamente!');
}

// ── Aggiunge movimento personale ──
function aggiungiPersonale(params) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const id        = params.id        || Date.now().toString();
    const data      = params.data      || '';
    const categoria = params.categoria || '';
    const tipo      = params.tipo      || '';
    const importo   = parseFloat(params.importo) || 0;
    const nota      = params.nota      || '';

    if (!data || !categoria || !tipo || importo <= 0)
      return { success: false, error: 'Parametri mancanti o non validi' };

    const foglio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_FOGLIO_PERSONALE);
    if (!foglio) return { success: false, error: 'Foglio "Personale" non trovato. Esegui setupFoglioPersonale().' };

    // stesso id già presente → niente seconda riga (vedi trovaRigaPerId)
    if (trovaRigaPerId(foglio, id) !== -1) {
      return { success: true, id: id, duplicato: true };
    }

    foglio.appendRow([id, data, categoria, tipo, importo, nota]);
    const ultimaRiga = foglio.getLastRow();
    foglio.getRange(ultimaRiga, 1, 1, 6).setBackground(tipo === 'entrata' ? '#E8F5E9' : '#FFEBEE');
    foglio.getRange(ultimaRiga, 5).setNumberFormat('€#,##0.00');

    return { success: true, id: id };
  } catch (err) {
    return { success: false, error: err.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ── Legge movimenti personali ──
function leggiPersonale() {
  try {
    const foglio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_FOGLIO_PERSONALE);
    if (!foglio) return { success: true, data: [] };

    const dati = foglio.getDataRange().getValues();
    if (dati.length <= 1) return { success: true, data: [] };

    const movimenti = dati.slice(1)
      .filter(r => r[0] !== '')
      .map(r => ({
        id:        r[0].toString(),
        data:      formattaData(r[1]),
        categoria: r[2].toString(),
        tipo:      r[3].toString(),
        importo:   parseFloat(r[4]) || 0,
        nota:      r[5] ? r[5].toString() : ''
      }))
      .sort((a, b) => {
        const da = new Date(a.data), db = new Date(b.data);
        return db - da || b.id.localeCompare(a.id);
      });

    return { success: true, data: movimenti };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════
//  PRESTITO AZIENDA → conto temporaneo di rientro del debito
//  Foglio: ID | Data | Tipo (rata|extra) | Importo (€) | Nota
//  Il foglio si crea da solo al primo utilizzo: nessun setup manuale.
// ═══════════════════════════════════════════════════════════════════

function getFoglioPrestito(creaSeManca) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let foglio = ss.getSheetByName(NOME_FOGLIO_PRESTITO);
  if (!foglio && creaSeManca) {
    foglio = ss.insertSheet(NOME_FOGLIO_PRESTITO);
    const intestazioni = ['ID', 'Data', 'Tipo', 'Importo (€)', 'Nota'];
    const primaRiga = foglio.getRange(1, 1, 1, intestazioni.length);
    primaRiga.setValues([intestazioni]);
    primaRiga.setFontWeight('bold');
    primaRiga.setBackground('#6A1B9A');
    primaRiga.setFontColor('white');
    foglio.setColumnWidth(1, 160);
    foglio.setColumnWidth(2, 110);
    foglio.setColumnWidth(3, 90);
    foglio.setColumnWidth(4, 110);
    foglio.setColumnWidth(5, 250);
    foglio.setFrozenRows(1);
  }
  return foglio;
}

// ── Aggiunge un versamento di rientro ──
function aggiungiPrestito(params) {
  try {
    const id      = params.id   || Date.now().toString();
    const data    = params.data || '';
    const tipo    = (params.tipo === 'extra') ? 'extra' : 'rata';
    const importo = parseFloat(params.importo) || 0;
    const nota    = params.nota || '';

    if (!data || importo <= 0)
      return { success: false, error: 'Parametri mancanti o non validi' };

    const foglio = getFoglioPrestito(true);
    foglio.appendRow([id, data, tipo, importo, nota]);
    const ultimaRiga = foglio.getLastRow();
    foglio.getRange(ultimaRiga, 1, 1, 5).setBackground(tipo === 'extra' ? '#E1F5FE' : '#F3E5F5');
    foglio.getRange(ultimaRiga, 4).setNumberFormat('€#,##0.00');

    return { success: true, id: id };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ── Legge i versamenti di rientro ──
function leggiPrestito() {
  try {
    const foglio = getFoglioPrestito(false);
    if (!foglio) return { success: true, data: [] };

    const dati = foglio.getDataRange().getValues();
    if (dati.length <= 1) return { success: true, data: [] };

    const versamenti = dati.slice(1)
      .filter(r => r[0] !== '')
      .map(r => ({
        id:      r[0].toString(),
        data:    formattaData(r[1]),
        tipo:    r[2] ? r[2].toString() : 'rata',
        importo: parseFloat(r[3]) || 0,
        nota:    r[4] ? r[4].toString() : ''
      }))
      .sort((a, b) => {
        const da = new Date(a.data), db = new Date(b.data);
        return db - da || b.id.localeCompare(a.id);
      });

    return { success: true, data: versamenti };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ── Elimina un versamento ──
function eliminaPrestito(id) {
  try {
    const foglio = getFoglioPrestito(false);
    if (!foglio) return { success: false, error: 'Foglio "Prestito" non trovato' };
    const dati = foglio.getDataRange().getValues();
    for (let i = 1; i < dati.length; i++) {
      if (dati[i][0].toString() === id.toString()) {
        foglio.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Record non trovato' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ── Setup foglio Categorie ──
function setupFoglioCategorie() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let foglio = ss.getSheetByName(NOME_FOGLIO_CATEGORIE);
  if (!foglio) foglio = ss.insertSheet(NOME_FOGLIO_CATEGORIE);
  const primaRiga = foglio.getRange(1, 1);
  primaRiga.setValue('Categoria');
  primaRiga.setFontWeight('bold');
  primaRiga.setBackground('#6A1B9A');
  primaRiga.setFontColor('white');
  foglio.setColumnWidth(1, 260);
  foglio.setFrozenRows(1);
  if (foglio.getLastRow() <= 1) {
    const defCats = ['Abbigliamento','Assicurazione Vita','Auto','Casa','Cane','Cerimonie','Cultura / ChatGPT','Fotografia','Informatica','Riccardo','Senza Categoria','Moto','Multe','Regali','Ristorante / Asporti / Bar','Salute','Spesa Cibo','Sport','Viaggi'];
    defCats.forEach((c, i) => foglio.getRange(i + 2, 1).setValue(c));
  }
  SpreadsheetApp.getUi().alert('Foglio "Categorie" configurato!');
}

// ── Leggi categorie ──
function leggiCategorie() {
  try {
    const foglio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_FOGLIO_CATEGORIE);
    if (!foglio || foglio.getLastRow() <= 1) return { success: true, data: [] };
    const dati = foglio.getRange(2, 1, foglio.getLastRow() - 1, 1).getValues();
    return { success: true, data: dati.map(r => r[0].toString()).filter(c => c !== '') };
  } catch(err) { return { success: false, error: err.toString() }; }
}

// ── Aggiungi categoria ──
function aggiungiCategoria(nome) {
  try {
    if (!nome || nome.trim() === '') return { success: false, error: 'Nome vuoto' };
    const foglio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_FOGLIO_CATEGORIE);
    if (!foglio) return { success: false, error: 'Foglio non trovato. Esegui setupFoglioCategorie().' };
    foglio.appendRow([nome.trim()]);
    return { success: true };
  } catch(err) { return { success: false, error: err.toString() }; }
}

// ── Elimina categoria ──
function eliminaCategoria(nome) {
  try {
    const foglio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_FOGLIO_CATEGORIE);
    if (!foglio) return { success: false, error: 'Foglio non trovato' };
    const dati = foglio.getDataRange().getValues();
    for (let i = 1; i < dati.length; i++) {
      if (dati[i][0].toString().trim() === nome.trim()) {
        foglio.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Categoria non trovata' };
  } catch(err) { return { success: false, error: err.toString() }; }
}

function formattaData(valore) {
  if (!valore) return '';
  if (valore instanceof Date) {
    const y = valore.getFullYear();
    const m = String(valore.getMonth() + 1).padStart(2, '0');
    const g = String(valore.getDate()).padStart(2, '0');
    return `${y}-${m}-${g}`;
  }
  return valore.toString();
}

// ═══════════════════════════════════════════════════════════════════
//  BUDGET FAMILIARE — Setup, lettura, scrittura
// ═══════════════════════════════════════════════════════════════════

// Setup completo: crea i 2 fogli, popola dati 2025 + budget 2025/2026
// Eseguibile sia dall'editor (con alert UI) sia via webapp (action=setup_fam)
function setupFoglioBudgetFamiliare() {
  const res = setupFamCore();
  try { SpreadsheetApp.getUi().alert(res.message); } catch(e) { /* contesto webapp, niente UI */ }
  return res;
}

function setupFamCore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Foglio BUDGET ANNUALE
  let fb = ss.getSheetByName(NOME_FOGLIO_BUDGET_FAM);
  if (!fb) fb = ss.insertSheet(NOME_FOGLIO_BUDGET_FAM);
  fb.clear();
  const hb = fb.getRange(1, 1, 1, 3);
  hb.setValues([['Categoria','Anno','BudgetAnnuale']]);
  hb.setFontWeight('bold').setBackground('#7E57C2').setFontColor('white');
  fb.setColumnWidth(1, 180); fb.setColumnWidth(2, 80); fb.setColumnWidth(3, 140);
  fb.setFrozenRows(1);

  const righeBudget = [];
  [2025, 2026].forEach(anno => {
    CATEGORIE_FAM.forEach(cat => {
      righeBudget.push([cat, anno, BUDGET_FAM_DEFAULT[anno][cat] || 0]);
    });
  });
  fb.getRange(2, 1, righeBudget.length, 3).setValues(righeBudget);
  fb.getRange(2, 3, righeBudget.length, 1).setNumberFormat('€#,##0');

  // Foglio SPESE MENSILI (con breakdown E/L per 2026+)
  let fs = ss.getSheetByName(NOME_FOGLIO_SPESE_FAM);
  if (!fs) fs = ss.insertSheet(NOME_FOGLIO_SPESE_FAM);
  fs.clear();
  const hs = fs.getRange(1, 1, 1, 6);
  hs.setValues([['Categoria','Anno','Mese','Importo','ImportoE','ImportoL']]);
  hs.setFontWeight('bold').setBackground('#FF7043').setFontColor('white');
  fs.setColumnWidth(1, 180); fs.setColumnWidth(2, 80); fs.setColumnWidth(3, 80);
  fs.setColumnWidth(4, 120); fs.setColumnWidth(5, 100); fs.setColumnWidth(6, 100);
  fs.setFrozenRows(1);

  const righeSpese = [];
  CATEGORIE_FAM.forEach(cat => {
    const arr = SPESE_FAM_2025_DEFAULT[cat] || [];
    for (let m = 0; m < 12; m++) {
      if (arr[m] != null) righeSpese.push([cat, 2025, m + 1, arr[m], '', '']);
    }
  });
  if (righeSpese.length) {
    fs.getRange(2, 1, righeSpese.length, 6).setValues(righeSpese);
    fs.getRange(2, 4, righeSpese.length, 3).setNumberFormat('€#,##0');
  }

  return {
    success: true,
    righeBudget: righeBudget.length,
    righeSpese: righeSpese.length,
    message: 'Budget Familiare configurato!\n' +
      '- ' + righeBudget.length + ' righe budget (2025 + 2026)\n' +
      '- ' + righeSpese.length + ' righe spese 2025'
  };
}

// Migra schema esistente: aggiunge colonne ImportoE/ImportoL se mancanti
function migraSpeseFamSchema() {
  const fs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_FOGLIO_SPESE_FAM);
  if (!fs) return { success: false, error: 'Foglio non trovato' };
  const lastCol = fs.getLastColumn();
  if (lastCol >= 6) return { success: true, message: 'Schema gia esteso (' + lastCol + ' colonne)' };
  fs.getRange(1, 5, 1, 2).setValues([['ImportoE','ImportoL']])
    .setFontWeight('bold').setBackground('#FF7043').setFontColor('white');
  fs.setColumnWidth(5, 100); fs.setColumnWidth(6, 100);
  const lastRow = fs.getLastRow();
  if (lastRow > 1) {
    fs.getRange(2, 5, lastRow - 1, 2).setNumberFormat('€#,##0');
  }
  return { success: true, message: 'Schema esteso: aggiunte ImportoE/ImportoL' };
}

// Legge tutto: budget annuali + spese mensili (con breakdown E/L se presente)
function leggiBudgetFam() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fb = ss.getSheetByName(NOME_FOGLIO_BUDGET_FAM);
    const fs = ss.getSheetByName(NOME_FOGLIO_SPESE_FAM);

    const budget = [];
    if (fb && fb.getLastRow() > 1) {
      const dati = fb.getRange(2, 1, fb.getLastRow() - 1, 3).getValues();
      dati.forEach(r => {
        if (r[0]) budget.push({
          categoria: r[0].toString(),
          anno: parseInt(r[1]) || 0,
          importo: parseFloat(r[2]) || 0
        });
      });
    }

    const spese = [];
    if (fs && fs.getLastRow() > 1) {
      const ncol = Math.max(4, fs.getLastColumn());
      const dati = fs.getRange(2, 1, fs.getLastRow() - 1, ncol).getValues();
      dati.forEach(r => {
        if (r[0]) spese.push({
          categoria: r[0].toString(),
          anno: parseInt(r[1]) || 0,
          mese: parseInt(r[2]) || 0,
          importo: parseFloat(r[3]) || 0,
          importoE: r.length > 4 && r[4] !== '' ? (parseFloat(r[4]) || 0) : null,
          importoL: r.length > 5 && r[5] !== '' ? (parseFloat(r[5]) || 0) : null
        });
      });
    }

    return { success: true, budget: budget, spese: spese };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// Upsert spesa (categoria + anno + mese identificano univocamente)
// Accetta opzionalmente importoE e importoL (quote pagate da Erika/Luca)
function salvaSpesaFam(params) {
  try {
    const cat      = (params.categoria || '').toString();
    const anno     = parseInt(params.anno) || 0;
    const mese     = parseInt(params.mese) || 0;
    const importo  = parseFloat(params.importo) || 0;
    const hasE     = params.importoE !== undefined && params.importoE !== '';
    const hasL     = params.importoL !== undefined && params.importoL !== '';
    const importoE = hasE ? (parseFloat(params.importoE) || 0) : null;
    const importoL = hasL ? (parseFloat(params.importoL) || 0) : null;

    if (!cat || !anno || !mese || mese < 1 || mese > 12) {
      return { success: false, error: 'Parametri non validi' };
    }

    let fs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_FOGLIO_SPESE_FAM);
    if (!fs) return { success: false, error: 'Foglio "SpeseFamiliari" non trovato. Esegui setupFoglioBudgetFamiliare().' };

    // Auto-migra schema se servono le colonne E/L
    if ((hasE || hasL) && fs.getLastColumn() < 6) {
      fs.getRange(1, 5, 1, 2).setValues([['ImportoE','ImportoL']])
        .setFontWeight('bold').setBackground('#FF7043').setFontColor('white');
      fs.setColumnWidth(5, 100); fs.setColumnWidth(6, 100);
    }

    const dati = fs.getDataRange().getValues();
    for (let i = 1; i < dati.length; i++) {
      if (dati[i][0].toString() === cat && parseInt(dati[i][1]) === anno && parseInt(dati[i][2]) === mese) {
        fs.getRange(i + 1, 4).setValue(importo).setNumberFormat('€#,##0');
        if (hasE) fs.getRange(i + 1, 5).setValue(importoE).setNumberFormat('€#,##0');
        if (hasL) fs.getRange(i + 1, 6).setValue(importoL).setNumberFormat('€#,##0');
        return { success: true, updated: true };
      }
    }
    const row = [cat, anno, mese, importo];
    if (hasE) row.push(importoE); else if (hasL) row.push('');
    if (hasL) row.push(importoL);
    fs.appendRow(row);
    const lr = fs.getLastRow();
    fs.getRange(lr, 4, 1, Math.max(1, row.length - 3)).setNumberFormat('€#,##0');
    return { success: true, created: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// Upsert budget annuale (categoria + anno univoci)
function salvaBudgetFam(params) {
  try {
    const cat     = (params.categoria || '').toString();
    const anno    = parseInt(params.anno) || 0;
    const importo = parseFloat(params.importo) || 0;

    if (!cat || !anno) return { success: false, error: 'Parametri non validi' };

    const fb = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_FOGLIO_BUDGET_FAM);
    if (!fb) return { success: false, error: 'Foglio "BudgetFamiliare" non trovato. Esegui setupFoglioBudgetFamiliare().' };

    const dati = fb.getDataRange().getValues();
    for (let i = 1; i < dati.length; i++) {
      if (dati[i][0].toString() === cat && parseInt(dati[i][1]) === anno) {
        fb.getRange(i + 1, 3).setValue(importo).setNumberFormat('€#,##0');
        return { success: true, updated: true };
      }
    }
    fb.appendRow([cat, anno, importo]);
    fb.getRange(fb.getLastRow(), 3).setNumberFormat('€#,##0');
    return { success: true, created: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// Elimina spesa (cat + anno + mese)
function eliminaSpesaFam(params) {
  try {
    const cat  = (params.categoria || '').toString();
    const anno = parseInt(params.anno) || 0;
    const mese = parseInt(params.mese) || 0;

    const fs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_FOGLIO_SPESE_FAM);
    if (!fs) return { success: false, error: 'Foglio non trovato' };

    const dati = fs.getDataRange().getValues();
    for (let i = 1; i < dati.length; i++) {
      if (dati[i][0].toString() === cat && parseInt(dati[i][1]) === anno && parseInt(dati[i][2]) === mese) {
        fs.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Record non trovato' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════════
//  SPLITWISE SYNC — Aggiornamento automatico dati 2026
// ═══════════════════════════════════════════════════════════════════

// Mapping gruppi Splitwise → categorie del Budget Familiare
const SPLITWISE_MAPPING = {
  'Auto':       [92066736],
  'Casa':       [92066790],
  'Regali':     [92066837, 93886585],
  'Riccardo':   [92066860],
  'Ristoranti': [92004523],
  'Spesa Cibo': [92066881],
  'Spritz':     [92066901],
  'Viaggi':     [92066921]
};
const SPLITWISE_ID_LUCA  = 8058310;
const SPLITWISE_ID_ERIKA = 43041864;

// Salva token Splitwise nelle Script Properties (guard: solo se non esiste o force=true)
function setSplitwiseTokenSafe(params) {
  const props = PropertiesService.getScriptProperties();
  const cur = props.getProperty('SPLITWISE_TOKEN');
  if (cur && params.force !== 'true') {
    return { success: false, error: 'Token già impostato. Usa force=true per sovrascrivere.' };
  }
  const t = (params.token || '').toString().trim();
  if (!t || t.length < 20) return { success: false, error: 'Token mancante o troppo corto.' };
  props.setProperty('SPLITWISE_TOKEN', t);
  return { success: true, message: 'Token salvato (' + t.length + ' caratteri).' };
}

// Aggiorna foglio SpeseFamiliari leggendo da Splitwise
function aggiornaDaSplitwise() {
  try {
    const token = PropertiesService.getScriptProperties().getProperty('SPLITWISE_TOKEN');
    if (!token) return { success: false, error: 'Token Splitwise non configurato.' };

    const anno = new Date().getFullYear();
    const datedAfter = (anno - 1) + '-12-31T00:00:00Z';
    const datedBefore = (anno + 1) + '-01-01T00:00:00Z';

    const headers = { 'Authorization': 'Bearer ' + token };
    const aggregato = {};

    for (const cat of Object.keys(SPLITWISE_MAPPING)) {
      aggregato[cat] = {
        tot: new Array(12).fill(0),
        E:   new Array(12).fill(0),
        L:   new Array(12).fill(0)
      };
      for (const gid of SPLITWISE_MAPPING[cat]) {
        const url = 'https://secure.splitwise.com/api/v3.0/get_expenses?group_id=' + gid +
                    '&dated_after=' + encodeURIComponent(datedAfter) +
                    '&dated_before=' + encodeURIComponent(datedBefore) +
                    '&limit=500';
        const resp = UrlFetchApp.fetch(url, { headers: headers, muteHttpExceptions: true });
        if (resp.getResponseCode() !== 200) {
          return { success: false, error: 'API Splitwise HTTP ' + resp.getResponseCode() + ' su gruppo ' + gid };
        }
        const data = JSON.parse(resp.getContentText());
        for (const e of (data.expenses || [])) {
          if (e.deleted_at) continue;
          if (e.payment) continue;
          const d = new Date(e.date);
          if (d.getFullYear() !== anno) continue;
          const m = d.getMonth(); // 0-based
          const cost = parseFloat(e.cost) || 0;
          aggregato[cat].tot[m] += cost;
          for (const u of (e.users || [])) {
            const paid = parseFloat(u.paid_share) || 0;
            if (u.user_id === SPLITWISE_ID_LUCA) aggregato[cat].L[m] += paid;
            else if (u.user_id === SPLITWISE_ID_ERIKA) aggregato[cat].E[m] += paid;
          }
        }
      }
    }

    // Scrivi nel foglio SpeseFamiliari
    let aggiornate = 0;
    for (const cat of Object.keys(aggregato)) {
      for (let m = 0; m < 12; m++) {
        const tot = Math.round(aggregato[cat].tot[m] * 100) / 100;
        const E   = Math.round(aggregato[cat].E[m]   * 100) / 100;
        const L   = Math.round(aggregato[cat].L[m]   * 100) / 100;
        // Skip mesi senza spese se anche il record nel foglio è 0/assente
        if (tot === 0 && E === 0 && L === 0) {
          // Comunque azzera (per consistenza) — passa solo se la cella non esiste
        }
        salvaSpesaFam({
          categoria: cat,
          anno: anno,
          mese: m + 1,
          importo: tot,
          importoE: E,
          importoL: L
        });
        aggiornate++;
      }
    }

    // Salva ultima sync
    PropertiesService.getScriptProperties().setProperty('SPLITWISE_LAST_SYNC', new Date().toISOString());

    return {
      success: true,
      anno: anno,
      righe_aggiornate: aggiornate,
      timestamp: new Date().toISOString()
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// Installa Time-driven Trigger giornaliero per aggiornaDaSplitwise
function installaTriggerSplitwise() {
  // Rimuovi eventuali trigger esistenti per evitare duplicati
  const existing = ScriptApp.getProjectTriggers();
  let removed = 0;
  existing.forEach(t => {
    if (t.getHandlerFunction() === 'aggiornaDaSplitwise') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  // Crea trigger ogni giorno alle 03:00 (timezone del progetto = Europe/Rome)
  ScriptApp.newTrigger('aggiornaDaSplitwise')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
  return { success: true, message: 'Trigger installato (giornaliero alle 03:00). Rimossi ' + removed + ' duplicati.' };
}

function elencoTriggers() {
  const triggers = ScriptApp.getProjectTriggers().map(t => ({
    handler: t.getHandlerFunction(),
    type: t.getEventType().toString(),
    source: t.getTriggerSource().toString()
  }));
  const last = PropertiesService.getScriptProperties().getProperty('SPLITWISE_LAST_SYNC');
  return { success: true, triggers: triggers, last_sync: last };
}

/* ══════════════════════════════════════════════════════════════════════════
   MANUTENZIONE — pulizia dei movimenti duplicati (25/08/2026)
   Fino ad oggi 'add' faceva appendRow senza controllare l'id: ogni re-invio
   (retry del client o re-push del sync) creava una riga in più con lo STESSO id.
   Ora l'add è idempotente; questa funzione ripulisce lo storico già sporcato.
   Uso: dedupDryRun() per vedere il conto, dedupEsegui() per cancellare davvero
   (fa prima una copia di sicurezza dei fogli toccati).
   ══════════════════════════════════════════════════════════════════════════ */

function dedupAnalizza() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const report = [];
  [NOME_FOGLIO, NOME_FOGLIO_PERSONALE, NOME_FOGLIO_PRESTITO].forEach(nome => {
    const foglio = ss.getSheetByName(nome);
    if (!foglio || foglio.getLastRow() < 2) return;
    const righe = foglio.getDataRange().getValues();
    const visti = {}, daEliminare = [];
    for (let i = 1; i < righe.length; i++) {
      const id = righe[i][0] ? righe[i][0].toString() : '';
      if (!id) continue;
      if (visti[id]) daEliminare.push({ riga: i + 1, id: id, valori: righe[i] });
      else visti[id] = i + 1;
    }
    report.push({ foglio: nome, righe: righe.length - 1, duplicati: daEliminare.length, dettaglio: daEliminare });
  });
  return report;
}

function dedupDryRun() {
  const report = dedupAnalizza();
  const testo = report.map(r =>
    r.foglio + ': ' + r.duplicati + ' righe doppie su ' + r.righe +
    (r.duplicati ? '\n  ' + r.dettaglio.map(d => 'riga ' + d.riga + ' — id ' + d.id + ' — ' + d.valori.slice(1).join(' · ')).join('\n  ') : '')
  ).join('\n\n');
  Logger.log(testo);
  return { success: true, report: report, testo: testo };
}

function dedupEsegui() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(60000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const report = dedupAnalizza();
    const stamp = Utilities.formatDate(new Date(), 'Europe/Rome', 'yyyyMMdd-HHmm');
    const esito = [];

    report.forEach(r => {
      if (!r.duplicati) { esito.push(r.foglio + ': niente da fare'); return; }
      const foglio = ss.getSheetByName(r.foglio);
      // copia di sicurezza prima di cancellare
      foglio.copyTo(ss).setName('BACKUP-' + r.foglio + '-' + stamp);
      // si cancella dal basso verso l'alto, altrimenti gli indici slittano
      r.dettaglio.map(d => d.riga).sort((a, b) => b - a).forEach(riga => foglio.deleteRow(riga));
      esito.push(r.foglio + ': eliminate ' + r.duplicati + ' righe doppie (backup: BACKUP-' + r.foglio + '-' + stamp + ')');
    });

    const testo = esito.join('\n');
    Logger.log(testo);
    return { success: true, testo: testo };
  } catch (err) {
    return { success: false, error: err.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
