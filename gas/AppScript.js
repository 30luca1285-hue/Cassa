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

// ── Aggiunge un nuovo movimento ──
function aggiungiMovimento(params) {
  try {
    const id      = params.id      || Date.now().toString();
    const data    = params.data    || '';
    const tipo    = params.tipo    || '';
    const importo = parseFloat(params.importo) || 0;
    const nota    = params.nota    || '';

    if (!data || !tipo || importo <= 0) {
      return { success: false, error: 'Parametri mancanti o non validi' };
    }

    const foglio = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_FOGLIO);
    foglio.appendRow([id, data, tipo, importo, nota]);

    const ultimaRiga = foglio.getLastRow();
    const coloreRiga = tipo === 'incasso' ? '#E8F5E9' : '#FFEBEE';
    foglio.getRange(ultimaRiga, 1, 1, 5).setBackground(coloreRiga);
    foglio.getRange(ultimaRiga, 4).setNumberFormat('€#,##0.00');

    return { success: true, id: id };

  } catch (err) {
    return { success: false, error: err.toString() };
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
  try {
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

    foglio.appendRow([id, data, categoria, tipo, importo, nota]);
    const ultimaRiga = foglio.getLastRow();
    foglio.getRange(ultimaRiga, 1, 1, 6).setBackground(tipo === 'entrata' ? '#E8F5E9' : '#FFEBEE');
    foglio.getRange(ultimaRiga, 5).setNumberFormat('€#,##0.00');

    return { success: true, id: id };
  } catch (err) {
    return { success: false, error: err.toString() };
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
    const defCats = ['Abbigliamento','Assicurazione Vita','Auto','Casa','Camper','Cane','Cerimonie','Cultura / ChatGPT','Fotografia','Informatica','Riccardo','Senza Categoria','Moto','Multe','Regali','Ristorante / Asporti / Bar','Salute','Spesa Cibo','Sport','Viaggi'];
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
