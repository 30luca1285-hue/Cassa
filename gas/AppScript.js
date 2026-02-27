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

const NOME_FOGLIO           = 'Movimenti';
const NOME_FOGLIO_PERSONALE = 'Personale';
const NOME_FOGLIO_CATEGORIE = 'Categorie';

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
