/**
 * ============================================================================
 * TITOLO:        Gmail to Sheets - Auto Extractor
 * DESCRIZIONE:   Estrae automaticamente dati (Ordini, Prezzi, Artisti) dalle email
 *                identificate da una specifica Etichetta Gmail.
 * AUTORE:        [Il tuo Nome/Github]
 * SETUP:         Vedi README.md per istruzioni dettagliate.
 * ============================================================================
 */

/**
 * CONFIGURAZIONE
 * Sostituisci i valori qui sotto con i tuoi dati reali.
 */
const CONFIG = {
  // L'URL completo del tuo Foglio Google
  SHEET_URL: 'INSERISCI_QUI_URL_TUO_FOGLIO_GOOGLE', 
  
  // Il nome esatto del TAB (es. 'Foglio1' o 'Ordini')
  SHEET_NAME: 'Foglio1', 
  
  // L'Etichetta Gmail da processare (es. 'DaProcessare')
  LABEL_TO_PROCESS: 'Nemo-COP',
  
  // L'Etichetta da applicare dopo aver finito (per non riprocessare le stesse mail)
  LABEL_PROCESSED: 'Nemo-Processed'
};

/**
 * Funzione Principale da impostare come Trigger (es. ogni ora)
 */
function processEmails() {
  const sheet = getSheet();
  if (!sheet) {
    Logger.log('Errore: Foglio non trovato. Controlla URL e Nome Tab.');
    return;
  }

  const processedLabel = getOrCreateLabel(CONFIG.LABEL_PROCESSED);
  const existingOrderIds = getExistingOrderIds(sheet); // Evita duplicati

  // Cerca solo le email con l'etichetta "Da Processare" e senza "Processato"
  const query = `label:${CONFIG.LABEL_TO_PROCESS} -label:${CONFIG.LABEL_PROCESSED}`;
  const threads = GmailApp.search(query, 0, 20); // Processa max 20 thread alla volta

  if (threads.length === 0) {
    Logger.log('Nessuna nuova email da processare.');
    return;
  }

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
       const data = extractData(message);
       
       // CONTROLLO DUPLICATI: Se ID Ordine esiste già, SALTA.
       if (data.orderId && data.orderId.length > 3 && existingOrderIds.has(data.orderId.toString())) {
         Logger.log(`[SKIP] Ordine duplicato: ${data.orderId}`);
         return; 
       }
       
       // PREPARAZIONE RIGA (Personalizza qui le colonne)
       const rowData = [
         data.date,           // Data
         data.artist || '',   // Artista
         data.sector || '',   // Settore
         data.location || '', // Luogo
         data.eventDateRaw || '', // Data Evento
         data.quantity || 1,  // Quantità
         data.amountEUR,      // Importo
         data.vendor,         // Venditore (TicketOne, Ticketmaster, Amazon...)
         data.orderId,        // ID Ordine
         `https://mail.google.com/mail/u/0/#inbox/${message.getId()}` // Link Email
       ];
       
       // Scrittura nella prossima riga libera
       // (Qui usa una logica "Smart" che controlla l'ultima riga piena della colonna Artista)
       const nextRow = getNextEmptyRowByColumn(sheet, 2); // Esempio: Colonna B (2)
       sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
       
       Logger.log(`[OK] Inserito: ${data.vendor} - ID: ${data.orderId}`);
       
       if (data.orderId) existingOrderIds.add(data.orderId.toString());
    });
    
    // Segna la CONVERSAZIONE come processata
    thread.addLabel(processedLabel);
  });
}

/**
 * Logica di estrazione (Cuore dello script)
 * Qui usiamo Regex per "pescare" i dati dal testo delle email.
 */
function extractData(message) {
  const body = message.getPlainBody(); 
  const htmlBody = message.getBody(); // Utile per strutture complesse HTML
  const from = message.getFrom();
  const dateObj = message.getDate() || new Date();
  const dateStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd/MM/yyyy");

  let data = {
    date: dateStr,
    vendor: '',
    amountEUR: '',
    orderId: '',
    artist: '',
    eventDateRaw: '',
    sector: '',
    location: '',
    quantity: ''
  };

  // --- ESEMPIO 1: TICKETONE ---
  if (/ticketone/i.test(from) || /ticketone/i.test(body)) {
    data.vendor = 'TicketOne';
    
    // ID Ordine
    const orderMatch = body.match(/Numero ordine:\s*(\d+)/i);
    if (orderMatch) data.orderId = orderMatch[1];
    
    // Importo
    const amountMatch = body.match(/Importo:\s*€\s?(\d+[.,]\d{2})/i);
    if (amountMatch) data.amountEUR = amountMatch[1];
    
    // Artista (Esempio Parsing HTML per precisione)
    const artistHtmlMatch = htmlBody.match(/<th[^>]*>[\s\S]*?<strong>\s*([^<]+?)\s*<\/strong>/i);
    if (artistHtmlMatch) data.artist = artistHtmlMatch[1].trim();

    // Settore (Keywords)
    const sectorMatch = body.match(/(Parterre|Tribuna|Distinti|Posto Unico)/i);
    if (sectorMatch) data.sector = sectorMatch[0].trim();
  }
  
  // --- ESEMPIO 2: TICKETMASTER ---
  else if (/ticketmaster/i.test(from) || /ticketmaster/i.test(body)) {
    data.vendor = 'Ticketmaster';
    
    // ID Ordine (HTML robusto)
    const orderMatch = htmlBody.match(/Numero d(?:’|')ordine[\s\S]{0,100}?(\d{5,})/i);
    if (orderMatch) data.orderId = orderMatch[1];
    
    // Importo
    const amountMatch = htmlBody.match(/Totale[\s\S]*?<b>\s*(\d+[.,]\d{2})/i);
    if (amountMatch) data.amountEUR = amountMatch[1].replace('.', ','); // Fix formato IT

    // Quantità "2 x Settore"
    const qtyMatch = htmlBody.match(/>\s*(\d+)\s*x\s*([^<]+)/i);
    if (qtyMatch) {
       data.quantity = qtyMatch[1];
       if (!data.sector) data.sector = qtyMatch[2].trim();
    }
  }
  
  // --- ESEMPIO 3: AMAZON ---
  else if (/amazon/i.test(from)) {
    data.vendor = 'Amazon';
    const orderMatch = body.match(/Ordine #\s?([0-9-]{15,})/i);
    if (orderMatch) data.orderId = orderMatch[1];
    const amountMatch = body.match(/Totale:?\s?EUR\s?(\d+[.,]\d{2})/i);
    if (amountMatch) data.amountEUR = amountMatch[1];
  }

  // Fallback generico
  if (!data.vendor) {
    data.vendor = from; 
  }

  return data;
}

/**
 * =======================
 * HELPER FUNCTIONS
 * =======================
 */

function getOrCreateLabel(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) label = GmailApp.createLabel(name);
  return label;
}

// Lettura ID esistenti per evitare duplicati
function getExistingOrderIds(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set(); 
  // Modifica l'indice colonna (es. 9 per colonna I) in base a dove salvi l'ID Ordine
  const values = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
  return new Set(values.flat().map(id => String(id).trim()));
}

function getSheet() {
  let ss;
  if (CONFIG.SHEET_URL && CONFIG.SHEET_URL.startsWith('http')) {
    ss = SpreadsheetApp.openByUrl(CONFIG.SHEET_URL);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  return ss ? ss.getSheetByName(CONFIG.SHEET_NAME) : null;
}

// Trova la prima riga vuota basandosi su una colonna specifica (utile se hai righe semivuote)
function getNextEmptyRowByColumn(sheet, colIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;
  const values = sheet.getRange(1, colIndex, lastRow, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] && values[i][0] !== "") return i + 2;
  }
  return 2;
}

