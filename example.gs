/**
 * CONFIGURAZIONE (DA MODIFICARE)
 */
const CONFIG = {
  // URL del Foglio Google
  // Esempio: 'https://docs.google.com/spreadsheets/d/1aBcD.../edit'
  SHEET_URL: 'INSERISCI_QUI_L_URL_DEL_TUO_FOGLIO_GOOGLE', 
  
  // Il nome del TAB (foglio) specifico dove inserire i dati.
  SHEET_NAME: 'Biglietti Acquistati', 
  
  // Etichette Gmail
  LABEL_TO_PROCESS: 'ETICHETTA_DA_PROCESSARE', // Es. 'Nemo-COP'
  LABEL_PROCESSED: 'ETICHETTA_ARCHIVIO'       // Es. 'Nemo-Processed'
};

/**
 * Funzione Principale
 */
function processEmails() {
  const sheet = getSheet();
  if (!sheet) return;

  const processedLabel = getOrCreateLabel(CONFIG.LABEL_PROCESSED);
  const existingOrderIds = getExistingOrderIds(sheet); // Carica ID esistenti da Col L

  // Cerca thread
  const query = `label:${CONFIG.LABEL_TO_PROCESS} -label:${CONFIG.LABEL_PROCESSED}`;
  const threads = GmailApp.search(query, 0, 20);

  if (threads.length === 0) {
    Logger.log('Nessuna nuova email da processare.');
    return;
  }

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
       const data = extractData(message);
       
       // CONTROLLO DUPLICATI: Se ID Ordine esiste già (ed è valido), SALTA.
       if (data.orderId && data.orderId.length > 3 && existingOrderIds.has(data.orderId.toString())) {
         Logger.log(`[SKIP] Ordine duplicato: ${data.orderId}`);
         return; // Salta questo messaggio
       }
       
       // Preparazione Riga (Mappatura aggiornata: Inizia da Col C)
       // A=Check, B=Check, C=DataAcq, D=Artista, E=Settore, F=Luogo, G=DataEv, H=Qtà, I=Eur, J=Gbp, K=Usd, L=AcqDa, M=Pay, N=IdOrd, O=Vend, P=Acc, Q=Link
       const rowData = [
         data.date,           // Col C: DATA ACQUISTO
         data.artist || '',   // Col D: ARTISTA 
         data.sector || '',   // Col E: SETTORE
         data.location || '', // Col F: LUOGO
         data.eventDateRaw || '', // Col G: DATA EVENTO
         data.quantity || 1,  // Col H: Q.tà
         data.amountEUR,      // Col I: € SPESI
         data.amountGBP,      // Col J: £ SPESI
         data.amountUSD,      // Col K: $ SPESI
         'Nemo',              // Col L: ACQUISTATO DA
         data.paymentMethod,  // Col M: PAGAMENTO
         data.orderId,        // Col N: ID ORDINE
         data.vendor,         // Col O: VENDITORE
         data.account,        // Col P: ACCOUNT
         `https://mail.google.com/mail/u/0/#inbox/${message.getId()}` // Col Q: Link
       ];
       
       // Scrittura nella prossima riga libera basata sulla Colonna D (Artista)
       const nextRow = getNextEmptyRowByColumn(sheet, 4); // 4 = Colonna D
       Logger.log(`Scrivo alla riga ${nextRow} (Basato su l'ultimo Artista in Col D)`);
       sheet.getRange(nextRow, 3, 1, rowData.length).setValues([rowData]);
       
       Logger.log(`[OK] Inserito: ${data.vendor} - ID: ${data.orderId}`);
       
       // Aggiungiamo l'ID al set locale così se c'è un altro messaggio uguale nello stesso batch lo becchiamo
       if (data.orderId) existingOrderIds.add(data.orderId.toString());
    });
    
    // Segna la CONVERSAZIONE come processata
    thread.addLabel(processedLabel);
  });
}

/**
 * Logica di estrazione (Regex)
 */
function extractData(message) {
  const body = message.getPlainBody(); 
  const htmlBody = message.getBody(); // Usiamo anche l'HTML per estrazioni più precise (TicketOne)
  
  const from = message.getFrom();
  const to = message.getTo();
  const dateObj = message.getDate() || new Date();
  const dateStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd/MM/yyyy");

  let data = {
    date: dateStr,
    vendor: '',
    amountEUR: '',
    amountUSD: '',
    amountGBP: '',
    orderId: '',
    paymentMethod: '',
    account: extractEmailAddress(to),
    artist: '',
    eventDateRaw: '',
    sector: '',
    location: '',
    quantity: ''
  };

  // --- RICONOSCIMENTO VENDITORE E PATTERN SPECIFICI ---
  
  // TICKETONE
  if (/ticketone/i.test(from) || /ticketone/i.test(body)) {
    data.vendor = 'TicketOne';
    
    // ID Ordine
    const orderMatch = body.match(/Numero ordine:\s*(\d+)/i) || body.match(/ID Ordine è:\s*(\d+)/i);
    if (orderMatch) data.orderId = orderMatch[1];
    
    // Importo
    const amountMatch = body.match(/(?:Totale ordine|Importo):\s*€\s?(\d+[.,]\d{2})/i);
    if (amountMatch) data.amountEUR = amountMatch[1];
    
    // Data Evento
    const eventDateMatch = body.match(/(\d{2}\/\d{2}\/\d{4}), \d{2}:\d{2}/i);
    if (eventDateMatch) data.eventDateRaw = eventDateMatch[0]; 

    // Artista: USIAMO HTML (Più preciso, evita "Olly Data")
    // Cerca pattern: <th ...> ... <strong>Nome Artista</strong> ... </th>
    const artistHtmlMatch = htmlBody.match(/<th[^>]*>(?:[\s\S]*?<p[^>]*>)?[\s\S]*?<strong>\s*([^<]+?)\s*<\/strong>/i);
    
    if (artistHtmlMatch && artistHtmlMatch[1].trim().length < 50 && !/Dettagli|Ordine|Data|Ora/i.test(artistHtmlMatch[1])) {
       // Pulisce entità HTML base
       data.artist = artistHtmlMatch[1].trim().replace(/&amp;/g, '&').replace(/&quot;/g, '"');
    } else {
       // Fallback su Plain Text se HTML fallisce, ma proviamo a pulire "Data"
       const artistMatch = body.match(/Dettagli ordine\s+([\w\s\-]+?)(?:\sData|\n|$)/i);
       if (artistMatch && artistMatch[1].trim().length < 50) {
          data.artist = artistMatch[1].trim(); 
       }
    }

    // Settore/Posto - Logica Raffinata
    // Cerca stringhe comuni nei biglietti
    const sectorSpecificStr = body.match(/(Parterre in Piedi|Parterre|Tribuna|Distinti|Curva|Platea|Settore Numerato|Posto Unico)(?:,\s*[A-Za-z\s]+)?/i);
    if (sectorSpecificStr) {
      data.sector = sectorSpecificStr[0].trim();
    } else {
      // Fallback standard
      const sectorMatch = body.match(/(?:Settore|Posto|Ingresso):\s*([^\n\r]+)/i);
      if (sectorMatch) data.sector = sectorMatch[1].trim();
    }

    // Luogo
    const locationMatch = body.match(/(?:Presso|Luogo|Location):\s*([^\n\r]+)/i);
    if (locationMatch) data.location = locationMatch[1].trim();
    else {
        const knownVenues = /(San Siro|Stadio|Arena|Forum|Palazzo|Teatro|Unipol|Ippodromo)[^\n\r]*/i;
        const venueMatch = body.match(knownVenues);
        if (venueMatch) data.location = venueMatch[0].trim();
    }
  }
  
  // TICKETMASTER
  else if (/ticketmaster/i.test(from) || /ticketmaster/i.test(body)) {
    data.vendor = 'Ticketmaster';
    
    // ID Ordine (HTML)
    const orderMatch = htmlBody.match(/Numero d(?:’|')ordine[\s\S]{0,100}?(\d{5,})/i);
    if (orderMatch) data.orderId = orderMatch[1];
    
    // Importo (HTML)
    const amountMatch = htmlBody.match(/Totale[\s\S]*?<b>\s*(\d+[.,]\d{2})/i) ||
                        body.match(/Totale[\s\r\n]*(\d+[.,]\d{2})\s?€/i);
    if (amountMatch) {
        // Sostituisci punto con virgola per formato italiano
        data.amountEUR = amountMatch[1].replace('.', ',');
    }

    // Artista (HTML)
    const artistMatch = htmlBody.match(/dettagli relativi al tuo ordine[\s\S]*?<b>\s*([^<]+?)\s*<\/b>/i);
    if (artistMatch) {
       data.artist = artistMatch[1].trim();
    }

    // Luogo e Data (HTML)
    const infoMatches = htmlBody.match(/<font style="color:#262626!important">\s*([^<]+?)\s*<\/font>/gi);
    if (infoMatches) {
       infoMatches.forEach(matchContent => {
           const cleanText = matchContent.replace(/<[^>]+>/g, '').trim();
           if (/\d{4}/.test(cleanText) && /:/.test(cleanText)) {
               data.eventDateRaw = cleanText;
           } else {
               if (!data.location) data.location = cleanText;
           }
       });
    }

    // Quantità e Settore
    // 1. Cerca pattern "2 x SETTORE" (che sta fuori dal font grigio)
    // Snippet: <td...>2 x POSTO UNICO / Intero<br/>
    const qtySectorMatch = htmlBody.match(/>\s*(\d+)\s*x\s*([^<]+)/i);
    if (qtySectorMatch) {
        data.quantity = qtySectorMatch[1];
        // Se non abbiamo ancora un settore, usiamo questo testo parziale (es. POSTO UNICO / Intero)
        if (!data.sector) data.sector = qtySectorMatch[2].trim();
    }

    // 2. Cerca Settore specifico (grigio) - Sovrascrive se trovato perché più preciso
    const sectorHtmlMatch = htmlBody.match(/<font[^>]*style="[^"]*color:#646464[^"]*"[^>]*>([\s\S]*?)<\/font>/i);
    if (sectorHtmlMatch) {
         data.sector = sectorHtmlMatch[1].replace(/&nbsp;/g, ' ').trim();
    }
    
    // Tentativo quantità da stringa settore se presente "1 x"
    if (data.sector && /^(\d+)\s*x/.test(data.sector)) {
         data.quantity = data.sector.match(/^(\d+)/)[1];
    }
  }
  
  // AMAZON
  else if (/amazon/i.test(from)) {
    data.vendor = 'Amazon';
    const orderMatch = body.match(/Ordine #\s?([0-9-]{15,})/i);
    if (orderMatch) data.orderId = orderMatch[1];
     const amountMatch = body.match(/Totale:?\s?EUR\s?(\d+[.,]\d{2})/i);
     if (amountMatch) data.amountEUR = amountMatch[1];
     data.location = 'Online';
     data.sector = '-';
  }

  // --- FALLBACK GENERICI ---
  if (!data.vendor) {
    const vendorMatch = from.match(/^"?(.*?)"? <.*>$/);
    data.vendor = vendorMatch ? vendorMatch[1] : from;
  }
  if (!data.amountEUR && !data.amountUSD && !data.amountGBP) {
     const eurMatch = body.match(/Totale.*?(\d+[.,]\d{2})\s?€|€\s?(\d+[.,]\d{2})/i);
     if (eurMatch) data.amountEUR = eurMatch[1] || eurMatch[2];
  }
  if (!data.orderId) {
     const orderMatch = body.match(/(?:Ordine|Order|ID|Riferimento)\s?(?:#|n\.|numero|id)?\s?[:.]?\s?([A-Z0-9\-_]{5,})/i);
     if (orderMatch) data.orderId = orderMatch[1];
  }

  // Pulisci l'ID Ordine
  if (data.orderId) data.orderId = data.orderId.trim();

  return data;
}

/**
 * Helpers
 */
function extractEmailAddress(str) {
  const match = str.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/);
  return match ? match[1] : str;
}

function getOrCreateLabel(name) {
  let label = GmailApp.getUserLabelByName(name);
  if (!label) label = GmailApp.createLabel(name);
  return label;
}

// Aggiornato: Cerca duplicati in Colonna N (14)
function getExistingOrderIds(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set(); 
  // Colonna N è la 14
  const values = sheet.getRange(2, 14, lastRow - 1, 1).getValues();
  return new Set(values.flat().map(id => String(id).trim()));
}

function getSheet() {
  let ss;
  if (CONFIG.SHEET_URL) ss = SpreadsheetApp.openByUrl(CONFIG.SHEET_URL);
  else ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss ? ss.getSheetByName(CONFIG.SHEET_NAME) : null;
}

function setup() {
  getOrCreateLabel(CONFIG.LABEL_TO_PROCESS);
  getOrCreateLabel(CONFIG.LABEL_PROCESSED);
  SpreadsheetApp.getUi().alert('Setup completato.');
}

// Nuova funzione per trovare la prima riga vuota basandosi su una colonna specifica
function getNextEmptyRowByColumn(sheet, colIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2; // Foglio vuoto
  
  // Prendi tutti i valori della colonna specifica (es. D = 4)
  // nota: getRange(row, col, numRows, numCols)
  const values = sheet.getRange(1, colIndex, lastRow, 1).getValues();
  
  // Cerca dal basso verso l'alto l'ultima cella piena
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] && values[i][0] !== "") {
      return i + 2; // indice + 1 (perché array è 0-based) + 1 (per andare alla riga successiva)
    }
  }
  return 2; // Se tutto vuoto, inizia dalla 2
}

/**
 * Trigger giornaliero
 */
function createDailyTrigger() {
  // Cancella trigger esistenti per evitare duplicati
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // Crea nuovo trigger giornaliero alle ore 9 del mattino
  ScriptApp.newTrigger('processEmails')
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();
      
  Logger.log('Trigger giornaliero creato correttamente.');
}
