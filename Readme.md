📧 Gmail to Google Sheets - Auto Cleaner & Extractor
Questo progetto mostra come costruire un sistema automatico (senza costi) per catturare dati importanti dalle email (come Conferme Ordine, Biglietti, Fatture) e salvarli ordinatamente in un Foglio Google.

Il tutto funziona tramite Google Apps Script e le Etichette (Labels) di Gmail.

🎯 Obiettivo
Creare uno script che:

Controlla periodicamente se ci sono nuove email con una certa etichetta (es. DaProcessare).
Legge l'email e usa Regex/HTML Parsing per estrarre: ID Ordine, Prezzo, Oggetto/Artista, Quantità.
Scrive i dati in una riga di Google Sheets.
Sposta l'email in un'etichetta Processato per non leggerla due volte.
🚀 Guida all'Installazione (Step-by-Step)
1. Prepara il Foglio Google
Crea un nuovo Google Sheet (o usane uno esistente).
Crea una riga di intestazione (Data, Artista, Settore, Prezzo, ID Ordine, Link, ecc.).
Copia l'URL completo del foglio (dalla barra degli indirizzi). Ti servirà dopo.
2. Configura Gmail
Crea un'etichetta (Label) in Gmail dove sposterai le email da processare (es. Nemo-COP).
(Opzionale) Crea un filtro Gmail che applica automaticamente questa etichetta alle email di TicketOne/Ticketmaster/Amazon.
3. Installa lo Script
Dal tuo Foglio Google, vai su menù: Estensioni > Apps Script.
Si aprirà un editor di codice.
Cancella tutto il codice presente nel file Codice.gs.
Copia e incolla il contenuto del file 
Code_Example.gs
 fornito in questa repo.
4. Personalizza il Codice
Vai in alto nel file incollato, nella sezione CONFIG:

javascript
const CONFIG = {
  SHEET_URL: 'INCOLLA_QUI_IL_TUO_URL_DI_GOOGLE_SHEET',
  SHEET_NAME: 'Foglio1', // Nome esatto della tab in basso
  LABEL_TO_PROCESS: 'Nemo', // Nome della tua etichetta Gmail
  LABEL_PROCESSED: 'Dory' // Etichetta di archivio (creata in automatico)
};
5. Prova manuale
Assicurati di avere almeno un'email con l'etichetta Nemo-COP.
Nell'editor Apps Script, seleziona la funzione 
processEmails
 dal menù a tendina in alto.
Premi Esegui.
La prima volta ti chiederà i permessi (clicca su "Vedi dettagli" > "Vai a ... (non sicuro)" > "Consenti").
Controlla il tuo foglio: dovresti vedere i dati estratti!
6. Automazione (Opzionale)
Se vuoi che lo script giri da solo ogni ora:

Clicca sull'icona Sveglia (Attivatori/Triggers) nel menu a sinistra.
Clicca Aggiungi Attivatore.
Funzione: 
processEmails
.
Fonte evento: Vincolato al tempo.
Tipo timer: Timer orario -> Ogni ora.
🛠️ Come personalizzare l'estrazione dati
La parte "intelligente" è nella funzione 
extractData()
. Qui usiamo Espressioni Regolari (Regex) per trovare i dati.

Esempio per trovare un prezzo (es. "Importo: € 20,00"):

javascript
// Cerca "Importo", eventuali spazi, simbolo €, numeri, virgola, decimali
const priceMatch = body.match(/Importo:\s*€\s?(\d+[.,]\d{2})/i);
if (priceMatch) data.amount = priceMatch[1];
Puoi aggiungere tutti i else if che vuoi per nuovi venditori (eBay, PayPal, Trenitalia, ecc.).

🔒 Sicurezza e Privacy
Questo script gira esclusivamente sul tuo account Google.
Nessun dato viene inviato a server esterni.
Il codice è visibile e modificabile solo da te.
Nota: Non condividere mai con nessuno l'URL del tuo foglio se contiene dati sensibili.

GitHub Repository
Se questo tutorial ti è stato utile, lascia una stella ⭐ alla repo!