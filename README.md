# scheda-palestra-app-script

# Allenamento Tracker

## Scopo del Progetto
Il progetto `Allenamento Tracker` è stato sviluppato per fornire una soluzione semplice ed efficace per tracciare e analizzare gli allenamenti utilizzando Google Sheets e Google Apps Script. L'obiettivo è offrire un sistema flessibile e automatizzato che permetta agli utenti di:

- Creare e gestire schede di allenamento personalizzate.
- Monitorare i progressi e registrare gli allenamenti.
- Visualizzare analisi e statistiche direttamente nel foglio di calcolo (TO DO).

## Struttura del Progetto
Il progetto è organizzato nei seguenti file principali:

1. **`workoutManager.gs`**: Gestisce la logica principale per il tracciamento degli allenamenti, inclusa la gestione della cache e il salvataggio dei dati in batch.
2. **`config.gs`**: Contiene la configurazione globale, come i nomi dei fogli di Google Sheets e le chiavi di cache.
3. **`ui.gs`**: Gestisce la formattazione e l'interfaccia utente dei fogli di calcolo.
4. **`utils.gs`**: Include funzioni di utilità per la validazione, il calcolo e la formattazione dei dati.
5. **`main.gs`**: Punto di ingresso per le funzioni principali, come la gestione dei menu e degli eventi di modifica nei fogli.
6. **`workoutBuilder.gs`**: Si occupa della costruzione dinamica delle schede di allenamento, formattando e proteggendo i campi di input.

### Descrizione dei Fogli Google Utilizzati
- **SCHEDA**: Contiene la pianificazione degli allenamenti.
- **ALLENAMENTOODIERNO**: Foglio dedicato all'allenamento giornaliero.
- **STORICO**: Archivia tutti gli allenamenti passati.

## Uso dei Modelli di Linguaggio
Per lo sviluppo di questo progetto, è stato adottato un approccio iterativo utilizzando costantemente modelli di linguaggio avanzati come ChatGPT 4o e Claude 3.5 Sonnet. Questi strumenti hanno supportato:

- **Creazione e ottimizzazione del codice**: generazione di soluzioni rapide e scalabili per problemi complessi.
- **Validazione del design**: analisi dei requisiti e proposta di strutture dati e flussi logici.
- **Miglioramento dell'usabilità**: perfezionamento dell'interfaccia utente e delle funzionalità per ottimizzare l'esperienza mobile.
- **Scrittura di questo README**: ebbene sì!

## Funzionalità Principali
- **Tracciamento Allenamenti**: Registrazione automatizzata degli esercizi, set, ripetizioni e carichi utilizzati.
- **Gestione Dati**: Utilizzo di operazioni batch e cache per migliorare le performance.
- **Protezione dei Dati**: Implementazione di protezioni nei fogli per evitare modifiche non autorizzate.

## Come Iniziare
1. Caricare i file `.gs` nel progetto Apps Script associato a un Google Sheet.
2. Configurare i nomi dei fogli nella sezione `config.gs` in base alle proprie esigenze.
3. Configurare la scheda `STORICO` con le colonne `DATA`, `BLOCCO`, `OBIETTIVO`, `SEDUTA`, `ESERCIZIO`, `SERIE_PREVISTE`, `REP_PREVISTE`, `SERIE_EFFETTIVE`, `REPS_EFFETTIVE`, `CARICO`, `NOTE`.
4. Configurare la scheda `SCHEDA` con le colonne `BLOCCO`, `SETTIMANE`, `OBIETTIVO`, `SEDUTA`, `ESERCIZIO`, `% 1RM`, `SERIE`, `REPS`, `RECUPERO`, `TIPO`, `COMMENTO`, `ATTIVA`, `PROGRESSIONI`; compilare la scheda di allenamento.
5. Eseguire le funzioni del menu custom per accedere alle funzionalità principali, oppure premere sulla checkbox della scheda `ALLENAMENTOODIERNO` per salvare la seduta.

## Contributi
Il progetto è stato sviluppato e perfezionato grazie alla collaborazione tra sviluppatori e strumenti di AI avanzati. Feedback e miglioramenti sono sempre ben accetti!


