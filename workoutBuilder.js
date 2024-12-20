/**
 * Oggetto che gestisce la costruzione della scheda allenamento
 */
class WorkoutBuilder {
  constructor(sheet) {
    if (!sheet) {
      throw new Error('Foglio non specificato per WorkoutBuilder');
    }
    this.sheet = sheet;
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.ui = this.formatWorkoutSheet();
    this.currentWorkout = null; // Aggiungiamo questa proprietà
  }


  resetSheet() {
    const sheet = this.sheet;
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();

    // Rimuove protezioni
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(protection => protection.remove());

    // Rimuove validazioni
    sheet.getRange(1, 1, maxRows, maxCols).setDataValidation(null);

    // Clear totale con valori default espliciti
    const range = sheet.getRange(1, 1, maxRows, maxCols);
    range.clear({
      contentsOnly: true,
      formatOnly: true,
      validationsOnly: true,
      skipFilteredRows: false
    });

    range.setBackground("white")
         .setFontColor("black")
         .setFontWeight("normal")
         .setFontStyle("normal")
         .setFontSize(10)
         .setHorizontalAlignment('left')
         .setVerticalAlignment('bottom')
         .setBorder(false, false, false, false, false, false);

    // Reset dimensioni colonne
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(2, 210);
  }


  /**
   * Recupera i dati dello storico dalla cache
   * @return {Array} Dati dello storico
   */
  getCachedStoricoData() {
    if (!this.cachedStoricoData) {
      const storicoSheet = this.ss.getSheetByName(CONFIG.SHEET_NAMES.STORICO);
      if (!storicoSheet) {
        throw new Error(`Foglio ${CONFIG.SHEET_NAMES.STORICO} non trovato`);
      }
      this.cachedStoricoData = storicoSheet.getDataRange().getValues();
    }
    return this.cachedStoricoData;
  }




/**
 * Costruisce la scheda allenamento con operazioni batch
 * @param {Array} exercises - Array degli esercizi
 */
buildWorkoutBatch(exercises) {
  if (!exercises || exercises.length === 0) {
    throw new Error('Nessun esercizio fornito per la costruzione della scheda');
  }

  // Reset del foglio
  this.resetSheet();

  // Calcolo dinamico delle righe
  const rowsPerExercise = 6; // Numero di righe per ogni esercizio (titolo, parametri, campi input)
  const separatorRows = exercises.length > 1 ? exercises.length - 1 : 0; // Separatori tra gli esercizi
  const totalRows = 2 + exercises.length * rowsPerExercise + separatorRows; // 2 righe iniziali + righe per esercizi + separatori

  const seduta = exercises[0].seduta;

  // Inizializzazione arrays
  const values = Array.from({ length: totalRows }, () => ['', '']);
  const styles = {
    backgrounds: Array.from({ length: totalRows }, () => ['', '']),
    fontColors: Array.from({ length: totalRows }, () => ['', '']),
    fontWeights: Array.from({ length: totalRows }, () => ['', ''])
  };
  const merges = [];
  let currentRow = 0;

  // Pulsante salva con checkbox
  this.addRowBatch(values, styles, merges, currentRow, {
    value: ['💾 salva allenamento', false],
    background: ['#2e7d32', '#2e7d32'],
    fontColor: ['#ffffff', '#ffffff'],
    fontWeight: ['bold', 'bold']
  });

  // Validazione per la checkbox
  const checkboxRange = this.sheet.getRange(currentRow + 1, 2);
  const checkboxValidation = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();
  checkboxRange.setDataValidation(checkboxValidation);

  currentRow++;

  // Header della seduta
  this.addRowBatch(values, styles, merges, currentRow, {
    value: [`Seduta ${seduta}`, ''],
    background: ['#000000', '#000000'],
    fontColor: ['#ffffff', '#ffffff'],
    fontWeight: ['bold', 'bold'],
    mergeCols: 2
  });
  currentRow++;

  // Costruzione dinamica degli esercizi
  exercises.forEach((exercise, index) => {
    // Titolo esercizio
    this.addRowBatch(values, styles, merges, currentRow, {
      value: [exercise.esercizio, ''],
      background: ['#e8eaf6', '#e8eaf6'],
      fontColor: ['#1a237e', '#1a237e'],
      fontWeight: ['bold', 'bold']
    });
    currentRow++;

    // Parametri previsti
    const paramsText = `Serie: ${exercise.serie} | Rep: ${exercise.reps} | Rec: ${exercise.recupero}`;
    this.addRowBatch(values, styles, merges, currentRow, {
      value: [paramsText, ''],
      background: ['#ffffff', '#ffffff']
    });
    currentRow++;

    // Campi input con logica di recupero storico
    currentRow = this.addInputFieldsWithHistory(values, styles, merges, exercise, currentRow, seduta);

    // Separatore
    if (index < exercises.length - 1) {
      this.addRowBatch(values, styles, merges, currentRow, {
        background: ['#f5f5f5', '#f5f5f5']
      });
      currentRow++;
    }
  });

  // Applicazione batch al foglio
  const range = this.sheet.getRange(1, 1, currentRow, 2);
  range.setValues(values);
  range.setBackgrounds(styles.backgrounds);
  range.setFontColors(styles.fontColors);
  range.setFontWeights(styles.fontWeights);

  // Gestione dei merge
  merges.forEach(merge => {
    this.sheet.getRange(merge.start, 1, 1, merge.cols).merge();
  });

  // Protezione e validazione
  this.protectInputCellsBatch();
  this.hideUnusedRowsAndColumns(currentRow);
}

/**
 * Aggiunge i campi input per un esercizio con logica di recupero storico
 * @param {Array} values - Array dei valori
 * @param {Object} styles - Stili di riga (backgrounds, fontColors, fontWeights)
 * @param {Array} merges - Array dei merge
 * @param {Object} exercise - Dati dell'esercizio
 * @param {number} currentRow - Riga corrente
 * @return {number} Prossima riga disponibile
 */
addInputFieldsWithHistory(values, styles, merges, exercise, currentRow, workoutName) {
  const storicoSheet = this.ss.getSheetByName(CONFIG.SHEET_NAMES.STORICO);
  if (!storicoSheet) {
    throw new Error(`Foglio ${CONFIG.SHEET_NAMES.STORICO} non trovato`);
  }

  const schedaSheet = this.ss.getSheetByName(CONFIG.SHEET_NAMES.SCHEDA);

  const storicoData = this.getCachedStoricoData();
  const columns = ColumnCache.getColumns(storicoSheet);
  const lastSession = this.getLastSessionData(exercise.esercizio, workoutName, storicoData, columns);

  const fields = [
    { label: 'Serie Effettive:', key: 'serie' },
    { label: 'Ripetizioni Effettive:', key: 'reps' },
    { label: 'Carico (kg):', key: 'carico' },
    { label: 'Note:', key: 'note' }
  ];

  fields.forEach(field => {
      this.addRowBatch(values, styles, merges, currentRow, {
          value: [field.label, lastSession && lastSession[field.key] ? lastSession[field.key] : ''],
          background: [
              '#ffffff',
              this.checkRepeatedParameters(exercise.esercizio, workoutName, storicoSheet, schedaSheet)
                  ? this.ui.styles.input.repeatedBackground
                  : this.ui.styles.input.background
          ]
      });
      currentRow++;
  });


  return currentRow;
}




  /**
   * Recupera i dati dell'ultima sessione per un esercizio
   * @param {string} exerciseName - Nome dell'esercizio
   * @param {Array} storicoData - Dati dello storico
   * @param {Object} columns - Colonne dello storico
   * @return {Object|null} Dati ultima sessione o null se non trovata
   */
  getLastSessionData(exerciseName, workoutName, storicoData, columns) {
    for (let i = storicoData.length - 1; i > 0; i--) {
      const row = storicoData[i];
      if (row[columns['ESERCIZIO'].index - 1] === exerciseName &&
          row[columns['SEDUTA'].index - 1] === workoutName) {
        return {
          serie: row[columns['SERIE_EFFETTIVE'].index - 1],
          reps: row[columns['REPS_EFFETTIVE'].index - 1],
          carico: row[columns['CARICO'].index - 1],
          note: row[columns['NOTE'].index - 1] || ''
        };
      }
    }
    return null;
  }


  /**
   * Aggiunge una riga con configurazione batch
   * @param {Array} values - Array dei valori
   * @param {Object} styles - Stili di riga (backgrounds, fontColors, fontWeights)
   * @param {Array} merges - Array dei merge
   * @param {number} rowIndex - Indice della riga
   * @param {Object} options - Configurazioni per la riga
   */
  addRowBatch(values, styles, merges, rowIndex, options) {
    if (options.value) {
      values[rowIndex] = options.value;
    }
    if (options.background) {
      styles.backgrounds[rowIndex] = options.background;
    }
    if (options.fontColor) {
      styles.fontColors[rowIndex] = options.fontColor;
    }
    if (options.fontWeight) {
      styles.fontWeights[rowIndex] = options.fontWeight;
    }
    if (options.mergeCols) {
      merges.push({ start: rowIndex + 1, end: rowIndex + 1, cols: options.mergeCols });
    }
  }

  /**
   * Calcola il numero totale di righe necessarie per costruire la scheda
   * @param {Array} exercises - Array degli esercizi
   * @return {number} Numero totale di righe richieste
   */
  calculateTotalRows(exercises) {
    // 1 riga per il pulsante salva
    // 1 riga per l'header seduta
    // 1 riga vuota dopo l'header
    let total = 3;

    exercises.forEach((_, index) => {
      // 1 riga per il titolo
      // 1 riga per i parametri
      // 4 righe per i campi input
      total += 6;

      // 1 riga per il separatore (se non è l'ultimo esercizio)
      if (index < exercises.length - 1) {
        total += 1;
      }
    });

    return total;
  }

  /**
   * Applica le protezioni alle celle in batch
   */
  protectInputCellsBatch() {
    const ranges = [];
    const sheet = this.sheet;
    const lastRow = sheet.getLastRow();
    let currentRow = 1;

    // Raccogli tutti i range da proteggere
    while (currentRow <= lastRow) {
      const cell = sheet.getRange(currentRow, 2);
      if (cell.getBackground() === this.ui.styles.input.background) {
        ranges.push(cell);
      }
      currentRow++;
    }

    // Applica la protezione in un'unica operazione
    if (ranges.length > 0) {
      const protection = sheet.protect();
      protection.setDescription('Protezione foglio allenamento');
      protection.setUnprotectedRanges(ranges);

      const me = Session.getEffectiveUser();
      protection.addEditor(me);
      protection.removeEditors(protection.getEditors());
      if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
      }
    }
  }


  extractWorkoutData() {
    const data = [];
    const lastRow = this.sheet.getLastRow();
    let currentRow = 4;

    const skip = 8;

    while (currentRow < lastRow) {
      const exerciseName = this.sheet.getRange(currentRow, 1).getValue();

      if (!exerciseName) {
        currentRow += skip;
        continue;
      }

      const serieEffettive = this.sheet.getRange(currentRow + 2, 2).getValue();
      const repsEffettive = this.sheet.getRange(currentRow + 3, 2).getValue();
      const carico = this.sheet.getRange(currentRow + 4, 2).getValue();
      const note = this.sheet.getRange(currentRow + 5, 2).getValue();

      data.push({
        esercizio: exerciseName,
        serieEffettive: serieEffettive,
        repsEffettive: repsEffettive,
        carico: carico,
        note: note || ''
      });

      currentRow += skip;
    }

    return data;
  }

  /**
   * Formatta il foglio di lavoro
   * @return {Object} Oggetto con metodi di formattazione
   */
  formatWorkoutSheet() {
    this.sheet.setFrozenRows(1);
    this.sheet.setColumnWidth(1, 140);
    this.sheet.setColumnWidth(2, 210);

    const styles = {
      header: {
        background: '#1a237e',
        color: '#ffffff',
        fontSize: 14,
        bold: true,
        vertical: 'middle',
        horizontal: 'center'
      },
      exerciseTitle: {
        background: '#e8eaf6',
        color: '#1a237e',
        fontSize: 12,
        bold: true,
        vertical: 'middle'
      },
      parameterRow: {
        background: '#ffffff',
        fontSize: 11
      },
      lastSession: {
        background: '#f5f5f5',
        italic: true,
        fontSize: 11
      },
      input: {
        background: '#e3f2fd',
        repeatedBackground: '#f5b342', // Colore per parametri ripetuti
        border: true,
        fontSize: 10
      },
      saveButton: {
        background: '#2e7d32',
        color: '#ffffff',
        fontSize: 14,
        bold: true,
        vertical: 'middle',
        horizontal: 'center'
      }
    };

    const headerRange = this.sheet.getRange('A1:B1');
    headerRange.merge()
      .setBackground(styles.header.background)
      .setFontColor(styles.header.color)
      .setFontSize(styles.header.fontSize)
      .setFontWeight('bold')
      .setVerticalAlignment(styles.header.vertical)
      .setHorizontalAlignment(styles.header.horizontal);

    return {
      styles,
      applyExerciseTitleStyle: (range) => {
        range.setBackground(styles.exerciseTitle.background)
          .setFontColor(styles.exerciseTitle.color)
          .setFontSize(styles.exerciseTitle.fontSize)
          .setFontWeight('bold')
          .setVerticalAlignment(styles.exerciseTitle.vertical);
      },
      applyParameterRowStyle: (range) => {
        range.setBackground(styles.parameterRow.background)
          .setFontSize(styles.parameterRow.fontSize)
          .setFontWeight('bold');
      },
      applyInputStyle: (range) => {
        range.setBackground(styles.input.background)
          .setBorder(true, true, true, true, false, false)
          .setFontSize(styles.input.fontSize)
          .setHorizontalAlignment('left');
      },
      applySaveButtonStyle: (range) => {
        range.merge()
          .setBackground(styles.saveButton.background)
          .setFontColor(styles.saveButton.color)
          .setFontSize(styles.saveButton.fontSize)
          .setFontWeight('bold')
          .setVerticalAlignment(styles.saveButton.vertical)
          .setHorizontalAlignment(styles.saveButton.horizontal);
      }
    };
  }

  checkRepeatedParameters(exerciseName, currentWorkout, storicoSheet, schedaSheet) {
    if (!storicoSheet) return false;

    const schedaColumns = ColumnCache.getColumns(schedaSheet);
    const schedaData = schedaSheet.getDataRange().getValues();

    // Verifica se PROGRESSIONI è TRUE per l'esercizio nella scheda
    const exerciseRow = schedaData.find(row =>
        row[schedaColumns['ESERCIZIO'].index - 1] === exerciseName &&
        row[schedaColumns['SEDUTA'].index - 1] === currentWorkout
    );

    if (!exerciseRow || exerciseRow[schedaColumns['PROGRESSIONI'].index - 1] !== true) {
        return false; // Se PROGRESSIONI non è TRUE, non applicare lo stile
    }

    // Otteniamo i riferimenti alle colonne usando il nuovo sistema
    const columns = ColumnCache.getColumns(storicoSheet);

    // Verifichiamo che tutte le colonne necessarie esistano
    const requiredColumns = ['ESERCIZIO', 'SEDUTA', 'SERIE_EFFETTIVE', 'REPS_EFFETTIVE', 'CARICO'];
    const missingColumns = requiredColumns.filter(col => !columns[col] || !columns[col].index);
    if (missingColumns.length > 0) {
      Logger.log(`Attenzione: colonne mancanti nel foglio STORICO: ${missingColumns.join(', ')}`);
      return false;
    }

    const data = storicoSheet.getDataRange().getValues();

    // Filtriamo le ultime due sessioni per l'esercizio specificato
    const lastTwoSessions = data
      .filter(row =>
        row[columns['ESERCIZIO'].index - 1] === exerciseName &&
        row[columns['SEDUTA'].index - 1] === currentWorkout
      )
      .slice(-2);

    // Se non abbiamo due sessioni da confrontare, ritorniamo false
    if (lastTwoSessions.length < 2) return false;

    // Confrontiamo i parametri delle due sessioni
    return lastTwoSessions[0][columns['SERIE_EFFETTIVE'].index - 1] === lastTwoSessions[1][columns['SERIE_EFFETTIVE'].index - 1] &&
          lastTwoSessions[0][columns['REPS_EFFETTIVE'].index - 1] === lastTwoSessions[1][columns['REPS_EFFETTIVE'].index - 1] &&
          lastTwoSessions[0][columns['CARICO'].index - 1] === lastTwoSessions[1][columns['CARICO'].index - 1];
  }

  hideUnusedRowsAndColumns(lastUsedRow) {
    // Mostra tutte le righe e colonne prima di nascondere quelle inutilizzate
    const maxRows = this.sheet.getMaxRows();
    const maxCols = this.sheet.getMaxColumns();

    // Mostra tutte le righe e colonne
    this.sheet.showRows(1, maxRows);
    this.sheet.showColumns(1, maxCols);

    // Nascondi righe inutilizzate
    if (lastUsedRow < maxRows) {
      const unusedRows = maxRows - lastUsedRow;
      if (unusedRows > 0) {
        this.sheet.hideRows(lastUsedRow + 1, unusedRows);
      }
    }

    // Nascondi colonne inutilizzate (dopo la colonna B)
    if (maxCols > 2) {
      const unusedCols = maxCols - 2;
      if (unusedCols > 0) {
        this.sheet.hideColumns(3, unusedCols);
      }
    }
  }






  /**
   * Aggiunge un separatore tra esercizi
   * @param {number} row - Riga dove aggiungere il separatore
   */
  addExerciseSeparator(row) {
    const range = this.sheet.getRange(row, 1, 1, 2);
    range.setBackground('#f5f5f5')
      .setBorder(true, false, true, false, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  }





  protectInputCells() {
    const lastRow = this.sheet.getLastRow();
    const inputRanges = [];

    let currentRow = 1;
    while (currentRow < lastRow) {
      const cell = this.sheet.getRange(currentRow, 2);
      if (cell.getBackground() === this.ui.styles.input.background) {
        inputRanges.push(cell);
      }
      currentRow++;
    }

    if (inputRanges.length > 0) {
      const protection = this.sheet.protect();
      protection.setDescription('Protezione foglio allenamento');
      protection.setUnprotectedRanges(inputRanges);

      const me = Session.getEffectiveUser();
      protection.addEditor(me);
      protection.removeEditors(protection.getEditors());
      if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
      }
    }
  }
}

/**
 * Funzione principale per costruire la scheda allenamento
 * @param {Sheet} sheet - Foglio ALLENAMENTOODIERNO
 * @param {Array} exercises - Array degli esercizi
 */
function buildCurrentWorkout(sheet, exercises) {
  if (!sheet || !exercises) {
    throw new Error('Parametri mancanti per buildCurrentWorkout');
  }
  const builder = new WorkoutBuilder(sheet);
  builder.buildWorkout(exercises);
}
