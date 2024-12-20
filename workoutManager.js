class WorkoutManager {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.cache = CacheService.getScriptCache();
  }


  /**
   * Salva l'allenamento utilizzando operazioni batch
   * @param {Array} workoutData - Dati dell'allenamento da salvare
   */
  saveWorkoutBatch(workoutData) {
    const storicoSheet = this.ss.getSheetByName(CONFIG.SHEET_NAMES.STORICO);
    if (!storicoSheet) {
      throw new Error(`Foglio ${CONFIG.SHEET_NAMES.STORICO} non trovato`);
    }

    const columns = ColumnCache.getColumns(storicoSheet);
    const currentDate = new Date();
    const currentInfo = this.getCurrentTrainingInfo();
    const seduta = this.getCurrentWorkout();
    const cachedData = this.getCachedSheetData(storicoSheet); // Usa cache
    const lastRow = cachedData.length; // Calcola ultima riga dalla cache

    // Prepara i dati per l'operazione batch
    const rowsToAppend = workoutData.map(exercise => {
      const schedaData = this.getExerciseScheduleData(exercise.esercizio);
      return [
        currentDate,
        currentInfo.blocco,
        currentInfo.obiettivo,
        seduta,
        exercise.esercizio,
        schedaData ? schedaData.serie : '',
        schedaData ? schedaData.reps : '',
        exercise.serieEffettive,
        exercise.repsEffettive,
        exercise.carico,
        exercise.note || ''
      ];
    });

    // Esegui l'operazione batch per lo storico
    const targetRange = storicoSheet.getRange(
      lastRow + 1,
      1,
      rowsToAppend.length,
      rowsToAppend[0].length
    );

    targetRange.setValues(rowsToAppend);

    // Aggiorna la cache
    this.cache.put('currentWorkout', seduta);

    // Aggiorna la formattazione in batch
    const formatRange = targetRange.setNumberFormat('@');
    const dateRange = storicoSheet.getRange(
      storicoSheet.getLastRow() - rowsToAppend.length + 1,
      1,
      rowsToAppend.length,
      1
    );
    dateRange.setNumberFormat('dd/mm/yyyy');
  }

getCurrentTrainingInfo() {
    const schedaSheet = this.ss.getSheetByName(CONFIG.SHEET_NAMES.SCHEDA);
    if (!schedaSheet) {
      throw new Error(`Foglio ${CONFIG.SHEET_NAMES.SCHEDA} non trovato.`);
    }

    const columns = ColumnCache.getColumns(schedaSheet);
    const data = this.getCachedSheetData(schedaSheet); // Utilizza cache temporanea

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[columns['ATTIVA'].index - 1] === true) {
        return {
          blocco: row[columns['BLOCCO'].index - 1],
          obiettivo: row[columns['OBIETTIVO'].index - 1]
        };
      }
    }
    throw new Error('Nessun allenamento attivo trovato nel foglio SCHEDA');
  }



  getCachedSheetData(sheet) {
    if (!this.cachedData) {
      this.cachedData = sheet.getDataRange().getValues();
    }
    return this.cachedData;
  }


  /**
   * Ottiene la prossima seduta nel cycle
   * @param {string} currentWorkout - Seduta corrente
   * @return {string} Prossima seduta
   */
  getNextWorkoutFromCurrent(currentWorkout) {
    const cycle = this.getWorkoutCycle();
    const currentIndex = cycle.indexOf(currentWorkout);

    if (currentIndex === -1) {
      // Se la seduta corrente non è nel cycle, ritorna la prima
      return cycle[0];
    }

    // Passa alla prossima seduta ciclicamente
    return cycle[(currentIndex + 1) % cycle.length];
  }


  /**
   * Ottiene il cycle delle sedute dalle righe attive della SCHEDA
   * @return {Array<string>} Array ordinato delle sedute uniche
   */
  getWorkoutCycle() {
    const schedaSheet = this.ss.getSheetByName(CONFIG.SHEET_NAMES.SCHEDA);
    if (!schedaSheet) {
      throw new Error(`Foglio ${CONFIG.SHEET_NAMES.SCHEDA} non trovato`);
    }

    const columns = ColumnCache.getColumns(schedaSheet);
    const currentInfo = this.getCurrentTrainingInfo();
    const data = schedaSheet.getDataRange().getValues();

    // Filtra le righe attive e estrae le sedute uniche
    const uniqueWorkouts = [...new Set(
      data
        .slice(1) // Salta l'header
        .filter(row =>
          row[columns['ATTIVA'].index - 1] === true &&
          row[columns['BLOCCO'].index - 1] === currentInfo.blocco &&
          row[columns['OBIETTIVO'].index - 1] === currentInfo.obiettivo &&
          row[columns['SEDUTA'].index - 1] // Verifica che la seduta non sia vuota
        )
        .map(row => row[columns['SEDUTA'].index - 1])
    )].sort();

    if (uniqueWorkouts.length === 0) {
      throw new Error('Nessuna seduta attiva trovata nella SCHEDA');
    }

    return uniqueWorkouts;
  }

  /**
   * Ottiene la seduta corrente
   * @return {string} Seduta corrente/successiva
   */
  getCurrentWorkout() {
    const cycle = this.getWorkoutCycle();
    const currentWorkout = this.cache.get('currentWorkout');

    if (!currentWorkout) {
      // Se non c'è cache, chiedi all'utente
      return this.promptWorkoutSelection(cycle);
    }

    // Altrimenti passa alla successiva nel cycle
    return this.getNextWorkoutFromCurrent(currentWorkout);
  }

  /**
   * Mostra un prompt per la selezione della seduta
   * @param {Array} availableWorkouts - Array delle sedute disponibili
   * @return {string} Seduta selezionata
   */
  promptWorkoutSelection(availableWorkouts) {
    const ui = SpreadsheetApp.getUi();

    const message = 'Seleziona la seduta da generare:\n\n' +
      availableWorkouts.map((workout, index) =>
        `${index + 1}) Seduta ${workout}`
      ).join('\n');

    const response = ui.prompt(
      'Selezione Seduta',
      message,
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() === ui.Button.OK) {
      const selectedIndex = parseInt(response.getResponseText()) - 1;

      if (selectedIndex >= 0 && selectedIndex < availableWorkouts.length) {
        const selectedWorkout = availableWorkouts[selectedIndex];
        this.cache.put('currentWorkout', selectedWorkout);
        return selectedWorkout;
      }
    }

    // Default alla prima seduta disponibile
    const defaultWorkout = availableWorkouts[0];
    this.cache.put('currentWorkout', defaultWorkout);

    ui.alert(
      'Selezione Default',
      `Input non valido. Verrà generata la Seduta ${defaultWorkout}`,
      ui.ButtonSet.OK
    );

    return defaultWorkout;
  }

  /**
   * Recupera gli esercizi attivi per una seduta specifica
   * @param {string} workoutName - Nome della seduta
   * @return {Array} Array di esercizi attivi
   */
  getActiveExercises(workoutName) {
    const schedaSheet = this.ss.getSheetByName(CONFIG.SHEET_NAMES.SCHEDA);
    if (!schedaSheet) {
      throw new Error(`Foglio "${CONFIG.SHEET_NAMES.SCHEDA}" non trovato`);
    }

    const columns = ColumnCache.getColumns(schedaSheet);
    const currentInfo = this.getCurrentTrainingInfo();
    const data = schedaSheet.getDataRange().getValues();

    return data
      .slice(1) // Salta l'intestazione
      .filter(row =>
        row[columns['SEDUTA'].index - 1] === workoutName &&
        row[columns['BLOCCO'].index - 1] === currentInfo.blocco &&
        row[columns['ATTIVA'].index - 1] === true
      )
      .map(row => ({
        seduta: row[columns['SEDUTA'].index - 1],
        esercizio: row[columns['ESERCIZIO'].index - 1],
        serie: row[columns['SERIE'].index - 1],
        reps: row[columns['REPS'].index - 1],
        recupero: row[columns['RECUPERO'].index - 1],
        tipo: row[columns['TIPO'].index - 1],
        percentuale: row[columns['% 1RM'].index - 1],
        progressioni: row[columns['PROGRESSIONI'].index - 1]
      }));
  }



  getExerciseScheduleData(exerciseName) {
    const schedaSheet = this.ss.getSheetByName(CONFIG.SHEET_NAMES.SCHEDA);
    if (!schedaSheet) return null;

    const columns = ColumnCache.getColumns(schedaSheet);
    const currentInfo = this.getCurrentTrainingInfo();
    const data = schedaSheet.getDataRange().getValues();

    const exerciseRow = data.find(row =>
      row[columns['ESERCIZIO'].index - 1] === exerciseName &&
      row[columns['BLOCCO'].index - 1] === currentInfo.blocco &&
      row[columns['OBIETTIVO'].index - 1] === currentInfo.obiettivo &&
      row[columns['ATTIVA'].index - 1] === true
    );

    if (!exerciseRow) return null;

    return {
      serie: exerciseRow[columns['SERIE'].index - 1],
      reps: exerciseRow[columns['REPS'].index - 1]
    };
  }
}
