/**
 * Inizializza i manager necessari
 * @return {Object} Oggetto contenente le istanze dei manager
 */
function initializeManagers() {
  return {
    workout: new WorkoutManager()
  };
}

/**
 * Inizializza l'applicazione
 */
function onOpen() {
  const managers = initializeManagers();

  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏋️ Allenamento')
    .addItem('Aggiorna Scheda Odierna', 'updateCurrentWorkout')
    .addItem('Salva Allenamento', 'saveWorkout')
    .addItem('Genera scheda specifica', 'generateSpecificWorkoutPrompt')
    .addToUi();
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAMES.ALLENAMENTO_ODIERNO) return;

  const range = e.range;
  const saveButtonRow = 1;//sheet.getLastRow();
  const saveButtonCol = 2;

  if (range.getRow() !== saveButtonRow || range.getColumn() !== saveButtonCol) return;

  const dataValidation = range.getDataValidation();
  if (!dataValidation || dataValidation.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX) return;

  if (e.value !== 'TRUE') return;

  try {
    const builder = new WorkoutBuilder(sheet);
    const data = builder.extractWorkoutData();
    if (data && data.length > 0) {
      const managers = initializeManagers();
      managers.workout.saveWorkoutBatch(data)
      SpreadsheetApp.getUi().alert('Allenamento salvato con successo!');
      updateCurrentWorkout();
    } else {
      throw new Error('Nessun dato da salvare');
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Errore: ' + error.message);
    Logger.log(error);
  } finally {
    range.setValue(false);
  }
}

/**
 * Prompt per selezionare una seduta specifica e generare la scheda
 */
function generateSpecificWorkoutPrompt() {
  const managers = initializeManagers();
  const workoutManager = managers.workout;

  try {
    // Recupera il ciclo delle sedute
    const cycle = workoutManager.getWorkoutCycle();

    // Costruisci il messaggio con l'elenco delle sedute
    const ui = SpreadsheetApp.getUi();
    const message = 'Seleziona la seduta da generare:\n\n' +
      cycle.map((workout, index) => `${index + 1}) ${workout}`).join('\n');

    // Mostra il prompt
    const response = ui.prompt(
      'Selezione Seduta',
      message,
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() === ui.Button.OK) {
      const selectedIndex = parseInt(response.getResponseText().trim()) - 1;

      if (selectedIndex >= 0 && selectedIndex < cycle.length) {
        const selectedWorkout = cycle[selectedIndex];
        generateSpecificWorkout(selectedWorkout); // Passa la seduta selezionata
        ui.alert('Successo', `La scheda per la seduta "${selectedWorkout}" è stata generata con successo.`, ui.ButtonSet.OK);
      } else {
        ui.alert('Errore', 'Indice non valido. Nessuna scheda generata.', ui.ButtonSet.OK);
      }
    }
  } catch (error) {
    ui.alert('Errore', `Errore durante la selezione della seduta: ${error.message}`, ui.ButtonSet.OK);
  }
}



/**
 * Genera una scheda specifica basata sul nome della seduta
 * @param {string} workoutName - Nome della seduta
 */
function generateSpecificWorkout(workoutName) {
  const managers = initializeManagers();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ALLENAMENTO_ODIERNO);

  if (!currentSheet) {
    throw new Error(`Foglio "${CONFIG.SHEET_NAMES.ALLENAMENTO_ODIERNO}" non trovato.`);
  }

  // Recupera esercizi specifici per la seduta
  const exercises = managers.workout.getActiveExercises(workoutName);
  if (!exercises || exercises.length === 0) {
    throw new Error(`Nessun esercizio trovato per la seduta "${workoutName}".`);
  }

  // Genera la scheda
  const builder = new WorkoutBuilder(currentSheet);
  builder.buildWorkoutBatch(exercises);
}



/**
 * Genera una scheda specifica basata sul nome della seduta
 * @param {string} workoutName - Nome della seduta
 */
function generateSpecificWorkout(workoutName) {
  const managers = initializeManagers();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ALLENAMENTO_ODIERNO);

  if (!currentSheet) {
    throw new Error(`Foglio "${CONFIG.SHEET_NAMES.ALLENAMENTO_ODIERNO}" non trovato.`);
  }

  // Recupera esercizi specifici per la seduta
  const exercises = managers.workout.getActiveExercises(workoutName);
  if (!exercises || exercises.length === 0) {
    throw new Error(`Nessun esercizio trovato per la seduta "${workoutName}".`);
  }

  // Genera la scheda
  const builder = new WorkoutBuilder(currentSheet);
  builder.buildWorkoutBatch(exercises);
}




function saveWorkout() {
  const managers = initializeManagers();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ALLENAMENTO_ODIERNO);

  try {
    const builder = new WorkoutBuilder(currentSheet);
    const data = builder.extractWorkoutData();
    if (data && data.length > 0) {
      const managers = initializeManagers();
      managers.workout.saveWorkoutBatch(data)
      SpreadsheetApp.getUi().alert('Allenamento salvato con successo!');
    } else {
      throw new Error('Nessun dato da salvare');
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Errore: ' + error.message);
    Logger.log(error);
  } finally {
    range.setValue(false);
  }
}

/**
 * Aggiorna il foglio ALLENAMENTOODIERNO
 */
function updateCurrentWorkout() {
  const managers = initializeManagers();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ALLENAMENTO_ODIERNO);

  const nextWorkout = managers.workout.getCurrentWorkout();
  const exercises = managers.workout.getActiveExercises(nextWorkout);

  const builder = new WorkoutBuilder(currentSheet);
  builder.buildWorkoutBatch(exercises);
}




/**
 * Aggiorna i grafici nel foglio ANALISI
 */
function updateAnalytics() {
  const managers = initializeManagers();
  managers.analytics.updateAnalytics();
}


