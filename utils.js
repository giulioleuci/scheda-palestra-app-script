/**
 * Utility functions per la gestione dell'app di tracciamento allenamenti
 */

/**
 * Genera un ID univoco per nuovi record
 * @return {string} ID univoco
 */
function generateUniqueId() {
  return Utilities.getUuid();
}

/**
 * Validazione dati esercizio
 */
const ExerciseValidator = {
  /**
   * Valida i parametri di un esercizio
   * @param {Object} exercise - Dati esercizio
   * @return {Object} Risultato validazione {isValid: boolean, errors: string[]}
   */
  validate(exercise) {
    const errors = [];

    if (!this.validateSeries(exercise.serie)) {
      errors.push('Numero serie non valido');
    }

    if (!this.validateReps(exercise.reps)) {
      errors.push('Numero ripetizioni non valido');
    }

    if (!this.validateLoad(exercise.carico)) {
      errors.push('Carico non valido');
    }

    if (!this.validateRest(exercise.recupero)) {
      errors.push('Tempo recupero non valido');
    }

    return {
      isValid: errors.length === 0,
      errors
    };
  },

  validateSeries(series) {
    if (!series) return false;
    const num = parseInt(series);
    return num >= VALIDATION.MIN_SERIE && num <= VALIDATION.MAX_SERIE;
  },

  validateReps(reps) {
    if (!reps) return false;
    const num = parseInt(reps);
    return num >= VALIDATION.MIN_REPS && num <= VALIDATION.MAX_REPS;
  },

  validateLoad(load) {
    if (!load) return false;
    const num = parseFloat(load);
    return num >= VALIDATION.MIN_CARICO && num <= VALIDATION.MAX_CARICO;
  },

  validateRest(rest) {
    if (!rest) return false;
    const num = parseInt(rest);
    return num >= VALIDATION.MIN_RECUPERO && num <= VALIDATION.MAX_RECUPERO;
  }
};

/**
 * Formattazione dati
 */
const Formatter = {
  /**
   * Formatta una data nel formato locale IT
   * @param {Date} date - Data da formattare
   * @return {string} Data formattata
   */
  formatDate(date) {
    return date.toLocaleDateString('it-IT', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric'
    });
  },

  /**
   * Formatta un peso in kg
   * @param {number} weight - Peso da formattare
   * @return {string} Peso formattato
   */
  formatWeight(weight) {
    return `${weight.toFixed(1)}kg`;
  },

  /**
   * Formatta una percentuale
   * @param {number} value - Valore da formattare
   * @return {string} Percentuale formattata
   */
  formatPercent(value) {
    return `${value.toFixed(1)}%`;
  },

  /**
   * Formatta il tempo di recupero
   * @param {number} seconds - Secondi di recupero
   * @return {string} Tempo formattato
   */
  formatRest(seconds) {
    if (seconds < 60) {
      return `${seconds}s`;
    }
    const minutes = Math.floor(seconds / 60);
    const remainingSeconds = seconds % 60;
    return remainingSeconds > 0 ?
      `${minutes}m ${remainingSeconds}s` :
      `${minutes}m`;
  }
};

/**
 * Utility per calcoli statistici
 */
const StatsCalculator = {
  /**
   * Calcola il volume di allenamento
   * @param {number} series - Numero serie
   * @param {number} reps - Ripetizioni
   * @param {number} load - Carico
   * @return {number} Volume totale
   */
  calculateVolume(series, reps, load) {
    return series * reps * load;
  },

  /**
   * Calcola l'intensità relativa
   * @param {number} load - Carico utilizzato
   * @param {number} oneRM - Massimale
   * @return {number} Percentuale intensità
   */
  calculateIntensity(load, oneRM) {
    return (load / oneRM) * 100;
  },

  /**
   * Calcola il tempo sotto tensione
   * @param {string} som - Tempo concentrica-isometrica-eccentrica (es. "2-1-2")
   * @param {number} reps - Ripetizioni
   * @return {number} Tempo totale in secondi
   */
  calculateTimeUnderTension(som, reps) {
    if (!som) return 0;
    const [concentrica, isometrica, eccentrica] = som.split('-').map(Number);
    return (concentrica + isometrica + eccentrica) * reps;
  }
};

/**
 * Utility per manipolazione range settimane
 */
const WeekRangeUtil = {
  /**
   * Converte un range settimane in array
   * @param {string} range - Range settimane (es. "1-6")
   * @return {number[]} Array settimane
   */
  rangeToArray(range) {
    if (!range) return [];
    const [start, end] = range.split('-').map(Number);
    return Array.from(
      {length: end - start + 1},
      (_, i) => start + i
    );
  },

  /**
   * Verifica se una settimana è in un range
   * @param {number} week - Numero settimana
   * @param {string} range - Range settimane
   * @return {boolean} True se la settimana è nel range
   */
  isWeekInRange(week, range) {
    if (!range) return false;
    const weeks = this.rangeToArray(range);
    return weeks.includes(week);
  },

  /**
   * Calcola settimane totali in un range
   * @param {string} range - Range settimane
   * @return {number} Numero totale settimane
   */
  getTotalWeeks(range) {
    if (!range) return 0;
    const [start, end] = range.split('-').map(Number);
    return end - start + 1;
  }
};


/**
 * Converte una stringa in camelCase
 * @param {string} str - Stringa da convertire
 * @return {string} Stringa in camelCase
 */
function toCamelCase(str) {
  return str.toLowerCase()
    .replace(/[^a-zA-Z0-9]+(.)/g, (match, chr) => chr.toUpperCase());
}

/**
 * Legge le intestazioni delle colonne da un foglio e crea un oggetto con i nomi in camelCase
 * @param {Sheet} sheet - Foglio da cui leggere le intestazioni
 * @param {number} [headerRow=1] - Riga delle intestazioni (default: 1)
 * @return {Object} Oggetto con mapping nome originale -> camelCase
 */
function getSheetColumns(sheet, headerRow = 1) {
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const columns = {};

  headers.forEach((header, index) => {
    if (header) { // Consideriamo solo celle non vuote
      columns[header] = {
        name: toCamelCase(header),
        index: index + 1
      };
    }
  });

  return columns;
}

/**
 * Cache per le colonne dei fogli per evitare letture ripetute
 */
const ColumnCache = {
  cache: {},

  /**
   * Ottiene le colonne di un foglio, usando la cache se disponibile
   * @param {Sheet} sheet - Foglio di cui ottenere le colonne
   * @return {Object} Mapping delle colonne
   */
  getColumns(sheet) {
    const sheetName = sheet.getName();
    if (!this.cache[sheetName]) {
      this.cache[sheetName] = getSheetColumns(sheet);
    }
    return this.cache[sheetName];
  },

  /**
   * Invalida la cache per un foglio specifico o per tutti
   * @param {string} [sheetName] - Nome del foglio (se omesso, invalida tutta la cache)
   */
  invalidate(sheetName) {
    if (sheetName) {
      delete this.cache[sheetName];
    } else {
      this.cache = {};
    }
  }
};
