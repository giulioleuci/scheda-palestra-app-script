/**
 * Applica la formattazione al foglio ALLENAMENTOODIERNO
 * @param {Sheet} sheet - Foglio da formattare
 */
function formatWorkoutSheet(sheet) {
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 120);

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
      border: true,
      fontSize: 12
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

  const headerRange = sheet.getRange('A1:B1');
  headerRange.merge()
    .setBackground(styles.header.background)
    .setFontColor(styles.header.color)
    .setFontSize(styles.header.fontSize)
    .setFontWeight('bold')
    .setVerticalAlignment(styles.header.vertical)
    .setHorizontalAlignment(styles.header.horizontal);

  return {
    applyExerciseTitleStyle(range) {
      range.setBackground(styles.exerciseTitle.background)
        .setFontColor(styles.exerciseTitle.color)
        .setFontSize(styles.exerciseTitle.fontSize)
        .setFontWeight('bold')
        .setVerticalAlignment(styles.exerciseTitle.vertical);
    },

    applyParameterRowStyle(range) {
      range.setBackground(styles.parameterRow.background)
        .setFontSize(styles.parameterRow.fontSize);
    },

    applyLastSessionStyle(range) {
      range.setBackground(styles.lastSession.background)
        .setFontStyle('italic')
        .setFontSize(styles.lastSession.fontSize);
    },

    applyInputStyle(range) {
      range.setBackground(styles.input.background)
        .setBorder(true, true, true, true, false, false)
        .setFontSize(styles.input.fontSize);
    },

    applySaveButtonStyle(range) {
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

/**
 * Crea un separatore visivo tra esercizi
 */
function addExerciseSeparator(sheet, row) {
  const range = sheet.getRange(row, 1, 1, 2);
  range.setBackground('#f5f5f5')
    .setBorder(true, false, true, false, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
}


/**
 * Protegge il foglio lasciando modificabili solo i campi input
 */
function protectSheet(sheet, inputRanges) {
  const protection = sheet.protect();
  protection.setDescription('Protezione foglio allenamento');
  protection.setUnprotectedRanges(inputRanges);

  const me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}
