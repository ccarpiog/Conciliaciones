/**
 * Bank Reconciliation System for Google Sheets
 * Reconciles accounting movements with bank statements
 */

// Configuration constants
const CONFIG = {
  SOURCE_SHEET: 'Origen',
  OUTPUT_SHEET: 'Salida',
  DATE_TOLERANCE_DAYS: 3, // Days of difference allowed between dates
  MIN_SIMILARITY_SCORE: 0.3, // Minimum similarity score for concept matching

  // Column mappings for accounting data (A-D)
  ACCOUNTING: {
    DATE_COL: 0,      // Column A
    ENTRY_COL: 1,     // Column B
    CONCEPT_COL: 2,   // Column C
    AMOUNT_COL: 3     // Column D
  },

  // Column mappings for bank data (F-J)
  BANK: {
    DATE_COL: 5,      // Column F
    VALUE_DATE_COL: 6,// Column G
    CONCEPT_COL: 7,   // Column H
    ADDITIONAL_COL: 8,// Column I
    AMOUNT_COL: 9     // Column J
  }
};

/**
 * Creates custom menu on spreadsheet open
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Conciliación')
    .addItem('Ejecutar conciliación automática', 'runReconciliation')
    .addItem('Revisar conflictos', 'showConflictsSidebar')
    .addItem('Configurar parámetros', 'showConfigDialog')
    .addSeparator()
    .addItem('Limpiar hoja de salida', 'clearOutputSheet')
    .addItem('Borrar conciliaciones manuales', 'clearManualMatchesWithConfirm')
    .addToUi();
}

/**
 * Main reconciliation function
 */
function runReconciliation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  const outputSheet = ss.getSheetByName(CONFIG.OUTPUT_SHEET);

  if (!sourceSheet || !outputSheet) {
    SpreadsheetApp.getUi().alert('Error: No se encontraron las hojas "Origen" o "Salida"');
    return;
  }

  // Load user configuration
  const userProperties = PropertiesService.getUserProperties();
  const dateToleranceDays = Number(userProperties.getProperty('dateToleranceDays')) || CONFIG.DATE_TOLERANCE_DAYS;
  const minSimilarityScore = Number(userProperties.getProperty('minSimilarityScore')) || CONFIG.MIN_SIMILARITY_SCORE;

  // Update CONFIG with user settings
  CONFIG.DATE_TOLERANCE_DAYS = dateToleranceDays;
  CONFIG.MIN_SIMILARITY_SCORE = minSimilarityScore;

  // Get data from source sheet
  const accountingData = getAccountingData(sourceSheet);
  const bankData = getBankData(sourceSheet);

  // Perform reconciliation
  const reconciliationResults = reconcileMovements(accountingData, bankData);

  // Output results
  outputReconciliationResults(outputSheet, reconciliationResults);

  // Show summary
  showReconciliationSummary(reconciliationResults);
}

/**
 * Gets accounting data from columns A-D
 */
function getAccountingData(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const range = sheet.getRange(2, 1, lastRow - 1, 4); // Skip header row
  const values = range.getValues();

  return values
    .filter(row => row[CONFIG.ACCOUNTING.DATE_COL] && row[CONFIG.ACCOUNTING.AMOUNT_COL])
    .map((row, index) => {
      // Create stable ID based on date, entry number, and amount
      const date = new Date(row[CONFIG.ACCOUNTING.DATE_COL]);
      const dateStr = date.getTime();
      const entry = String(row[CONFIG.ACCOUNTING.ENTRY_COL] || '');
      const amount = Number(row[CONFIG.ACCOUNTING.AMOUNT_COL]);
      const id = `ACC_${dateStr}_${entry}_${amount}`;

      return {
        id,
        date,
        entryNumber: row[CONFIG.ACCOUNTING.ENTRY_COL],
        concept: String(row[CONFIG.ACCOUNTING.CONCEPT_COL] || ''),
        amount,
        rowNumber: index + 2,
        matched: false,
        bankMatches: []
      };
    });
}

/**
 * Gets bank data from columns F-J
 */
function getBankData(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const range = sheet.getRange(2, 6, lastRow - 1, 5); // Columns F-J, skip header
  const values = range.getValues();

  return values
    .filter(row => row[0] && row[4]) // Check date and amount exist
    .map((row, index) => {
      // Create stable ID based on date, concept, and amount
      const date = new Date(row[0]);
      const dateStr = date.getTime();
      const concept = String(row[2] || '');
      const amount = Number(row[4]);
      // Include first 20 chars of concept to differentiate same-day same-amount transactions
      const conceptKey = concept.substring(0, 20).replace(/[^a-zA-Z0-9]/g, '');
      const id = `BANK_${dateStr}_${conceptKey}_${amount}`;

      return {
        id,
        date,
        valueDate: row[1] ? new Date(row[1]) : null, // Value date (G)
        concept, // Concept (H)
        additional: String(row[3] || ''), // Additional data (I)
        amount,
        rowNumber: index + 2,
        matched: false
      };
    });
}

/**
 * Main reconciliation logic
 */
function reconcileMovements(accountingData, bankData) {
  const results = {
    matched: [],
    conflicts: [],
    unmatchedAccounting: [],
    unmatchedBank: [...bankData] // Start with all bank movements as unmatched
  };

  // Get manual matches
  const manualMatches = getManualMatches();

  // Process each accounting movement
  accountingData.forEach(accMovement => {
    // Check if there's a manual match for this accounting movement
    if (manualMatches[accMovement.id]) {
      const manualBankId = manualMatches[accMovement.id];
      const manualMatch = bankData.find(b => b.id === manualBankId);

      if (manualMatch && !manualMatch.matched) {
        // Apply manual match
        results.matched.push({
          accounting: accMovement,
          bank: manualMatch,
          score: 1.0, // Manual matches have perfect score
          autoMatched: false,
          manualMatch: true
        });

        // Mark as matched
        accMovement.matched = true;
        manualMatch.matched = true;

        // Remove from unmatched bank list
        const index = results.unmatchedBank.findIndex(b => b.id === manualMatch.id);
        if (index > -1) results.unmatchedBank.splice(index, 1);

        return; // Skip automatic matching
      }
    }

    // Find potential matches in bank data (same amount is mandatory)
    // Round to 2 decimals for exact comparison (handling floating-point precision)
    const accAmount = Math.round(accMovement.amount * 100) / 100;
    const potentialMatches = bankData.filter(bankMovement => {
      const bankAmount = Math.round(bankMovement.amount * 100) / 100;
      return !bankMovement.matched && accAmount === bankAmount;
    });

    if (potentialMatches.length === 0) {
      // No matches found
      results.unmatchedAccounting.push(accMovement);
      return;
    }

    // Score each potential match
    const scoredMatches = potentialMatches.map(bankMovement => ({
      bankMovement,
      score: calculateMatchScore(accMovement, bankMovement)
    }));

    // Sort by score (highest first)
    scoredMatches.sort((a, b) => b.score - a.score);

    if (scoredMatches.length === 1 ||
        (scoredMatches[0].score > CONFIG.MIN_SIMILARITY_SCORE &&
         scoredMatches[0].score - (scoredMatches[1]?.score || 0) > 0.2)) {
      // Clear winner - automatic match
      const bestMatch = scoredMatches[0].bankMovement;
      results.matched.push({
        accounting: accMovement,
        bank: bestMatch,
        score: scoredMatches[0].score,
        autoMatched: true
      });

      // Mark as matched
      accMovement.matched = true;
      bestMatch.matched = true;

      // Remove from unmatched bank list
      const index = results.unmatchedBank.findIndex(b => b.id === bestMatch.id);
      if (index > -1) results.unmatchedBank.splice(index, 1);
    } else {
      // Multiple candidates or low confidence - needs manual review
      results.conflicts.push({
        accounting: accMovement,
        candidates: scoredMatches.slice(0, 5), // Keep top 5 candidates
        reason: scoredMatches.length > 1 ? 'Múltiples candidatos' : 'Baja confianza'
      });
    }
  });

  return results;
}

/**
 * Calculates match score between accounting and bank movements
 */
function calculateMatchScore(accMovement, bankMovement) {
  let score = 0;
  const weights = {
    dateMatch: 0.3,
    conceptSimilarity: 0.7
  };

  // Date scoring
  const dateDiff = Math.abs(accMovement.date - bankMovement.date) / (1000 * 60 * 60 * 24);
  if (dateDiff === 0) {
    score += weights.dateMatch;
  } else if (dateDiff <= CONFIG.DATE_TOLERANCE_DAYS) {
    score += weights.dateMatch * (1 - dateDiff / (CONFIG.DATE_TOLERANCE_DAYS * 2));
  }

  // Concept similarity scoring
  const conceptSimilarity = calculateConceptSimilarity(
    accMovement.concept,
    bankMovement.concept + ' ' + bankMovement.additional
  );
  score += weights.conceptSimilarity * conceptSimilarity;

  return score;
}

/**
 * Calculates similarity between two concept strings
 */
function calculateConceptSimilarity(concept1, concept2) {
  if (!concept1 || !concept2) return 0;

  // Normalize strings
  const norm1 = normalizeString(concept1);
  const norm2 = normalizeString(concept2);

  // Exact match
  if (norm1 === norm2) return 1;

  // Check if one contains the other
  if (norm1.includes(norm2) || norm2.includes(norm1)) {
    return 0.8;
  }

  // Extract numbers and check if they match (including partial matches)
  const numbers1 = extractNumbers(concept1);
  const numbers2 = extractNumbers(concept2);

  if (numbers1.length > 0 && numbers2.length > 0) {
    // First check for exact matches
    const commonNumbers = numbers1.filter(n => numbers2.includes(n));
    if (commonNumbers.length > 0) {
      return 0.6 + (0.3 * commonNumbers.length / Math.max(numbers1.length, numbers2.length));
    }

    // Check for partial matches (e.g., "661112" contains "1112")
    for (const num1 of numbers1) {
      for (const num2 of numbers2) {
        // Check if one number is a substring of the other (minimum 3 digits)
        if (num1.length >= 3 && num2.length >= 3) {
          if (num1.includes(num2) || num2.includes(num1)) {
            return 0.65;
          }
          // Check for trailing match (e.g., "661112" ends with "1112")
          if (num1.length >= 4 && num2.length >= 3) {
            if (num1.endsWith(num2) || num2.endsWith(num1)) {
              return 0.7;
            }
          }
        }
      }
    }
  }

  // Token-based similarity
  const tokens1 = norm1.split(/\s+/);
  const tokens2 = norm2.split(/\s+/);
  const commonTokens = tokens1.filter(t => tokens2.includes(t));

  if (commonTokens.length > 0) {
    return 0.3 + (0.4 * commonTokens.length / Math.max(tokens1.length, tokens2.length));
  }

  // Levenshtein distance for short strings
  if (norm1.length < 20 && norm2.length < 20) {
    const distance = levenshteinDistance(norm1, norm2);
    const maxLength = Math.max(norm1.length, norm2.length);
    return Math.max(0, 1 - distance / maxLength) * 0.5;
  }

  return 0;
}

/**
 * Normalizes a string for comparison
 */
function normalizeString(str) {
  return str.toLowerCase()
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Extracts numbers from a string
 */
function extractNumbers(str) {
  const matches = str.match(/\d+/g);
  return matches ? matches.map(n => n.replace(/^0+/, '')) : [];
}

/**
 * Calculates Levenshtein distance between two strings
 */
function levenshteinDistance(str1, str2) {
  const matrix = [];

  for (let i = 0; i <= str2.length; i++) {
    matrix[i] = [i];
  }

  for (let j = 0; j <= str1.length; j++) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= str2.length; i++) {
    for (let j = 1; j <= str1.length; j++) {
      if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // substitution
          matrix[i][j - 1] + 1,     // insertion
          matrix[i - 1][j] + 1      // deletion
        );
      }
    }
  }

  return matrix[str2.length][str1.length];
}

/**
 * Outputs reconciliation results to the output sheet
 */
function outputReconciliationResults(sheet, results) {
  // Clear existing content
  sheet.clear();

  // Set up headers - write separately to avoid dimension mismatch
  const numColumns = 10;

  // Title row
  sheet.getRange(1, 1, 1, numColumns).merge()
    .setValue('MOVIMIENTOS CONCILIADOS')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Column headers
  sheet.getRange(2, 1, 1, numColumns).setValues([[
    'Fecha Cont.', 'Asiento', 'Concepto Cont.', 'Importe', 'Estado',
    'Fecha Banco', 'Fecha Valor', 'Concepto Banco', 'Datos Adic.', 'Puntuación'
  ]])
    .setBackground('#c9daf8')
    .setFontWeight('bold');

  let currentRow = 4;

  // Combine and sort all movements
  const allMovements = [];

  // Add matched movements
  results.matched.forEach(match => {
    const status = match.manualMatch ? '✓ Manual' : '✓ Conciliado';
    const scoreDisplay = match.manualMatch ? 'Manual' : Math.round(match.score * 100) + '%';

    allMovements.push({
      date: match.accounting.date,
      entryNumber: match.accounting.entryNumber,
      row: [
        formatDate(match.accounting.date),
        match.accounting.entryNumber,
        match.accounting.concept,
        match.accounting.amount,
        status,
        formatDate(match.bank.date),
        match.bank.valueDate ? formatDate(match.bank.valueDate) : '',
        match.bank.concept,
        match.bank.additional,
        scoreDisplay
      ],
      type: match.manualMatch ? 'manual' : 'matched'
    });
  });

  // Add conflicts
  results.conflicts.forEach(conflict => {
    allMovements.push({
      date: conflict.accounting.date,
      entryNumber: conflict.accounting.entryNumber,
      row: [
        formatDate(conflict.accounting.date),
        conflict.accounting.entryNumber,
        conflict.accounting.concept,
        conflict.accounting.amount,
        '⚠ ' + conflict.reason,
        conflict.candidates.length + ' candidatos',
        '',
        conflict.candidates[0].bankMovement.concept,
        conflict.candidates[0].bankMovement.additional,
        Math.round(conflict.candidates[0].score * 100) + '%'
      ],
      type: 'conflict'
    });
  });

  // Add unmatched accounting movements
  results.unmatchedAccounting.forEach(movement => {
    allMovements.push({
      date: movement.date,
      entryNumber: movement.entryNumber,
      row: [
        formatDate(movement.date),
        movement.entryNumber,
        movement.concept,
        movement.amount,
        '✗ Sin conciliar',
        '',
        '',
        '',
        '',
        ''
      ],
      type: 'unmatched'
    });
  });

  // Sort by date and entry number
  allMovements.sort((a, b) => {
    const dateDiff = a.date - b.date;
    if (dateDiff !== 0) return dateDiff;

    // Handle entry number sorting (can be text or number)
    const entryA = a.entryNumber;
    const entryB = b.entryNumber;

    // Try to parse as numbers first
    const numA = Number(entryA);
    const numB = Number(entryB);

    if (!isNaN(numA) && !isNaN(numB)) {
      return numA - numB;
    }

    // Fall back to string comparison
    return String(entryA || '').localeCompare(String(entryB || ''));
  });

  // Output sorted movements
  allMovements.forEach(movement => {
    sheet.getRange(currentRow, 1, 1, movement.row.length).setValues([movement.row]);

    // Apply formatting based on type
    const rowRange = sheet.getRange(currentRow, 1, 1, movement.row.length);

    switch(movement.type) {
      case 'matched':
        rowRange.setBackground('#d9ead3');
        break;
      case 'manual':
        rowRange.setBackground('#b7e1cd'); // Slightly darker green for manual matches
        break;
      case 'conflict':
        rowRange.setBackground('#fff2cc');
        break;
      case 'unmatched':
        rowRange.setBackground('#f4cccc');
        break;
    }

    currentRow++;
  });

  // Add summary section for unmatched bank movements
  if (results.unmatchedBank.length > 0) {
    currentRow += 2;
    sheet.getRange(currentRow, 1).setValue('MOVIMIENTOS BANCARIOS NO CONCILIADOS')
      .setBackground('#ea4335')
      .setFontColor('#ffffff')
      .setFontWeight('bold');

    sheet.getRange(currentRow, 1, 1, 5).merge();
    currentRow++;

    sheet.getRange(currentRow, 1, 1, 5).setValues([
      ['Fecha Mov.', 'Fecha Valor', 'Concepto', 'Datos Adic.', 'Importe']
    ]).setBackground('#f4cccc').setFontWeight('bold');
    currentRow++;

    results.unmatchedBank.forEach(bank => {
      sheet.getRange(currentRow, 1, 1, 5).setValues([[
        formatDate(bank.date),
        bank.valueDate ? formatDate(bank.valueDate) : '',
        bank.concept,
        bank.additional,
        bank.amount
      ]]);
      currentRow++;
    });
  }

  // Auto-resize columns
  sheet.autoResizeColumns(1, 10);
}

/**
 * Formats a date for display
 */
function formatDate(date) {
  if (!date) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

/**
 * Shows reconciliation summary
 */
function showReconciliationSummary(results) {
  const ui = SpreadsheetApp.getUi();

  const summary = `Resumen de Conciliación:

  ✓ Movimientos conciliados: ${results.matched.length}
  ⚠ Conflictos (requieren revisión): ${results.conflicts.length}
  ✗ Movimientos contables sin conciliar: ${results.unmatchedAccounting.length}
  ✗ Movimientos bancarios sin conciliar: ${results.unmatchedBank.length}

  Total procesado: ${results.matched.length + results.conflicts.length + results.unmatchedAccounting.length}`;

  ui.alert('Conciliación Completada', summary, ui.ButtonSet.OK);
}

/**
 * Shows conflicts sidebar for manual resolution
 */
function showConflictsSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ConflictsSidebar')
    .setTitle('Resolver Conflictos')
    .setWidth(400);

  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Gets conflicts data for sidebar
 */
function getConflictsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);

  const accountingData = getAccountingData(sourceSheet);
  const bankData = getBankData(sourceSheet);
  const results = reconcileMovements(accountingData, bankData);

  return results.conflicts.map(conflict => ({
    accounting: {
      id: conflict.accounting.id,
      date: formatDate(conflict.accounting.date),
      entry: conflict.accounting.entryNumber,
      concept: conflict.accounting.concept,
      amount: conflict.accounting.amount
    },
    candidates: conflict.candidates.map(c => ({
      id: c.bankMovement.id,
      date: formatDate(c.bankMovement.date),
      concept: c.bankMovement.concept,
      additional: c.bankMovement.additional,
      amount: c.bankMovement.amount,
      score: Math.round(c.score * 100)
    }))
  }));
}

/**
 * Resolves a conflict by manually matching movements
 */
function resolveConflict(accountingId, bankId) {
  // Store the manual match in document properties (shared across all users)
  const documentProperties = PropertiesService.getDocumentProperties();
  const manualMatches = JSON.parse(documentProperties.getProperty('manualMatches') || '{}');

  // Store the match (accounting ID -> bank ID)
  manualMatches[accountingId] = bankId;

  // Save back to properties
  documentProperties.setProperty('manualMatches', JSON.stringify(manualMatches));

  SpreadsheetApp.getActiveSpreadsheet().toast('Conflicto resuelto', 'Éxito', 3);
  return true;
}

/**
 * Gets all manual matches
 */
function getManualMatches() {
  const documentProperties = PropertiesService.getDocumentProperties();
  return JSON.parse(documentProperties.getProperty('manualMatches') || '{}');
}

/**
 * Clears all manual matches
 */
function clearManualMatches() {
  const documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.deleteProperty('manualMatches');
}

/**
 * Clears manual matches with user confirmation
 */
function clearManualMatchesWithConfirm() {
  const ui = SpreadsheetApp.getUi();
  const manualMatches = getManualMatches();
  const matchCount = Object.keys(manualMatches).length;

  if (matchCount === 0) {
    ui.alert('No hay conciliaciones manuales', 'No hay conciliaciones manuales guardadas para borrar.', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    'Borrar conciliaciones manuales',
    `Hay ${matchCount} ${matchCount === 1 ? 'conciliación manual' : 'conciliaciones manuales'} guardadas.\n\n¿Desea borrarlas? Esta acción no se puede deshacer.`,
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    clearManualMatches();
    ui.alert('Conciliaciones borradas', 'Las conciliaciones manuales han sido borradas. Ejecute la conciliación de nuevo para ver los cambios.', ui.ButtonSet.OK);
  }
}

/**
 * Shows configuration dialog
 */
function showConfigDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ConfigDialog')
    .setWidth(400)
    .setHeight(300);

  SpreadsheetApp.getUi().showModalDialog(html, 'Configuración de Conciliación');
}

/**
 * Gets current configuration
 */
function getConfig() {
  const userProperties = PropertiesService.getUserProperties();
  return {
    dateToleranceDays: userProperties.getProperty('dateToleranceDays') || CONFIG.DATE_TOLERANCE_DAYS,
    minSimilarityScore: userProperties.getProperty('minSimilarityScore') || CONFIG.MIN_SIMILARITY_SCORE
  };
}

/**
 * Saves configuration
 */
function saveConfig(config) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('dateToleranceDays', config.dateToleranceDays);
  userProperties.setProperty('minSimilarityScore', config.minSimilarityScore);

  // Update CONFIG object
  CONFIG.DATE_TOLERANCE_DAYS = Number(config.dateToleranceDays);
  CONFIG.MIN_SIMILARITY_SCORE = Number(config.minSimilarityScore);

  return true;
}

/**
 * Clears the output sheet
 */
function clearOutputSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheet = ss.getSheetByName(CONFIG.OUTPUT_SHEET);

  if (outputSheet) {
    outputSheet.clear();
    SpreadsheetApp.getActiveSpreadsheet().toast('Hoja de salida limpiada', 'Éxito', 2);
  }
}

/**
 * Manual match function for sidebar
 */
function manualMatch(accountingRow, bankRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);

  // Here you would implement the logic to mark these as manually matched
  // and store the match for future reference

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Movimiento contable fila ${accountingRow} conciliado con movimiento bancario fila ${bankRow}`,
    'Conciliación Manual',
    3
  );

  return true;
}
