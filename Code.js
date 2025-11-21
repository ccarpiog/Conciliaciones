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
    .addItem('Limpiar todo', 'clearEverything')
    .addToUi();
}

/**
 * Internal reconciliation function without UI actions
 * Used by sidebar to update sheet without triggering UI elements
 */
function runReconciliationInternal() {
  Logger.log('>> runReconciliationInternal: START');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  const outputSheet = ss.getSheetByName(CONFIG.OUTPUT_SHEET);

  if (!sourceSheet || !outputSheet) {
    throw new Error('No se encontraron las hojas "Origen" o "Salida"');
  }

  Logger.log('>> Sheets found: Origen and Salida');

  // Load user configuration
  const userProperties = PropertiesService.getUserProperties();
  const dateToleranceDays = Number(userProperties.getProperty('dateToleranceDays')) || CONFIG.DATE_TOLERANCE_DAYS;
  const minSimilarityScore = Number(userProperties.getProperty('minSimilarityScore')) || CONFIG.MIN_SIMILARITY_SCORE;

  // Update CONFIG with user settings
  CONFIG.DATE_TOLERANCE_DAYS = dateToleranceDays;
  CONFIG.MIN_SIMILARITY_SCORE = minSimilarityScore;

  Logger.log('>> Config - Date tolerance: ' + dateToleranceDays + ' days, Min similarity: ' + minSimilarityScore);

  // Get data from source sheet
  Logger.log('>> Loading accounting data...');
  const accountingData = getAccountingData(sourceSheet);
  Logger.log('>> Accounting data loaded: ' + accountingData.length + ' rows');

  Logger.log('>> Loading bank data...');
  const bankData = getBankData(sourceSheet);
  Logger.log('>> Bank data loaded: ' + bankData.length + ' rows');

  // Perform reconciliation
  Logger.log('>> Starting reconciliation...');
  const reconciliationResults = reconcileMovements(accountingData, bankData);
  Logger.log('>> Reconciliation complete - Matched: ' + reconciliationResults.matched.length +
             ', Conflicts: ' + reconciliationResults.conflicts.length +
             ', Unmatched Accounting: ' + reconciliationResults.unmatchedAccounting.length +
             ', Unmatched Bank: ' + reconciliationResults.unmatchedBank.length);

  // Output results
  Logger.log('>> Writing results to output sheet...');
  outputReconciliationResults(outputSheet, reconciliationResults);
  Logger.log('>> Results written successfully');

  Logger.log('>> runReconciliationInternal: END');
  return reconciliationResults;
}

/**
 * Main reconciliation function
 * Called from menu - shows UI elements (summary and sidebar)
 */
function runReconciliation() {
  try {
    Logger.log('=== START runReconciliation ===');
    const startTime = new Date();
    Logger.log('Start time: ' + startTime.toISOString());

    const reconciliationResults = runReconciliationInternal();

    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    Logger.log('Reconciliation completed in ' + duration + ' seconds');

    // Show summary
    showReconciliationSummary(reconciliationResults);

    // Automatically open conflicts sidebar to show results
    showConflictsSidebar();

    Logger.log('=== END runReconciliation ===');
  } catch (error) {
    Logger.log('ERROR in runReconciliation: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    SpreadsheetApp.getUi().alert('Error: ' + error.message + '\n\nRevise el registro de ejecución (View > Logs) para más detalles.');
  }
}

/**
 * Parses Spanish number format to standard number
 * Converts Spanish format like "-1.750,00" to -1750.00
 * @param {string|number} value - The value to parse
 * @return {number} The parsed number
 */
function parseSpanishNumber(value) {
  // If already a number, return it
  if (typeof value === 'number') {
    return value;
  }

  // Convert to string and trim
  const str = String(value).trim();

  // Remove thousands separators (dots)
  // Replace decimal separator (comma) with dot
  const normalized = str.replace(/\./g, '').replace(',', '.');

  // Parse as float
  const parsed = parseFloat(normalized);

  // Return 0 if parsing failed
  return isNaN(parsed) ? 0 : parsed;
}

/**
 * Test function for Spanish number parsing
 * Run this from the script editor to verify the parser works
 */
function testSpanishNumberParser() {
  const tests = [
    { input: '-1.750,00', expected: -1750.00 },
    { input: '1.750,00', expected: 1750.00 },
    { input: '750,00', expected: 750.00 },
    { input: '-750,50', expected: -750.50 },
    { input: '1.234.567,89', expected: 1234567.89 },
    { input: 1750, expected: 1750 },
    { input: '1750', expected: 1750 },
    { input: '0,00', expected: 0 }
  ];

  let allPassed = true;
  tests.forEach(test => {
    const result = parseSpanishNumber(test.input);
    const passed = Math.abs(result - test.expected) < 0.01;
    if (!passed) {
      Logger.log(`FAILED: parseSpanishNumber("${test.input}") = ${result}, expected ${test.expected}`);
      allPassed = false;
    } else {
      Logger.log(`PASSED: parseSpanishNumber("${test.input}") = ${result}`);
    }
  });

  if (allPassed) {
    Logger.log('\n✓ All tests passed!');
    SpreadsheetApp.getUi().alert('Test completado', 'Todos los tests de parseSpanishNumber pasaron correctamente.', SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    Logger.log('\n✗ Some tests failed');
    SpreadsheetApp.getUi().alert('Test fallido', 'Algunos tests de parseSpanishNumber fallaron. Revise los logs.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
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
    .filter(row => row[CONFIG.ACCOUNTING.DATE_COL] && row[CONFIG.ACCOUNTING.AMOUNT_COL] !== '' && row[CONFIG.ACCOUNTING.AMOUNT_COL] !== null)
    .map((row, index) => {
      // Create stable ID based on date, entry number, and amount
      const date = new Date(row[CONFIG.ACCOUNTING.DATE_COL]);
      const dateStr = date.getTime();
      const entry = String(row[CONFIG.ACCOUNTING.ENTRY_COL] || '');
      const amount = parseSpanishNumber(row[CONFIG.ACCOUNTING.AMOUNT_COL]);
      const id = `ACC_${dateStr}_${entry}_${amount}`;

      const concept = String(row[CONFIG.ACCOUNTING.CONCEPT_COL] || '');

      return {
        id,
        date,
        entryNumber: row[CONFIG.ACCOUNTING.ENTRY_COL],
        concept,
        amount,
        rowNumber: index + 2,
        matched: false,
        bankMatches: [],
        // Pre-compute normalized values for faster matching
        _normalized: normalizeString(concept),
        _numbers: extractNumbers(concept)
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
    .filter(row => row[0] && row[4] !== '' && row[4] !== null) // Check date and amount exist (allow 0)
    .map((row, index) => {
      // Create stable ID based on date, concept, and amount
      const date = new Date(row[0]);
      const dateStr = date.getTime();
      const concept = String(row[2] || '');
      const additional = String(row[3] || '');
      const amount = parseSpanishNumber(row[4]);
      // Include first 20 chars of concept to differentiate same-day same-amount transactions
      const conceptKey = concept.substring(0, 20).replace(/[^a-zA-Z0-9]/g, '');
      const id = `BANK_${dateStr}_${conceptKey}_${amount}`;

      // Combine concept and additional for matching
      const fullConcept = concept + ' ' + additional;

      return {
        id,
        date,
        valueDate: row[1] ? new Date(row[1]) : null, // Value date (G)
        concept, // Concept (H)
        additional, // Additional data (I)
        amount,
        rowNumber: index + 2,
        matched: false,
        // Pre-compute normalized values for faster matching
        _fullConcept: fullConcept,
        _normalized: normalizeString(fullConcept),
        _numbers: extractNumbers(fullConcept)
      };
    });
}

/**
 * Main reconciliation logic - optimized for large datasets
 */
function reconcileMovements(accountingData, bankData) {
  const results = {
    matched: [],
    conflicts: [],
    unmatchedAccounting: [],
    unmatchedBank: new Set(bankData.map(b => b.id)) // Use Set for O(1) deletion
  };

  // Get manual matches
  const manualMatches = getManualMatches();

  // Build index: amount -> array of bank movements (O(n) preprocessing)
  const bankByAmount = new Map();
  const bankById = new Map();

  bankData.forEach(bankMovement => {
    // Pre-calculate rounded amount and store it
    bankMovement._roundedAmount = Math.round(bankMovement.amount * 100) / 100;

    // Index by amount
    const amountKey = bankMovement._roundedAmount;
    if (!bankByAmount.has(amountKey)) {
      bankByAmount.set(amountKey, []);
    }
    bankByAmount.get(amountKey).push(bankMovement);

    // Index by ID for O(1) manual match lookup
    bankById.set(bankMovement.id, bankMovement);
  });

  // Process each accounting movement
  accountingData.forEach(accMovement => {
    // Check if there's a manual match for this accounting movement
    if (manualMatches[accMovement.id]) {
      const manualBankId = manualMatches[accMovement.id];
      const manualMatch = bankById.get(manualBankId);

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

        // Remove from unmatched set (O(1))
        results.unmatchedBank.delete(manualMatch.id);

        return; // Skip automatic matching
      }
    }

    // Find potential matches using amount index (O(k) where k is # with same amount)
    const accAmount = Math.round(accMovement.amount * 100) / 100;
    const candidatesForAmount = bankByAmount.get(accAmount) || [];

    const potentialMatches = candidatesForAmount.filter(bankMovement => !bankMovement.matched);

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

    // Auto-match only if best score meets threshold AND is a clear winner
    const bestScore = scoredMatches[0].score;
    const secondScore = scoredMatches[1]?.score || 0;
    const isAboveThreshold = bestScore > CONFIG.MIN_SIMILARITY_SCORE;
    const isClearWinner = scoredMatches.length === 1 || (bestScore - secondScore > 0.2);

    if (isAboveThreshold && isClearWinner) {
      // Clear winner with good score - automatic match
      const bestMatch = scoredMatches[0].bankMovement;
      results.matched.push({
        accounting: accMovement,
        bank: bestMatch,
        score: bestScore,
        autoMatched: true
      });

      // Mark as matched
      accMovement.matched = true;
      bestMatch.matched = true;

      // Remove from unmatched set (O(1))
      results.unmatchedBank.delete(bestMatch.id);
    } else {
      // Multiple candidates or low confidence - needs manual review
      let reason;
      if (!isAboveThreshold) {
        reason = 'Baja confianza';
      } else if (scoredMatches.length > 1) {
        reason = 'Múltiples candidatos';
      } else {
        reason = 'Baja confianza';
      }

      results.conflicts.push({
        accounting: accMovement,
        candidates: scoredMatches.slice(0, 5), // Keep top 5 candidates
        reason
      });
    }
  });

  // Convert unmatchedBank Set back to array
  results.unmatchedBank = bankData.filter(b => results.unmatchedBank.has(b.id));

  return results;
}

/**
 * Calculates match score between accounting and bank movements
 * Uses pre-computed normalized values for performance
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

  // Concept similarity scoring using cached normalized values
  const conceptSimilarity = calculateConceptSimilarityCached(accMovement, bankMovement);
  score += weights.conceptSimilarity * conceptSimilarity;

  return score;
}

/**
 * Calculates similarity using pre-computed normalized values
 * Much faster than calculateConceptSimilarity for large datasets
 */
function calculateConceptSimilarityCached(accMovement, bankMovement) {
  // Use pre-computed normalized strings
  const norm1 = accMovement._normalized;
  const norm2 = bankMovement._normalized;

  if (!norm1 || !norm2) return 0;

  // Exact match
  if (norm1 === norm2) return 1;

  // Check if one contains the other
  if (norm1.includes(norm2) || norm2.includes(norm1)) {
    return 0.8;
  }

  // Use pre-computed numbers
  const numbers1 = accMovement._numbers;
  const numbers2 = bankMovement._numbers;

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

  // Token-based similarity (already normalized)
  const tokens1 = norm1.split(/\s+/);
  const tokens2 = norm2.split(/\s+/);
  const commonTokens = tokens1.filter(t => tokens2.includes(t));

  if (commonTokens.length > 0) {
    return 0.3 + (0.4 * commonTokens.length / Math.max(tokens1.length, tokens2.length));
  }

  // Levenshtein distance for short strings only
  if (norm1.length < 20 && norm2.length < 20) {
    const distance = levenshteinDistance(norm1, norm2);
    const maxLength = Math.max(norm1.length, norm2.length);
    return Math.max(0, 1 - distance / maxLength) * 0.5;
  }

  return 0;
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
  Logger.log('>>> outputReconciliationResults: START');
  const startTime = new Date();
  let currentRow = 4; // Declare at function scope for error logging

  try {
    // Clear existing content
    Logger.log('>>> Clearing sheet...');
    sheet.clear();

    // Set up headers - write separately to avoid dimension mismatch
    const numColumns = 10;

    // Title row
    Logger.log('>>> Writing title row...');
    const titleRange = sheet.getRange(1, 1, 1, numColumns);
    titleRange.merge();
    titleRange.setValue('MOVIMIENTOS CONCILIADOS');
    titleRange.setBackground('#4a86e8');
    titleRange.setFontColor('#ffffff');
    titleRange.setFontWeight('bold');
    titleRange.setHorizontalAlignment('center');

    // Column headers
    Logger.log('>>> Writing column headers...');
    const headerRange = sheet.getRange(2, 1, 1, numColumns);
    headerRange.setValues([[
      'Fecha Cont.', 'Asiento', 'Concepto Cont.', 'Importe', 'Estado',
      'Fecha Banco', 'Fecha Valor', 'Concepto Banco', 'Datos Adic.', 'Puntuación'
    ]]);
    headerRange.setBackground('#c9daf8');
    headerRange.setFontWeight('bold');

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
        '',
        '',
        ''
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

  // Output sorted movements - BATCH WRITE for performance
  Logger.log('>>> Processing ' + allMovements.length + ' total movements for output');
  if (allMovements.length > 0) {
    // Build data and background color arrays
    Logger.log('>>> Building movement data arrays...');
    const movementData = [];
    const movementBackgrounds = [];

    allMovements.forEach(movement => {
      movementData.push(movement.row);

      // Determine background color based on type
      let bgColor;
      switch(movement.type) {
        case 'matched':
          bgColor = '#d9ead3';
          break;
        case 'manual':
          bgColor = '#b7e1cd'; // Slightly darker green for manual matches
          break;
        case 'conflict':
          bgColor = '#fff2cc';
          break;
        case 'unmatched':
          bgColor = '#f4cccc';
          break;
        default:
          bgColor = '#ffffff';
      }

      // Create array of same color for all columns in this row
      const rowColors = new Array(movement.row.length).fill(bgColor);
      movementBackgrounds.push(rowColors);
    });

    // Write all data in ONE operation
    Logger.log('>>> Writing ' + movementData.length + ' rows of movement data to sheet...');
    sheet.getRange(currentRow, 1, movementData.length, numColumns).setValues(movementData);

    // Apply all backgrounds in ONE operation
    Logger.log('>>> Applying background colors to ' + movementBackgrounds.length + ' rows...');
    sheet.getRange(currentRow, 1, movementBackgrounds.length, numColumns).setBackgrounds(movementBackgrounds);

    currentRow += allMovements.length;
    Logger.log('>>> Movement data written. Current row: ' + currentRow);
  }

  // Add summary section for unmatched bank movements
  Logger.log('>>> Processing unmatched bank movements: ' + results.unmatchedBank.length);
  if (results.unmatchedBank.length > 0) {
    currentRow += 2;
    Logger.log('>>> Writing unmatched bank movements section at row ' + currentRow);

    // Add a delay to let API recover from previous operations
    Utilities.sleep(1000); // 1 second delay

    try {
      // Prepare all data including headers
      const unmatchedTitleRow = currentRow;
      const unmatchedHeaderRow = currentRow + 1;
      const unmatchedDataStartRow = currentRow + 2;

      // Write title text only
      Logger.log('>>> Writing unmatched title at row ' + unmatchedTitleRow);
      sheet.getRange(unmatchedTitleRow, 1).setValue('MOVIMIENTOS BANCARIOS NO CONCILIADOS');

      // Write header row
      Logger.log('>>> Writing unmatched header at row ' + unmatchedHeaderRow);
      sheet.getRange(unmatchedHeaderRow, 1, 1, 5).setValues([
        ['Fecha Mov.', 'Fecha Valor', 'Concepto', 'Datos Adic.', 'Importe']
      ]);

      // Write data
      Logger.log('>>> Writing ' + results.unmatchedBank.length + ' unmatched bank rows...');
      const bankData = results.unmatchedBank.map(bank => [
        formatDate(bank.date),
        bank.valueDate ? formatDate(bank.valueDate) : '',
        bank.concept,
        bank.additional,
        bank.amount
      ]);
      sheet.getRange(unmatchedDataStartRow, 1, bankData.length, 5).setValues(bankData);

      // Now apply formatting in batch operations
      Logger.log('>>> Applying formatting to unmatched bank section...');
      Utilities.sleep(500);

      // Format title row
      const unmatchedTitleRange = sheet.getRange(unmatchedTitleRow, 1, 1, 5);
      unmatchedTitleRange.merge();
      unmatchedTitleRange.setBackground('#ea4335');
      unmatchedTitleRange.setFontColor('#ffffff');
      unmatchedTitleRange.setFontWeight('bold');
      unmatchedTitleRange.setHorizontalAlignment('center');

      // Format header row
      const unmatchedHeaderRange = sheet.getRange(unmatchedHeaderRow, 1, 1, 5);
      unmatchedHeaderRange.setBackground('#f4cccc');
      unmatchedHeaderRange.setFontWeight('bold');

      currentRow = unmatchedDataStartRow + bankData.length;
      Logger.log('>>> Unmatched bank section complete. Current row: ' + currentRow);
    } catch (unmatchedError) {
      Logger.log('>>> ERROR writing unmatched bank section: ' + unmatchedError.toString());
      // Try to continue without formatting
      currentRow += 2 + results.unmatchedBank.length;
    }
  }

  // Set fixed column widths (much faster than auto-resize for large datasets)
  Logger.log('>>> Setting column widths...');

  // Add delay before column width operations
  Utilities.sleep(500);

  // Set widths individually (Google Sheets API doesn't support batch width setting)
  const columnWidths = [100, 80, 250, 100, 150, 100, 100, 250, 150, 100];
  try {
    columnWidths.forEach((width, index) => {
      sheet.setColumnWidth(index + 1, width);
    });
    Logger.log('>>> Column widths set');
  } catch (widthError) {
    Logger.log('>>> Warning: Could not set column widths: ' + widthError.toString());
    // Non-critical error, continue execution
  }

  const endTime = new Date();
  const totalDuration = (endTime - startTime) / 1000;
  Logger.log('>>> outputReconciliationResults: END (took ' + totalDuration + ' seconds)');

  } catch (error) {
    Logger.log('>>> ERROR in outputReconciliationResults: ' + error.toString());
    Logger.log('>>> Error stack: ' + error.stack);
    Logger.log('>>> Error occurred at currentRow: ' + currentRow);
    throw error; // Re-throw to be caught by caller
  }
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
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      SpreadsheetApp.getUi().alert(
        'Error de acceso',
        'No se pudo acceder a la hoja de cálculo. Por favor, intente cerrar y volver a abrir el documento.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    const outputSheet = ss.getSheetByName(CONFIG.OUTPUT_SHEET);

    // Check if reconciliation has been run (output sheet has content beyond row 1)
    if (!outputSheet || outputSheet.getLastRow() <= 1) {
      SpreadsheetApp.getUi().alert(
        'Conciliación no ejecutada',
        'Debe ejecutar "Ejecutar conciliación automática" antes de revisar conflictos.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    const html = HtmlService.createHtmlOutputFromFile('ConflictsSidebar')
      .setTitle('Resolver Conflictos')
      .setWidth(400);

    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    Logger.log('Error in showConflictsSidebar: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'Error',
      'Error al abrir la barra lateral: ' + error.message + '\n\nPor favor, intente:\n1. Ejecutar "Ejecutar conciliación automática" primero\n2. Reautorizar el script desde el editor (Extensions > Apps Script)',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Gets conflicts data for sidebar
 */
function getConflictsData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      throw new Error('No se pudo acceder a la hoja de cálculo activa');
    }

    const sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);

    if (!sourceSheet) {
      throw new Error('No se encontró la hoja "' + CONFIG.SOURCE_SHEET + '"');
    }

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
  } catch (error) {
    Logger.log('Error in getConflictsData: ' + error.toString());
    throw new Error('Error al cargar los conflictos: ' + error.message);
  }
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
 * Resolves multiple conflicts at once (batch operation)
 * @param {Array<{accountingId: string, bankId: string}>} matches - Array of matches to resolve
 */
function resolveConflictsBatch(matches) {
  // Validation
  if (!matches || !Array.isArray(matches) || matches.length === 0) {
    return { success: false, message: 'No hay conciliaciones para aplicar' };
  }

  // De-duplicate matches by accountingId (keep last occurrence)
  const uniqueMatches = {};
  matches.forEach(match => {
    if (match && match.accountingId && match.bankId) {
      uniqueMatches[match.accountingId] = match.bankId;
    }
  });

  if (Object.keys(uniqueMatches).length === 0) {
    return { success: false, message: 'No hay conciliaciones válidas para aplicar' };
  }

  // Acquire lock to prevent race conditions
  const lock = LockService.getDocumentLock();
  try {
    // Wait up to 30 seconds for the lock
    lock.waitLock(30000);
  } catch (lockError) {
    return { success: false, message: 'No se pudo obtener el bloqueo. Otro usuario podría estar aplicando conciliaciones. Intente de nuevo.' };
  }

  try {
    // Store all manual matches in a single operation
    const documentProperties = PropertiesService.getDocumentProperties();
    const manualMatches = JSON.parse(documentProperties.getProperty('manualMatches') || '{}');

    // Add all new matches
    Object.keys(uniqueMatches).forEach(accountingId => {
      manualMatches[accountingId] = uniqueMatches[accountingId];
    });

    // Save back to properties (single write operation)
    documentProperties.setProperty('manualMatches', JSON.stringify(manualMatches));

    return { success: true, count: Object.keys(uniqueMatches).length };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    // Always release the lock
    lock.releaseLock();
  }
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
 * Clears everything: output sheet and manual matches with user confirmation
 */
function clearEverything() {
  const ui = SpreadsheetApp.getUi();
  const manualMatches = getManualMatches();
  const matchCount = Object.keys(manualMatches).length;

  // Build confirmation message
  let message = '¿Desea realizar las siguientes acciones?\n\n';
  message += '• Limpiar la hoja de salida "Salida"\n';
  message += `• Borrar ${matchCount} ${matchCount === 1 ? 'conciliación manual' : 'conciliaciones manuales'}\n\n`;
  message += 'Esta acción no se puede deshacer.';

  const response = ui.alert(
    'Limpiar todo',
    message,
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    // Clear output sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const outputSheet = ss.getSheetByName(CONFIG.OUTPUT_SHEET);
    if (outputSheet) {
      outputSheet.clear();
    }

    // Clear manual matches
    clearManualMatches();

    // Hide sidebar if it's open
    try {
      ui.showSidebar(HtmlService.createHtmlOutput(''));
    } catch (e) {
      // Sidebar may not be open, ignore error
    }

    // Show success message
    ui.alert('Limpieza completada', 'La hoja de salida y las conciliaciones manuales han sido eliminadas.', ui.ButtonSet.OK);
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
