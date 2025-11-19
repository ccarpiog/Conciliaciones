# Bank Reconciliation System - Technical Documentation

## Project Overview

This is an automated bank reconciliation system built for Google Sheets using Google Apps Script. It intelligently matches accounting entries with bank statement transactions using fuzzy matching algorithms, concept similarity scoring, and configurable parameters.

## Architecture

### Core Components

1. **Code.js** - Main reconciliation engine
   - Matching algorithm with configurable thresholds
   - Batch processing optimized for large datasets
   - Manual match persistence using Document Properties
   - Spanish number format parsing support

2. **ConflictsSidebar.html** - Interactive conflict resolution UI
   - Displays conflicts requiring manual review
   - Allows batch conflict resolution
   - Real-time reconciliation updates

3. **ConfigDialog.html** - User configuration interface
   - Date tolerance settings (0-10 days)
   - Minimum similarity score threshold (0-100%)

## Data Model

### Input Sheet: "Origen" (Source)

**Accounting Data (Columns A-D):**
- Column A: Transaction date
- Column B: Entry number
- Column C: Concept/description
- Column D: Amount (supports Spanish format: "1.750,00")

**Bank Data (Columns F-J):**
- Column F: Transaction date
- Column G: Value date (optional)
- Column H: Concept
- Column I: Additional data
- Column J: Amount (supports Spanish format)

### Output Sheet: "Salida" (Output)

Generated automatically with reconciliation results, color-coded by status:
- Green: Auto-matched movements
- Yellow: Conflicts requiring manual review
- Red: Unmatched movements

## Matching Algorithm

### Phase 1: Exact Amount Match (Required)
- Amounts must match exactly (rounded to 2 decimals)
- Uses Map-based indexing for O(1) lookup by amount
- No fuzzy matching on amounts

### Phase 2: Scoring System

Each potential match receives a score (0-1) based on:

**Date Similarity (30% weight):**
- Exact date match: Full score
- Within tolerance: Prorated score based on distance
- Beyond tolerance: No points

**Concept Similarity (70% weight):**

Uses multiple techniques with different scoring levels:

1. **Exact Match (100%)**: Normalized strings are identical
2. **Containment (80%)**: One concept contains the other
3. **Number Matching (60-70%)**:
   - Exact number matches: 60-90%
   - Partial substring matches: 65%
   - Trailing number matches: 70%
4. **Token Matching (30-70%)**: Common words between concepts
5. **Levenshtein Distance (0-50%)**: Character-level similarity for short strings

### Phase 3: Decision Logic

**Auto-match conditions:**
```
isAboveThreshold = bestScore > MIN_SIMILARITY_SCORE (default: 0.3)
isClearWinner = (scoredMatches.length === 1) || (bestScore - secondScore > 0.2)

if (isAboveThreshold && isClearWinner) {
  → Auto-match
}
```

**Conflict types:**

1. **Low Confidence ("Baja confianza")**:
   - `bestScore <= MIN_SIMILARITY_SCORE`
   - No candidate has sufficient similarity
   - Could be 1+ candidates

2. **Multiple Candidates ("Múltiples candidatos")**:
   - `bestScore > MIN_SIMILARITY_SCORE` BUT
   - `bestScore - secondScore <= 0.2`
   - Multiple good matches, can't distinguish winner

## Performance Optimizations

### Pre-computation Strategy
- Normalized strings and extracted numbers are cached during data loading
- Reduces redundant string processing in the matching loop

### Indexing
- Bank movements indexed by amount using Map
- O(1) lookup instead of O(n) iteration
- Only candidates with matching amounts are scored

### Batch Operations
- Single write operation for all output rows
- Single background color application for all rows
- Prevents multiple Google Sheets API calls

### Manual Match Storage
- Uses Document Properties (shared across all users)
- Lock-based concurrency control for batch operations
- Prevents race conditions during simultaneous edits

## String Normalization

All concept strings are normalized for comparison:
```javascript
str.toLowerCase()
   .replace(/[^\w\s]/g, ' ')  // Remove punctuation
   .replace(/\s+/g, ' ')       // Collapse whitespace
   .trim()
```

Numbers are extracted separately: `/\d+/g` with leading zeros removed

## Spanish Number Format Support

The system handles Spanish number formatting:
- Thousands separator: `.` (dot)
- Decimal separator: `,` (comma)
- Example: `-1.750,00` → `-1750.00`

Parsing function: `parseSpanishNumber(value)`

## Configuration System

### User Properties (per-user settings)
- `dateToleranceDays`: Days of difference allowed (default: 3)
- `minSimilarityScore`: Minimum score for auto-match (default: 0.3)

### Document Properties (shared settings)
- `manualMatches`: JSON object mapping accounting IDs to bank IDs
- Persists manual conflict resolutions across sessions

## Conflict Resolution Workflow

### v1.1 Behavior (Current)

Both conflict types now behave identically:
1. **No pre-selection**: Bank columns (F-J) are empty in output
2. **Manual review required**: User must open sidebar
3. **Candidate selection**: User chooses from available candidates
4. **Batch application**: Multiple conflicts can be resolved at once

### Sidebar Features

- **Individual Resolution**: Select and confirm one conflict at a time
- **Batch Resolution**: Select multiple, apply all at once
- **Skip Option**: Leave conflict unresolved for later
- **Automatic Re-run**: Reconciliation re-executes after applying matches

## ID Generation

Stable IDs prevent duplicate matches across re-runs:

**Accounting ID:**
```javascript
`ACC_${dateTimestamp}_${entryNumber}_${amount}`
```

**Bank ID:**
```javascript
`BANK_${dateTimestamp}_${conceptFirst20Chars}_${amount}`
```

## Menu Structure

```
Conciliación (Reconciliation)
├── Ejecutar conciliación automática (Run automatic reconciliation)
├── Revisar conflictos (Review conflicts)
├── Configurar parámetros (Configure parameters)
└── Limpiar todo (Clear everything)
```

## Error Handling

- Sheet validation on startup
- Authorization checks for first-time users
- Lock timeout handling (30 seconds) for batch operations
- Empty candidate array protection

## Changelog

### v1.1 (November 2025)
- **Breaking Change**: Conflicts no longer pre-select first candidate
- Low confidence items now behave like multiple candidate items
- Empty bank columns (F-J) for all conflicts
- Improved documentation distinguishing conflict types

### v1.0 (2024)
- Initial release with automatic reconciliation
- Fuzzy matching with configurable parameters
- Manual conflict resolution via sidebar
- Spanish number format support

## Code Structure

### Main Functions

- `onOpen()`: Creates custom menu
- `runReconciliation()`: Main entry point, orchestrates full reconciliation
- `reconcileMovements()`: Core matching algorithm
- `calculateMatchScore()`: Scoring logic
- `outputReconciliationResults()`: Generates output sheet with color coding
- `getConflictsData()`: Provides conflict data to sidebar
- `resolveConflict()`: Saves single manual match
- `resolveConflictsBatch()`: Saves multiple manual matches atomically

### Helper Functions

- `parseSpanishNumber()`: Parses Spanish number format
- `normalizeString()`: String normalization for comparison
- `extractNumbers()`: Extracts numeric sequences
- `levenshteinDistance()`: Character-level similarity
- `formatDate()`: Consistent date formatting
- `calculateConceptSimilarityCached()`: Optimized similarity using pre-computed values

## Best Practices

### For Users

1. **Start with defaults**: 3 days tolerance, 30% threshold
2. **Review auto-matches**: Verify accuracy before trusting system
3. **Adjust parameters**: Based on data quality and consistency
4. **Standardize concepts**: More consistent naming improves matching
5. **Include reference numbers**: Numbers in concepts improve matching accuracy

### For Developers

1. **Batch operations**: Always use batch writes/reads with Google Sheets API
2. **Pre-compute**: Cache normalized values to avoid redundant processing
3. **Index strategically**: Use Map for O(1) lookups on high-cardinality fields
4. **Stable IDs**: Ensure IDs are deterministic across runs
5. **Lock critical sections**: Use LockService for concurrent operations
6. **Escape HTML**: Always escape user input in HTML output

## Testing

Manual test function available: `testSpanishNumberParser()`

Tests various Spanish number formats and displays results via UI alert.

## Future Enhancements

Potential improvements:
- Machine learning-based similarity scoring
- Multi-currency support
- Date format auto-detection
- Undo functionality for manual matches
- Export reconciliation reports
- Audit trail for all matches
- Performance metrics dashboard

## Technical Constraints

- Google Apps Script execution time limit: 6 minutes
- Property value size limit: 9KB (affects manual match storage)
- Sheet API rate limits apply to batch operations
- Browser-side UI has standard web limitations

## Dependencies

- Google Apps Script runtime
- Google Sheets API (implicit)
- No external libraries required

## Security Considerations

- Script requires Google Sheets edit permissions
- Document Properties are shared across all users with access
- No sensitive data stored in properties
- HTML output is XSS-protected via escaping
- No external API calls or data transmission

## License

[Add license information if applicable]

## Contact

[Add contact/support information if applicable]
