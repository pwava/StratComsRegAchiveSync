/**
 * Create a normalized **name key** for a person, robust to:
 * - middle names
 * - swapped first/last name (Japanese style)
 *
 * Steps:
 * 1) Take FIRST WORD of lastName and firstName.
 * 2) Keep only letters, lowercase.
 * 3) If both exist: sort the two tokens alphabetically and join with "|".
 *    If only one exists: use that single token.
 */
function createNameKeyForArchived_(lastName, firstName) {
  function firstWordLettersOnly(s) {
    if (!s) return '';
    var w = String(s).toLowerCase().trim().split(/\s+/)[0]; // first word
    return w.replace(/[^a-z]/g, ''); // letters only
  }

  var t1 = firstWordLettersOnly(lastName);
  var t2 = firstWordLettersOnly(firstName);

  // If only one part exists, allow it
  if (t1 && !t2) return t1;
  if (t2 && !t1) return t2;

  if (!t1 || !t2) return '';

  // order-independent key (handles swapped names)
  var parts = [t1, t2].sort();
  return parts[0] + '|' + parts[1];
}

/**
 * Create a **person key** using:
 * - smart name key (handles swapped/middle names)
 * - birth date (as string)
 *
 * Same name but different birth date = different person.
 */
function createPersonKeyForArchived_(lastName, firstName, birthDate) {
  var nameKey = createNameKeyForArchived_(lastName, firstName);
  if (!nameKey) return '';

  var bd = (birthDate != null && birthDate !== '') ? String(birthDate) : '';
  return nameKey + '||' + bd;
}

/**
 * Build a validation map for one row:
 * For each column, if it has a "list of items" validation, store allowed values.
 */
function buildValidationMap_(sheet, rowIndex, startCol, numCols) {
  var dv = sheet.getRange(rowIndex, startCol, 1, numCols).getDataValidations();
  if (!dv || dv.length === 0) return new Array(numCols).fill(null);

  var validationsRow = dv[0];
  var map = new Array(numCols);

  for (var c = 0; c < numCols; c++) {
    var rule = validationsRow[c];
    if (!rule) {
      map[c] = null;
      continue;
    }

    var critType = rule.getCriteriaType();
    if (critType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      var critValues = rule.getCriteriaValues();
      var items = (critValues && critValues[0]) ? critValues[0] : [];
      var allowedSet = {};
      for (var i = 0; i < items.length; i++) {
        var v = items[i];
        if (v != null && v !== '') {
          var norm = String(v).toLowerCase().trim();
          allowedSet[norm] = true;
        }
      }
      map[c] = { type: 'LIST', allowed: allowedSet };
    } else {
      map[c] = null; // only handle explicit list-of-items
    }
  }

  return map;
}

/**
 * Clean a row before writing:
 * - Any string starting with "#" (e.g. #REF!, #VALUE!) → blank.
 * - For columns with list-of-items validation:
 *     if value (lowercased/trimmed) not in allowed list → blank.
 */
function cleanRowForWrite_(row, validationMap) {
  for (var c = 0; c < row.length; c++) {
    var v = row[c];

    // Clear any error-like values
    if (typeof v === 'string' && v.length > 0 && v.charAt(0) === '#') {
      row[c] = '';
      continue;
    }

    // Validation-based cleaning
    var info = validationMap && validationMap[c] ? validationMap[c] : null;
    if (info && info.type === 'LIST') {
      if (v != null && String(v).trim() !== '') {
        var norm = String(v).toLowerCase().trim();
        if (!info.allowed[norm]) {
          row[c] = ''; // invalid dropdown value → blank
        }
      }
    }
  }
  return row;
}

/**
 * Ensure a row matches the target sheet width to avoid column mismatch.
 * - Truncate if too long
 * - Pad with blanks if too short
 */
function normalizeRowToWidth_(row, width) {
  var out = row.slice(0, width);
  while (out.length < width) out.push('');
  return out;
}

/**
 * MAIN
 *
 * 1) Read Attendance Stats → map nameKey -> activity (active/core/inactive/archived/etc).
 *    - SAFETY: if the same nameKey appears with conflicting activities, mark it ambiguous
 *      and DO NOT use Attendance Stats for archiving decisions on that nameKey.
 * 2) Read Archived → build personKey map, mark existing duplicate rows as
 *    "Duplicate Archived Record - ignore this row, keep the first one".
 * 3) Directory -> Archived:
 *      - If Activity J = "archived" OR Column I is not blank (ascended),
 *        OR Attendance Stats activity is "archived" (ONLY if not ambiguous)
 *        AND personKey not yet in Archived → move row (one copy).
 *      - If personKey already in Archived → do NOT append, just remove from Directory.
 *        - If ascended, mark that Archived row as
 *          "Permanently Archived because ascended".
 *      - EXCEPTION for NEW MEMBERS:
 *        If Directory Column U date is within the last 6 weeks, then
 *        **do NOT archive** based on "archived" status. (Ascended still archives.)
 *
 * NOTE (per update request):
 * - Removed Archived -> Directory return logic completely.
 *   Archived stays archived permanently (except duplicate/ascended flags).
 *
 * UPDATE (as requested):
 * - Directory is NO LONGER rewritten (so formulas like Column H are preserved).
 *   Instead, rows to be archived are deleted from Directory (bottom-up).
 */
function syncArchivedMembers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var directorySheet = ss.getSheetByName('Directory');
  var archivedSheet = ss.getSheetByName('Archived');
  var configSheet = ss.getSheetByName('Config');

  if (!directorySheet || !archivedSheet || !configSheet) {
    throw new Error('Missing Directory, Archived, or Config sheet.');
  }

  // ===== Layout constants =====
  var DIR_START_ROW   = 4; // Directory data starts at row 4
  var ARCH_START_ROW  = 4; // Archived data starts at row 4
  var STATS_START_ROW = 3; // Attendance Stats data starts at row 3

  // 0-based array indexes (A=0, B=1, C=2, ...)
  var COL_STATUS_A        = 0;  // Column A in Archived
  var COL_LAST_NAME       = 2;  // C
  var COL_FIRST_NAME      = 3;  // D
  var COL_BIRTH_DATE      = 6;  // G
  var COL_ASCENDED        = 8;  // I (Directory only)
  var COL_ACTIVITY        = 9;  // J
  var COL_NEW_MEMBER_DATE = 20; // U (Directory/Archived)

  var lastColDir  = directorySheet.getLastColumn();
  var lastColArch = archivedSheet.getLastColumn();

  // Build validation map once for Directory's first data row
  var validationMapDir = buildValidationMap_(directorySheet, DIR_START_ROW, 1, lastColDir);

  // Precompute today's date (start of day) for 6-week comparison
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var SIX_WEEKS_IN_DAYS = 42;

  // ===== Counters for logging =====
  var lastRowDirInitial = directorySheet.getLastRow();
  var countBeforeDir = Math.max(0, lastRowDirInitial - (DIR_START_ROW - 1));
  var movedToArchived = 0;
  var duplicateMarked = 0;
  var ambiguousStatsKeys = 0;
  var removedFromDirectory = 0;

  // ============================================================
  // 0) Attendance Stats (external file) → nameKey → activity
  // ============================================================
  var attendanceRef = configSheet.getRange('B2').getValue();
  if (!attendanceRef) {
    throw new Error('Config!B2 must contain Attendance sheet URL or ID.');
  }

  var url = String(attendanceRef);
  if (!url.startsWith('http')) {
    url = 'https://docs.google.com/spreadsheets/d/' + url + '/edit';
  }

  var attendanceSs = SpreadsheetApp.openByUrl(url);
  var statsSheet = attendanceSs.getSheetByName('Attendance Stats');
  if (!statsSheet) {
    throw new Error('Attendance Stats tab not found.');
  }

  var statsMap = {}; // nameKey -> activity OR '__AMBIG__'
  var lastRowStats = statsSheet.getLastRow();

  if (lastRowStats >= STATS_START_ROW) {
    var statsValues = statsSheet
      .getRange(STATS_START_ROW, 3, lastRowStats - STATS_START_ROW + 1, 4)
      .getValues(); // C–F

    for (var i = 0; i < statsValues.length; i++) {
      var sRow = statsValues[i];
      var sLast = sRow[0];
      var sFirst = sRow[1];
      var sAct = String(sRow[3] || '').toLowerCase().trim();
      var sNameKey = createNameKeyForArchived_(sLast, sFirst);
      if (!sNameKey) continue;

      if (!statsMap.hasOwnProperty(sNameKey)) {
        statsMap[sNameKey] = sAct;
      } else {
        var existing = String(statsMap[sNameKey] || '').toLowerCase().trim();
        if (existing === '__ambig__') {
          // already ambiguous
        } else if (existing !== sAct) {
          statsMap[sNameKey] = '__AMBIG__';
        }
      }
    }

    for (var key in statsMap) {
      if (statsMap.hasOwnProperty(key) && String(statsMap[key]).toUpperCase() === '__AMBIG__') {
        ambiguousStatsKeys++;
      }
    }
  }

  // ============================================================
  // 1) Read existing Archived → build personKey map, mark DUP
  // ============================================================
  var lastRowArchExisting = archivedSheet.getLastRow();
  var archData = [];
  var archPersonIndex = {}; // personKey -> first index

  if (lastRowArchExisting >= ARCH_START_ROW) {
    archData = archivedSheet
      .getRange(ARCH_START_ROW, 1, lastRowArchExisting - ARCH_START_ROW + 1, lastColArch)
      .getValues();

    for (var j = 0; j < archData.length; j++) {
      var rowA = archData[j];
      var pKeyA = createPersonKeyForArchived_(
        rowA[COL_LAST_NAME],
        rowA[COL_FIRST_NAME],
        rowA[COL_BIRTH_DATE]
      );
      if (!pKeyA) continue;

      if (archPersonIndex.hasOwnProperty(pKeyA)) {
        if (!rowA[COL_STATUS_A]) {
          rowA[COL_STATUS_A] = 'Duplicate Archived Record - ignore this row, keep the first one';
          duplicateMarked++;
        }
      } else {
        archPersonIndex[pKeyA] = j;
      }
    }
  }

  // Track which existing Archived rows need status updates (Column A only)
  var archivedStatusUpdates = [];

  // ============================================================
  // 2) Directory → decide archive + collect rows to delete (NO rewrite)
  // ============================================================
  var rowsToDelete = []; // sheet row numbers (1-based)
  var newArchivedRows = []; // rows to append to next blank row

  if (lastRowDirInitial >= DIR_START_ROW) {
    var dirNumRows = lastRowDirInitial - DIR_START_ROW + 1;
    var dirData = directorySheet
      .getRange(DIR_START_ROW, 1, dirNumRows, lastColDir)
      .getValues();

    for (var r = 0; r < dirData.length; r++) {
      var rowD = dirData[r];
      var ln  = rowD[COL_LAST_NAME];
      var fn  = rowD[COL_FIRST_NAME];
      var bd  = rowD[COL_BIRTH_DATE];
      var asc = rowD[COL_ASCENDED];
      var act = String(rowD[COL_ACTIVITY] || '').toLowerCase().trim();
      var newMemberDate = rowD[COL_NEW_MEMBER_DATE];

      var isNewMember = false;
      if (newMemberDate) {
        var nmDate = new Date(newMemberDate);
        if (!isNaN(nmDate.getTime())) {
          nmDate.setHours(0, 0, 0, 0);
          var diffMs = today.getTime() - nmDate.getTime();
          var diffDays = diffMs / (1000 * 60 * 60 * 24);
          if (diffDays <= SIX_WEEKS_IN_DAYS) isNewMember = true;
        }
      }

      var statsAct = '';
      var nameKeyDir = createNameKeyForArchived_(ln, fn);
      if (nameKeyDir && statsMap[nameKeyDir]) {
        var mapped = String(statsMap[nameKeyDir]).toLowerCase().trim();
        if (mapped !== '__ambig__') statsAct = mapped;
      }

      var hasName =
        (fn && String(fn).trim() !== '') ||
        (ln && String(ln).trim() !== '');

      var ascended = asc != null && String(asc).trim() !== '';
      var shouldArchive = ascended || (
        !isNewMember && (act === 'archived' || statsAct === 'archived')
      );

      if (!hasName || !shouldArchive) continue;

      rowD = cleanRowForWrite_(rowD, validationMapDir);

      var pKey = createPersonKeyForArchived_(ln, fn, bd);

      if (pKey && archPersonIndex.hasOwnProperty(pKey)) {
        if (ascended) {
          var idx = archPersonIndex[pKey];
          var desired = 'Permanently Archived because ascended';
          if (archData[idx][COL_STATUS_A] !== desired) {
            archData[idx][COL_STATUS_A] = desired;
            archivedStatusUpdates.push({ rowNumber: ARCH_START_ROW + idx, value: desired });
          }
        }
      } else {
        var newArchRow = rowD.slice();
        newArchRow[COL_STATUS_A] = ascended
          ? 'Permanently Archived because ascended'
          : (newArchRow[COL_STATUS_A] || '');

        newArchRow = normalizeRowToWidth_(newArchRow, lastColArch);

        newArchivedRows.push(newArchRow);

        if (pKey) {
          archPersonIndex[pKey] = (archData.length + newArchivedRows.length - 1);
        }
        movedToArchived++;
      }

      rowsToDelete.push(DIR_START_ROW + r);
    }
  }

  if (rowsToDelete.length > 0) {
    rowsToDelete.sort(function(a, b) { return b - a; });
    for (var d = 0; d < rowsToDelete.length; d++) {
      directorySheet.deleteRow(rowsToDelete[d]);
      removedFromDirectory++;
    }
  }

  var finalDirCount = Math.max(0, directorySheet.getLastRow() - (DIR_START_ROW - 1));

  // ============================================================
  // 3) Update Archived:
  //    A) Update Column A statuses for existing rows (duplicates + ascended)
  //    B) Append new rows to FIRST TRUE BLANK ROW (not getLastRow)
  // ============================================================

  // A) Update Column A for rows where status text exists and differs
  if (lastRowArchExisting >= ARCH_START_ROW && archData.length > 0) {
    var archAValues = archivedSheet
      .getRange(ARCH_START_ROW, 1, archData.length, 1)
      .getValues();

    var updates = [];
    for (var x = 0; x < archData.length; x++) {
      var desiredA = archData[x][COL_STATUS_A];
      if (!desiredA) continue;

      var currentA = archAValues[x] ? archAValues[x][0] : '';
      if (currentA !== desiredA) {
        updates.push({ rowNumber: ARCH_START_ROW + x, value: desiredA });
      }
    }

    var seen = {};
    for (var u = 0; u < updates.length; u++) seen[String(updates[u].rowNumber)] = true;
    for (var u2 = 0; u2 < archivedStatusUpdates.length; u2++) {
      var rn = String(archivedStatusUpdates[u2].rowNumber);
      if (!seen[rn]) updates.push(archivedStatusUpdates[u2]);
    }

    for (var w = 0; w < updates.length; w++) {
      archivedSheet.getRange(updates[w].rowNumber, 1).setValue(updates[w].value);
    }
  }

  function findFirstBlankArchivedRow_() {
    var maxRows = archivedSheet.getMaxRows();
    if (maxRows < ARCH_START_ROW) return ARCH_START_ROW;

    var numRows = maxRows - ARCH_START_ROW + 1;

    // Check columns C:D (Last/First) to find the first truly empty row
    var vals = archivedSheet.getRange(ARCH_START_ROW, 3, numRows, 2).getValues();
    for (var i = 0; i < vals.length; i++) {
      var last = vals[i][0];
      var first = vals[i][1];
      var isEmpty =
        (last == null || String(last).trim() === '') &&
        (first == null || String(first).trim() === '');
      if (isEmpty) return ARCH_START_ROW + i;
    }
    return maxRows + 1;
  }

  // B) Append new archived rows to first blank row
  if (newArchivedRows.length > 0) {
    var startAppendRow = findFirstBlankArchivedRow_();
    archivedSheet
      .getRange(startAppendRow, 1, newArchivedRows.length, lastColArch)
      .setValues(newArchivedRows);
  }

  var finalArchCount = Math.max(0, archivedSheet.getLastRow() - (ARCH_START_ROW - 1));

  // ============================================================
  // 4) Final counts and log
  // ============================================================
  Logger.log("=== ARCHIVE SYNC SUMMARY ===");
  Logger.log("Directory before: " + countBeforeDir);
  Logger.log("Removed from Directory (archived or already in Archived): " + removedFromDirectory);
  Logger.log("Moved to Archived: " + movedToArchived);
  Logger.log("Duplicates marked in Archived (Duplicate Archived Record - ignore this row, keep the first one): " + duplicateMarked);
  Logger.log("Ambiguous Attendance Stats nameKeys ignored (conflicting statuses): " + ambiguousStatsKeys);
  Logger.log("Directory after: " + finalDirCount);
  Logger.log("Archived total now: " + finalArchCount);
  Logger.log("============================");
}



/**
 * Install a trigger to run syncArchivedMembers() every 6 hours.
 * Run this ONCE manually.
 */
function setupTrigger_syncArchivedEverySixHours() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'syncArchivedMembers') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('syncArchivedMembers')
    .timeBased()
    .everyHours(6)
    .create();
}
