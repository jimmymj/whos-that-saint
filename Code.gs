/**
 * ============================================================================
 * "WHO’S THAT SAINT?" — SECURE GOOGLE SHEETS CLOUD BACKEND (Code.gs)
 * ============================================================================
 * Tabs Managed:
 *   1. Roster_Saints_Kids             (Master 100-Kid Input Table + Outfit_Clue)
 *   2. Quiz_Participants_Scores       (Centralized Participant Points Log)
 *   3. Live_Leaderboard_And_Kid_Stats (Automated Summary & Appreciation Dashboard)
 *
 * IMPORTANT — HOW TO UPDATE YOUR GOOGLE SHEET & WEB APP:
 *   1. In your Google Sheet, click Extensions -> Apps Script.
 *   2. Replace ALL code in Code.gs with this entire file and press Ctrl + S (Save).
 *   3. Select "setupSheetStructure" in the top dropdown and click "Run".
 *      -> This immediately adds the "Outfit_Clue" column (Column F) to Roster_Saints_Kids
 *         and populates the Saint Outfit Clues right inside your Google Sheet!
 *   4. Click "Deploy" (top right) -> "New deployment" -> Select type: "Web app"
 *      -> Execute as: "Me", Who has access: "Anyone" -> Click "Deploy".
 *   5. Copy the NEW Web App URL and paste it into the App's Admin Tab 1 (PIN: 1101).
 * ============================================================================
 */

const TAB_ROSTER = 'Roster_Saints_Kids';
const TAB_SCORES = 'Quiz_Participants_Scores';
const TAB_DASHBOARD = 'Live_Leaderboard_And_Kid_Stats';

const ROSTER_HEADERS = [
  'Bib_Number',
  'Kid_Name',
  'Saint_Name',
  'Category_Grade',
  'Saint_Clue',
  'Outfit_Clue',
  'Secret_Pin'
];

const STARTER_KIDS_WITH_OUTFIT_CLUES = [
  [1, 'Noah Thomas', 'St. Francis of Assisi', 'PreK-2', 'Patron saint of animals, birds, and ecology who loved Lady Poverty', 'Brown Franciscan friar robe tied with a white 3-knot rope belt, plush bird or wolf prop', '101'],
  [2, 'Emma Garcia', 'St. Thérèse of Lisieux', 'PreK-2', 'Known as the Little Flower who practiced the Little Way of love', 'Brown and black Carmelite habit with white veil, holding red roses and a crucifix', '102'],
  [3, 'Liam Miller', 'St. Michael the Archangel', 'PreK-2', 'Leader of the heavenly armies who defeated the dragon', 'Shining Roman breastplate armor, large angel wings, red cape, sword and shield', '103'],
  [4, 'Olivia Davis', 'St. Joan of Arc', 'Grades 3-5', 'Teenage French heroine who led troops into battle', 'Silver knight armor over tunic, carrying a white French fleur-de-lis banner and sword', '104'],
  [5, 'Lucas Rodriguez', 'St. Patrick', 'Grades 3-5', 'Used the three-leaf shamrock to teach the Holy Trinity in Ireland', 'Green bishop vestments, tall gold mitre hat, pastoral staff, and a green shamrock', '105'],
  [6, 'Sophia Martinez', 'St. Teresa of Calcutta', 'Grades 3-5', 'Founder of the Missionaries of Charity serving the poorest of the poor', 'White cotton sari with three royal-blue border stripes and a small shoulder crucifix', '106'],
  [7, 'Mateo Hernandez', 'St. Juan Diego', 'Grades 3-5', 'Humble peasant whose tilma holds the miraculous image of Our Lady of Guadalupe', 'Beige woven tilma poncho painted with Our Lady of Guadalupe and red roses inside', '107'],
  [8, 'Isabella Lopez', 'St. Kateri Tekakwitha', 'Grades 6-8', 'Lily of the Mohawks and first Native American canonized saint', 'Fringed buckskin tunic, braided hair, woven shawl, holding a handmade wooden cross and lily', '108'],
  [9, 'Ethan Gonzalez', 'St. Carlo Acutis', 'Grades 6-8', 'Modern computer-coding teenager who cataloged Eucharistic miracles online', 'Red polo shirt or zip hoodie, jeans, sneakers, backpack, and a laptop prop', '109'],
  [10, 'Mia Wilson', 'St. Cecilia', 'Grades 6-8', 'Early Roman martyr and patron saint of musicians and church music', 'Flowing Roman gown with flower crown, holding a small harp, violin, flute, or keyboard', '110']
];

/**
 * Security Helper: Prevents Google Sheets Formula Injection (CSV/Sheet Injection)
 * and strips control characters / excessive length.
 */
function sanitizeCellForSheet(rawVal, maxLen) {
  const limit = maxLen || 220;
  let str = String(rawVal === null || rawVal === undefined ? '' : rawVal).trim();
  str = str.replace(/[\x00-\x1F\x7F]/g, '');
  if (/^[=+\-@]/.test(str)) {
    str = "'" + str;
  }
  return str.substring(0, limit);
}

/**
 * Helper: Normalizes Participant_Name for case-insensitive uniqueness check.
 */
function normalizeParticipantName(name) {
  return String(name || '')
    .replace(/^'+/, '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

/**
 * 1. ONE-CLICK INITIALIZER & UPGRADER:
 *    Creates or upgrades all 3 tabs, adds the "Outfit_Clue" column (Column F)
 *    to Roster_Saints_Kids, and populates missing Outfit Clues automatically!
 */
function setupSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Please open Apps Script from inside your Google Sheet (Extensions -> Apps Script).');
  }

  // --- TAB 1: Roster_Saints_Kids ---
  let rosterSheet = ss.getSheetByName(TAB_ROSTER);
  if (!rosterSheet) {
    rosterSheet = ss.insertSheet(TAB_ROSTER);
  }

  // Write the 7 official headers (Bib_Number, Kid_Name, Saint_Name, Category_Grade, Saint_Clue, Outfit_Clue, Secret_Pin)
  rosterSheet.getRange(1, 1, 1, ROSTER_HEADERS.length)
    .setValues([ROSTER_HEADERS])
    .setFontWeight('bold')
    .setBackground('#4a1c11')
    .setFontColor('#fde68a');
  rosterSheet.setFrozenRows(1);

  const lastRow = rosterSheet.getLastRow();
  if (lastRow <= 11) {
    // Populate or upgrade the 10 starter rows with full Outfit_Clue data in Column F
    rosterSheet.getRange(2, 1, STARTER_KIDS_WITH_OUTFIT_CLUES.length, ROSTER_HEADERS.length)
      .setValues(STARTER_KIDS_WITH_OUTFIT_CLUES);
  } else {
    // Upgrade any existing rows where Outfit_Clue (Column F) was previously Secret_Pin or blank
    const existingRange = rosterSheet.getRange(2, 1, lastRow - 1, ROSTER_HEADERS.length);
    const rows = existingRange.getValues();
    const updatedRows = rows.map(function(row, idx) {
      const bib = row[0] || (idx + 1);
      const kid = row[1] || '';
      const saint = row[2] || '';
      const cat = row[3] || 'Grades 3-5';
      const bioClue = row[4] || '';
      let outfitClue = String(row[5] || '').trim();
      let pin = String(row[6] || '').trim();

      // If Column F looks like a 3-digit PIN from the older 6-column format, move it to Column G!
      if (/^\d{3,4}$/.test(outfitClue) && !pin) {
        pin = outfitClue;
        outfitClue = '';
      }
      if (!outfitClue) {
        const matchStarter = STARTER_KIDS_WITH_OUTFIT_CLUES.find(function(s) { return s[2] === saint; });
        outfitClue = matchStarter ? matchStarter[5] : ('Traditional saint vestments and patron symbol prop for ' + saint);
      }
      return [bib, kid, saint, cat, bioClue, outfitClue, pin || String(100 + Number(bib)).slice(-3)];
    });
    existingRange.setValues(updatedRows);
  }

  // Auto-resize columns A..G so Outfit_Clue is clearly visible in Google Sheets
  rosterSheet.autoResizeColumns(1, ROSTER_HEADERS.length);

  // --- TAB 2: Quiz_Participants_Scores ---
  let scoresSheet = ss.getSheetByName(TAB_SCORES);
  if (!scoresSheet) {
    scoresSheet = ss.insertSheet(TAB_SCORES);
  }
  const scoreHeaders = [
    'Submission_ID',
    'Timestamp',
    'Participant_Name',
    'Family_or_Phone',
    'Game_Mode',
    'Correct_Score',
    'Total_Attempted',
    'Accuracy_Percent',
    'Time_Taken_Sec',
    'Matched_Bib_IDs'
  ];
  scoresSheet.getRange(1, 1, 1, scoreHeaders.length)
    .setValues([scoreHeaders])
    .setFontWeight('bold')
    .setBackground('#4a1c11')
    .setFontColor('#fde68a');
  scoresSheet.setFrozenRows(1);

  // --- TAB 3: Live_Leaderboard_And_Kid_Stats ---
  let dashSheet = ss.getSheetByName(TAB_DASHBOARD);
  if (!dashSheet) {
    dashSheet = ss.insertSheet(TAB_DASHBOARD);
  }
  dashSheet.clear();

  dashSheet.getRange('A1').setValue('🏆 TOP 10 PARTICIPANTS (10 PTS/MATCH, -2 PTS CLUE)').setFontWeight('bold').setFontSize(12);
  const lbHeaders = ['Participant_Name', 'Family_or_Phone', 'Game_Mode', 'Correct_Score', 'Accuracy_Percent', 'Time_Taken_Sec'];
  dashSheet.getRange(2, 1, 1, lbHeaders.length)
    .setValues([lbHeaders])
    .setFontWeight('bold')
    .setBackground('#6b2818')
    .setFontColor('#ffffff');

  dashSheet.getRange('A3').setFormula(
    '=IFERROR(QUERY(Quiz_Participants_Scores!A2:J, "SELECT C, D, E, F, H, I WHERE C IS NOT NULL ORDER BY F DESC, I ASC LIMIT 10", 0), "No submissions yet")'
  );

  dashSheet.getRange('H1').setValue('✞ 100-KID SAINT APPRECIATION COUNTER').setFontWeight('bold').setFontSize(12);
  const kidStatHeaders = ['Bib_Number', 'Kid_Name', 'Saint_Name', 'Times_Identified_By_Community'];
  dashSheet.getRange(2, 8, 1, kidStatHeaders.length)
    .setValues([kidStatHeaders])
    .setFontWeight('bold')
    .setBackground('#6b2818')
    .setFontColor('#fde68a');

  dashSheet.getRange('H3').setFormula('=ARRAYFORMULA(Roster_Saints_Kids!A2:C101)');

  const formulas = [];
  for (let r = 3; r <= 102; r++) {
    formulas.push([`=IF(H${r}="","",COUNTIF(ARRAYFORMULA(","&Quiz_Participants_Scores!$J$2:$J&","), "*,"&H${r}&",*"))`]);
  }
  dashSheet.getRange(3, 11, 100, 1).setFormulas(formulas);

  SpreadsheetApp.flush();
}

/**
 * 2. HTTP GET: Returns the ~100-kid Roster (with Outfit_Clue) + Leaderboard as JSON.
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return createJsonResponse({
        status: 'error',
        message: 'Script is not bound to a Google Sheet. Open your Google Sheet -> Extensions -> Apps Script.'
      });
    }
    let rosterSheet = ss.getSheetByName(TAB_ROSTER);
    let scoresSheet = ss.getSheetByName(TAB_SCORES);
    if (!rosterSheet || !scoresSheet) {
      setupSheetStructure();
      rosterSheet = ss.getSheetByName(TAB_ROSTER);
      scoresSheet = ss.getSheetByName(TAB_SCORES);
    }

    const roster = readSheetAsObjects(rosterSheet);
    const allScores = readSheetAsObjects(scoresSheet);

    allScores.sort(function(a, b) {
      const scoreDiff = Number(b.Correct_Score || 0) - Number(a.Correct_Score || 0);
      if (scoreDiff !== 0) return scoreDiff;
      return Number(a.Time_Taken_Sec || 999999) - Number(b.Time_Taken_Sec || 999999);
    });

    return createJsonResponse({
      status: 'success',
      timestamp: new Date().toISOString(),
      roster: roster,
      leaderboard: allScores.slice(0, 50)
    });
  } catch (err) {
    return createJsonResponse({
      status: 'error',
      message: err.toString()
    });
  }
}

/**
 * 3. HTTP POST: Handles Participant Score Submissions AND Admin Roster Add/Edit/Delete/Sync.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const rawBody = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const body = JSON.parse(rawBody);
    const action = body.action || 'submitScore';
    const p = body.payload || body;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let rosterSheet = ss.getSheetByName(TAB_ROSTER);
    let scoresSheet = ss.getSheetByName(TAB_SCORES);
    if (!rosterSheet || !scoresSheet) {
      setupSheetStructure();
      rosterSheet = ss.getSheetByName(TAB_ROSTER);
      scoresSheet = ss.getSheetByName(TAB_SCORES);
    }

    // Ensure Roster_Saints_Kids header row includes Outfit_Clue (Columns A..G)
    rosterSheet.getRange(1, 1, 1, ROSTER_HEADERS.length).setValues([ROSTER_HEADERS]);

    // ====================================================================
    // ACTION A: ADD OR EDIT A KID & SAINT ROW IN GOOGLE SHEET
    // ====================================================================
    if (action === 'saveKidRow') {
      const bibNum = Math.max(1, Math.min(500, Math.floor(Number(p.Bib_Number || 1))));
      const kidRowValues = [
        bibNum,
        sanitizeCellForSheet(p.Kid_Name || '', 80),
        sanitizeCellForSheet(p.Saint_Name || '', 80),
        sanitizeCellForSheet(p.Category_Grade || 'Grades 3-5', 40),
        sanitizeCellForSheet(p.Saint_Clue || '', 220),
        sanitizeCellForSheet(p.Outfit_Clue || '', 220),
        sanitizeCellForSheet(p.Secret_Pin || '', 15)
      ];

      const data = rosterSheet.getDataRange().getValues();
      let updatedRowIndex = -1;
      for (let r = 1; r < data.length; r++) {
        if (Number(data[r][0]) === bibNum) {
          updatedRowIndex = r + 1;
          break;
        }
      }

      if (updatedRowIndex > 0) {
        rosterSheet.getRange(updatedRowIndex, 1, 1, ROSTER_HEADERS.length).setValues([kidRowValues]);
      } else {
        rosterSheet.appendRow(kidRowValues);
      }

      // Sort Roster by Bib_Number ascending
      if (rosterSheet.getLastRow() > 2) {
        rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, ROSTER_HEADERS.length).sort({ column: 1, ascending: true });
      }
      SpreadsheetApp.flush();

      return createJsonResponse({
        status: 'success',
        action: 'saveKidRow',
        bibNumber: bibNum,
        roster: readSheetAsObjects(rosterSheet)
      });
    }

    // ====================================================================
    // ACTION B: DELETE A KID & SAINT ROW FROM GOOGLE SHEET
    // ====================================================================
    if (action === 'deleteKidRow') {
      const bibToDelete = Number(p.Bib_Number || 0);
      const data = rosterSheet.getDataRange().getValues();
      let deleted = false;
      for (let r = data.length - 1; r >= 1; r--) {
        if (Number(data[r][0]) === bibToDelete) {
          rosterSheet.deleteRow(r + 1);
          deleted = true;
        }
      }
      SpreadsheetApp.flush();
      return createJsonResponse({
        status: 'success',
        action: 'deleteKidRow',
        deleted: deleted,
        bibNumber: bibToDelete,
        roster: readSheetAsObjects(rosterSheet)
      });
    }

    // ====================================================================
    // ACTION C: BULK SYNC FULL ROSTER TO GOOGLE SHEET
    // ====================================================================
    if (action === 'syncFullRoster' && Array.isArray(p.roster)) {
      const rows = p.roster.map(function(item, idx) {
        return [
          Math.max(1, Math.floor(Number(item.Bib_Number || (idx + 1)))),
          sanitizeCellForSheet(item.Kid_Name || '', 80),
          sanitizeCellForSheet(item.Saint_Name || '', 80),
          sanitizeCellForSheet(item.Category_Grade || 'Grades 3-5', 40),
          sanitizeCellForSheet(item.Saint_Clue || '', 220),
          sanitizeCellForSheet(item.Outfit_Clue || '', 220),
          sanitizeCellForSheet(item.Secret_Pin || '', 15)
        ];
      });

      if (rosterSheet.getLastRow() > 1) {
        rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, ROSTER_HEADERS.length).clearContent();
      }
      if (rows.length > 0) {
        rosterSheet.getRange(2, 1, rows.length, ROSTER_HEADERS.length).setValues(rows);
      }
      SpreadsheetApp.flush();

      return createJsonResponse({
        status: 'success',
        action: 'syncFullRoster',
        count: rows.length
      });
    }

    // ====================================================================
    // ACTION D: PARTICIPANT SCORE SUBMISSION (DEFAULT)
    // ====================================================================
    const submissionId = sanitizeCellForSheet(p.Submission_ID || ('SUB-' + new Date().getTime()), 50);
    const cleanName = sanitizeCellForSheet(p.Participant_Name || '', 60);
    const normCandidateName = normalizeParticipantName(cleanName);

    if (!normCandidateName || normCandidateName.length < 2) {
      return createJsonResponse({
        status: 'error',
        code: 'INVALID_NAME',
        message: 'Participant_Name must be at least 2 characters.'
      });
    }

    const existingData = scoresSheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      if (String(existingData[i][0]) === submissionId) {
        return createJsonResponse({
          status: 'success',
          duplicate: true,
          rank: calculateParticipantRank(existingData, Number(p.Correct_Score || 0), Number(p.Time_Taken_Sec || 0))
        });
      }
      if (normalizeParticipantName(existingData[i][2]) === normCandidateName) {
        return createJsonResponse({
          status: 'error',
          code: 'DUPLICATE_NAME',
          message: 'Participant name is already registered. Duplicate names are not allowed.'
        });
      }
    }

    const totalAttempted = Math.min(100, Math.max(1, Math.floor(Number(p.Total_Attempted || 1))));
    const maxPossiblePoints = totalAttempted * 10;
    const pointsScore = Math.min(maxPossiblePoints, Math.max(0, Math.floor(Number(p.Correct_Score || 0))));
    const accuracyPct = Math.min(100, Math.max(0, Number(Number(p.Accuracy_Percent || 0).toFixed(1))));
    const timeTakenSec = Math.min(86400, Math.max(1, Math.floor(Number(p.Time_Taken_Sec || 1))));

    const validBibs = String(p.Matched_Bib_IDs || '')
      .split(',')
      .map(function(s) { return parseInt(s.trim(), 10); })
      .filter(function(n) { return !isNaN(n) && n >= 1 && n <= 500; });

    const rowToAppend = [
      submissionId,
      sanitizeCellForSheet(p.Timestamp || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'), 30),
      cleanName,
      sanitizeCellForSheet(p.Family_or_Phone || '', 60),
      sanitizeCellForSheet(p.Game_Mode || 'Find Saint by Kid Name', 40),
      pointsScore,
      totalAttempted,
      accuracyPct,
      timeTakenSec,
      validBibs.join(',')
    ];

    scoresSheet.appendRow(rowToAppend);
    SpreadsheetApp.flush();

    const updatedData = scoresSheet.getDataRange().getValues();
    const calculatedRank = calculateParticipantRank(updatedData, pointsScore, timeTakenSec);

    return createJsonResponse({
      status: 'success',
      submissionId: submissionId,
      rank: calculatedRank
    });
  } catch (err) {
    return createJsonResponse({
      status: 'error',
      message: err.toString()
    });
  } finally {
    lock.releaseLock();
  }
}

function readSheetAsObjects(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const rowObj = {};
    let hasContent = false;
    for (let c = 0; c < headers.length; c++) {
      rowObj[headers[c]] = values[r][c];
      if (values[r][c] !== '' && values[r][c] !== null) {
        hasContent = true;
      }
    }
    if (hasContent) rows.push(rowObj);
  }
  return rows;
}

function calculateParticipantRank(sheetValues, myScore, myTime) {
  let rank = 1;
  for (let i = 1; i < sheetValues.length; i++) {
    const rowScore = Number(sheetValues[i][5] || 0);
    const rowTime = Number(sheetValues[i][8] || 999999);
    if (rowScore > myScore || (rowScore === myScore && rowTime < myTime)) {
      rank++;
    }
  }
  return rank;
}

function createJsonResponse(payloadObj) {
  return ContentService
    .createTextOutput(JSON.stringify(payloadObj))
    .setMimeType(ContentService.MimeType.JSON);
}
