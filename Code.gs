/**
 * ============================================================================
 * "WHO’S THAT SAINT?" — SECURE GOOGLE SHEETS CLOUD BACKEND (Code.gs)
 * ============================================================================
 * Tabs Managed:
 *   1. Roster_Saints_Kids             (Master 100-Kid Input Table + Outfit_Clue)
 *   2. Quiz_Participants_Scores       (Centralized Participant Points Log)
 *   3. Live_Leaderboard_And_Kid_Stats (Automated Summary, Appreciation & App Config)
 *
 * SUPPORTED ACTIONS:
 *   - GET                          : Returns Roster + Leaderboard + Live App Config
 *   - POST/GET action="submitScore"    : Validates & appends unique participant score
 *   - POST/GET action="saveKidRow"     : Adds or updates Kid + Saint + Outfit_Clue
 *   - POST/GET action="deleteKidRow"   : Deletes Kid row by Bib_Number
 *   - POST/GET action="syncFullRoster" : Bulk-syncs full 100-kid roster
 *   - POST/GET action="saveAppConfig"  : Saves enabled/disabled Saint Find Options in Sheet
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

function sanitizeCellForSheet(rawVal, maxLen) {
  const limit = maxLen || 220;
  let str = String(rawVal === null || rawVal === undefined ? '' : rawVal).trim();
  str = str.replace(/[\x00-\x1F\x7F]/g, '');
  if (/^[=+\-@]/.test(str)) {
    str = "'" + str;
  }
  return str.substring(0, limit);
}

function normalizeParticipantName(name) {
  return String(name || '')
    .replace(/^'+/, '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

/**
 * Reads App Config (which Saint Find Options are enabled/disabled) from Sheet & ScriptProperties
 */
function readAppConfigFromSheet(ss) {
  const defaultCfg = {
    option1_KidToSaint: true,
    option2_SaintToKid: true,
    requireSecretPin: false
  };
  try {
    const props = PropertiesService.getScriptProperties();
    const savedJson = props.getProperty('APP_CONFIG_JSON');
    if (savedJson) {
      const parsed = JSON.parse(savedJson);
      defaultCfg.option1_KidToSaint = parsed.option1_KidToSaint !== false;
      defaultCfg.option2_SaintToKid = parsed.option2_SaintToKid !== false;
      defaultCfg.requireSecretPin = Boolean(parsed.requireSecretPin);
    }
    const dashSheet = ss ? ss.getSheetByName(TAB_DASHBOARD) : null;
    if (dashSheet) {
      const cfgValues = dashSheet.getRange('M3:N5').getValues();
      if (cfgValues[0][0] === 'option1_KidToSaint' && cfgValues[0][1] !== '') {
        defaultCfg.option1_KidToSaint = String(cfgValues[0][1]).toLowerCase() !== 'false';
      }
      if (cfgValues[1][0] === 'option2_SaintToKid' && cfgValues[1][1] !== '') {
        defaultCfg.option2_SaintToKid = String(cfgValues[1][1]).toLowerCase() !== 'false';
      }
      if (cfgValues[2][0] === 'requireSecretPin' && cfgValues[2][1] !== '') {
        defaultCfg.requireSecretPin = String(cfgValues[2][1]).toLowerCase() === 'true';
      }
    }
  } catch (e) {}
  // Ensure at least one option is enabled
  if (!defaultCfg.option1_KidToSaint && !defaultCfg.option2_SaintToKid) {
    defaultCfg.option1_KidToSaint = true;
  }
  return defaultCfg;
}

/**
 * Writes App Config to both Live_Leaderboard_And_Kid_Stats (M2:N5) and ScriptProperties
 */
function writeAppConfigToSheet(ss, cfg) {
  const cleanCfg = {
    option1_KidToSaint: cfg.option1_KidToSaint !== false,
    option2_SaintToKid: cfg.option2_SaintToKid !== false,
    requireSecretPin: Boolean(cfg.requireSecretPin)
  };
  if (!cleanCfg.option1_KidToSaint && !cleanCfg.option2_SaintToKid) {
    cleanCfg.option1_KidToSaint = true;
  }
  try {
    PropertiesService.getScriptProperties().setProperty('APP_CONFIG_JSON', JSON.stringify(cleanCfg));
  } catch (e) {}

  let dashSheet = ss.getSheetByName(TAB_DASHBOARD);
  if (dashSheet) {
    dashSheet.getRange('M1').setValue('⚙ APP SETTINGS').setFontWeight('bold');
    dashSheet.getRange('M2:N2').setValues([['Setting_Key', 'Setting_Value']]).setFontWeight('bold').setBackground('#4a1c11').setFontColor('#fde68a');
    dashSheet.getRange('M3:N5').setValues([
      ['option1_KidToSaint', cleanCfg.option1_KidToSaint],
      ['option2_SaintToKid', cleanCfg.option2_SaintToKid],
      ['requireSecretPin', cleanCfg.requireSecretPin]
    ]);
  }
  return cleanCfg;
}

/**
 * 1. ONE-CLICK INITIALIZER & UPGRADER
 */
function setupSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Please open Apps Script from inside your Google Sheet (Extensions -> Apps Script).');
  }

  let rosterSheet = ss.getSheetByName(TAB_ROSTER);
  if (!rosterSheet) {
    rosterSheet = ss.insertSheet(TAB_ROSTER);
  }

  rosterSheet.getRange(1, 1, 1, ROSTER_HEADERS.length)
    .setValues([ROSTER_HEADERS])
    .setFontWeight('bold')
    .setBackground('#4a1c11')
    .setFontColor('#fde68a');
  rosterSheet.setFrozenRows(1);
  rosterSheet.autoResizeColumns(1, ROSTER_HEADERS.length);

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

  let dashSheet = ss.getSheetByName(TAB_DASHBOARD);
  if (!dashSheet) {
    dashSheet = ss.insertSheet(TAB_DASHBOARD);
  }
  const existingCfg = readAppConfigFromSheet(ss);
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

  writeAppConfigToSheet(ss, existingCfg);

  SpreadsheetApp.flush();
}

/**
 * 2. HTTP GET: Returns Roster + Leaderboard + Config OR processes GET-encoded write actions
 */
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action) {
      const action = String(e.parameter.action).trim();
      let payload = {};
      if (e.parameter.payload) {
        try {
          payload = JSON.parse(e.parameter.payload);
        } catch (pe) {
          payload = e.parameter;
        }
      } else {
        payload = e.parameter;
      }
      return handleWriteAction(action, payload);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return createJsonResponse({
        status: 'error',
        message: 'Google Sheet database not bound. Run setupSheetStructure from Apps Script.'
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
    const appConfig = readAppConfigFromSheet(ss);

    allScores.sort(function(a, b) {
      const scoreDiff = Number(b.Correct_Score || 0) - Number(a.Correct_Score || 0);
      if (scoreDiff !== 0) return scoreDiff;
      return Number(a.Time_Taken_Sec || 999999) - Number(b.Time_Taken_Sec || 999999);
    });

    return createJsonResponse({
      status: 'success',
      timestamp: new Date().toISOString(),
      config: appConfig,
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
 * 3. HTTP POST: Processes POST write actions
 */
function doPost(e) {
  try {
    let action = 'submitScore';
    let p = {};

    if (e && e.postData && e.postData.contents) {
      try {
        const body = JSON.parse(e.postData.contents);
        action = body.action || (e.parameter ? e.parameter.action : '') || 'submitScore';
        p = body.payload || body;
      } catch (parseErr) {
        action = (e.parameter ? e.parameter.action : '') || 'submitScore';
        p = e.parameter || {};
      }
    } else if (e && e.parameter) {
      action = e.parameter.action || 'submitScore';
      if (e.parameter.payload) {
        try {
          p = JSON.parse(e.parameter.payload);
        } catch (pe) {
          p = e.parameter;
        }
      } else {
        p = e.parameter;
      }
    }

    return handleWriteAction(action, p);
  } catch (err) {
    return createJsonResponse({
      status: 'error',
      message: err.toString()
    });
  }
}

/**
 * Core Write Handler (used by both doPost and doGet fallback)
 */
function handleWriteAction(action, p) {
  if (action === 'ping') {
    return createJsonResponse({
      status: 'success',
      pong: true,
      timestamp: new Date().toISOString()
    });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return createJsonResponse({
        status: 'error',
        message: 'Spreadsheet connection not found.'
      });
    }

    let rosterSheet = ss.getSheetByName(TAB_ROSTER);
    let scoresSheet = ss.getSheetByName(TAB_SCORES);
    if (!rosterSheet || !scoresSheet) {
      setupSheetStructure();
      rosterSheet = ss.getSheetByName(TAB_ROSTER);
      scoresSheet = ss.getSheetByName(TAB_SCORES);
    }

    // ACTION: SAVE APP CONFIG (Enable/Disable Saint Find Options)
    if (action === 'saveAppConfig') {
      const savedCfg = writeAppConfigToSheet(ss, p);
      SpreadsheetApp.flush();
      return createJsonResponse({
        status: 'success',
        action: 'saveAppConfig',
        config: savedCfg
      });
    }

    // ACTION: CLEAR ALL SCORES
    if (action === 'clearAllScores') {
      if (scoresSheet.getLastRow() > 1) {
        scoresSheet.getRange(2, 1, scoresSheet.getLastRow() - 1, 10).clearContent();
      }
      SpreadsheetApp.flush();
      return createJsonResponse({
        status: 'success',
        action: 'clearAllScores'
      });
    }

    rosterSheet.getRange(1, 1, 1, ROSTER_HEADERS.length).setValues([ROSTER_HEADERS]);

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

      if (rosterSheet.getLastRow() > 2) {
        rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, ROSTER_HEADERS.length).sort({ column: 1, ascending: true });
      }
      SpreadsheetApp.flush();

      return createJsonResponse({
        status: 'success',
        action: 'saveKidRow',
        bibNumber: bibNum
      });
    }

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
        bibNumber: bibToDelete
      });
    }

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

    // Default: submitScore (Optimized single-pass validation and rank calculation)
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

    const totalAttempted = Math.min(100, Math.max(1, Math.floor(Number(p.Total_Attempted || 1))));
    const maxPossiblePoints = totalAttempted * 10;
    const pointsScore = Math.min(maxPossiblePoints, Math.max(0, Math.floor(Number(p.Correct_Score || 0))));
    const accuracyPct = Math.min(100, Math.max(0, Number(Number(p.Accuracy_Percent || 0).toFixed(1))));
    const timeTakenSec = Math.min(86400, Math.max(1, Math.floor(Number(p.Time_Taken_Sec || 1))));

    const existingData = scoresSheet.getDataRange().getValues();
    let calculatedRank = 1;

    for (let i = 1; i < existingData.length; i++) {
      if (String(existingData[i][0]) === submissionId) {
        return createJsonResponse({
          status: 'success',
          duplicate: true,
          rank: calculateParticipantRank(existingData, pointsScore, timeTakenSec)
        });
      }
      if (normalizeParticipantName(existingData[i][2]) === normCandidateName) {
        return createJsonResponse({
          status: 'error',
          code: 'DUPLICATE_NAME',
          message: 'Participant name is already registered. Duplicate names are not allowed.'
        });
      }
      const rowScore = Number(existingData[i][5] || 0);
      const rowTime = Number(existingData[i][8] || 999999);
      if (rowScore > pointsScore || (rowScore === pointsScore && rowTime < timeTakenSec)) {
        calculatedRank++;
      }
    }

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
