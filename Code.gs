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
const TAB_NOTIFICATIONS = 'Parish_Notifications';
const TAB_BONUS_QUESTIONS = 'Bonus_Questions';
const TAB_BONUS_LOGS = 'Bonus_Question_Logs';

const NOTIFICATION_HEADERS = [
  'Notification_ID',
  'Timestamp',
  'Message',
  'Sender',
  'Status'
];

const BONUS_QUESTION_HEADERS = [
  'Question_ID',
  'Question_Text',
  'Option_A',
  'Option_B',
  'Option_C',
  'Option_D',
  'Correct_Answer',
  'Status'
];

const BONUS_LOG_HEADERS = [
  'Log_ID',
  'Timestamp',
  'Participant_Name',
  'Family_or_Phone',
  'Question_ID',
  'Question_Text',
  'Selected_Option',
  'Is_Correct',
  'Bonus_Points_Earned',
  'Cumulative_Bonus_Points'
];

const DEFAULT_BONUS_QUESTIONS = [
  {
    Question_ID: 'BONUS-001',
    Question_Text: 'In Catholic tradition, how many Holy Sacraments are instituted by Christ?',
    Option_A: '7 Sacraments',
    Option_B: '3 Sacraments',
    Option_C: '10 Sacraments',
    Option_D: '12 Sacraments',
    Correct_Answer: '7 Sacraments',
    Status: 'ACTIVE'
  },
  {
    Question_ID: 'BONUS-002',
    Question_Text: 'Which liturgical color is traditionally worn by priests during Ordinary Time?',
    Option_A: 'Green',
    Option_B: 'Red',
    Option_C: 'Purple',
    Option_D: 'White',
    Correct_Answer: 'Green',
    Status: 'ACTIVE'
  },
  {
    Question_ID: 'BONUS-003',
    Question_Text: 'What is the very first book of the Holy Bible?',
    Option_A: 'Genesis',
    Option_B: 'Exodus',
    Option_C: 'Matthew',
    Option_D: 'Psalms',
    Correct_Answer: 'Genesis',
    Status: 'ACTIVE'
  },
  {
    Question_ID: 'BONUS-004',
    Question_Text: 'What do the four candles on an Advent wreath symbolize?',
    Option_A: 'The four weeks of Advent waiting for Christ',
    Option_B: 'The four Gospel Evangelists',
    Option_C: 'The four seasons of the year',
    Option_D: 'The four cardinal virtues',
    Correct_Answer: 'The four weeks of Advent waiting for Christ',
    Status: 'ACTIVE'
  },
  {
    Question_ID: 'BONUS-005',
    Question_Text: 'Which prayer did Jesus Himself teach His disciples in the Gospels?',
    Option_A: 'The Lord\'s Prayer (Our Father)',
    Option_B: 'The Hail Mary',
    Option_C: 'The Apostles\' Creed',
    Option_D: 'The Glory Be',
    Correct_Answer: 'The Lord\'s Prayer (Our Father)',
    Status: 'ACTIVE'
  },
  {
    Question_ID: 'BONUS-006',
    Question_Text: 'On what date does the Catholic Church celebrate All Saints\' Day every year?',
    Option_A: 'November 1st',
    Option_B: 'October 31st',
    Option_C: 'December 25th',
    Option_D: 'January 1st',
    Correct_Answer: 'November 1st',
    Status: 'ACTIVE'
  },
  {
    Question_ID: 'BONUS-007',
    Question_Text: 'To which Apostle did Jesus say, "You are Peter, and upon this rock I will build my Church"?',
    Option_A: 'Saint Peter',
    Option_B: 'Saint Paul',
    Option_C: 'Saint John',
    Option_D: 'Saint Andrew',
    Correct_Answer: 'Saint Peter',
    Status: 'ACTIVE'
  },
  {
    Question_ID: 'BONUS-008',
    Question_Text: 'How many decades are prayed in one standard mystery of the Holy Rosary?',
    Option_A: '5 decades',
    Option_B: '3 decades',
    Option_C: '7 decades',
    Option_D: '10 decades',
    Correct_Answer: '5 decades',
    Status: 'ACTIVE'
  },
  {
    Question_ID: 'BONUS-009',
    Question_Text: 'Which liturgical season of prayer and fasting directly precedes Easter?',
    Option_A: 'Lent',
    Option_B: 'Advent',
    Option_C: 'Ordinary Time',
    Option_D: 'Christmastide',
    Correct_Answer: 'Lent',
    Status: 'ACTIVE'
  },
  {
    Question_ID: 'BONUS-010',
    Question_Text: 'What gift of the Holy Spirit helps us judge rightly between good and evil in God\'s light?',
    Option_A: 'Wisdom',
    Option_B: 'Courage',
    Option_C: 'Piety',
    Option_D: 'Knowledge',
    Correct_Answer: 'Wisdom',
    Status: 'ACTIVE'
  },
  {
    Question_ID: 'BONUS-011',
    Question_Text: 'Which Archangel greeted the Blessed Virgin Mary at the Annunciation?',
    Option_A: 'Archangel Gabriel',
    Option_B: 'Archangel Michael',
    Option_C: 'Archangel Raphael',
    Option_D: 'Archangel Uriel',
    Correct_Answer: 'Archangel Gabriel',
    Status: 'ACTIVE'
  },
  {
    Question_ID: 'BONUS-012',
    Question_Text: 'What is the sacred vessel in church where the Blessed Sacrament is reserved called?',
    Option_A: 'Tabernacle',
    Option_B: 'Altar',
    Option_C: 'Ambo',
    Option_D: 'Chalice',
    Correct_Answer: 'Tabernacle',
    Status: 'ACTIVE'
  },
  {
    Question_ID: 'BONUS-013',
    Question_Text: 'How many days did Jesus fast in the desert before beginning His public ministry?',
    Option_A: '40 days',
    Option_B: '12 days',
    Option_C: '7 days',
    Option_D: '3 days',
    Correct_Answer: '40 days',
    Status: 'ACTIVE'
  },
  {
    Question_ID: 'BONUS-014',
    Question_Text: 'Which of the four Gospel writers was a physician and companion of Saint Paul?',
    Option_A: 'Saint Luke',
    Option_B: 'Saint Matthew',
    Option_C: 'Saint Mark',
    Option_D: 'Saint John',
    Correct_Answer: 'Saint Luke',
    Status: 'ACTIVE'
  },
  {
    Question_ID: 'BONUS-015',
    Question_Text: 'What feast day immediately follows All Saints\' Day on November 2nd?',
    Option_A: 'All Souls\' Day',
    Option_B: 'Feast of Christ the King',
    Option_C: 'Epiphany',
    Option_D: 'Pentecost',
    Correct_Answer: 'All Souls\' Day',
    Status: 'ACTIVE'
  }
];

/**
 * Eastern Standard Time (EST/EDT) Helper for all sheet records
 */
const TIMEZONE_EST = 'America/New_York';

function formatEstDateForSheet(d) {
  try {
    return Utilities.formatDate(d || new Date(), TIMEZONE_EST, 'yyyy-MM-dd HH:mm:ss') + ' EST';
  } catch (e) {
    return Utilities.formatDate(d || new Date(), 'GMT-5', 'yyyy-MM-dd HH:mm:ss') + ' EST';
  }
}

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
    requireSecretPin: false,
    eventStatus: 'RUNNING',
    eventEndTime: 0,
    bonusQuestionsEnabled: true,
    bonusPointsPerQuestion: 5,
    themeStartupMode: 'dynamic_month',
    themeManualMonth: 10,
    eventDate: '2026-11-01'
  };
  try {
    const props = PropertiesService.getScriptProperties();
    const savedJson = props.getProperty('APP_CONFIG_JSON');
    if (savedJson) {
      const parsed = JSON.parse(savedJson);
      defaultCfg.option1_KidToSaint = parsed.option1_KidToSaint !== false;
      defaultCfg.option2_SaintToKid = parsed.option2_SaintToKid !== false;
      defaultCfg.requireSecretPin = Boolean(parsed.requireSecretPin);
      if (parsed.eventStatus) defaultCfg.eventStatus = String(parsed.eventStatus).toUpperCase();
      if (parsed.eventEndTime !== undefined) defaultCfg.eventEndTime = Number(parsed.eventEndTime) || 0;
      if (parsed.bonusQuestionsEnabled !== undefined) defaultCfg.bonusQuestionsEnabled = Boolean(parsed.bonusQuestionsEnabled);
      if (parsed.bonusPointsPerQuestion !== undefined) defaultCfg.bonusPointsPerQuestion = Number(parsed.bonusPointsPerQuestion) || 5;
      if (parsed.themeStartupMode) defaultCfg.themeStartupMode = String(parsed.themeStartupMode);
      if (parsed.themeManualMonth !== undefined) defaultCfg.themeManualMonth = Number(parsed.themeManualMonth);
      if (parsed.eventDate) defaultCfg.eventDate = String(parsed.eventDate);
    }
    const dashSheet = ss ? ss.getSheetByName(TAB_DASHBOARD) : null;
    if (dashSheet) {
      const cfgValues = dashSheet.getRange('M3:N12').getValues();
      if (cfgValues[0] && cfgValues[0][0] === 'option1_KidToSaint' && cfgValues[0][1] !== '') {
        defaultCfg.option1_KidToSaint = String(cfgValues[0][1]).toLowerCase() !== 'false';
      }
      if (cfgValues[1] && cfgValues[1][0] === 'option2_SaintToKid' && cfgValues[1][1] !== '') {
        defaultCfg.option2_SaintToKid = String(cfgValues[1][1]).toLowerCase() !== 'false';
      }
      if (cfgValues[2] && cfgValues[2][0] === 'requireSecretPin' && cfgValues[2][1] !== '') {
        defaultCfg.requireSecretPin = String(cfgValues[2][1]).toLowerCase() === 'true';
      }
      if (cfgValues[3] && cfgValues[3][0] === 'eventStatus' && cfgValues[3][1] !== '') {
        defaultCfg.eventStatus = String(cfgValues[3][1]).toUpperCase();
      }
      if (cfgValues[4] && cfgValues[4][0] === 'eventEndTime' && cfgValues[4][1] !== '') {
        defaultCfg.eventEndTime = Number(cfgValues[4][1]) || 0;
      }
      if (cfgValues[5] && cfgValues[5][0] === 'bonusQuestionsEnabled' && cfgValues[5][1] !== '') {
        defaultCfg.bonusQuestionsEnabled = String(cfgValues[5][1]).toLowerCase() !== 'false';
      }
      if (cfgValues[6] && cfgValues[6][0] === 'bonusPointsPerQuestion' && cfgValues[6][1] !== '') {
        defaultCfg.bonusPointsPerQuestion = Number(cfgValues[6][1]) || 5;
      }
      if (cfgValues[7] && cfgValues[7][0] === 'themeStartupMode' && cfgValues[7][1] !== '') {
        defaultCfg.themeStartupMode = String(cfgValues[7][1]);
      }
      if (cfgValues[8] && cfgValues[8][0] === 'themeManualMonth' && cfgValues[8][1] !== '') {
        defaultCfg.themeManualMonth = Number(cfgValues[8][1]);
      }
      if (cfgValues[9] && cfgValues[9][0] === 'eventDate' && cfgValues[9][1] !== '') {
        defaultCfg.eventDate = String(cfgValues[9][1]);
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
 * Writes App Config to both Live_Leaderboard_And_Kid_Stats (M2:N11) and ScriptProperties
 */
function writeAppConfigToSheet(ss, cfg) {
  const cleanCfg = {
    option1_KidToSaint: cfg.option1_KidToSaint !== false,
    option2_SaintToKid: cfg.option2_SaintToKid !== false,
    requireSecretPin: Boolean(cfg.requireSecretPin),
    eventStatus: cfg.eventStatus === 'STOPPED' ? 'STOPPED' : 'RUNNING',
    eventEndTime: cfg.eventEndTime ? Number(cfg.eventEndTime) : 0,
    bonusQuestionsEnabled: cfg.bonusQuestionsEnabled !== false,
    bonusPointsPerQuestion: cfg.bonusPointsPerQuestion ? Number(cfg.bonusPointsPerQuestion) : 5,
    themeStartupMode: cfg.themeStartupMode || 'dynamic_month',
    themeManualMonth: cfg.themeManualMonth !== undefined ? Number(cfg.themeManualMonth) : 10,
    eventDate: cfg.eventDate || '2026-11-01'
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
    dashSheet.getRange('M3:N12').setValues([
      ['option1_KidToSaint', cleanCfg.option1_KidToSaint],
      ['option2_SaintToKid', cleanCfg.option2_SaintToKid],
      ['requireSecretPin', cleanCfg.requireSecretPin],
      ['eventStatus', cleanCfg.eventStatus],
      ['eventEndTime', cleanCfg.eventEndTime],
      ['bonusQuestionsEnabled', cleanCfg.bonusQuestionsEnabled],
      ['bonusPointsPerQuestion', cleanCfg.bonusPointsPerQuestion],
      ['themeStartupMode', cleanCfg.themeStartupMode],
      ['themeManualMonth', cleanCfg.themeManualMonth],
      ['eventDate', cleanCfg.eventDate]
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
    'Matched_Bib_IDs',
    'Bonus_Points',
    'Total_Score'
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

  dashSheet.getRange('A1').setValue('🏆 TOP 10 PARTICIPANTS (BASE PTS + BONUS TRIVIA PTS)').setFontWeight('bold').setFontSize(12);
  const lbHeaders = ['Participant_Name', 'Family_or_Phone', 'Game_Mode', 'Base_Points', 'Bonus_Points', 'Total_Score', 'Accuracy_Percent', 'Time_Taken_Sec'];
  dashSheet.getRange(2, 1, 1, lbHeaders.length)
    .setValues([lbHeaders])
    .setFontWeight('bold')
    .setBackground('#6b2818')
    .setFontColor('#ffffff');

  dashSheet.getRange('A3').setFormula(
    '=IFERROR(QUERY(Quiz_Participants_Scores!A2:L, "SELECT C, D, E, F, K, L, H, I WHERE C IS NOT NULL ORDER BY L DESC, I ASC LIMIT 10", 0), "No submissions yet")'
  );

  dashSheet.getRange('J1').setValue('✞ 100-KID SAINT APPRECIATION COUNTER').setFontWeight('bold').setFontSize(12);
  const kidStatHeaders = ['Bib_Number', 'Kid_Name', 'Saint_Name', 'Times_Identified_By_Community'];
  dashSheet.getRange(2, 10, 1, kidStatHeaders.length)
    .setValues([kidStatHeaders])
    .setFontWeight('bold')
    .setBackground('#6b2818')
    .setFontColor('#fde68a');

  dashSheet.getRange('J3').setFormula('=ARRAYFORMULA(Roster_Saints_Kids!A2:C101)');

  const formulas = [];
  for (let r = 3; r <= 102; r++) {
    formulas.push([`=IF(J${r}="","",COUNTIF(ARRAYFORMULA(","&Quiz_Participants_Scores!$J$2:$J&","), "*,"&J${r}&",*"))`]);
  }
  dashSheet.getRange(3, 13, 100, 1).setFormulas(formulas);

  writeAppConfigToSheet(ss, existingCfg);

  let notifSheet = ss.getSheetByName(TAB_NOTIFICATIONS);
  if (!notifSheet) {
    notifSheet = ss.insertSheet(TAB_NOTIFICATIONS);
  }
  notifSheet.getRange(1, 1, 1, NOTIFICATION_HEADERS.length)
    .setValues([NOTIFICATION_HEADERS])
    .setFontWeight('bold')
    .setBackground('#4a1c11')
    .setFontColor('#fde68a');
  notifSheet.setFrozenRows(1);
  notifSheet.autoResizeColumns(1, NOTIFICATION_HEADERS.length);

  // Bonus Questions Bank Sheet
  let bonusSheet = ss.getSheetByName(TAB_BONUS_QUESTIONS);
  if (!bonusSheet) {
    bonusSheet = ss.insertSheet(TAB_BONUS_QUESTIONS);
  }
  bonusSheet.getRange(1, 1, 1, BONUS_QUESTION_HEADERS.length)
    .setValues([BONUS_QUESTION_HEADERS])
    .setFontWeight('bold')
    .setBackground('#4a1c11')
    .setFontColor('#fde68a');
  bonusSheet.setFrozenRows(1);

  if (bonusSheet.getLastRow() < 2) {
    const qRows = DEFAULT_BONUS_QUESTIONS.map(function(q) {
      return [
        q.Question_ID,
        q.Question_Text,
        q.Option_A,
        q.Option_B,
        q.Option_C,
        q.Option_D,
        q.Correct_Answer,
        q.Status
      ];
    });
    bonusSheet.getRange(2, 1, qRows.length, BONUS_QUESTION_HEADERS.length).setValues(qRows);
  }
  bonusSheet.autoResizeColumns(1, BONUS_QUESTION_HEADERS.length);

  // Bonus Question Logs Sheet
  let bonusLogsSheet = ss.getSheetByName(TAB_BONUS_LOGS);
  if (!bonusLogsSheet) {
    bonusLogsSheet = ss.insertSheet(TAB_BONUS_LOGS);
  }
  bonusLogsSheet.getRange(1, 1, 1, BONUS_LOG_HEADERS.length)
    .setValues([BONUS_LOG_HEADERS])
    .setFontWeight('bold')
    .setBackground('#4a1c11')
    .setFontColor('#fde68a');
  bonusLogsSheet.setFrozenRows(1);
  bonusLogsSheet.autoResizeColumns(1, BONUS_LOG_HEADERS.length);

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
      const scoreA = Number(a.Total_Score !== undefined && a.Total_Score !== '' ? a.Total_Score : (Number(a.Correct_Score || 0) + Number(a.Bonus_Points || 0)));
      const scoreB = Number(b.Total_Score !== undefined && b.Total_Score !== '' ? b.Total_Score : (Number(b.Correct_Score || 0) + Number(b.Bonus_Points || 0)));
      const scoreDiff = scoreB - scoreA;
      if (scoreDiff !== 0) return scoreDiff;
      return Number(a.Time_Taken_Sec || 999999) - Number(b.Time_Taken_Sec || 999999);
    });

    function isScoreOutOfDate(sub, officialEventDate) {
      const targetDate = String(officialEventDate || '2026-11-01').trim();
      const ts = String(sub.Timestamp || '');
      const match = ts.match(/^(\d{4}-\d{2}-\d{2})/);
      let subDate = '';
      if (match) {
        subDate = match[1];
      } else {
        const mdyMatch = ts.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (mdyMatch) {
          subDate = mdyMatch[3] + '-' + mdyMatch[1].padStart(2, '0') + '-' + mdyMatch[2].padStart(2, '0');
        }
      }
      if (!subDate) {
        const idMatch = String(sub.Submission_ID || '').match(/SUB-(\d{10,13})/i);
        if (idMatch) {
          const ep = Number(idMatch[1]);
          if (ep > 1600000000000 && ep < 2500000000000) {
            subDate = Utilities.formatDate(new Date(ep), TIMEZONE_EST, 'yyyy-MM-dd');
          }
        }
      }
      if (!subDate) return false;
      const allowed = targetDate.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
      return allowed.indexOf(subDate) === -1;
    }

    allScores.forEach(function(s) {
      s.isOutOfDate = isScoreOutOfDate(s, appConfig.eventDate);
    });

    let bonusQuestions = [];
    let bonusSheet = ss.getSheetByName(TAB_BONUS_QUESTIONS);
    if (bonusSheet && bonusSheet.getLastRow() > 1) {
      bonusQuestions = readSheetAsObjects(bonusSheet).filter(function(q) {
        return String(q.Status || 'ACTIVE').toUpperCase() !== 'INACTIVE';
      });
    }
    if (!bonusQuestions || bonusQuestions.length === 0) {
      bonusQuestions = DEFAULT_BONUS_QUESTIONS;
    }

    let latestNotification = null;
    try {
      const props = PropertiesService.getScriptProperties();
      const notifJson = props.getProperty('LATEST_NOTIFICATION_JSON');
      if (notifJson) {
        latestNotification = JSON.parse(notifJson);
      }
    } catch (ne) {}

    let notifSheet = ss.getSheetByName(TAB_NOTIFICATIONS);
    let allNotifications = [];
    if (notifSheet && notifSheet.getLastRow() > 1) {
      allNotifications = readSheetAsObjects(notifSheet);
      if (!latestNotification && allNotifications.length > 0) {
        latestNotification = allNotifications[allNotifications.length - 1];
      }
    }

    function sanitizeNotifEstTimestamp(notif) {
      if (!notif) return notif;
      const idMatch = String(notif.Notification_ID || '').match(/NOTIF-(\d{10,13})/i);
      if (idMatch) {
        const ep = Number(idMatch[1]);
        if (ep > 1600000000000 && ep < 2500000000000) {
          notif.Timestamp = formatEstDateForSheet(new Date(ep));
        }
      }
      return notif;
    }

    if (latestNotification) {
      latestNotification = sanitizeNotifEstTimestamp(latestNotification);
      try {
        PropertiesService.getScriptProperties().setProperty('LATEST_NOTIFICATION_JSON', JSON.stringify(latestNotification));
      } catch (e) {}
    }
    allNotifications = allNotifications.map(sanitizeNotifEstTimestamp);

    return createJsonResponse({
      status: 'success',
      timestamp: formatEstDateForSheet(new Date()),
      config: appConfig,
      roster: roster,
      leaderboard: allScores.slice(0, 50),
      bonusQuestions: bonusQuestions,
      latestNotification: latestNotification,
      notifications: allNotifications.slice(-15).reverse()
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
 * Core Write Handler (Optimized for High Concurrency & Parallel Submissions)
 */
function handleWriteAction(action, p) {
  if (action === 'ping') {
    return createJsonResponse({
      status: 'success',
      pong: true,
      timestamp: formatEstDateForSheet(new Date())
    });
  }

  // Pre-process and sanitize payload OUTSIDE the lock to minimize lock hold duration
  const submissionId = sanitizeCellForSheet(p.Submission_ID || ('SUB-' + new Date().getTime()), 50);
  const cleanName = sanitizeCellForSheet(p.Participant_Name || '', 60);
  const normCandidateName = normalizeParticipantName(cleanName);

  if (action === 'submitScore' && (!normCandidateName || normCandidateName.length < 2)) {
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

  const validBibs = String(p.Matched_Bib_IDs || '')
    .split(',')
    .map(function(s) { return parseInt(s.trim(), 10); })
    .filter(function(n) { return !isNaN(n) && n >= 1 && n <= 500; });

  let subDate = new Date();
  const subMatch = String(submissionId).match(/SUB-(\d{10,13})/i);
  if (subMatch) {
    const ep = Number(subMatch[1]);
    if (ep > 1600000000000 && ep < 2500000000000) {
      subDate = new Date(ep);
    }
  }
  const scoreTimestamp = (typeof p.Timestamp === 'string' && p.Timestamp.includes('EST'))
    ? sanitizeCellForSheet(p.Timestamp, 35)
    : formatEstDateForSheet(subDate);

  const bonusPoints = Math.max(0, Math.floor(Number(p.Bonus_Points || 0)));
  const totalScore = pointsScore + bonusPoints;

  const rowToAppend = [
    submissionId,
    scoreTimestamp,
    cleanName,
    sanitizeCellForSheet(p.Family_or_Phone || '', 60),
    sanitizeCellForSheet(p.Game_Mode || 'Find Saint by Kid Name', 40),
    pointsScore,
    totalAttempted,
    accuracyPct,
    timeTakenSec,
    validBibs.join(','),
    bonusPoints,
    totalScore
  ];

  // Up to 30-second queue wait for high-concurrency bursts (e.g. 100+ simultaneous submissions)
  const lock = LockService.getScriptLock();
  try {
    const gotLock = lock.tryLock(30000);
    if (!gotLock) {
      return createJsonResponse({
        status: 'busy',
        code: 'SERVER_BUSY',
        message: 'High simultaneous submission traffic. Automatic retry scheduled.'
      });
    }

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
      lock.releaseLock(); // Release lock immediately
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
      lock.releaseLock(); // Release lock immediately
      return createJsonResponse({
        status: 'success',
        action: 'clearAllScores'
      });
    }

    // ACTION: SEND BROADCAST NOTIFICATION (Store permanently in Google Sheet + broadcast to all devices)
    if (action === 'sendNotification') {
      let notifSheet = ss.getSheetByName(TAB_NOTIFICATIONS);
      if (!notifSheet) {
        notifSheet = ss.insertSheet(TAB_NOTIFICATIONS);
        notifSheet.getRange(1, 1, 1, NOTIFICATION_HEADERS.length)
          .setValues([NOTIFICATION_HEADERS])
          .setFontWeight('bold')
          .setBackground('#4a1c11')
          .setFontColor('#fde68a');
        notifSheet.setFrozenRows(1);
      }

      const notifId = sanitizeCellForSheet(p.Notification_ID || ('NOTIF-' + new Date().getTime() + '-' + Math.random().toString(36).substring(2, 6).toUpperCase()), 50);

      let notifDate = new Date();
      const notifMatch = String(notifId).match(/NOTIF-(\d{10,13})/i);
      if (notifMatch) {
        const ep = Number(notifMatch[1]);
        if (ep > 1600000000000 && ep < 2500000000000) {
          notifDate = new Date(ep);
        }
      }
      const timestamp = formatEstDateForSheet(notifDate);
      const message = sanitizeCellForSheet(p.Message || '', 500);
      const sender = sanitizeCellForSheet(p.Sender || 'Parish Admin', 60);
      const status = 'ACTIVE';

      const notifObj = {
        Notification_ID: notifId,
        Timestamp: timestamp,
        Message: message,
        Sender: sender,
        Status: status
      };

      notifSheet.appendRow([notifId, timestamp, message, sender, status]);

      try {
        PropertiesService.getScriptProperties().setProperty('LATEST_NOTIFICATION_JSON', JSON.stringify(notifObj));
      } catch (e) {}

      lock.releaseLock();
      return createJsonResponse({
        status: 'success',
        action: 'sendNotification',
        notification: notifObj
      });
    }

    // ACTION: CLEAR / EXPIRE ACTIVE NOTIFICATION
    if (action === 'clearNotification') {
      try {
        PropertiesService.getScriptProperties().deleteProperty('LATEST_NOTIFICATION_JSON');
      } catch (e) {}

      let notifSheet = ss.getSheetByName(TAB_NOTIFICATIONS);
      if (notifSheet && notifSheet.getLastRow() > 1) {
        notifSheet.getRange(notifSheet.getLastRow(), 5).setValue('EXPIRED');
      }

      lock.releaseLock();
      return createJsonResponse({
        status: 'success',
        action: 'clearNotification'
      });
    }

    // ACTION: RECORD USER'S BONUS QUESTION ANSWER (Real-time tracking to Google Sheet)
    if (action === 'recordBonusAnswer') {
      let bonusLogsSheet = ss.getSheetByName(TAB_BONUS_LOGS);
      if (!bonusLogsSheet) {
        bonusLogsSheet = ss.insertSheet(TAB_BONUS_LOGS);
        bonusLogsSheet.getRange(1, 1, 1, BONUS_LOG_HEADERS.length)
          .setValues([BONUS_LOG_HEADERS])
          .setFontWeight('bold')
          .setBackground('#4a1c11')
          .setFontColor('#fde68a');
        bonusLogsSheet.setFrozenRows(1);
      }

      const logId = sanitizeCellForSheet(p.Log_ID || ('BLOG-' + Date.now() + '-' + Math.random().toString(36).substring(2, 6).toUpperCase()), 40);
      const timestamp = formatEstDateForSheet(new Date());
      const participantName = sanitizeCellForSheet(p.Participant_Name || 'Anonymous', 50);
      const familyOrPhone = sanitizeCellForSheet(p.Family_or_Phone || '', 50);
      const questionId = sanitizeCellForSheet(p.Question_ID || '', 40);
      const questionText = sanitizeCellForSheet(p.Question_Text || '', 200);
      const selectedOption = sanitizeCellForSheet(p.Selected_Option || '', 100);
      const isCorrect = p.Is_Correct === true || String(p.Is_Correct).toLowerCase() === 'true';
      const pointsEarned = Number(p.Bonus_Points_Earned || 0);
      const cumulativeBonus = Number(p.Cumulative_Bonus_Points || 0);

      bonusLogsSheet.appendRow([
        logId,
        timestamp,
        participantName,
        familyOrPhone,
        questionId,
        questionText,
        selectedOption,
        isCorrect ? 'CORRECT' : 'INCORRECT',
        pointsEarned,
        cumulativeBonus
      ]);

      lock.releaseLock();
      return createJsonResponse({
        status: 'success',
        action: 'recordBonusAnswer',
        logId: logId
      });
    }

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
      lock.releaseLock();

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
      lock.releaseLock();
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
      lock.releaseLock();

      return createJsonResponse({
        status: 'success',
        action: 'syncFullRoster',
        count: rows.length
      });
    }

    // Default: submitScore
    // Verify Event is open for official score submission (grace period of 75 seconds allowed for active players auto-submitting)
    const currentConfig = readAppConfigFromSheet(ss);
    const nowMs = Date.now();
    const hasGracePeriod = currentConfig.eventEndTime > 0 && nowMs <= (currentConfig.eventEndTime + 75000);
    const isEventClosed = (currentConfig.eventStatus === 'STOPPED' && !hasGracePeriod) || (currentConfig.eventEndTime > 0 && nowMs > (currentConfig.eventEndTime + 75000));
    if (isEventClosed) {
      lock.releaseLock();
      return createJsonResponse({
        status: 'error',
        code: 'EVENT_CLOSED',
        message: 'The parish event is closed. Official score submissions are disabled. Users may play in Practice & Learn Mode.'
      });
    }

    // Ultra-fast critical section: Read only Columns A to C (Submission_ID, Timestamp, Participant_Name)
    const lastRow = scoresSheet.getLastRow();
    if (lastRow > 1) {
      const idAndNames = scoresSheet.getRange(2, 1, lastRow - 1, 3).getValues();
      for (let i = 0; i < idAndNames.length; i++) {
        if (String(idAndNames[i][0]) === submissionId) {
          lock.releaseLock();
          return createJsonResponse({
            status: 'success',
            duplicate: true,
            submissionId: submissionId
          });
        }
        if (normalizeParticipantName(idAndNames[i][2]) === normCandidateName) {
          lock.releaseLock();
          return createJsonResponse({
            status: 'error',
            code: 'DUPLICATE_NAME',
            message: 'Participant name is already registered. Duplicate names are not allowed.'
          });
        }
      }
    }

    // Append row atomically
    scoresSheet.appendRow(rowToAppend);

    // Release lock IMMEDIATELY so parallel users waiting in queue acquire lock instantly (<40ms hold time)
    lock.releaseLock();

    return createJsonResponse({
      status: 'success',
      submissionId: submissionId,
      correctScore: pointsScore,
      bonusPoints: bonusPoints,
      totalScore: totalScore,
      timeTaken: timeTakenSec
    });
  } catch (err) {
    try { lock.releaseLock(); } catch (e) {}
    return createJsonResponse({
      status: 'error',
      message: err.toString()
    });
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
      let cellVal = values[r][c];
      if (cellVal instanceof Date) {
        cellVal = formatEstDateForSheet(cellVal);
      }
      rowObj[headers[c]] = cellVal;
      if (cellVal !== '' && cellVal !== null && cellVal !== undefined) {
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
    const rowScore = sheetValues[i][11] !== undefined && sheetValues[i][11] !== ''
      ? Number(sheetValues[i][11] || 0)
      : (Number(sheetValues[i][5] || 0) + Number(sheetValues[i][10] || 0));
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
