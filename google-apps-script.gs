// Paste this script into Google Apps Script attached to the spreadsheet:
// https://docs.google.com/spreadsheets/d/1y2gspAeawDir9VmmSHi_rqaIvKor9qWXQg_hFvwSdIU/edit?usp=sharing
//
// Then deploy as a Web App and use the web app URL in your frontend if needed.

const SPREADSHEET_ID = '1y2gspAeawDir9VmmSHi_rqaIvKor9qWXQg_hFvwSdIU';
const SCORE_SHEET_NAME = 'Scores';
const APP_DATA_SHEET_NAME = 'AppData';
const SETTINGS_SHEET_NAME = 'Settings';
const DEFAULT_TEACHER_PASSWORD = '123456';

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getOrCreateSheet_(name) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function ensureScoreHeaders_(sheet) {
  const headers = ['Timestamp', 'Student Name', 'Class', 'Exercise', 'Score'];
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeader = !firstRow || firstRow[0] !== headers[0];
  if (needsHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function getSettingsValue_(key, fallbackValue) {
  const settingsSheet = getOrCreateSheet_(SETTINGS_SHEET_NAME);
  const values = settingsSheet.getDataRange().getValues();
  for (let i = 0; i < values.length; i += 1) {
    if (String(values[i][0] || '').trim() === String(key)) {
      return values[i][1] !== undefined ? values[i][1] : fallbackValue;
    }
  }
  return fallbackValue;
}

function setSettingsValue_(key, value) {
  const settingsSheet = getOrCreateSheet_(SETTINGS_SHEET_NAME);
  const values = settingsSheet.getDataRange().getValues();
  let found = false;

  for (let i = 0; i < values.length; i += 1) {
    if (String(values[i][0] || '').trim() === String(key)) {
      settingsSheet.getRange(i + 1, 2).setValue(value);
      found = true;
      break;
    }
  }

  if (!found) {
    settingsSheet.appendRow([key, value]);
  }
}

function defaultExercisesSeed_() {
  return [
    {
      id: 'ex_seed_1',
      title: 'លំហាត់ទី 1',
      questions: [
        { id: 'q1', text: '១. ត្រីកោណកែងមួយមានអ៊ីប៉ូតេនុសប្រវែង 10cm និងមុំមួយ 30°។ តើជ្រុងឈមនៃមុំនោះមានប្រវែងប៉ុន្មាន?', options: { A: '5cm', B: '10cm', C: '8.66cm', D: '20cm' }, correct: 'A', optionPoints: { A: 1, B: 0, C: 0, D: 0 } },
        { id: 'q2', text: '២. ត្រីកោណកែងមួយមានជ្រុងជាប់មុំ 45° ប្រវែង 8cm។ គណនាប្រវែងជ្រុងឈមនៃមុំនោះ។', options: { A: '4cm', B: '8cm', C: '16cm', D: '11.3cm' }, correct: 'B', optionPoints: { A: 0, B: 1, C: 0, D: 0 } },
        { id: 'q3', text: '៣. ត្រីកោណកែងមួយមានជrokkenឈមប្រវែង 6cm និងអ៊ីប៉ូតេនុសប្រវែង 12cm។ គណនារង្វាស់មុំឈមនឹងជ្រុងនោះ។', options: { A: '30°', B: '45°', C: '60°', D: '90°' }, correct: 'A', optionPoints: { A: 1, B: 0, C: 0, D: 0 } },
        { id: 'q4', text: '៤. ត្រីកោណកែង ABC កែងត្រង់ A មានជ្រុងជាប់ AB = 3cm និងជ្រុងឈម AC = 4cm (ធៀបនឹងមុំ B)។ តើតម្លៃនៃ tan(B) ស្មើនឹងប៉ុន្មាន?', options: { A: '3/4', B: '4/3', C: '3/5', D: '4/5' }, correct: 'B', optionPoints: { A: 0, B: 1, C: 0, D: 0 } },
        { id: 'q5', text: '៥. ត្រីកោណកែងមួយមានមុំមួយ 60° និងជ្រុងឈមមានប្រវែង 10√3 cm។ រកប្រវែងជ្រុងជាប់។', options: { A: '10 cm', B: '20 cm', C: '5 cm', D: '10√2 cm' }, correct: 'A', optionPoints: { A: 1, B: 0, C: 0, D: 0 } }
      ],
      createdAt: new Date().toISOString(),
      targetClass: 'all'
    }
  ];
}

function getExercisesServer() {
  const appDataSheet = getOrCreateSheet_(APP_DATA_SHEET_NAME);
  const firstCell = appDataSheet.getRange(1, 1).getValue();

  if (!firstCell) {
    const seed = defaultExercisesSeed_();
    appDataSheet.getRange(1, 1).setValue(JSON.stringify(seed));
    return { ok: true, exercises: seed };
  }

  try {
    const parsed = JSON.parse(String(firstCell));
    return { ok: true, exercises: Array.isArray(parsed) ? parsed : defaultExercisesSeed_() };
  } catch (err) {
    return { ok: true, exercises: defaultExercisesSeed_() };
  }
}

function saveExercisesServer(exercises) {
  const appDataSheet = getOrCreateSheet_(APP_DATA_SHEET_NAME);
  const safeExercises = Array.isArray(exercises) ? exercises : [];
  appDataSheet.clearContents();
  appDataSheet.getRange(1, 1).setValue(JSON.stringify(safeExercises));
  return { ok: true, exercises: safeExercises };
}

function getTeacherPasswordServer() {
  const password = getSettingsValue_('teacher_password', DEFAULT_TEACHER_PASSWORD);
  return { ok: true, password: String(password || DEFAULT_TEACHER_PASSWORD) };
}

function saveTeacherPasswordServer(password) {
  const value = String(password || '').trim();
  setSettingsValue_('teacher_password', value || DEFAULT_TEACHER_PASSWORD);
  return { ok: true, password: value || DEFAULT_TEACHER_PASSWORD };
}

function saveScoreServer(studentName, exerciseTitle, score, studentClass) {
  const sheet = getOrCreateSheet_(SCORE_SHEET_NAME);
  ensureScoreHeaders_(sheet);

  const row = [
    new Date(),
    String(studentName || '').trim(),
    String(studentClass || '').trim(),
    String(exerciseTitle || '').trim(),
    Number(score) || 0
  ];

  sheet.appendRow(row);
  return { ok: true, message: 'Score saved to Google Sheet.' };
}

function checkTeacherPassword(password) {
  return String(password || '') === String(DEFAULT_TEACHER_PASSWORD);
}

function getScoresServer() {
  const sheet = getOrCreateSheet_(SCORE_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return { ok: true, scores: [] };
  }

  const headers = values[0];
  const rows = values.slice(1).map((row) => {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = row[index] || '';
    });
    return record;
  });

  return { ok: true, scores: rows };
}
