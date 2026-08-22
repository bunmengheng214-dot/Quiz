const { google } = require('googleapis');

const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID;
const SERVICE_ACCOUNT_EMAIL = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
const SERVICE_ACCOUNT_PRIVATE_KEY = (process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY || '').replace(/\\n/g, '\n');

function getAuth() {
  if (!SERVICE_ACCOUNT_EMAIL || !SERVICE_ACCOUNT_PRIVATE_KEY || !SPREADSHEET_ID) {
    throw new Error('Missing Google Sheet environment variables.');
  }

  return new google.auth.GoogleAuth({
    credentials: {
      client_email: SERVICE_ACCOUNT_EMAIL,
      private_key: SERVICE_ACCOUNT_PRIVATE_KEY,
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

function normalizeRow(row = []) {
  return row.map((value) => (value === undefined ? '' : value));
}

function toObject(rows, headers) {
  return rows.map((row) => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = row[index] || '';
    });
    return item;
  });
}

async function getSheetValues(range) {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  const response = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range });
  return response.data.values || [];
}

async function appendSheetRow(range, row) {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [normalizeRow(row)] },
  });
}

async function saveScore(req, res) {
  try {
    const { studentName, exerciseTitle, score, studentClass } = req.body || {};
    const values = await getSheetValues('Scores!A1:E');
    if (!values.length || values[0][0] !== 'Timestamp') {
      await appendSheetRow('Scores!A1:E1', ['Timestamp', 'Student Name', 'Class', 'Exercise', 'Score']);
    }
    await appendSheetRow('Scores!A:E', [new Date().toISOString(), String(studentName || '').trim(), String(studentClass || '').trim(), String(exerciseTitle || '').trim(), Number(score) || 0]);
    res.status(200).json({ ok: true, message: 'Score saved to Google Sheet.' });
  } catch (error) {
    res.status(500).json({ ok: false, message: error.message || 'Unable to save score.' });
  }
}

async function getScores(req, res) {
  try {
    const values = await getSheetValues('Scores!A1:E');
    if (!values.length || values[0][0] !== 'Timestamp') {
      res.status(200).json({ ok: true, scores: [] });
      return;
    }
    const headers = values[0];
    const rows = values.slice(1);
    res.status(200).json({ ok: true, scores: toObject(rows, headers) });
  } catch (error) {
    res.status(500).json({ ok: false, message: error.message || 'Unable to fetch scores.' });
  }
}

async function checkTeacherPassword(req, res) {
  try {
    const { password } = req.body || {};
    const values = await getSheetValues('Settings!A1:B');
    const passwordRow = (values || []).find((row) => String(row[0] || '').trim() === 'teacher_password');
    const expected = passwordRow && passwordRow[1] ? String(passwordRow[1]) : '123456';
    res.status(200).json({ ok: true, valid: String(password || '') === expected });
  } catch (error) {
    res.status(500).json({ ok: false, message: error.message || 'Unable to check teacher password.' });
  }
}

async function getTeacherPassword(req, res) {
  try {
    const values = await getSheetValues('Settings!A1:B');
    const passwordRow = (values || []).find((row) => String(row[0] || '').trim() === 'teacher_password');
    const password = passwordRow && passwordRow[1] ? String(passwordRow[1]) : '123456';
    res.status(200).json({ ok: true, password });
  } catch (error) {
    res.status(500).json({ ok: false, message: error.message || 'Unable to fetch teacher password.' });
  }
}

async function saveTeacherPassword(req, res) {
  try {
    const { password } = req.body || {};
    const values = await getSheetValues('Settings!A1:B');
    const rowIndex = (values || []).findIndex((row) => String(row[0] || '').trim() === 'teacher_password');
    const nextPassword = String(password || '').trim() || '123456';
    if (rowIndex >= 0) {
      const auth = await getAuth();
      const sheets = google.sheets({ version: 'v4', auth });
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `Settings!B${rowIndex + 1}`,
        valueInputOption: 'RAW',
        requestBody: { values: [[nextPassword]] },
      });
    } else {
      await appendSheetRow('Settings!A:B', ['teacher_password', nextPassword]);
    }
    res.status(200).json({ ok: true, password: nextPassword });
  } catch (error) {
    res.status(500).json({ ok: false, message: error.message || 'Unable to save teacher password.' });
  }
}

async function getExercises(req, res) {
  try {
    const values = await getSheetValues('AppData!A1:A');
    if (!values.length || !values[0][0]) {
      res.status(200).json({ ok: true, exercises: [] });
      return;
    }
    const parsed = JSON.parse(String(values[0][0]));
    res.status(200).json({ ok: true, exercises: Array.isArray(parsed) ? parsed : [] });
  } catch (error) {
    res.status(200).json({ ok: true, exercises: [] });
  }
}

async function saveExercises(req, res) {
  try {
    const { exercises } = req.body || {};
    const auth = await getAuth();
    const sheets = google.sheets({ version: 'v4', auth });
    await sheets.spreadsheets.values.clear({ spreadsheetId: SPREADSHEET_ID, range: 'AppData!A1:A' });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: 'AppData!A1',
      valueInputOption: 'RAW',
      requestBody: { values: [[JSON.stringify(Array.isArray(exercises) ? exercises : [])]] },
    });
    res.status(200).json({ ok: true, exercises: Array.isArray(exercises) ? exercises : [] });
  } catch (error) {
    res.status(500).json({ ok: false, message: error.message || 'Unable to save exercises.' });
  }
}

async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  const action = req.body && req.body.action ? req.body.action : (req.query && req.query.action ? req.query.action : 'getScores');

  switch (action) {
    case 'saveScore':
      return saveScore(req, res);
    case 'getScores':
      return getScores(req, res);
    case 'checkTeacherPassword':
      return checkTeacherPassword(req, res);
    case 'getTeacherPassword':
      return getTeacherPassword(req, res);
    case 'saveTeacherPassword':
      return saveTeacherPassword(req, res);
    case 'getExercises':
      return getExercises(req, res);
    case 'saveExercises':
      return saveExercises(req, res);
    default:
      return res.status(400).json({ ok: false, message: 'Unsupported action.' });
  }
}

module.exports = {
  saveScore,
  getScores,
  checkTeacherPassword,
  getTeacherPassword,
  saveTeacherPassword,
  getExercises,
  saveExercises,
  default: handler,
};
