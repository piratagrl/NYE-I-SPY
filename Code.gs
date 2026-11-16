/**
 * =======================================================================
 * Google Apps Script Backend for "Midnight Sparkle: Office I Spy"
 * Database: Sheet 2 named 'NYE High Scores'
 * =======================================================================
 */

const SHEET_NAME = 'NYE High Scores';

/**
 * Serves the web app UI when someone visits the deployed URL
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Midnight Sparkle: New Year's Office I Spy")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Locates or creates Sheet 2 ('NYE High Scores') in the active spreadsheet.
 * Automatically initializes columns and headers on first run.
 */
function getHighScoresSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  // If not found by name, check if sheet at index 1 (Sheet 2) exists
  if (!sheet) {
    const allSheets = ss.getSheets();
    if (allSheets.length >= 2) {
      sheet = allSheets[1];
      sheet.setName(SHEET_NAME);
    } else {
      sheet = ss.insertSheet(SHEET_NAME, 2);
    }
  }
  
  // Initialize Header Row if empty
  if (sheet.getLastRow() === 0) {
    const headers = [
      'Entry ID',
      'Timestamp',
      'Coworker Name',
      'Office (DTW / PHX)',
      'High Score',
      'Time (Seconds)',
      'Items Found',
      'Difficulty',
      'Puzzle Flavor',
      'Display Date'
    ];
    sheet.appendRow(headers);
    
    // Style the header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#0F172A'); // Dark navy
    headerRange.setFontColor('#F59E0B');  // Festive amber
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(10);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
  
  return sheet;
}

/**
 * Returns all leaderboard records from Sheet 2 sorted descending by score.
 * Called automatically from client-side via:
 *   google.script.run.getSheetLeaderboard()
 */
function getSheetLeaderboard() {
  const sheet = getHighScoresSheet();
  const lastRow = sheet.getLastRow();
  
  // If only header row exists
  if (lastRow <= 1) {
    return [];
  }
  
  const range = sheet.getRange(2, 1, lastRow - 1, 10);
  const rows = range.getValues();
  
  const entries = rows.map(function(row) {
    return {
      id: String(row[0] || 'id-' + Math.random().toString(36).substring(2, 7)),
      name: String(row[2] || 'Anonymous'),
      team: String(row[3] || 'DTW'),
      score: Number(row[4] || 0),
      timeSeconds: Number(row[5] || 0),
      itemsFound: Number(row[6] || 0),
      difficulty: String(row[7] || 'challenger'),
      themeId: String(row[8] || 'times-square'),
      date: String(row[9] || 'Dec 31, 2026')
    };
  });
  
  // Sort by highest score first
  entries.sort(function(a, b) {
    return b.score - a.score;
  });
  
  return entries;
}

/**
 * Appends a new game result to Sheet 2 ('NYE High Scores').
 * Called automatically from client-side via:
 *   google.script.run.recordSheetScore(entry)
 */
function recordSheetScore(entry) {
  const sheet = getHighScoresSheet();
  const timestamp = new Date().toISOString();
  
  sheet.appendRow([
    entry.id || ('entry-' + new Date().getTime()),
    timestamp,
    entry.name || 'Anonymous',
    entry.team || 'DTW',
    Number(entry.score || 0),
    Number(entry.timeSeconds || 0),
    Number(entry.itemsFound || 0),
    entry.difficulty || 'challenger',
    entry.themeId || 'times-square',
    entry.date || 'Just now'
  ]);
  
  return { success: true };
}

/**
 * Quick Test function:
 * In Apps Script editor, click "Run" -> select "testSheetConnection" to verify Sheet 2!
 */
function testSheetConnection() {
  const sheet = getHighScoresSheet();
  Logger.log("Connected to Sheet: " + sheet.getName());
  Logger.log("Total Rows: " + sheet.getLastRow());
  
  // Add a sample test score
  recordSheetScore({
    id: 'test-1',
    name: 'Bill',
    team: 'PHX',
    score: 5500,
    timeSeconds: 65,
    itemsFound: 7,
    difficulty: 'challenger',
    themeId: 'times-square',
    date: 'Dec 31, 2026'
  });
  
  const scores = getSheetLeaderboard();
  Logger.log("Retrieved scores count: " + scores.length);
}
