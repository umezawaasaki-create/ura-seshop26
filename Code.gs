/**
 * 世小P裏 URA SSP 2026 - トーナメント対戦管理アプリ
 * Google Apps Script (バインドスクリプト) + Google スプレッドシート
 */

var TEAMS_SHEET = 'Teams';
var MATCHES_SHEET = 'Matches';

// logo_url is optional: leave blank in the sheet to keep the colored-initial
// badge, or paste an image URL to use that team's own logo instead.
// kana is the team name's reading (hiragana), used only to sort the team
// list (hamburger menu) in real a-i-u-e-o order - kanji alone can't be
// sorted phonetically. Leave blank to fall back to sorting by name as-is.
var TEAM_HEADERS = ['team_id', 'name', 'initial', 'color', 'logo_url', 'kana'];

var MATCH_HEADERS = [
  'match_id', 'round', 'order', 'team1_id', 'team2_id',
  'date', 'time', 'venue',
  'score1', 'score2', 'winner_id', 'status', 'win_type',
  'next_match_id', 'next_slot', 'loser_next_match_id', 'loser_next_slot'
];

// ---------------------------------------------------------------------------
// Web App entry point
// ---------------------------------------------------------------------------

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('URA SSP 2026 トーナメント管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------------------------------------------------------------------------
// Sheet helpers
// ---------------------------------------------------------------------------

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name, headers) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sheetToObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  return values.slice(1)
    .filter(function (row) { return row.some(function (cell) { return cell !== ''; }); })
    .map(function (row) {
      var obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });
      return obj;
    });
}

function findRowIndexByMatchId_(sheet, matchId) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === matchId) return i + 1; // 1-based row number
  }
  return -1;
}

function findRowIndexByTeamId_(sheet, teamId) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === teamId) return i + 1; // 1-based row number
  }
  return -1;
}

function formatDateIfNeeded_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value;
}

// ---------------------------------------------------------------------------
// Public API (called from client via google.script.run)
// ---------------------------------------------------------------------------

function getBracketData() {
  var teamsSheet = getSheet_(TEAMS_SHEET, TEAM_HEADERS);
  var matchesSheet = getSheet_(MATCHES_SHEET, MATCH_HEADERS);

  var teams = sheetToObjects_(teamsSheet);
  var matches = sheetToObjects_(matchesSheet).map(function (m) {
    m.date = formatDateIfNeeded_(m.date);
    return m;
  });

  return { teams: teams, matches: matches };
}

function saveSchedule(matchId, date, time, venue) {
  var sheet = getSheet_(MATCHES_SHEET, MATCH_HEADERS);
  var rowIdx = findRowIndexByMatchId_(sheet, matchId);
  if (rowIdx === -1) throw new Error('試合が見つかりません: ' + matchId);

  setCell_(sheet, rowIdx, 'date', date || '');
  setCell_(sheet, rowIdx, 'time', time || '');
  setCell_(sheet, rowIdx, 'venue', venue || '');

  return getBracketData();
}

/**
 * チームアイコン画像をGoogleドライブに保存し、Teamsシートのlogo_urlに反映する。
 * dataUrl: クライアントでFileReaderにより読み込んだ "data:image/xxx;base64,...." 形式の文字列
 */
function uploadTeamLogo(teamId, dataUrl, fileName) {
  var sheet = getSheet_(TEAMS_SHEET, TEAM_HEADERS);
  var rowIdx = findRowIndexByTeamId_(sheet, teamId);
  if (rowIdx === -1) throw new Error('チームが見つかりません: ' + teamId);

  var match = /^data:(.+?);base64,(.*)$/.exec(dataUrl || '');
  if (!match) throw new Error('画像データの形式が正しくありません。');
  var mimeType = match[1];
  var base64Data = match[2];

  var bytes = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(bytes, mimeType, fileName || (teamId + '_logo'));

  var folder = getTeamLogoFolder_();
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var url = 'https://drive.google.com/uc?export=view&id=' + file.getId();
  var colIdx = TEAM_HEADERS.indexOf('logo_url') + 1;
  sheet.getRange(rowIdx, colIdx).setValue(url);

  return getBracketData();
}

/**
 * チームアイコンを未設定に戻す（色付きイニシャルのバッジ表示に戻る）
 */
function clearTeamLogo(teamId) {
  var sheet = getSheet_(TEAMS_SHEET, TEAM_HEADERS);
  var rowIdx = findRowIndexByTeamId_(sheet, teamId);
  if (rowIdx === -1) throw new Error('チームが見つかりません: ' + teamId);

  var colIdx = TEAM_HEADERS.indexOf('logo_url') + 1;
  sheet.getRange(rowIdx, colIdx).setValue('');

  return getBracketData();
}

function getTeamLogoFolder_() {
  var folderName = 'URA_SSP_2026_チームアイコン';
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

/**
 * winType: 'normal' | 'fusen1' (team1不戦勝) | 'fusen2' (team2不戦勝)
 */
function saveResult(matchId, score1, score2, winType) {
  var sheet = getSheet_(MATCHES_SHEET, MATCH_HEADERS);
  var rowIdx = findRowIndexByMatchId_(sheet, matchId);
  if (rowIdx === -1) throw new Error('試合が見つかりません: ' + matchId);

  var rowValues = sheet.getRange(rowIdx, 1, 1, MATCH_HEADERS.length).getValues()[0];
  var match = {};
  MATCH_HEADERS.forEach(function (h, i) { match[h] = rowValues[i]; });

  if (!match.team1_id || !match.team2_id) {
    throw new Error('対戦チームがまだ確定していないため、結果を入力できません。');
  }

  var winnerId, loserId, s1, s2;

  if (winType === 'fusen1' || winType === 'fusen2') {
    winnerId = (winType === 'fusen1') ? match.team1_id : match.team2_id;
    loserId = (winType === 'fusen1') ? match.team2_id : match.team1_id;
    s1 = '';
    s2 = '';
  } else {
    s1 = Number(score1);
    s2 = Number(score2);
    if (isNaN(s1) || isNaN(s2)) throw new Error('得点は数値で入力してください。');
    if (s1 === s2) throw new Error('同点は保存できません。勝敗が決まってから入力してください。');
    winnerId = (s1 > s2) ? match.team1_id : match.team2_id;
    loserId = (s1 > s2) ? match.team2_id : match.team1_id;
    winType = 'normal';
  }

  setCell_(sheet, rowIdx, 'score1', s1);
  setCell_(sheet, rowIdx, 'score2', s2);
  setCell_(sheet, rowIdx, 'winner_id', winnerId);
  setCell_(sheet, rowIdx, 'status', 'completed');
  setCell_(sheet, rowIdx, 'win_type', winType);

  if (match.next_match_id) {
    advanceTeam_(match.next_match_id, match.next_slot, winnerId);
  }
  if (match.loser_next_match_id) {
    advanceTeam_(match.loser_next_match_id, match.loser_next_slot, loserId);
  }

  return getBracketData();
}

/**
 * 結果の取り消し。次ラウンド（準決勝の場合は3位決定戦も）に進出済みの場合は、
 * 進出先の試合がまだ未決着であることを確認した上で、進出も一緒に巻き戻す。
 */
function clearResult(matchId) {
  var sheet = getSheet_(MATCHES_SHEET, MATCH_HEADERS);
  var rowIdx = findRowIndexByMatchId_(sheet, matchId);
  if (rowIdx === -1) throw new Error('試合が見つかりません: ' + matchId);

  var rowValues = sheet.getRange(rowIdx, 1, 1, MATCH_HEADERS.length).getValues()[0];
  var match = {};
  MATCH_HEADERS.forEach(function (h, i) { match[h] = rowValues[i]; });

  if (!match.winner_id) {
    throw new Error('この試合はまだ結果が入力されていません。');
  }

  assertNextMatchNotDecided_(match.next_match_id);
  assertNextMatchNotDecided_(match.loser_next_match_id);

  setCell_(sheet, rowIdx, 'score1', '');
  setCell_(sheet, rowIdx, 'score2', '');
  setCell_(sheet, rowIdx, 'winner_id', '');
  setCell_(sheet, rowIdx, 'status', '');
  setCell_(sheet, rowIdx, 'win_type', '');

  if (match.next_match_id) {
    retractTeam_(match.next_match_id, match.next_slot);
  }
  if (match.loser_next_match_id) {
    retractTeam_(match.loser_next_match_id, match.loser_next_slot);
  }

  return getBracketData();
}

function assertNextMatchNotDecided_(nextMatchId) {
  if (!nextMatchId) return;
  var sheet = getSheet_(MATCHES_SHEET, MATCH_HEADERS);
  var rowIdx = findRowIndexByMatchId_(sheet, nextMatchId);
  if (rowIdx === -1) return;
  var winnerIdx = MATCH_HEADERS.indexOf('winner_id');
  var winnerId = sheet.getRange(rowIdx, winnerIdx + 1).getValue();
  if (winnerId) {
    throw new Error('進出先の試合（' + nextMatchId + '）の結果が既に入力されているため取り消せません。先にそちらの結果を取り消してください。');
  }
}

function setCell_(sheet, rowIdx, columnName, value) {
  var colIdx = MATCH_HEADERS.indexOf(columnName) + 1;
  sheet.getRange(rowIdx, colIdx).setValue(value);
}

function advanceTeam_(nextMatchId, slot, teamId) {
  var sheet = getSheet_(MATCHES_SHEET, MATCH_HEADERS);
  var rowIdx = findRowIndexByMatchId_(sheet, nextMatchId);
  if (rowIdx === -1) return;
  var column = (String(slot) === '1') ? 'team1_id' : 'team2_id';
  setCell_(sheet, rowIdx, column, teamId);
}

function retractTeam_(nextMatchId, slot) {
  var sheet = getSheet_(MATCHES_SHEET, MATCH_HEADERS);
  var rowIdx = findRowIndexByMatchId_(sheet, nextMatchId);
  if (rowIdx === -1) return;
  var column = (String(slot) === '1') ? 'team1_id' : 'team2_id';
  setCell_(sheet, rowIdx, column, '');
}

// ---------------------------------------------------------------------------
// Initial data setup - run manually once from the Apps Script editor
// (Teams/Matches シートを画像のトーナメント表の内容でリセットします)
// ---------------------------------------------------------------------------

function setupSheet() {
  var ss = getSpreadsheet_();

  var teamsSheet = ss.getSheetByName(TEAMS_SHEET) || ss.insertSheet(TEAMS_SHEET);
  teamsSheet.clear();
  teamsSheet.appendRow(TEAM_HEADERS);
  teamsSheet.setFrozenRows(1);
  // logo_url column is left blank - fill it in per team to swap in a real
  // logo image; leave blank to keep the colored-initial placeholder badge.
  var teams = [
    ['T1', '玉川', 'T', '#c0392b', '', 'たまがわ'],
    ['T2', '二子玉川・瀬田', 'FS', '#1a2634', '', 'ふたこたまがわせた'],
    ['T3', '三軒茶屋', 'S', '#2e7d4f', '', 'さんげんぢゃや'],
    ['T4', '祖師谷', '祖', '#d4b483', '', 'そしがや'],
    ['T5', '若林', 'W', '#5b3a29', '', 'わかばやし'],
    ['T6', '深沢', 'F', '#b8935a', '', 'ふかさわ'],
    ['T7', '経堂', 'K', '#c9a227', '', 'きょうどう'],
    ['T8', '砧', '砧', '#2b2b2b', '', 'きぬた'],
    ['T9', '中里', 'N', '#e07b39', '', 'なかざと'],
    ['T10', '用賀', '用', '#3f7a6d', '', 'ようが'],
    ['T11', '桜丘', '桜', '#e8a0b4', '', 'さくらがおか'],
    ['T12', '塚戸', 'T', '#1f2a4a', '', 'つかど'],
    ['T13', '駒繋', 'K', '#2f5233', '', 'こまつなぎ'],
    ['T14', '多聞', '多', '#2f6e4f', '', 'たもん'],
    ['T15', '松沢', '松', '#8a6d3b', '', 'まつざわ'],
    ['T16', '松原', 'M', '#8e1b3c', '', 'まつばら']
  ];
  teamsSheet.getRange(2, 1, teams.length, TEAM_HEADERS.length).setValues(teams);

  var matchesSheet = ss.getSheetByName(MATCHES_SHEET) || ss.insertSheet(MATCHES_SHEET);
  matchesSheet.clear();
  matchesSheet.appendRow(MATCH_HEADERS);
  matchesSheet.setFrozenRows(1);

  // match_id, round, order, team1_id, team2_id, date, time, venue,
  // score1, score2, winner_id, status, win_type,
  // next_match_id, next_slot, loser_next_match_id, loser_next_slot
  var matches = [
    ['R16-1', 'R16', 1, 'T1', 'T2', '', '', '', '', '', 'T1', 'completed', 'fusen1', 'QF-1', 1, '', ''],
    ['R16-2', 'R16', 2, 'T3', 'T4', '', '', '', '', '', '', '', '', 'QF-1', 2, '', ''],
    ['R16-3', 'R16', 3, 'T5', 'T6', '2026-08-23', '', '', '', '', '', '', '', 'QF-2', 1, '', ''],
    ['R16-4', 'R16', 4, 'T7', 'T8', '', '', '', '', '', '', '', '', 'QF-2', 2, '', ''],
    ['R16-5', 'R16', 5, 'T9', 'T10', '', '', '', '', '', '', '', '', 'QF-3', 1, '', ''],
    ['R16-6', 'R16', 6, 'T11', 'T12', '', '', '', '', '', '', '', '', 'QF-3', 2, '', ''],
    ['R16-7', 'R16', 7, 'T13', 'T14', '2026-07-18', '', '', '', '', '', '', '', 'QF-4', 1, '', ''],
    ['R16-8', 'R16', 8, 'T15', 'T16', '', '', '', '', '', '', '', '', 'QF-4', 2, '', ''],

    ['QF-1', 'QF', 1, 'T1', '', '', '', '', '', '', '', '', '', 'SF-1', 1, '', ''],
    ['QF-2', 'QF', 2, '', '', '', '', '', '', '', '', '', '', 'SF-1', 2, '', ''],
    ['QF-3', 'QF', 3, '', '', '', '', '', '', '', '', '', '', 'SF-2', 1, '', ''],
    ['QF-4', 'QF', 4, '', '', '', '', '', '', '', '', '', '', 'SF-2', 2, '', ''],

    ['SF-1', 'SF', 1, '', '', '', '', '', '', '', '', '', '', 'F', 1, '3RD', 1],
    ['SF-2', 'SF', 2, '', '', '', '', '', '', '', '', '', '', 'F', 2, '3RD', 2],

    ['F', 'F', 1, '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['3RD', '3RD', 1, '', '', '', '', '', '', '', '', '', '', '', '', '', '']
  ];
  matchesSheet.getRange(2, 1, matches.length, MATCH_HEADERS.length).setValues(matches);

  SpreadsheetApp.flush();
  Logger.log('setupSheet: Teams/Matches シートを初期化しました。');
}
