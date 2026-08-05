// ============================================================
// 環境2課 中期経営計画ダッシュボード — データ保存スクリプト v3
// 更新内容:
//   ⑥ パスワードハッシュをスプレッドシートで管理
//   ⑧ 保存のたびにバックアップ履歴を記録（最新30件）
//   ⑩ 勤怠データ（Excel）のGoogleドライブ自動取込
// ============================================================

const SHEET_JSON     = 'ダッシュボードデータ';
const SHEET_OT       = '残業実績';
const SHEET_LEAVE    = '有休取得';
const SHEET_BACKUP   = 'バックアップ履歴';
const SHEET_SETTINGS = '設定';
const SHEET_SALES    = '売上実績';
const SHEET_ROSTER   = 'メンバー設定'; // ⑨ 氏名・グループ・付与日数・取得数はここが正
const MONTHS = ['4月','5月','6月','7月','8月','9月','10月','11月','12月','1月','2月','3月'];
const MAX_BACKUPS = 30;
// デフォルトパスワード: 3150 の SHA-256ハッシュ
const DEFAULT_PIN_HASH = '4d364fbb3786fc31157cc1e2a2671aac0e36348ae9d6b4ba4459ee883c240fe8';
// ② ブルートフォース対策: 連続失敗回数としきい値
const MAX_PIN_ATTEMPTS = 5;
const LOCKOUT_MINUTES  = 5;

/* ─── GET: データ取得 / PIN照合 ─── */
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';

  if (action === 'verify_pin') {
    return verifyPin_(e.parameter.hash || '');
  }
  if (action === 'change_pin') {
    return changePin_(e.parameter.oldHash || '', e.parameter.newHash || '');
  }

  try {
    const sheet = getOrCreate_(SHEET_JSON);
    const json  = sheet.getRange('A1').getValue();
    const data  = json ? JSON.parse(json) : {};
    // ⑨ メンバー設定シートの読込に失敗してもメインデータは返す（機能を分離してダッシュボード全体を落とさない）
    try { data.roster = readRoster_(); } catch (e) { /* ロースター読込失敗時はメインデータのみ返す */ }
    // ⑪ メンバー別の残業実績（取込データ（月別）由来、あれば返す。失敗してもメインデータは返す）
    try { data.memberOvertime = readMemberOvertime_(); } catch (e) { /* 読込失敗時はメインデータのみ返す */ }
    return ok_(JSON.stringify(data));
  } catch (err) {
    return ok_(JSON.stringify({ error: err.message }));
  }
}

/* ─── POST: データ保存 ─── */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // ① 書き込みにもPIN認証を要求（②のロックアウトも共有）
    const auth = checkPinAuth_(body.pinHash || '');
    if (auth.locked) {
      return ok_(JSON.stringify({
        error: '試行回数が多すぎます。しばらくしてから再試行してください',
        locked: true, retryAfterSec: auth.retryAfterSec
      }));
    }
    if (!auth.ok) {
      return ok_(JSON.stringify({ error: '認証に失敗しました' }));
    }

    const data = body.payload || {};
    const json = JSON.stringify(data);

    const sheet = getOrCreate_(SHEET_JSON);
    sheet.getRange('A1').setValue(json);
    sheet.getRange('B1').setValue(
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
    );
    sheet.getRange('C1').setValue('最終保存日時');

    saveBackup_(json);
    writeOvertimeSheet_(data);
    writeLeaveSheet_(data);
    writeSalesSheet_(data);
    writeSalesTargetSheet_(data);

    return ok_(JSON.stringify({ status: 'saved' }));
  } catch (err) {
    return ok_(JSON.stringify({ error: err.message }));
  }
}

/* ─── ⑥ PIN照合 ─── */
function verifyPin_(inputHash) {
  try {
    const auth = checkPinAuth_(inputHash);
    if (auth.locked) {
      return ok_(JSON.stringify({ valid: false, locked: true, retryAfterSec: auth.retryAfterSec }));
    }
    return ok_(JSON.stringify({ valid: auth.ok }));
  } catch (err) {
    return ok_(JSON.stringify({ valid: false, error: err.message }));
  }
}

/* ─── ⑥ PIN変更 ─── */
function changePin_(oldHash, newHash) {
  try {
    const auth = checkPinAuth_(oldHash);
    if (auth.locked) {
      return ok_(JSON.stringify({
        success: false, locked: true, retryAfterSec: auth.retryAfterSec,
        error: '試行回数が多すぎます'
      }));
    }
    if (!auth.ok)
      return ok_(JSON.stringify({ success: false, error: '現在のパスワードが違います' }));
    if (!newHash || newHash.length !== 64)
      return ok_(JSON.stringify({ success: false, error: '新しいパスワードが無効です' }));
    const s = getOrCreate_(SHEET_SETTINGS);
    s.getRange('A1').setValue(newHash);
    s.getRange('B1').setValue('パスワードハッシュ（SHA-256）');
    s.getRange('A2').setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
    s.getRange('B2').setValue('最終変更日時');
    return ok_(JSON.stringify({ success: true }));
  } catch (err) {
    return ok_(JSON.stringify({ success: false, error: err.message }));
  }
}

/* ─── ⑥ 保存済みPINハッシュを取得 ─── */
function getStoredPinHash_() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  if (!s) return DEFAULT_PIN_HASH;
  return s.getRange('A1').getValue() || DEFAULT_PIN_HASH;
}

/* ─── ① ② PIN認証（ロックアウト付き） ─── */
// doPost / verifyPin_ / changePin_ が共有する認証窓口。
// スクリプトプロパティに失敗回数とロック解除時刻を記録し、連続失敗でロックする。
function checkPinAuth_(hash) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const props = PropertiesService.getScriptProperties();
    const now = Date.now();
    const lockedUntil = Number(props.getProperty('PIN_LOCKED_UNTIL') || 0);
    if (now < lockedUntil) {
      return { ok: false, locked: true, retryAfterSec: Math.ceil((lockedUntil - now) / 1000) };
    }
    const stored = getStoredPinHash_();
    if (stored === hash) {
      props.deleteProperty('PIN_FAIL_COUNT');
      props.deleteProperty('PIN_LOCKED_UNTIL');
      return { ok: true };
    }
    const fails = Number(props.getProperty('PIN_FAIL_COUNT') || 0) + 1;
    if (fails >= MAX_PIN_ATTEMPTS) {
      props.setProperty('PIN_LOCKED_UNTIL', String(now + LOCKOUT_MINUTES * 60 * 1000));
      props.deleteProperty('PIN_FAIL_COUNT');
      return { ok: false, locked: true, retryAfterSec: LOCKOUT_MINUTES * 60 };
    }
    props.setProperty('PIN_FAIL_COUNT', String(fails));
    return { ok: false, locked: false, remaining: MAX_PIN_ATTEMPTS - fails };
  } finally {
    lock.releaseLock();
  }
}

/* ─── ⑨ メンバー設定シート（氏名・グループ・付与日数・取得数はここが正） ─── */
// 列: A=氏名 B=グループ C=付与日数 D=取得数
function readRoster_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_ROSTER);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ROSTER);
    sheet.getRange(1,1,1,4).setValues([['氏名','グループ','付与日数','取得数']])
      .setFontWeight('bold').setBackground('#f1f5f9');
    sheet.setFrozenRows(1);
    seedRosterFromExisting_(sheet);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 4).getValues()
    .filter(r => String(r[0] || '').trim() !== '')
    .map(r => ({
      name: String(r[0]).trim(),
      group: String(r[1] || '').trim(),
      days: Number(r[2]) || 0,
      takenTotal: Number(r[3]) || 0
    }));
}

/* ─── ⑨ 初回作成時、既存のダッシュボードデータからメンバー設定シートへ移行 ─── */
function seedRosterFromExisting_(sheet) {
  try {
    const jsonSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_JSON);
    if (!jsonSheet) return;
    const raw = jsonSheet.getRange('A1').getValue();
    if (!raw) return;
    const data = JSON.parse(raw);
    const members = data.members || [];
    if (!members.length) return;
    const rows = members.map(m => {
      const tm = m.takenMonths || [];
      const taken = tm.reduce((a, v) => a + (v || 0), 0);
      return [m.name, m.group || '', m.days || 20, taken];
    });
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  } catch (e) { /* 移行失敗時は空のまま。管理者に手入力を促す */ }
}

/* ─── ⑧ バックアップ履歴（最新30件保持） ─── */
function saveBackup_(json) {
  try {
    const sheet = getOrCreate_(SHEET_BACKUP);
    if (sheet.getRange('A1').getValue() !== '保存日時')
      sheet.getRange('A1:B1').setValues([['保存日時','データ（JSON）']]).setFontWeight('bold');
    const ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    sheet.insertRowAfter(1);
    sheet.getRange(2,1,1,2).setValues([[ts, json]]);
    const last = sheet.getLastRow();
    if (last > MAX_BACKUPS + 1)
      sheet.deleteRows(MAX_BACKUPS + 2, last - MAX_BACKUPS - 1);
  } catch(e) { /* バックアップ失敗はメイン処理を止めない */ }
}

/* ─── 残業実績シート ─── */
function writeOvertimeSheet_(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_OT) || ss.insertSheet(SHEET_OT);
  sheet.clearContents();
  const OT25  = [239,148,245.5,234.5,229,283.75,276.5,351.25,366.5,289,297.7,401];
  const OT26T = OT25.map(v => Math.round(v*.50*100)/100);
  const OT27T = OT25.map(v => Math.round(v*.40*100)/100);
  const OT28T = OT25.map(v => Math.round(v*.30*100)/100);
  const header = ['月','2025年度実績(h)','2026実績(h)','2027実績(h)','2028実績(h)',
    '2026目標(h)','2027目標(h)','2028目標(h)'];
  const rows = MONTHS.map((m,i) => [m, OT25[i],
    data.actualData && data.actualData['2026'] ? (data.actualData['2026'][i] ?? '') : '',
    data.actualData && data.actualData['2027'] ? (data.actualData['2027'][i] ?? '') : '',
    data.actualData && data.actualData['2028'] ? (data.actualData['2028'][i] ?? '') : '',
    OT26T[i], OT27T[i], OT28T[i]]);
  const toSum = a => a.filter(v=>v!==''&&v!=null).reduce((s,v)=>s+v,0);
  rows.push(['合計',toSum(OT25),
    toSum((data.actualData||{})['2026']||[]),toSum((data.actualData||{})['2027']||[]),
    toSum((data.actualData||{})['2028']||[]),toSum(OT26T),toSum(OT27T),toSum(OT28T)]);
  sheet.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold').setBackground('#f1f5f9');
  sheet.getRange(2,1,rows.length,header.length).setValues(rows);
  sheet.setFrozenRows(1);
}

/* ─── 有休取得シート ─── */
function writeLeaveSheet_(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_LEAVE) || ss.insertSheet(SHEET_LEAVE);
  sheet.clearContents();
  const header = ['氏名','グループ','付与日数','取得合計(日)','取得率(%)','残日数',
    ...MONTHS.map(m=>m+'取得')];
  const members = data.members || [];
  const rows = members.map(m => {
    const tm    = m.takenMonths || new Array(12).fill(0);
    const taken = tm.reduce((a,v)=>a+(v||0),0);
    const rate  = m.days>0?Math.round(taken/m.days*100):0;
    return [m.name,m.group||'',m.days,taken,rate,Math.max(0,m.days-taken),...tm];
  });
  if (!rows.length) return;
  sheet.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold').setBackground('#f1f5f9');
  sheet.getRange(2,1,rows.length,header.length).setValues(rows);
  for (let r=2; r<=rows.length+1; r++) {
    const rate = sheet.getRange(r,5).getValue();
    sheet.getRange(r,5).setBackground(rate>=80?'#dcfce7':rate>=50?'#fef3c7':'#fee2e2');
  }
  sheet.setFrozenRows(1);
}

/* ─── 売上実績シート ─── */
function writeSalesSheet_(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_SALES) || ss.insertSheet(SHEET_SALES);
  sheet.clearContents();
  const header = ['年度', '月', 'ユニット2(万円)', 'ユニット3(万円)', '合算(万円)'];
  const rows = [];
  const sd = data.salesData || {};
  ['2025','2026','2027','2028'].forEach(yr => {
    const u2 = sd[yr]?.unit2 || new Array(12).fill(null);
    const u3 = sd[yr]?.unit3 || new Array(12).fill(null);
    MONTHS.forEach((m, i) => {
      const v2 = u2[i] ?? '', v3 = u3[i] ?? '';
      const tot = (u2[i] != null && u3[i] != null) ? u2[i] + u3[i] : '';
      rows.push([yr + '年度', m, v2, v3, tot]);
    });
  });
  sheet.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold').setBackground('#f1f5f9');
  if(rows.length) sheet.getRange(2,1,rows.length,header.length).setValues(rows);
  sheet.setFrozenRows(1);
}

/* ─── 売上目標シート ─── */
function writeSalesTargetSheet_(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName('売上目標') || ss.insertSheet('売上目標');
  sheet.clearContents();
  const header = ['年度','ユニット2目標(万円)','ユニット3目標(万円)','合算目標(万円)'];
  const st = data.salesTargets || {};
  const rows = ['2025','2026','2027','2028'].map(yr => [
    yr+'年度',
    st[yr]?.unit2 ?? '',
    st[yr]?.unit3 ?? '',
    st[yr]?.total ?? ''
  ]);
  sheet.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold').setBackground('#f1f5f9');
  sheet.getRange(2,1,rows.length,header.length).setValues(rows);
  sheet.setFrozenRows(1);
}

/* ─── ヘルパー ─── */
function getOrCreate_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function ok_(text) {
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// ⑩ 勤怠データ自動取込（Googleドライブ連携）
// 「勤怠データ取込」フォルダに勤怠システム出力のExcel(.xlsx)をドロップすると、
// ファイル内の「総残業時間」「有休取得日数」を読み取り、
//   - 残業実績（全メンバー分をその月で合算）→ ダッシュボードのactualDataへ
//   - 有休取得数（メンバーごと）→「メンバー設定」シートD列へ
// を自動反映する。処理済みファイルは自動でサブフォルダへ移動。
//
// 【事前準備（初回のみ、Apps Scriptエディタで実施）】
//   左メニュー「サービス」の＋ボタン → 「Drive API」を追加（識別子: Drive）
//   これを行わないと下記コードは動作しない（Driveが未定義というエラーになる）
// ============================================================
const IMPORT_FOLDER_NAME    = '勤怠データ取込';
const PROCESSED_FOLDER_NAME = '処理済み';
const ERROR_FOLDER_NAME     = 'エラー';
const SHEET_IMPORT_TRACK    = '取込データ（月別）';
const SHEET_IMPORT_LOG      = '取込ログ';
const VALID_ACTUAL_YEARS    = ['2026', '2027', '2028']; // actualDataで管理している残業実績の対象年度

/* ─── スプレッドシートを開いたときに操作用メニューを追加 ─── */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('勤怠データ取込')
    .addItem('今すぐ取込む', 'importAttendanceFiles')
    .addItem('取込フォルダのURLを確認（初回はここで作成されます）', 'showImportFolderUrl_')
    .addItem('自動実行（1日1回・6時頃）を有効化', 'setupImportTrigger')
    .addToUi();
}

/* ─── 毎日1回の自動実行トリガーを作成（すでにあれば何もしない） ─── */
function setupImportTrigger() {
  const exists = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'importAttendanceFiles');
  if (exists) {
    SpreadsheetApp.getUi().alert('すでに自動実行は設定されています（毎日6時頃に実行）');
    return;
  }
  ScriptApp.newTrigger('importAttendanceFiles')
    .timeBased().everyDays(1).atHour(6).nearMinute(0).inTimezone('Asia/Tokyo')
    .create();
  SpreadsheetApp.getUi().alert('毎日6時頃に自動取込を実行するよう設定しました');
}

/* ─── 取込フォルダ（なければ作成）を取得。処理済み/エラー用の子フォルダも用意 ─── */
function getImportFolders_() {
  const ssFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const parents = ssFile.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const getOrCreateChild = (base, name) => {
    const it = base.getFoldersByName(name);
    return it.hasNext() ? it.next() : base.createFolder(name);
  };
  const inbox     = getOrCreateChild(parent, IMPORT_FOLDER_NAME);
  const processed = getOrCreateChild(inbox, PROCESSED_FOLDER_NAME);
  const errorF    = getOrCreateChild(inbox, ERROR_FOLDER_NAME);
  return { inbox, processed, errorF };
}

function showImportFolderUrl_() {
  const { inbox } = getImportFolders_();
  SpreadsheetApp.getUi().alert('取込フォルダ:\n' + inbox.getUrl());
}

/* ─── メイン: 取込フォルダ直下のExcelファイルを処理する（手動メニュー／自動トリガー共通） ─── */
function importAttendanceFiles() {
  if (typeof Drive === 'undefined') {
    writeImportLog_([[new Date(), '(設定エラー)', 'エラー',
      'Drive APIが未設定です。Apps Scriptエディタの「サービス」から Drive API を追加してください。']]);
    return;
  }

  const { inbox, processed, errorF } = getImportFolders_();
  const files = inbox.getFiles(); // サブフォルダ（処理済み/エラー）は対象外
  const logRows = [];
  const touchedYearMonths = new Set(); // "年度-月インデックス"
  const touchedMembers    = new Set();
  let count = 0;

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (!/\.xlsx$/i.test(name)) continue;
    count++;
    try {
      const parsed = parseAttendanceFile_(file);
      upsertImportTrack_(parsed, name);
      touchedYearMonths.add(parsed.year + '-' + parsed.monthIdx);
      if (parsed.matchedMember) touchedMembers.add(parsed.matchedMember);
      file.moveTo(processed);
      logRows.push([new Date(), name, '成功',
        `${parsed.name} / ${parsed.year}年度 ${parsed.month}月 / 総残業${parsed.overtime}h / 有休${parsed.leaveDays}日` +
        (parsed.matchedMember ? '' : '（「メンバー設定」に氏名が見つからず有休は反映されていません）')]);
    } catch (err) {
      file.moveTo(errorF);
      logRows.push([new Date(), name, 'エラー', err.message]);
    }
  }

  if (count === 0) {
    writeImportLog_([[new Date(), '(対象ファイルなし)', '情報', '取込フォルダにxlsxファイルがありませんでした']]);
    return;
  }

  recalcFromTrack_(touchedYearMonths, touchedMembers);
  logRows.push(...checkCompleteness_(touchedYearMonths));
  writeImportLog_(logRows);
}

/* ─── 今回取込んだ年度・月について、「メンバー設定」に対して未取込のメンバーがいないか確認 ─── */
// アップし忘れがあっても合計が黙って少なくなるだけになるのを防ぐため、警告として出す
function checkCompleteness_(touchedYearMonths) {
  const roster = readRoster_();
  if (!roster.length) return [];
  const trackSheet = getOrCreate_(SHEET_IMPORT_TRACK);
  const lastRow = trackSheet.getLastRow();
  const rows = lastRow >= 2 ? trackSheet.getRange(2, 1, lastRow - 1, 3).getValues() : []; // 氏名,年度,月

  const out = [];
  touchedYearMonths.forEach(key => {
    const [year, idxStr] = key.split('-');
    const monthIdx = Number(idxStr);
    const month = monthIdx <= 8 ? monthIdx + 4 : monthIdx - 8; // 0=4月...11=3月 → カレンダー月
    const present = new Set(
      rows.filter(r => String(r[1]) === year && Number(r[2]) === month).map(r => r[0])
    );
    const missing = roster.filter(m => !present.has(m.name)).map(m => m.name);
    if (missing.length) {
      out.push([new Date(), '(未取込チェック)', '情報',
        `${year}年度 ${month}月: 未取込のメンバーが${missing.length}名います → ${missing.join('、')}`]);
    }
  });
  return out;
}

/* ─── 1ファイルを解析し、氏名・年月・総残業時間・有休取得日数を取り出す ─── */
function parseAttendanceFile_(file) {
  const values = readXlsxAsValues_(file);

  // 5〜10行目（0-indexで4〜9）はラベルと値が2列ずつ交互に並ぶ構造
  const labelMap = {};
  for (let r = 4; r <= 9 && r < values.length; r++) {
    const row = values[r] || [];
    for (let c = 0; c + 1 < row.length; c += 2) {
      const label = String(row[c] || '').trim();
      if (label) labelMap[label] = row[c + 1];
    }
  }

  // 2行目: 「名前 ： 阿部 凌太」「2026年7月」のように1セルに結合されている
  const row2 = (values[1] || []).map(v => String(v || ''));
  const nameCell  = row2.find(v => v.indexOf('名前') !== -1) || '';
  const nameMatch = nameCell.match(/名前\s*[:：]\s*(.+)/);
  const dateCell  = row2.find(v => /\d{4}年\d{1,2}月/.test(v)) || '';
  const dateMatch = dateCell.match(/(\d{4})年(\d{1,2})月/);

  if (!nameMatch || !dateMatch) {
    throw new Error('氏名または対象年月が読み取れませんでした（ファイル形式が想定と異なる可能性があります）');
  }
  const name  = nameMatch[1].trim();
  const year  = Number(dateMatch[1]);
  const month = Number(dateMatch[2]);

  const overtime  = Number(labelMap['総残業時間']);
  const leaveDays = Number(labelMap['有休取得日数']);
  if (isNaN(overtime))  throw new Error('「総残業時間」の値が読み取れませんでした');
  if (isNaN(leaveDays)) throw new Error('「有休取得日数」の値が読み取れませんでした');

  const fiscalYear = month >= 4 ? year : year - 1;      // 4月始まりの年度
  const monthIdx   = month >= 4 ? month - 4 : month + 8; // 0=4月...11=3月（ダッシュボードのMONTHSと同順）

  return { name, year: fiscalYear, month, monthIdx, overtime, leaveDays,
    matchedMember: findRosterMemberName_(name) };
}

/* ─── xlsxをGoogleスプレッドシートへ一時変換してセル値を読み取る（読了後は即削除） ─── */
function readXlsxAsValues_(file) {
  const tmp = Drive.Files.copy(
    { name: '__tmp_import_' + file.getId(), mimeType: MimeType.GOOGLE_SHEETS },
    file.getId()
  );
  try {
    const ss = SpreadsheetApp.openById(tmp.id);
    return ss.getSheets()[0].getDataRange().getValues();
  } finally {
    Drive.Files.remove(tmp.id);
  }
}

/* ─── ⑪ 「取込データ（月別）」から、メンバーごとの月別総残業時間を組み立てる ─── */
// 戻り値: { "氏名": { "2026": [12ヶ月分(h) or null, ...], "2027": [...], "2028": [...] }, ... }
// 手入力時代の月（取込データに存在しない月）はnullのまま＝グラフ側で欠損として扱う
function readMemberOvertime_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_IMPORT_TRACK);
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // 氏名,年度,月,総残業時間(h)
  const out = {};
  rows.forEach(r => {
    const name = String(r[0] || '').trim();
    const year = String(r[1]);
    const month = Number(r[2]);
    if (!name || VALID_ACTUAL_YEARS.indexOf(year) === -1) return;
    const monthIdx = month >= 4 ? month - 4 : month + 8; // 0=4月...11=3月
    if (!out[name]) out[name] = {};
    if (!out[name][year]) out[name][year] = new Array(12).fill(null);
    out[name][year][monthIdx] = Number(r[3]) || 0;
  });
  return out;
}

/* ─── 「メンバー設定」シートの氏名と突合（全角/半角スペースの差異は無視） ─── */
function findRosterMemberName_(rawName) {
  const norm = s => String(s || '').replace(/[\s　]/g, '');
  const target = norm(rawName);
  const hit = readRoster_().find(m => norm(m.name) === target);
  return hit ? hit.name : null;
}

/* ─── 「取込データ（月別）」シートへ反映。同一氏名+年度+月の行は上書き（再取込・修正に対応） ─── */
function upsertImportTrack_(parsed, fileName) {
  const sheet = getOrCreate_(SHEET_IMPORT_TRACK);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 7)
      .setValues([['氏名', '年度', '月', '総残業時間(h)', '有休取得日数', '取込日時', 'ファイル名']])
      .setFontWeight('bold').setBackground('#f1f5f9');
    sheet.setFrozenRows(1);
  }
  const memberKey = parsed.matchedMember || parsed.name;
  const key = memberKey + '|' + parsed.year + '|' + parsed.month;
  let targetRow = -1;
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < data.length; i++) {
      if ((data[i][0] + '|' + data[i][1] + '|' + data[i][2]) === key) { targetRow = i + 2; break; }
    }
  }
  const rowVals = [memberKey, parsed.year, parsed.month, parsed.overtime, parsed.leaveDays,
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'), fileName];
  const row = targetRow === -1 ? sheet.getLastRow() + 1 : targetRow;
  sheet.getRange(row, 1, 1, rowVals.length).setValues([rowVals]);
}

/* ─── 「取込データ（月別）」を集計し、残業実績（全社）・メンバー取得数（個人）へ反映 ─── */
function recalcFromTrack_(touchedYearMonths, touchedMembers) {
  const trackSheet = getOrCreate_(SHEET_IMPORT_TRACK);
  const lastRow = trackSheet.getLastRow();
  if (lastRow < 2) return;
  const rows = trackSheet.getRange(2, 1, lastRow - 1, 5).getValues(); // 氏名,年度,月,残業,有休

  const jsonSheet = getOrCreate_(SHEET_JSON);
  const raw  = jsonSheet.getRange('A1').getValue();
  const data = raw ? JSON.parse(raw) : {};
  data.actualData = data.actualData || {};
  VALID_ACTUAL_YEARS.forEach(y => { if (!data.actualData[y]) data.actualData[y] = new Array(12).fill(null); });

  // ① 残業実績: 該当年度・月について全メンバー分を合算して上書き
  touchedYearMonths.forEach(key => {
    const [year, idxStr] = key.split('-');
    const monthIdx = Number(idxStr);
    if (VALID_ACTUAL_YEARS.indexOf(year) === -1) return; // 対象外年度はスキップ（残業実績グラフの管理外）
    let sum = 0, any = false;
    rows.forEach(r => {
      const rMonthIdx = r[2] >= 4 ? r[2] - 4 : r[2] + 8;
      if (String(r[1]) === year && rMonthIdx === monthIdx) { sum += Number(r[3]) || 0; any = true; }
    });
    if (any) data.actualData[year][monthIdx] = Math.round(sum * 100) / 100;
  });

  // ② 有休取得数: 「メンバー設定」シートD列を、当該メンバーの今年度内合計で上書き
  const now = new Date();
  const curFiscalYear = String(now.getMonth() + 1 >= 4 ? now.getFullYear() : now.getFullYear() - 1);
  const rosterSheet = getOrCreate_(SHEET_ROSTER);
  const rLast = rosterSheet.getLastRow();
  if (rLast >= 2 && touchedMembers.size) {
    const rosterVals = rosterSheet.getRange(2, 1, rLast - 1, 4).getValues();
    touchedMembers.forEach(memberName => {
      let total = 0;
      rows.forEach(r => { if (r[0] === memberName && String(r[1]) === curFiscalYear) total += Number(r[4]) || 0; });
      for (let i = 0; i < rosterVals.length; i++) {
        if (rosterVals[i][0] === memberName) {
          rosterSheet.getRange(i + 2, 4).setValue(Math.round(total * 10) / 10);
          break;
        }
      }
    });
  }

  const json = JSON.stringify(data);
  jsonSheet.getRange('A1').setValue(json);
  jsonSheet.getRange('B1').setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
  jsonSheet.getRange('C1').setValue('最終保存日時（勤怠データ自動取込）');
  saveBackup_(json);
  writeOvertimeSheet_(data);
}

/* ─── 取込ログシートへ追記（最新が上に来るよう先頭挿入、最大500件保持） ─── */
function writeImportLog_(rows) {
  const sheet = getOrCreate_(SHEET_IMPORT_LOG);
  if (sheet.getRange('A1').getValue() !== '取込日時') {
    sheet.getRange('A1:D1').setValues([['取込日時', 'ファイル名', '結果', '詳細']])
      .setFontWeight('bold').setBackground('#f1f5f9');
    sheet.setFrozenRows(1);
  }
  rows.forEach(r => {
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, 4).setValues([r]);
    const cell = sheet.getRange(2, 3);
    cell.setBackground(r[2] === '成功' ? '#dcfce7' : r[2] === 'エラー' ? '#fee2e2' : '#f1f5f9');
  });
  const last = sheet.getLastRow();
  if (last > 501) sheet.deleteRows(502, last - 501);
}
