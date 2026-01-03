/**
 * 2ヶ月スケジュール（ホワイトボード形式）- Google Apps Script バックエンド
 *
 * スプレッドシートをDBとして、当月＋翌月の予定を管理するWebアプリ
 * - 権限制御: viewer（閲覧のみ）/ editor・admin（当月＋翌月のみCRUD可）
 * - 過去は閲覧のみ
 */

// =============================================================================
// 定数定義
// =============================================================================
const SHEET_SETTINGS = '01_Settings';
const SHEET_USERS    = '02_Users';
const SHEET_EVENTS   = '03_Events';
const SHEET_NOTES    = '05_Notes';

const SETTINGS_BASE_MONTH_CELL = 'B5';      // YYYY-MM
const SETTINGS_EDITABLE_MONTHS_CELL = 'B6'; // 2
const SETTINGS_TZ_CELL = 'B7';              // Asia/Tokyo

const EVENTS_HEADER_ROW = 5;
const EVENTS_DATA_START_ROW = 6;

const USERS_HEADER_ROW = 5;
const USERS_DATA_START_ROW = 6;

// =============================================================================
// Webアプリ エントリーポイント
// =============================================================================

/**
 * GET リクエスト時にHTMLを返す
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('社長スケジュール')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * HTMLテンプレートからファイルをインクルード
 */
function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// =============================================================================
// 初期化・データ取得
// =============================================================================

/**
 * 初回ロード時に必要なすべてのデータを返す
 * @returns {Object} { ok, user, settings, masters, range, events }
 */
function getBootstrap() {
  try {
    const settings = getSettings_();
    const user = getUserContext_();
    const masters = getMasters_();
    const range = getEditableRange_(settings);
    const events = listEvents_(range.fromISO, range.toISO);

    // 当月〜翌月の備考を取得
    const fromMonth = settings.baseMonth;
    const toMonth = nextMonthKeyServer_(settings.baseMonth);
    const notes = getNotesMap_(fromMonth, toMonth);

    return {
      ok: true,
      user,
      settings,
      masters,
      range,
      events,
      notes
    };
  } catch (e) {
    console.error('getBootstrap error:', e);
    return {
      ok: false,
      error: e.message || 'UNKNOWN_ERROR'
    };
  }
}

// =============================================================================
// Settings 取得
// =============================================================================

/**
 * 01_Settings から設定値を取得
 */
function getSettings_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_SETTINGS);

  if (!sh) {
    throw new Error('SHEET_NOT_FOUND: ' + SHEET_SETTINGS);
  }

  const tz = String(sh.getRange(SETTINGS_TZ_CELL).getValue() || 'Asia/Tokyo');
  let baseMonth = String(sh.getRange(SETTINGS_BASE_MONTH_CELL).getDisplayValue() || '').trim();
  const editableMonths = Number(sh.getRange(SETTINGS_EDITABLE_MONTHS_CELL).getValue() || 2);

  // baseMonthが空または不正な場合は今日の年月を自動採用
  if (!/^\d{4}-\d{2}$/.test(baseMonth)) {
    baseMonth = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  }

  return { tz, baseMonth, editableMonths };
}

/**
 * 編集可能期間を計算
 * @param {Object} settings
 * @returns {Object} { fromISO, toISO }
 */
function getEditableRange_(settings) {
  const tz = settings.tz;
  const [y, m] = settings.baseMonth.split('-').map(Number);

  // 当月1日から
  const from = new Date(y, m - 1, 1);
  // editableMonths ヶ月分の月末まで（2なら当月＋翌月末）
  const to = new Date(y, m - 1 + settings.editableMonths, 0);

  return {
    fromISO: Utilities.formatDate(from, tz, 'yyyy-MM-dd'),
    toISO: Utilities.formatDate(to, tz, 'yyyy-MM-dd')
  };
}

// =============================================================================
// User / ACL（権限制御）
// =============================================================================

/**
 * 現在のユーザー情報を取得
 */
function getUserContext_() {
  let email = '';
  try {
    const activeUser = Session.getActiveUser();
    if (activeUser) {
      email = activeUser.getEmail() || '';
    }
  } catch (e) {
    console.warn('Could not get active user email:', e);
  }

  const role = getRoleForEmail_(email);
  return { email, role };
}

/**
 * メールアドレスから権限を取得
 * @param {string} email
 * @returns {string} 'viewer' | 'editor' | 'admin'
 */
function getRoleForEmail_(email) {
  if (!email) return 'viewer';

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_USERS);

  if (!sh) {
    console.warn('Users sheet not found');
    return 'viewer';
  }

  const lastRow = sh.getLastRow();
  if (lastRow < USERS_DATA_START_ROW) return 'viewer';

  const numRows = lastRow - USERS_DATA_START_ROW + 1;
  // A: メールアドレス, B: 権限, C: 有効
  const values = sh.getRange(USERS_DATA_START_ROW, 1, numRows, 3).getValues();

  for (const [em, role, enabled] of values) {
    if (!em) continue;

    if (String(em).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      // 有効フラグがFALSEならviewerに降格
      if (enabled === false || String(enabled).toUpperCase() === 'FALSE') {
        return 'viewer';
      }

      const r = String(role || 'viewer').trim().toLowerCase();
      return (['viewer', 'editor', 'admin'].includes(r)) ? r : 'viewer';
    }
  }

  return 'viewer';
}

// =============================================================================
// Masters（type/status マスタ）
// =============================================================================

/**
 * 予定種別・ステータスのマスタデータを取得
 */
function getMasters_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_SETTINGS);

  if (!sh) {
    return { types: [], statuses: [] };
  }

  // type: 行13から（A:種別コード, B:表示名, C:色, D:並び順, E:有効）
  const types = readMaster_(sh, 13, 1, 5)
    .filter(r => r[0] && isTrue_(r[4]))
    .map(r => ({
      code: String(r[0]).trim(),
      label: String(r[1] || r[0]).trim(),
      color: String(r[2] || '#374151').trim(),
      sort: Number(r[3] || 999)
    }))
    .sort((a, b) => a.sort - b.sort);

  // status: 行25から（A:状態コード, B:表示名, C:並び順, D:有効）
  const statuses = readMaster_(sh, 25, 1, 4)
    .filter(r => r[0] && isTrue_(r[3]))
    .map(r => ({
      code: String(r[0]).trim(),
      label: String(r[1] || r[0]).trim(),
      sort: Number(r[2] || 999)
    }))
    .sort((a, b) => a.sort - b.sort);

  return { types, statuses };
}

/**
 * マスタデータを読み取るヘルパー
 */
function readMaster_(sh, startRow, startCol, width) {
  const last = sh.getLastRow();
  if (last < startRow) return [];

  const values = sh.getRange(startRow, startCol, last - startRow + 1, width).getValues();
  const out = [];

  for (const r of values) {
    if (r[0] === '' || r[0] == null) continue;
    out.push(r);
  }

  return out;
}

/**
 * TRUE判定ヘルパー
 */
function isTrue_(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}

// =============================================================================
// Events 読み取り
// =============================================================================

/**
 * 指定期間のイベントを取得（公開API）
 */
function listEvents(fromISO, toISO) {
  try {
    const events = listEvents_(fromISO, toISO);
    return { ok: true, events };
  } catch (e) {
    console.error('listEvents error:', e);
    return { ok: false, error: e.message };
  }
}

/**
 * 指定期間のイベントを取得（内部）
 */
function listEvents_(fromISO, toISO) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_EVENTS);

  if (!sh) {
    throw new Error('SHEET_NOT_FOUND: ' + SHEET_EVENTS);
  }

  const lastRow = sh.getLastRow();
  if (lastRow < EVENTS_DATA_START_ROW) return [];

  const numRows = lastRow - EVENTS_DATA_START_ROW + 1;
  // A〜O列（15列）
  const values = sh.getRange(EVENTS_DATA_START_ROW, 1, numRows, 15).getValues();

  const from = new Date(fromISO + 'T00:00:00');
  const to = new Date(toISO + 'T23:59:59');

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const eventId = r[0];
    const d = r[1];

    if (!eventId || !d) continue;

    const dd = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dd.getTime())) continue;
    if (dd < from || dd > to) continue;

    out.push({
      event_id: String(eventId),
      date: formatISODate_(dd),
      start_time: asTimeStr_(r[2]),
      end_time: asTimeStr_(r[3]),
      type: String(r[4] || ''),
      title: String(r[5] || ''),
      location: String(r[6] || ''),
      memo: String(r[7] || ''),
      display_order: Number(r[8] || 0),
      status: String(r[9] || 'CONFIRMED'),
      _row: EVENTS_DATA_START_ROW + i
    });
  }

  // 並び順: date → start_time（空は最後）→ display_order
  out.sort((a, b) => {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;

    const at = a.start_time || '99:99';
    const bt = b.start_time || '99:99';
    if (at !== bt) return at < bt ? -1 : 1;

    return (a.display_order || 0) - (b.display_order || 0);
  });

  return out;
}

// =============================================================================
// Events 書き込み（CRUD）
// =============================================================================

/**
 * 新規イベント作成
 */
function createEvent(payload) {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { ok: false, error: 'LOCK_TIMEOUT' };
  }

  try {
    const user = getUserContext_();
    const settings = getSettings_();

    // 権限チェック（サーバ側で確実に）
    assertCanEdit_(user, settings, payload.date);

    // バリデーション
    validateEventPayload_(payload, settings);

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_EVENTS);

    if (!sh) {
      throw new Error('SHEET_NOT_FOUND: ' + SHEET_EVENTS);
    }

    const now = new Date();
    const tz = settings.tz;
    const eventId = payload.event_id ? String(payload.event_id).trim() : Utilities.getUuid();

    // 時刻は明示的にテキストとして保存（タイムゾーン変換を防ぐ）
    const startTimeText = payload.start_time ? String(payload.start_time) : '';
    const endTimeText = payload.end_time ? String(payload.end_time) : '';

    const newRow = sh.getLastRow() + 1;
    sh.getRange(newRow, 1, 1, 15).setValues([[
      eventId,
      new Date(payload.date + 'T00:00:00'),
      startTimeText,
      endTimeText,
      payload.type || '',
      payload.title || '',
      payload.location || '',
      payload.memo || '',
      payload.display_order !== '' ? Number(payload.display_order) : '',
      payload.status || 'CONFIRMED',
      '', // 月キー（K列）はシート側で自動計算想定、ここは空
      Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'),
      user.email,
      Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'),
      user.email,
      ''
    ]]);
    // 時刻列（C, D）をテキスト形式に設定
    sh.getRange(newRow, 3, 1, 2).setNumberFormat('@');

    return { ok: true, event_id: eventId };
  } catch (e) {
    console.error('createEvent error:', e);
    return { ok: false, error: e.message || 'CREATE_FAILED' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * イベント更新
 */
function updateEvent(eventId, payload) {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { ok: false, error: 'LOCK_TIMEOUT' };
  }

  try {
    const user = getUserContext_();
    const settings = getSettings_();

    // 権限チェック（サーバ側で確実に）
    assertCanEdit_(user, settings, payload.date);

    // バリデーション
    validateEventPayload_(payload, settings);

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_EVENTS);

    if (!sh) {
      throw new Error('SHEET_NOT_FOUND: ' + SHEET_EVENTS);
    }

    const rowNum = findEventRowById_(sh, String(eventId));
    if (!rowNum) {
      return { ok: false, error: 'NOT_FOUND' };
    }

    const tz = settings.tz;
    const now = new Date();

    // 作成日時・作成者は保持
    const createdAt = sh.getRange(rowNum, 12).getValue();
    const createdBy = sh.getRange(rowNum, 13).getValue();

    // 時刻は明示的にテキストとして保存（タイムゾーン変換を防ぐ）
    const startTimeText = payload.start_time ? String(payload.start_time) : '';
    const endTimeText = payload.end_time ? String(payload.end_time) : '';

    sh.getRange(rowNum, 1, 1, 15).setValues([[
      String(eventId),
      new Date(payload.date + 'T00:00:00'),
      startTimeText,
      endTimeText,
      payload.type || '',
      payload.title || '',
      payload.location || '',
      payload.memo || '',
      payload.display_order !== '' ? Number(payload.display_order) : '',
      payload.status || 'CONFIRMED',
      '', // 月キー
      createdAt || '',
      createdBy || '',
      Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'),
      user.email
    ]]);
    // 時刻列（C, D）をテキスト形式に設定
    sh.getRange(rowNum, 3, 1, 2).setNumberFormat('@');

    return { ok: true };
  } catch (e) {
    console.error('updateEvent error:', e);
    return { ok: false, error: e.message || 'UPDATE_FAILED' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * イベント削除
 */
function deleteEvent(eventId) {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { ok: false, error: 'LOCK_TIMEOUT' };
  }

  try {
    const user = getUserContext_();
    const settings = getSettings_();

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_EVENTS);

    if (!sh) {
      throw new Error('SHEET_NOT_FOUND: ' + SHEET_EVENTS);
    }

    const rowNum = findEventRowById_(sh, String(eventId));
    if (!rowNum) {
      return { ok: false, error: 'NOT_FOUND' };
    }

    // 対象イベントの日付を取得して権限チェック
    const d = sh.getRange(rowNum, 2).getValue();
    const iso = formatISODate_(d instanceof Date ? d : new Date(d));

    // 権限チェック（サーバ側で確実に）
    assertCanEdit_(user, settings, iso);

    sh.deleteRow(rowNum);
    return { ok: true };
  } catch (e) {
    console.error('deleteEvent error:', e);
    return { ok: false, error: e.message || 'DELETE_FAILED' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * event_idから行番号を検索
 */
function findEventRowById_(sh, eventId) {
  const lastRow = sh.getLastRow();
  if (lastRow < EVENTS_DATA_START_ROW) return null;

  const numRows = lastRow - EVENTS_DATA_START_ROW + 1;
  const ids = sh.getRange(EVENTS_DATA_START_ROW, 1, numRows, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === eventId) {
      return EVENTS_DATA_START_ROW + i;
    }
  }

  return null;
}

// =============================================================================
// Archive（過去一覧 - 閲覧のみ）
// =============================================================================

/**
 * 指定月のイベント一覧を取得（閲覧のみ）
 */
function listArchive(monthYYYYMM) {
  try {
    const settings = getSettings_();
    const tz = settings.tz;

    if (!/^\d{4}-\d{2}$/.test(monthYYYYMM)) {
      return { ok: false, error: 'INVALID_MONTH' };
    }

    const [y, m] = monthYYYYMM.split('-').map(Number);
    const from = new Date(y, m - 1, 1);
    const to = new Date(y, m, 0); // 月末

    const fromISO = Utilities.formatDate(from, tz, 'yyyy-MM-dd');
    const toISO = Utilities.formatDate(to, tz, 'yyyy-MM-dd');

    return {
      ok: true,
      month: monthYYYYMM,
      events: listEvents_(fromISO, toISO)
    };
  } catch (e) {
    console.error('listArchive error:', e);
    return { ok: false, error: e.message };
  }
}

// =============================================================================
// Guards（権限・バリデーション）
// =============================================================================

/**
 * 編集権限チェック（サーバ側で確実に実施）
 * @throws {Error} 権限がない場合
 */
function assertCanEdit_(user, settings, isoDate) {
  // 認証チェック
  if (!user || !user.email) {
    throw new Error('AUTH_REQUIRED');
  }

  // ロールチェック（viewerは編集不可）
  if (!(user.role === 'editor' || user.role === 'admin')) {
    throw new Error('FORBIDDEN');
  }

  // 編集可能期間チェック
  const r = getEditableRange_(settings);
  const d = new Date(isoDate + 'T00:00:00');
  const from = new Date(r.fromISO + 'T00:00:00');
  const to = new Date(r.toISO + 'T23:59:59');

  if (d < from || d > to) {
    throw new Error('OUT_OF_EDITABLE_RANGE');
  }
}

/**
 * イベントペイロードのバリデーション
 * @throws {Error} 不正な値の場合
 */
function validateEventPayload_(p, settings) {
  if (!p) {
    throw new Error('INVALID_PAYLOAD');
  }

  // 必須: date
  if (!/^\d{4}-\d{2}-\d{2}$/.test(p.date || '')) {
    throw new Error('INVALID_DATE: 日付は YYYY-MM-DD 形式で入力してください');
  }

  // 必須: type
  if (!String(p.type || '').trim()) {
    throw new Error('TYPE_REQUIRED: 予定種別を選択してください');
  }

  // typeがマスタに存在するかチェック
  const masters = getMasters_();
  const validTypes = masters.types.map(t => t.code);
  if (validTypes.length > 0 && !validTypes.includes(p.type)) {
    throw new Error('INVALID_TYPE: 無効な予定種別です');
  }

  // 必須: title
  if (!String(p.title || '').trim()) {
    throw new Error('TITLE_REQUIRED: 件名を入力してください');
  }

  // 任意: start_time（HH:MM形式のみ許可）
  if (p.start_time && !/^\d{2}:\d{2}$/.test(p.start_time)) {
    throw new Error('INVALID_START_TIME: 開始時刻は HH:MM 形式で入力してください');
  }

  // 任意: end_time（HH:MM形式のみ許可）
  if (p.end_time && !/^\d{2}:\d{2}$/.test(p.end_time)) {
    throw new Error('INVALID_END_TIME: 終了時刻は HH:MM 形式で入力してください');
  }

  // 任意: status（マスタに存在するもののみ）
  if (p.status) {
    const validStatuses = masters.statuses.map(s => s.code);
    if (validStatuses.length > 0 && !validStatuses.includes(p.status)) {
      throw new Error('INVALID_STATUS: 無効な状態です');
    }
  }
}

// =============================================================================
// ユーティリティ
// =============================================================================

/**
 * DateをISO日付文字列に変換
 */
function formatISODate_(d) {
  const settings = getSettings_();
  return Utilities.formatDate(d, settings.tz, 'yyyy-MM-dd');
}

/**
 * 時刻値を HH:MM 文字列に変換
 */
function asTimeStr_(v) {
  if (!v) return '';
  if (typeof v === 'string') return v;
  if (v instanceof Date) {
    const settings = getSettings_();
    return Utilities.formatDate(v, settings.tz, 'HH:mm');
  }
  return String(v);
}

/**
 * 翌月キーを取得（サーバ側）
 */
function nextMonthKeyServer_(baseYYYYMM) {
  const [y, m] = baseYYYYMM.split('-').map(Number);
  const d = new Date(y, m - 1, 1);
  d.setMonth(d.getMonth() + 1);
  const settings = getSettings_();
  return Utilities.formatDate(d, settings.tz, 'yyyy-MM');
}

// =============================================================================
// Notes（月単位の備考）
// =============================================================================

/**
 * 05_Notesシートを確保（なければ作成）
 */
function ensureNotesSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NOTES);

  if (!sh) {
    sh = ss.insertSheet(SHEET_NOTES);
    sh.getRange(1, 1, 1, 4).setValues([['月キー', '備考', '更新日時', '更新者']]);
    sh.setFrozenRows(1);
  }

  return sh;
}

/**
 * 指定期間の月備考を取得
 */
function getNotesMap_(fromMonthYYYYMM, toMonthYYYYMM) {
  const sh = ensureNotesSheet_();
  const lastRow = sh.getLastRow();
  const map = {};

  if (lastRow < 2) return map;

  const values = sh.getRange(2, 1, lastRow - 1, 2).getValues(); // A:month, B:notes

  for (const [mk, notes] of values) {
    if (!mk) continue;
    const key = String(mk).trim();
    if (key >= fromMonthYYYYMM && key <= toMonthYYYYMM) {
      map[key] = String(notes || '');
    }
  }

  return map;
}

/**
 * 月備考を追加/更新
 */
function upsertNote(monthKey, text) {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { ok: false, error: 'LOCK_TIMEOUT' };
  }

  try {
    const user = getUserContext_();
    const settings = getSettings_();

    // 権限チェック
    if (!(user.role === 'editor' || user.role === 'admin')) {
      throw new Error('FORBIDDEN');
    }

    if (!/^\d{4}-\d{2}$/.test(monthKey)) {
      throw new Error('INVALID_MONTH');
    }

    // 当月・翌月のみ編集可
    const base = settings.baseMonth;
    const next = nextMonthKeyServer_(base);
    if (!(monthKey === base || monthKey === next)) {
      throw new Error('OUT_OF_EDITABLE_RANGE');
    }

    const sh = ensureNotesSheet_();
    const lastRow = sh.getLastRow();
    const tz = settings.tz;
    const now = new Date();

    // 既存行を探索
    if (lastRow >= 2) {
      const keys = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) {
        if (String(keys[i][0]).trim() === monthKey) {
          // 更新
          sh.getRange(2 + i, 2).setValue(String(text || ''));
          sh.getRange(2 + i, 3).setValue(Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'));
          sh.getRange(2 + i, 4).setValue(user.email);
          return { ok: true };
        }
      }
    }

    // 新規追加
    sh.appendRow([
      monthKey,
      String(text || ''),
      Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'),
      user.email
    ]);

    return { ok: true };
  } catch (e) {
    console.error('upsertNote error:', e);
    return { ok: false, error: e.message || 'SAVE_FAILED' };
  } finally {
    lock.releaseLock();
  }
}
