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
const SHEET_DAY_SETTINGS = '04_DaySettings';  // 出勤日・休日設定
const SHEET_NOTES    = '05_Notes';

const SETTINGS_BASE_MONTH_CELL = 'B5';      // YYYY-MM
const SETTINGS_EDITABLE_MONTHS_CELL = 'B6'; // 2
const SETTINGS_TZ_CELL = 'B7';              // Asia/Tokyo

const EVENTS_HEADER_ROW = 5;
const EVENTS_DATA_START_ROW = 6;

// Events列構成:
// A(1): event_id
// B(2): date
// C(3): start_time
// D(4): end_time
// E(5): type
// F(6): title
// G(7): location
// H(8): memo
// I(9): display_order
// J(10): status
// K(11): month_key
// L(12): created_at
// M(13): created_by
// N(14): updated_at
// O(15): updated_by
// P(16): gcal_event_id
// Q(17): gcal_calendar_id
// R(18): last_sync_at
// S(19): sync_source
// T(20): revision (applyPatch用)

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

    // 当月〜翌月の日付設定（出勤日・休日）を取得
    const daySettings = getDaySettingsMap_(range.fromISO, range.toISO);

    return {
      ok: true,
      user,
      settings,
      masters,
      range,
      events,
      notes,
      daySettings
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
  const editableMonths = Number(sh.getRange(SETTINGS_EDITABLE_MONTHS_CELL).getValue() || 2);

  // 常に今月を基準月として使用（自動進行）
  const baseMonth = Utilities.formatDate(new Date(), tz, 'yyyy-MM');

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
      user.email
    ]]);
    // 時刻列（C, D）をテキスト形式に設定
    sh.getRange(newRow, 3, 1, 2).setNumberFormat('@');

    // Googleカレンダーへ同期（非同期的に実行、エラーでも保存は成功扱い）
    try {
      const eventData = {
        event_id: eventId,
        date: payload.date,
        start_time: startTimeText,
        end_time: endTimeText,
        type: payload.type || '',
        title: payload.title || '',
        location: payload.location || '',
        memo: payload.memo || '',
        status: payload.status || 'CONFIRMED',
        is_all_day: !startTimeText && !endTimeText,
      };
      syncEventToGcalAfterCreate_(eventId, eventData);
    } catch (syncErr) {
      console.log('GCal sync error (non-fatal):', syncErr);
    }

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

    // 既存のgcal_event_idを取得
    const existingGcalEventId = sh.getRange(rowNum, 16).getValue() || '';

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

    // Googleカレンダーへ同期（非同期的に実行、エラーでも保存は成功扱い）
    try {
      const eventData = {
        event_id: eventId,
        date: payload.date,
        start_time: startTimeText,
        end_time: endTimeText,
        type: payload.type || '',
        title: payload.title || '',
        location: payload.location || '',
        memo: payload.memo || '',
        status: payload.status || 'CONFIRMED',
        is_all_day: !startTimeText && !endTimeText,
        gcal_event_id: existingGcalEventId,
      };
      syncEventToGcalAfterCreate_(eventId, eventData);
    } catch (syncErr) {
      console.log('GCal sync error (non-fatal):', syncErr);
    }

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

    // Googleカレンダーから削除（行削除前にgcal_event_idを取得）
    try {
      const gcalEventId = sh.getRange(rowNum, 16).getValue();
      if (gcalEventId) {
        const syncSettings = getGcalSyncSettings_();
        const calendarId = syncSettings[GCAL_SETTINGS_KEYS.CALENDAR_ID];
        const enabled = syncSettings[GCAL_SETTINGS_KEYS.SYNC_ENABLED];

        if ((enabled === true || enabled === 'TRUE' || enabled === 'true') && calendarId) {
          try {
            Calendar.Events.remove(calendarId, gcalEventId);
            console.log('GCal event deleted:', gcalEventId);
          } catch (gcalErr) {
            console.log('GCal delete error (may already deleted):', gcalErr.message);
          }
        }
      }
    } catch (syncErr) {
      console.log('GCal sync error (non-fatal):', syncErr);
    }

    sh.deleteRow(rowNum);
    return { ok: true };
  } catch (e) {
    console.error('deleteEvent error:', e);
    return { ok: false, error: e.message || 'DELETE_FAILED' };
  } finally {
    lock.releaseLock();
  }
}

// =============================================================================
// A＋C＋B自動保存用 パッチ適用API
// =============================================================================

/**
 * パッチを適用（upsert）- 自動保存用
 * @param {string} recordId - レコードID（新規はtmp-xxxの場合あり）
 * @param {Object} patch - 変更フィールド
 * @param {number} clientRevision - クライアント側のリビジョン
 * @returns {Object} {ok, recordId?, serverRevision?, error?}
 */
function applyPatch(recordId, patch, clientRevision) {
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
      return { ok: false, error: 'FORBIDDEN' };
    }

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_EVENTS);
    if (!sh) {
      throw new Error('SHEET_NOT_FOUND: ' + SHEET_EVENTS);
    }

    const tz = settings.tz;
    const now = new Date();
    const isNew = recordId.startsWith('tmp-');

    if (isNew) {
      // ========== 新規作成 ==========
      // 日付の権限チェック
      if (patch.date) {
        assertCanEdit_(user, settings, patch.date);
      }

      const newId = Utilities.getUuid();
      const newRevision = 1;

      // 時刻はテキスト形式で保存
      const startTimeText = patch.start_time ? String(patch.start_time) : '';
      const endTimeText = patch.end_time ? String(patch.end_time) : '';

      const newRow = sh.getLastRow() + 1;
      // A〜O列（15列）に書き込み
      sh.getRange(newRow, 1, 1, 15).setValues([[
        newId,
        patch.date ? new Date(patch.date + 'T00:00:00') : '',
        startTimeText,
        endTimeText,
        patch.type || '',
        patch.title || '',
        patch.location || '',
        patch.memo || '',
        patch.display_order !== undefined && patch.display_order !== '' ? Number(patch.display_order) : '',
        patch.status || 'CONFIRMED',
        '', // 月キー（K列）
        Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'),
        user.email,
        Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'),
        user.email
      ]]);
      // 時刻列をテキスト形式に
      sh.getRange(newRow, 3, 1, 2).setNumberFormat('@');
      // T列(20)にリビジョンを書き込み（P-S列はGCal同期用）
      sh.getRange(newRow, 20).setValue(newRevision);

      return { ok: true, recordId: newId, serverRevision: newRevision };

    } else {
      // ========== 既存更新 ==========
      const rowNum = findEventRowById_(sh, String(recordId));
      if (!rowNum) {
        return { ok: false, error: 'NOT_FOUND' };
      }

      // 日付の権限チェック（変更後の日付でチェック）
      const currentDate = sh.getRange(rowNum, 2).getValue();
      const targetDate = patch.date || formatISODate_(currentDate instanceof Date ? currentDate : new Date(currentDate));
      assertCanEdit_(user, settings, targetDate);

      // サーバ側リビジョン取得（T列(20)、なければ0）
      const serverRevision = Number(sh.getRange(rowNum, 20).getValue()) || 0;

      // リビジョンチェック（クライアントが古ければ競合）
      if (clientRevision < serverRevision) {
        return { ok: false, error: 'CONFLICT', serverRevision };
      }

      // パッチ適用
      const colMap = {
        date: 2,
        start_time: 3,
        end_time: 4,
        type: 5,
        title: 6,
        location: 7,
        memo: 8,
        display_order: 9,
        status: 10
      };

      for (const [field, value] of Object.entries(patch)) {
        const col = colMap[field];
        if (col) {
          if (field === 'date' && value) {
            sh.getRange(rowNum, col).setValue(new Date(value + 'T00:00:00'));
          } else if (field === 'start_time' || field === 'end_time') {
            sh.getRange(rowNum, col).setValue(value ? String(value) : '');
            sh.getRange(rowNum, col).setNumberFormat('@');
          } else if (field === 'display_order') {
            sh.getRange(rowNum, col).setValue(value !== '' ? Number(value) : '');
          } else {
            sh.getRange(rowNum, col).setValue(value || '');
          }
        }
      }

      // 更新日時・更新者・リビジョン更新
      const newRevision = serverRevision + 1;
      sh.getRange(rowNum, 14).setValue(Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'));
      sh.getRange(rowNum, 15).setValue(user.email);
      sh.getRange(rowNum, 20).setValue(newRevision); // T列(20)にリビジョン

      return { ok: true, serverRevision: newRevision };
    }

  } catch (e) {
    console.error('applyPatch error:', e);
    return { ok: false, error: e.message || 'PATCH_FAILED' };
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

// =============================================================================
// DaySettings（出勤日・休日設定）
// =============================================================================

/**
 * 04_DaySettingsシートを確保（なければ作成）
 */
function ensureDaySettingsSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_DAY_SETTINGS);

  if (!sh) {
    sh = ss.insertSheet(SHEET_DAY_SETTINGS);
    sh.getRange(1, 1, 1, 5).setValues([['日付', '種別', 'メモ', '更新日時', '更新者']]);
    sh.setFrozenRows(1);
  }

  return sh;
}

/**
 * 指定期間の日付設定を取得
 * @returns {Object} { 'YYYY-MM-DD': 'WORKDAY'|'HOLIDAY' }
 */
function getDaySettingsMap_(fromISO, toISO) {
  const sh = ensureDaySettingsSheet_();
  const lastRow = sh.getLastRow();
  const map = {};

  if (lastRow < 2) return map;

  const settings = getSettings_();
  const tz = settings.tz;
  const values = sh.getRange(2, 1, lastRow - 1, 2).getValues(); // A:date, B:type

  for (const [dateVal, type] of values) {
    if (!dateVal) continue;
    let key;
    if (dateVal instanceof Date) {
      // Date型の場合はタイムゾーンを考慮して変換
      key = Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
    } else {
      // 文字列の場合はそのまま使用（YYYY-MM-DD形式を期待）
      key = String(dateVal).trim();
    }
    const typeStr = String(type || '').trim();
    // 範囲チェックを緩めて、有効なデータは全て取得
    if (typeStr && key.match(/^\d{4}-\d{2}-\d{2}$/)) {
      map[key] = typeStr;  // 'WORKDAY' or 'HOLIDAY'
    }
  }

  console.log('getDaySettingsMap_ result:', JSON.stringify(map));
  return map;
}

/**
 * 日付設定を追加/更新/削除
 * @param {string} dateISO - YYYY-MM-DD
 * @param {string|null} type - 'WORKDAY'（出勤日）, 'HOLIDAY'（休日）, null（解除）
 * @param {string} memo - メモ（任意）
 */
function setDaySetting(dateISO, type, memo) {
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

    // 編集可能期間チェック
    assertCanEdit_(user, settings, dateISO);

    const sh = ensureDaySettingsSheet_();
    const lastRow = sh.getLastRow();
    const tz = settings.tz;
    const now = new Date();

    // 既存行を探索
    let existingRow = null;
    if (lastRow >= 2) {
      const dates = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < dates.length; i++) {
        let key;
        if (dates[i][0] instanceof Date) {
          key = formatISODate_(dates[i][0]);
        } else {
          key = String(dates[i][0]).trim();
        }
        if (key === dateISO) {
          existingRow = 2 + i;
          break;
        }
      }
    }

    if (type === null || type === '') {
      // 解除: 行を削除
      if (existingRow) {
        sh.deleteRow(existingRow);
      }
      return { ok: true, action: 'removed' };
    }

    if (existingRow) {
      // 更新
      sh.getRange(existingRow, 2).setValue(type);
      sh.getRange(existingRow, 3).setValue(memo || '');
      sh.getRange(existingRow, 4).setValue(Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'));
      sh.getRange(existingRow, 5).setValue(user.email);
      return { ok: true, action: 'updated' };
    }

    // 新規追加
    sh.appendRow([
      dateISO,
      type,
      memo || '',
      Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'),
      user.email
    ]);

    return { ok: true, action: 'created' };
  } catch (e) {
    console.error('setDaySetting error:', e);
    return { ok: false, error: e.message || 'SAVE_FAILED' };
  } finally {
    lock.releaseLock();
  }
}

// =============================================================================
// Googleカレンダー双方向同期
// =============================================================================

/**
 * 同期用定数
 */
const GCAL_SETTINGS_KEYS = {
  CALENDAR_ID: 'gcal_calendar_id',
  SYNC_ENABLED: 'gcal_sync_enabled',
  RANGE_START: 'gcal_sync_range_start',
  RANGE_END: 'gcal_sync_range_end',
  LAST_SYNC_TOKEN: 'gcal_sync_token',
  LAST_SYNC_AT: 'gcal_last_sync_at',
  LOOP_GUARD_SECONDS: 'gcal_loop_guard_seconds',
};

const SYNC_SOURCE = {
  WEBAPP: 'WEBAPP',
  GCAL: 'GCAL',
  SYSTEM: 'SYSTEM',
};

/**
 * 同期設定を取得
 */
function getGcalSyncSettings_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  if (!sh) return {};

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  const map = {};
  values.forEach(([k, v]) => {
    if (k) map[String(k).trim()] = v;
  });
  return map;
}

/**
 * 同期設定を保存
 */
function setGcalSyncSetting_(key, value) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    sh.getRange(2, 1, 1, 2).setValues([[key, value]]);
    return;
  }

  const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sh.getRange(2 + i, 2).setValue(value);
      return;
    }
  }
  // 無ければ末尾追加
  sh.getRange(lastRow + 1, 1, 1, 2).setValues([[key, value]]);
}

/**
 * イベントの同期用ハッシュを生成
 */
function buildSyncHash_(e) {
  const payload = {
    date: e.date || '',
    start_time: e.start_time || '',
    end_time: e.end_time || '',
    is_all_day: !!e.is_all_day,
    title: e.title || '',
    location: e.location || '',
    memo: e.memo || '',
    type: e.type || '',
    status: e.status || '',
  };
  const s = JSON.stringify(payload);
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

/**
 * Sheet予定をGCalイベントリソースに変換（新規作成用）
 */
function buildGcalEventResource_(e, forUpdate) {
  const tz = Session.getScriptTimeZone();

  const resource = {
    summary: e.title || '（件名なし）',
    location: e.location || '',
    description: e.memo || '',
  };

  // 新規作成時のみextendedPropertiesを設定（更新時は既存を保持）
  if (!forUpdate) {
    resource.extendedProperties = {
      private: {
        lw_event_id: String(e.event_id || ''),
        lw_origin: 'LW',
        lw_type: e.type || '',
      }
    };
  }

  // 終日 or 時間あり
  const isAllDay = e.is_all_day || (!e.start_time && !e.end_time);

  if (isAllDay) {
    // 終日：end.date は翌日が仕様
    const startDate = String(e.date);
    if (!startDate || !/^\d{4}-\d{2}-\d{2}$/.test(startDate)) {
      console.log('Invalid date format:', startDate);
      return null;
    }
    const endDateObj = new Date(startDate + 'T00:00:00');
    endDateObj.setDate(endDateObj.getDate() + 1);
    resource.start = { date: startDate };
    resource.end = { date: formatISODate_(endDateObj) };
  } else {
    const dateStr = String(e.date);
    if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      console.log('Invalid date format:', dateStr);
      return null;
    }
    const startTime = e.start_time || '00:00';
    const endTime = e.end_time || startTime;
    const start = `${dateStr}T${startTime}:00`;
    const end = `${dateStr}T${endTime}:00`;
    resource.start = { dateTime: start, timeZone: tz };
    resource.end = { dateTime: end, timeZone: tz };
  }

  // 種別→色（オプション）
  const colorMap = {
    'MEET': '9',      // 青
    'MEETING': '9',
    'VIP': '11',      // 赤
    'VISIT': '6',     // オレンジ
    'OUT': '2',       // 緑
    'TRAVEL': '5',    // 黄
    'DINNER': '3',    // 紫
    'HOLIDAY': '8',   // グレー
  };
  if (e.type && colorMap[e.type]) {
    resource.colorId = colorMap[e.type];
  }

  return resource;
}

/**
 * lw_event_idでGCalイベントを検索
 */
function findGcalEventByLwId_(calendarId, lwEventId) {
  if (!lwEventId) return null;

  try {
    // 直近3ヶ月を検索範囲とする
    const now = new Date();
    const timeMin = new Date(now.getFullYear(), now.getMonth() - 1, 1).toISOString();
    const timeMax = new Date(now.getFullYear(), now.getMonth() + 2, 0).toISOString();

    const events = Calendar.Events.list(calendarId, {
      privateExtendedProperty: 'lw_event_id=' + lwEventId,
      timeMin: timeMin,
      timeMax: timeMax,
      singleEvents: true,
      maxResults: 1
    });

    if (events.items && events.items.length > 0) {
      return events.items[0];
    }
  } catch (err) {
    console.log('Error searching for event by lw_event_id:', err.message);
  }
  return null;
}

/**
 * GCalイベントを安全に更新（get→update方式）
 */
function updateGcalEventSafe_(calendarId, gcalEventId, sheetEvent) {
  try {
    // 既存イベントを取得
    const existing = Calendar.Events.get(calendarId, gcalEventId);

    // 基本情報を更新
    existing.summary = sheetEvent.title || '（件名なし）';
    existing.location = sheetEvent.location || '';
    existing.description = sheetEvent.memo || '';

    // 日時を更新（形式を合わせる）
    const tz = Session.getScriptTimeZone();
    const isAllDay = sheetEvent.is_all_day || (!sheetEvent.start_time && !sheetEvent.end_time);
    const dateStr = String(sheetEvent.date);

    if (isAllDay) {
      const endDateObj = new Date(dateStr + 'T00:00:00');
      endDateObj.setDate(endDateObj.getDate() + 1);
      existing.start = { date: dateStr };
      existing.end = { date: formatISODate_(endDateObj) };
    } else {
      const startTime = sheetEvent.start_time || '00:00';
      const endTime = sheetEvent.end_time || startTime;
      existing.start = { dateTime: `${dateStr}T${startTime}:00`, timeZone: tz };
      existing.end = { dateTime: `${dateStr}T${endTime}:00`, timeZone: tz };
    }

    // 色を更新
    const colorMap = {
      'MEET': '9', 'MEETING': '9', 'VIP': '11', 'VISIT': '6',
      'OUT': '2', 'TRAVEL': '5', 'DINNER': '3', 'HOLIDAY': '8',
    };
    if (sheetEvent.type && colorMap[sheetEvent.type]) {
      existing.colorId = colorMap[sheetEvent.type];
    }

    // 更新を実行
    const updated = Calendar.Events.update(existing, calendarId, gcalEventId);
    return { success: true, id: updated.id };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * 1件のSheet予定をGCalへ同期（作成/更新/削除）
 */
function syncOneEventToGcal_(e, calendarId) {
  if (!calendarId) return { gcal_event_id: '' };

  // 入力データの検証
  if (!e.date || !/^\d{4}-\d{2}-\d{2}$/.test(String(e.date))) {
    console.log('Skipping event with invalid date:', e.event_id, e.date);
    return { gcal_event_id: e.gcal_event_id || '', error: 'Invalid date' };
  }

  // 削除フラグの場合
  if (e.is_deleted) {
    if (e.gcal_event_id) {
      try {
        Calendar.Events.remove(calendarId, e.gcal_event_id);
        console.log('GCal event deleted:', e.gcal_event_id);
      } catch (err) {
        console.log('GCal delete error (may already deleted):', err.message);
      }
    }
    return { gcal_event_id: e.gcal_event_id || '' };
  }

  // gcal_event_idがある場合は更新を試みる
  if (e.gcal_event_id) {
    const updateResult = updateGcalEventSafe_(calendarId, e.gcal_event_id, e);
    if (updateResult.success) {
      console.log('GCal event updated:', updateResult.id);
      return { gcal_event_id: updateResult.id };
    }

    // 404エラーの場合のみ、検索して再紐付けを試みる
    if (updateResult.error && updateResult.error.includes('404')) {
      console.log('GCal event not found (404), searching by lw_event_id...');
      const existing = findGcalEventByLwId_(calendarId, e.event_id);
      if (existing) {
        const retryResult = updateGcalEventSafe_(calendarId, existing.id, e);
        if (retryResult.success) {
          console.log('GCal event updated (found by lw_id):', retryResult.id);
          return { gcal_event_id: retryResult.id };
        }
      }
    } else {
      console.log('GCal update error:', updateResult.error);
    }
  }

  // lw_event_idで既存イベントを検索（重複防止）
  const existing = findGcalEventByLwId_(calendarId, e.event_id);
  if (existing) {
    console.log('Found existing GCal event by lw_event_id:', existing.id);
    const updateResult = updateGcalEventSafe_(calendarId, existing.id, e);
    if (updateResult.success) {
      console.log('GCal event updated (found):', updateResult.id);
      return { gcal_event_id: updateResult.id };
    }
  }

  // 新規作成
  const resource = buildGcalEventResource_(e, false);
  if (!resource) {
    return { gcal_event_id: '', error: 'Invalid event data' };
  }

  try {
    const created = Calendar.Events.insert(resource, calendarId);
    console.log('GCal event created:', created.id);
    return { gcal_event_id: created.id };
  } catch (err) {
    console.error('GCal insert error:', err.message, 'Event:', e.event_id);
    return { gcal_event_id: '', error: err.message };
  }
}

/**
 * イベント作成後にGCalへ同期（createEventから呼び出し）
 */
function syncEventToGcalAfterCreate_(eventId, eventData) {
  const syncSettings = getGcalSyncSettings_();
  const enabled = syncSettings[GCAL_SETTINGS_KEYS.SYNC_ENABLED];
  const calendarId = syncSettings[GCAL_SETTINGS_KEYS.CALENDAR_ID];

  if (enabled !== true && enabled !== 'TRUE' && enabled !== 'true') {
    console.log('GCal sync disabled');
    return null;
  }
  if (!calendarId) {
    console.log('GCal calendar ID not set');
    return null;
  }

  const result = syncOneEventToGcal_(eventData, calendarId);

  // gcal_event_id をシートに保存
  if (result.gcal_event_id) {
    updateEventGcalId_(eventId, result.gcal_event_id, calendarId);
  }

  return result;
}

/**
 * イベントのgcal_event_idを更新
 */
function updateEventGcalId_(eventId, gcalEventId, calendarId) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_EVENTS);
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < EVENTS_DATA_START_ROW) return;

  const ids = sh.getRange(EVENTS_DATA_START_ROW, 1, lastRow - EVENTS_DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(eventId).trim()) {
      const row = EVENTS_DATA_START_ROW + i;
      // P列(16): gcal_event_id, Q列(17): gcal_calendar_id, R列(18): last_sync_at, S列(19): sync_source
      sh.getRange(row, 16).setValue(gcalEventId);
      sh.getRange(row, 17).setValue(calendarId);
      sh.getRange(row, 18).setValue(new Date());
      sh.getRange(row, 19).setValue(SYNC_SOURCE.WEBAPP);
      console.log('Updated gcal_event_id for event:', eventId, '->', gcalEventId);
      return;
    }
  }
}

/**
 * GCalイベントをSheet形式に変換
 */
function gcalEventToSheetData_(ev, calendarId) {
  const tz = Session.getScriptTimeZone();
  const isDeleted = (ev.status === 'cancelled');

  const data = {
    gcal_event_id: ev.id,
    gcal_calendar_id: calendarId,
    title: ev.summary || '',
    location: ev.location || '',
    memo: ev.description || '',
    is_deleted: isDeleted,
    last_modified_at: ev.updated || new Date().toISOString(),
    last_modified_by: (ev.creator && ev.creator.email) ? ev.creator.email : '',
  };

  // 日付/時刻
  if (ev.start && ev.start.date) {
    data.is_all_day = true;
    data.date = ev.start.date;
    data.start_time = '';
    data.end_time = '';
  } else if (ev.start && ev.start.dateTime) {
    data.is_all_day = false;
    const d = new Date(ev.start.dateTime);
    data.date = formatISODate_(d);
    data.start_time = Utilities.formatDate(d, tz, 'HH:mm');
    if (ev.end && ev.end.dateTime) {
      const e = new Date(ev.end.dateTime);
      data.end_time = Utilities.formatDate(e, tz, 'HH:mm');
    } else {
      data.end_time = '';
    }
  }

  // extendedPropertiesからtype等を復元
  const priv = (ev.extendedProperties && ev.extendedProperties.private) || {};
  if (priv.lw_type) {
    data.type = priv.lw_type;
  }

  return data;
}

/**
 * GCal→Sheet 定期同期（トリガーで実行）
 */
function syncFromGcalToSheet() {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    console.log('Could not obtain lock for GCal sync');
    return;
  }

  try {
    const syncSettings = getGcalSyncSettings_();
    const enabled = syncSettings[GCAL_SETTINGS_KEYS.SYNC_ENABLED];
    const calendarId = syncSettings[GCAL_SETTINGS_KEYS.CALENDAR_ID];

    if (enabled !== true && enabled !== 'TRUE' && enabled !== 'true') {
      console.log('GCal sync disabled');
      return;
    }
    if (!calendarId) {
      console.log('GCal calendar ID not set');
      return;
    }

    const settings = getSettings_();
    const loopGuard = Number(syncSettings[GCAL_SETTINGS_KEYS.LOOP_GUARD_SECONDS] || 30);

    // 同期範囲を計算（基準月〜翌月末+7日）
    const [y, m] = settings.baseMonth.split('-').map(Number);
    const rangeStart = new Date(y, m - 1, 1);
    const rangeEnd = new Date(y, m + 1, 7); // 翌月末+7日程度

    const timeMin = rangeStart.toISOString();
    const timeMax = rangeEnd.toISOString();

    console.log('GCal sync range:', timeMin, '-', timeMax);

    // Calendar API で取得
    const collected = [];
    let pageToken = null;
    const syncToken = syncSettings[GCAL_SETTINGS_KEYS.LAST_SYNC_TOKEN] || '';

    try {
      do {
        const options = {
          singleEvents: true,
          showDeleted: true,
          timeMin: timeMin,
          timeMax: timeMax,
          maxResults: 500,
        };
        if (pageToken) options.pageToken = pageToken;
        if (syncToken) options.syncToken = syncToken;

        const res = Calendar.Events.list(calendarId, options);
        (res.items || []).forEach(it => collected.push(it));
        pageToken = res.nextPageToken;
        if (res.nextSyncToken) {
          setGcalSyncSetting_(GCAL_SETTINGS_KEYS.LAST_SYNC_TOKEN, res.nextSyncToken);
        }
      } while (pageToken);
    } catch (err) {
      // syncToken失効時はクリアして再取得
      console.log('SyncToken expired, clearing and retrying...');
      setGcalSyncSetting_(GCAL_SETTINGS_KEYS.LAST_SYNC_TOKEN, '');
      pageToken = null;
      do {
        const options = {
          singleEvents: true,
          showDeleted: true,
          timeMin: timeMin,
          timeMax: timeMax,
          maxResults: 500,
        };
        if (pageToken) options.pageToken = pageToken;

        const res = Calendar.Events.list(calendarId, options);
        (res.items || []).forEach(it => collected.push(it));
        pageToken = res.nextPageToken;
        if (res.nextSyncToken) {
          setGcalSyncSetting_(GCAL_SETTINGS_KEYS.LAST_SYNC_TOKEN, res.nextSyncToken);
        }
      } while (pageToken);
    }

    console.log('GCal events collected:', collected.length);

    if (collected.length === 0) {
      setGcalSyncSetting_(GCAL_SETTINGS_KEYS.LAST_SYNC_AT, new Date().toISOString());
      return;
    }

    // Sheet側のインデックス作成
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_EVENTS);
    const lastRow = sh.getLastRow();

    const idxByGcal = new Map();  // gcal_event_id -> row
    const idxById = new Map();    // event_id -> row
    const eventDataMap = new Map(); // event_id -> data

    if (lastRow >= EVENTS_DATA_START_ROW) {
      const values = sh.getRange(EVENTS_DATA_START_ROW, 1, lastRow - EVENTS_DATA_START_ROW + 1, 19).getValues();
      values.forEach((row, i) => {
        const rowNum = EVENTS_DATA_START_ROW + i;
        const eventId = row[0];   // A列
        const gcalId = row[15];   // P列 (gcal_event_id)
        const lastModified = row[13]; // N列 (updated_at)

        if (eventId) {
          idxById.set(String(eventId), rowNum);
          eventDataMap.set(String(eventId), {
            last_modified_at: lastModified,
            sync_source: row[18], // S列
          });
        }
        if (gcalId) {
          idxByGcal.set(String(gcalId), rowNum);
        }
      });
    }

    const now = new Date();
    const lastSyncAt = syncSettings[GCAL_SETTINGS_KEYS.LAST_SYNC_AT];
    const tz = settings.tz;

    // 各イベントを処理
    let updated = 0;
    let created = 0;
    let skipped = 0;

    collected.forEach(ev => {
      const evId = ev.id;
      const gcalUpdated = ev.updated ? new Date(ev.updated) : null;

      // extendedPropertiesからlw_event_idを取得
      const priv = (ev.extendedProperties && ev.extendedProperties.private) || {};
      const lwId = priv.lw_event_id || '';
      const lwOrigin = priv.lw_origin || '';

      // LW由来のイベントは、最近同期したばかりならスキップ（ループ防止）
      if (lwOrigin === 'LW' && lastSyncAt) {
        const dt = (now.getTime() - new Date(lastSyncAt).getTime()) / 1000;
        if (dt < loopGuard) {
          skipped++;
          return;
        }
      }

      // 紐付け
      let rowNum = null;
      if (lwId && idxById.has(lwId)) {
        rowNum = idxById.get(lwId);
      } else if (evId && idxByGcal.has(evId)) {
        rowNum = idxByGcal.get(evId);
      }

      const gcalData = gcalEventToSheetData_(ev, calendarId);

      if (rowNum) {
        // 既存レコード → 衝突チェック（Last Write Wins）
        const eventId = lwId || sh.getRange(rowNum, 1).getValue();
        const existing = eventDataMap.get(String(eventId));
        const sheetLast = existing && existing.last_modified_at ? new Date(existing.last_modified_at) : null;

        if (sheetLast && gcalUpdated && sheetLast.getTime() > gcalUpdated.getTime()) {
          // Sheetの方が新しい → スキップ
          skipped++;
          return;
        }

        // Sheetを更新
        // B列:日付, C列:開始, D列:終了, F列:件名, G列:場所, H列:メモ, N列:更新日時, O列:更新者, R列:last_sync_at, S列:sync_source
        if (gcalData.date) {
          sh.getRange(rowNum, 2).setValue(new Date(gcalData.date + 'T00:00:00'));
        }
        if (gcalData.start_time !== undefined) {
          sh.getRange(rowNum, 3).setValue(gcalData.start_time);
        }
        if (gcalData.end_time !== undefined) {
          sh.getRange(rowNum, 4).setValue(gcalData.end_time);
        }
        sh.getRange(rowNum, 6).setValue(gcalData.title);
        sh.getRange(rowNum, 7).setValue(gcalData.location);
        sh.getRange(rowNum, 8).setValue(gcalData.memo);
        sh.getRange(rowNum, 14).setValue(gcalData.last_modified_at);
        sh.getRange(rowNum, 15).setValue(gcalData.last_modified_by);
        sh.getRange(rowNum, 18).setValue(new Date());
        sh.getRange(rowNum, 19).setValue(SYNC_SOURCE.GCAL);

        updated++;
      } else {
        // 新規取り込み（LW由来でないイベント）
        if (lwOrigin === 'LW') {
          // LW由来だがSheetに無い → 削除された可能性、スキップ
          skipped++;
          return;
        }

        // GCalから新規取り込み
        const newId = Utilities.getUuid();
        const newRow = sh.getLastRow() + 1;

        sh.getRange(newRow, 1, 1, 19).setValues([[
          newId,                                              // A: event_id
          gcalData.date ? new Date(gcalData.date + 'T00:00:00') : '', // B: date
          gcalData.start_time || '',                         // C: start_time
          gcalData.end_time || '',                           // D: end_time
          gcalData.type || 'OTHER',                          // E: type
          gcalData.title || '',                              // F: title
          gcalData.location || '',                           // G: location
          gcalData.memo || '',                               // H: memo
          '',                                                 // I: display_order
          'CONFIRMED',                                        // J: status
          '',                                                 // K: month_key
          Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'), // L: created_at
          gcalData.last_modified_by || '',                   // M: created_by
          gcalData.last_modified_at || '',                   // N: updated_at
          gcalData.last_modified_by || '',                   // O: updated_by
          gcalData.gcal_event_id,                            // P: gcal_event_id
          gcalData.gcal_calendar_id,                         // Q: gcal_calendar_id
          new Date(),                                         // R: last_sync_at
          SYNC_SOURCE.GCAL                                   // S: sync_source
        ]]);

        // GCal側にlw_event_idを保存
        try {
          Calendar.Events.patch({
            extendedProperties: { private: { lw_event_id: newId, lw_origin: 'LW' } }
          }, calendarId, evId);
        } catch (e) {
          console.log('Could not update extendedProperties:', e.message);
        }

        created++;
      }
    });

    console.log('GCal sync completed - updated:', updated, 'created:', created, 'skipped:', skipped);
    setGcalSyncSetting_(GCAL_SETTINGS_KEYS.LAST_SYNC_AT, now.toISOString());

  } catch (e) {
    console.error('syncFromGcalToSheet error:', e);
  } finally {
    lock.releaseLock();
  }
}

/**
 * GCal同期トリガーをインストール
 */
function installGcalSyncTrigger() {
  // 既存トリガー削除
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncFromGcalToSheet') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 5分おきに実行
  ScriptApp.newTrigger('syncFromGcalToSheet')
    .timeBased()
    .everyMinutes(5)
    .create();

  console.log('GCal sync trigger installed (every 5 minutes)');
  return { ok: true, message: 'Trigger installed' };
}

/**
 * GCal同期トリガーを削除
 */
function uninstallGcalSyncTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncFromGcalToSheet') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  console.log('GCal sync trigger removed:', removed);
  return { ok: true, removed: removed };
}

/**
 * GCal同期の初期設定をSettingsシートに追加
 */
function setupGcalSyncSettings(calendarId) {
  if (!calendarId) {
    return { ok: false, error: 'Calendar ID is required' };
  }

  setGcalSyncSetting_(GCAL_SETTINGS_KEYS.CALENDAR_ID, calendarId);
  setGcalSyncSetting_(GCAL_SETTINGS_KEYS.SYNC_ENABLED, 'TRUE');
  setGcalSyncSetting_(GCAL_SETTINGS_KEYS.LOOP_GUARD_SECONDS, 30);

  console.log('GCal sync settings configured for calendar:', calendarId);
  return { ok: true, calendarId: calendarId };
}

/**
 * シートのgcal_event_idをすべてクリア
 */
function clearAllGcalEventIds() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_EVENTS);
  const lastRow = sh.getLastRow();

  if (lastRow < EVENTS_DATA_START_ROW) {
    return { ok: true, cleared: 0 };
  }

  const numRows = lastRow - EVENTS_DATA_START_ROW + 1;
  // P〜S列（16〜19）をクリア
  sh.getRange(EVENTS_DATA_START_ROW, 16, numRows, 4).clearContent();

  console.log('Cleared gcal_event_id for', numRows, 'rows');
  return { ok: true, cleared: numRows };
}

/**
 * GCal上のLW由来イベントをすべて削除
 */
function deleteAllLwEventsFromGcal() {
  const syncSettings = getGcalSyncSettings_();
  const calendarId = syncSettings[GCAL_SETTINGS_KEYS.CALENDAR_ID];

  if (!calendarId) {
    return { ok: false, error: 'Calendar ID not set' };
  }

  const settings = getSettings_();
  const [y, m] = settings.baseMonth.split('-').map(Number);
  const timeMin = new Date(y, m - 1, 1).toISOString();
  const timeMax = new Date(y, m + 2, 0).toISOString();

  // GCalから全イベント取得
  const events = [];
  let pageToken = null;
  do {
    const options = {
      singleEvents: true,
      timeMin: timeMin,
      timeMax: timeMax,
      maxResults: 500,
    };
    if (pageToken) options.pageToken = pageToken;

    const res = Calendar.Events.list(calendarId, options);
    (res.items || []).forEach(it => events.push(it));
    pageToken = res.nextPageToken;
  } while (pageToken);

  // LW由来のイベントを削除
  let deleted = 0;
  events.forEach(ev => {
    const priv = (ev.extendedProperties && ev.extendedProperties.private) || {};
    if (priv.lw_origin === 'LW' || priv.lw_event_id) {
      try {
        Calendar.Events.remove(calendarId, ev.id);
        console.log('Deleted LW event:', ev.id, ev.summary);
        deleted++;
      } catch (err) {
        console.log('Could not delete:', err.message);
      }
    }
  });

  console.log('Deleted', deleted, 'LW events from GCal');
  return { ok: true, deleted: deleted };
}

/**
 * 完全リセット＆再同期（重複問題を解消）
 * 1. GCalからLW由来イベントを削除
 * 2. シートのgcal_event_idをクリア
 * 3. シート→GCalへ新規同期
 */
function fullResetAndResync() {
  console.log('=== Starting full reset and resync ===');

  // Step 1: GCalからLW由来イベントを削除
  console.log('Step 1: Deleting LW events from GCal...');
  const deleteResult = deleteAllLwEventsFromGcal();
  console.log('Delete result:', JSON.stringify(deleteResult));

  // Step 2: シートのgcal_event_idをクリア
  console.log('Step 2: Clearing gcal_event_ids from sheet...');
  const clearResult = clearAllGcalEventIds();
  console.log('Clear result:', JSON.stringify(clearResult));

  // Step 3: 再同期
  console.log('Step 3: Syncing all events to GCal...');
  const syncResult = syncAllEventsToGcal();
  console.log('Sync result:', JSON.stringify(syncResult));

  console.log('=== Full reset and resync completed ===');
  return {
    ok: true,
    deleted: deleteResult.deleted || 0,
    cleared: clearResult.cleared || 0,
    synced: syncResult.synced || 0,
    errors: syncResult.errors || 0
  };
}

/**
 * GCal上の重複イベントをクリーンアップ（同じlw_event_idを持つイベント）
 */
function cleanupDuplicateGcalEvents() {
  const syncSettings = getGcalSyncSettings_();
  const calendarId = syncSettings[GCAL_SETTINGS_KEYS.CALENDAR_ID];

  if (!calendarId) {
    return { ok: false, error: 'Calendar ID not set' };
  }

  const settings = getSettings_();
  const [y, m] = settings.baseMonth.split('-').map(Number);
  const timeMin = new Date(y, m - 1, 1).toISOString();
  const timeMax = new Date(y, m + 2, 0).toISOString();

  // GCalから全イベント取得
  const events = [];
  let pageToken = null;
  do {
    const options = {
      singleEvents: true,
      timeMin: timeMin,
      timeMax: timeMax,
      maxResults: 500,
    };
    if (pageToken) options.pageToken = pageToken;

    const res = Calendar.Events.list(calendarId, options);
    (res.items || []).forEach(it => events.push(it));
    pageToken = res.nextPageToken;
  } while (pageToken);

  // lw_event_id毎にグルーピング
  const byLwId = new Map();
  events.forEach(ev => {
    const priv = (ev.extendedProperties && ev.extendedProperties.private) || {};
    const lwId = priv.lw_event_id || '';
    if (lwId) {
      if (!byLwId.has(lwId)) {
        byLwId.set(lwId, []);
      }
      byLwId.get(lwId).push(ev);
    }
  });

  // 重複を削除（最新のものを残す）
  let deleted = 0;
  byLwId.forEach((evList, lwId) => {
    if (evList.length > 1) {
      // updatedが最新のものを残す
      evList.sort((a, b) => new Date(b.updated) - new Date(a.updated));
      const keep = evList[0];
      console.log('Keeping event:', keep.id, 'for lw_event_id:', lwId);

      for (let i = 1; i < evList.length; i++) {
        try {
          Calendar.Events.remove(calendarId, evList[i].id);
          console.log('Deleted duplicate event:', evList[i].id);
          deleted++;
        } catch (err) {
          console.log('Could not delete duplicate:', err.message);
        }
      }
    }
  });

  console.log('Cleanup completed, deleted:', deleted, 'duplicates');
  return { ok: true, deleted: deleted };
}

/**
 * GCal同期の診断情報を取得
 */
function diagnoseGcalSync() {
  const syncSettings = getGcalSyncSettings_();
  const enabled = syncSettings[GCAL_SETTINGS_KEYS.SYNC_ENABLED];
  const calendarId = syncSettings[GCAL_SETTINGS_KEYS.CALENDAR_ID];

  const result = {
    settings: {
      enabled: enabled,
      calendarId: calendarId,
      loopGuardSeconds: syncSettings[GCAL_SETTINGS_KEYS.LOOP_GUARD_SECONDS],
      lastSyncAt: syncSettings[GCAL_SETTINGS_KEYS.LAST_SYNC_AT],
    },
    issues: []
  };

  if (!enabled || (enabled !== true && enabled !== 'TRUE' && enabled !== 'true')) {
    result.issues.push('同期が無効になっています（gcal_sync_enabledをTRUEに設定してください）');
  }

  if (!calendarId) {
    result.issues.push('カレンダーIDが設定されていません（gcal_calendar_idを設定してください）');
  } else {
    // カレンダーにアクセスできるか確認
    try {
      Calendar.Events.list(calendarId, { maxResults: 1 });
      result.calendarAccess = true;
    } catch (err) {
      result.calendarAccess = false;
      if (err.message.includes('403') || err.message.includes('forbidden')) {
        result.issues.push('カレンダーへの書き込み権限がありません。「予定の変更」権限を付与してください。');
      } else {
        result.issues.push('カレンダーにアクセスできません: ' + err.message);
      }
    }
  }

  // シートのデータ確認
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_EVENTS);
  const lastRow = sh.getLastRow();

  if (lastRow >= EVENTS_DATA_START_ROW) {
    const sampleRow = sh.getRange(EVENTS_DATA_START_ROW, 1, 1, 20).getValues()[0];
    result.sampleColumns = {
      'A(event_id)': sampleRow[0],
      'P(gcal_event_id)': sampleRow[15],
      'Q(gcal_calendar_id)': sampleRow[16],
      'R(last_sync_at)': sampleRow[17],
      'S(sync_source)': sampleRow[18],
      'T(revision)': sampleRow[19],
    };
  }

  result.ok = result.issues.length === 0;
  console.log('Diagnosis result:', JSON.stringify(result, null, 2));
  return result;
}

/**
 * 既存の全イベントをGCalへ一括同期
 */
function syncAllEventsToGcal() {
  const syncSettings = getGcalSyncSettings_();
  const enabled = syncSettings[GCAL_SETTINGS_KEYS.SYNC_ENABLED];
  const calendarId = syncSettings[GCAL_SETTINGS_KEYS.CALENDAR_ID];

  if (enabled !== true && enabled !== 'TRUE' && enabled !== 'true') {
    return { ok: false, error: 'GCal sync is disabled' };
  }
  if (!calendarId) {
    return { ok: false, error: 'Calendar ID not set' };
  }

  const settings = getSettings_();
  const range = getEditableRange_(settings);

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_EVENTS);
  const lastRow = sh.getLastRow();

  if (lastRow < EVENTS_DATA_START_ROW) {
    return { ok: true, synced: 0 };
  }

  const values = sh.getRange(EVENTS_DATA_START_ROW, 1, lastRow - EVENTS_DATA_START_ROW + 1, 19).getValues();
  let synced = 0;
  let errors = 0;

  values.forEach((row, i) => {
    const rowNum = EVENTS_DATA_START_ROW + i;
    const eventId = row[0];
    const dateVal = row[1];
    const gcalEventId = row[15];

    // 日付を文字列に変換
    let dateStr = '';
    if (dateVal instanceof Date) {
      dateStr = formatISODate_(dateVal);
    } else if (dateVal) {
      dateStr = String(dateVal).substring(0, 10);
    }

    // 編集可能範囲内のみ同期
    if (!dateStr || dateStr < range.fromISO || dateStr > range.toISO) {
      return;
    }

    const eventData = {
      event_id: eventId,
      date: dateStr,
      start_time: row[2] || '',
      end_time: row[3] || '',
      type: row[4] || '',
      title: row[5] || '',
      location: row[6] || '',
      memo: row[7] || '',
      status: row[9] || 'CONFIRMED',
      gcal_event_id: gcalEventId || '',
      is_all_day: !row[2] && !row[3],
    };

    const result = syncOneEventToGcal_(eventData, calendarId);

    if (result.gcal_event_id && result.gcal_event_id !== gcalEventId) {
      sh.getRange(rowNum, 16).setValue(result.gcal_event_id);
      sh.getRange(rowNum, 17).setValue(calendarId);
      sh.getRange(rowNum, 18).setValue(new Date());
      sh.getRange(rowNum, 19).setValue(SYNC_SOURCE.SYSTEM);
    }

    if (result.error) {
      errors++;
    } else {
      synced++;
    }
  });

  console.log('Bulk sync completed - synced:', synced, 'errors:', errors);
  return { ok: true, synced: synced, errors: errors };
}
