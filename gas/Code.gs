/**
 * 社長スケジュール - Google Apps Script バックエンド
 *
 * 【新設計】Googleカレンダー1方向同期 + BOARD編集
 * - Googleカレンダーは読み取り専用（書き込み禁止）
 * - 同期: GCal → Sheet（gcal_列のみ更新）
 * - BOARD編集: board_列のみ更新（同期で上書きされない）
 * - 表示: board_優先（空ならgcal_）
 */

// =============================================================================
// 定数定義
// =============================================================================
const SHEET_SETTINGS = '01_Settings';
const SHEET_USERS    = '02_Users';
const SHEET_EVENTS   = '03_Events';
const SHEET_DAY_SETTINGS = '04_DaySettings';
const SHEET_NOTES    = '05_Notes';

const SETTINGS_BASE_MONTH_CELL = 'B5';      // YYYY-MM
const SETTINGS_EDITABLE_MONTHS_CELL = 'B6'; // 2
const SETTINGS_TZ_CELL = 'B7';              // Asia/Tokyo

const EVENTS_HEADER_ROW = 5;
const EVENTS_DATA_START_ROW = 6;

// Events列構成（新設計）:
// A(1): gcal_event_id（主キー：GoogleカレンダーのイベントID）
// B(2): date
// C(3): start_time
// D(4): end_time
// E(5): is_all_day
// F(6): gcal_title
// G(7): gcal_location
// H(8): gcal_description
// I(9): gcal_updated
// J(10): gcal_status
// K(11): gcal_is_deleted
// L(12): board_title（BOARD上書き）
// M(13): board_location（BOARD上書き）
// N(14): board_memo（BOARD上書き）
// O(15): type（予定種別：BOARD編集可）
// P(16): display_order（BOARD編集可）
// Q(17): pin（BOARD編集可）
// R(18): hidden（BOARD編集可）
// S(19): board_updated_at
// T(20): board_updated_by
// U(21): gcal_imported_at

const USERS_HEADER_ROW = 5;
const USERS_DATA_START_ROW = 6;

// GCal同期設定キー
const GCAL_SETTINGS_KEYS = {
  CALENDAR_ID: 'gcal_calendar_id',
  SYNC_ENABLED: 'gcal_sync_enabled',
  SYNC_TOKEN: 'gcal_sync_token',
  LAST_SYNC_AT: 'gcal_last_sync_at',
  RANGE_MARGIN_DAYS: 'gcal_range_margin_days',
};

// =============================================================================
// イベント種別 自動分類
// =============================================================================

/**
 * タイトル・説明文からイベント種別を自動判定
 * @param {string} title - イベントタイトル
 * @param {string} description - イベント説明
 * @return {string} 種別コード
 */
function classifyEventType_(title, description) {
  const t = String(title || '');
  const d = String(description || '');
  const text = t + ' ' + d;

  // 優先順位順にマッチング
  if (/会議|MTG|ミーティング|打合せ|打ち合わせ|定例/i.test(text)) return 'MEETING';
  if (/来客|来社|ご来社/i.test(text)) return 'VISIT';
  if (/面会|面談|1on1|1:1/i.test(text)) return 'MEET';
  if (/VIP|役員|取締役/i.test(text)) return 'VIP';
  if (/出張/i.test(text) || /^★/i.test(t)) return 'TRAVEL';
  if (/外出|移動|訪問/i.test(text)) return 'OUT';
  if (/会食|懇親|食事|ディナー|ランチ/i.test(text)) return 'DINNER';
  if (/休|有給|休暇/i.test(text)) return 'HOLIDAY';

  return 'OTHER';
}

// =============================================================================
// Webアプリ エントリーポイント
// =============================================================================

function doGet() {
  const t = HtmlService.createTemplateFromFile('index');
  t.PORTAL_URL = 'https://script.google.com/a/macros/lineworks-local.info/s/AKfycbx2eyJMOYP9o--GPBuhY-pj071IIR6Kqb_0xALwwNzdLQZux0dIAlL3P9EoCucnzXA/exec';
  return t.evaluate()
    .setTitle('社長スケジュール')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// =============================================================================
// 初期化・データ取得
// =============================================================================

/**
 * 初回ロード時に必要なすべてのデータを返す
 */
function getBootstrap() {
  try {
    const settings = getSettings_();
    const user = getUserContext_();
    const masters = getMasters_();
    const range = getEditableRange_(settings);

    // 統合ビュー（board優先）でイベントを取得
    const events = listBoardEvents_(range.fromISO, range.toISO);

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
      daySettings,
      syncStatus: (() => {
        try {
          const ss = getGcalSyncSettings_();
          return {
            enabled: ss[GCAL_SETTINGS_KEYS.SYNC_ENABLED] === 'TRUE' || ss[GCAL_SETTINGS_KEYS.SYNC_ENABLED] === true,
            lastSyncAt: ss[GCAL_SETTINGS_KEYS.LAST_SYNC_AT] || null,
            calendarId: ss[GCAL_SETTINGS_KEYS.CALENDAR_ID] || null,
          };
        } catch(e) { return { enabled: false, lastSyncAt: null, calendarId: null }; }
      })()
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

function getSettings_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_SETTINGS);

  if (!sh) {
    throw new Error('SHEET_NOT_FOUND: ' + SHEET_SETTINGS);
  }

  const tzRaw = sh.getRange(SETTINGS_TZ_CELL).getValue();
  const tz = String(tzRaw || '').trim() || 'Asia/Tokyo';
  const editableMonths = Number(sh.getRange(SETTINGS_EDITABLE_MONTHS_CELL).getValue() || 2);

  // 常に今月を基準月として使用（自動進行）
  const baseMonth = Utilities.formatDate(new Date(), tz, 'yyyy-MM');

  return { tz, baseMonth, editableMonths };
}

function getEditableRange_(settings) {
  const tz = settings.tz;
  const [y, m] = settings.baseMonth.split('-').map(Number);

  const from = new Date(y, m - 1, 1);
  const to = new Date(y, m - 1 + settings.editableMonths, 0);

  return {
    fromISO: Utilities.formatDate(from, tz, 'yyyy-MM-dd'),
    toISO: Utilities.formatDate(to, tz, 'yyyy-MM-dd')
  };
}

// =============================================================================
// GCal同期設定
// =============================================================================

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
  sh.getRange(lastRow + 1, 1, 1, 2).setValues([[key, value]]);
}

// =============================================================================
// User / ACL（権限制御）
// =============================================================================

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
  const values = sh.getRange(USERS_DATA_START_ROW, 1, numRows, 3).getValues();

  for (const [em, role, enabled] of values) {
    if (!em) continue;

    if (String(em).trim().toLowerCase() === String(email).trim().toLowerCase()) {
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

function getMasters_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_SETTINGS);

  if (!sh) {
    return { types: [], statuses: [] };
  }

  const types = readMaster_(sh, 13, 1, 5)
    .filter(r => r[0] && isTrue_(r[4]))
    .map(r => ({
      code: String(r[0]).trim(),
      label: String(r[1] || r[0]).trim(),
      color: String(r[2] || '#374151').trim(),
      sort: Number(r[3] || 999)
    }))
    .sort((a, b) => a.sort - b.sort);

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

function isTrue_(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}

// =============================================================================
// Events 読み取り（統合ビュー：board優先）
// =============================================================================

/**
 * 指定期間のイベントを取得（統合ビュー）
 * - 全件必ず表示（hiddenによる除外なし）
 * - board_列は差し替え/追記として扱う（イベント自体は消さない）
 */
function listBoardEvents_(fromISO, toISO) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_EVENTS);

  if (!sh) {
    throw new Error('SHEET_NOT_FOUND: ' + SHEET_EVENTS);
  }

  const lastRow = sh.getLastRow();
  if (lastRow < EVENTS_DATA_START_ROW) return [];

  const numRows = lastRow - EVENTS_DATA_START_ROW + 1;
  const values = sh.getRange(EVENTS_DATA_START_ROW, 1, numRows, 21).getValues();

  const from = new Date(fromISO + 'T00:00:00');
  const to = new Date(toISO + 'T23:59:59');

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const gcalEventId = r[0];  // A: gcal_event_id
    const dateStr = r[1];      // B: date

    if (!gcalEventId || !dateStr) continue;

    // 日付チェック
    let dd;
    if (dateStr instanceof Date) {
      dd = dateStr;
    } else {
      dd = new Date(String(dateStr) + 'T00:00:00');
    }
    if (isNaN(dd.getTime())) continue;
    if (dd < from || dd > to) continue;

    // 削除済みチェック（GCalで削除されたものは非表示）
    const isDeleted = isTrue_(r[10]);  // K: gcal_is_deleted
    if (isDeleted) continue;

    // ※ hiddenによる除外は行わない（全件表示）

    // BOARD上書き
    const boardTitle = String(r[11] || '').trim();      // L: board_title
    const boardLocation = String(r[12] || '').trim();   // M: board_location
    const boardMemo = String(r[13] || '').trim();       // N: board_memo

    // GCal由来
    const gcalTitle = String(r[5] || '').trim();        // F: gcal_title
    const gcalLocation = String(r[6] || '').trim();     // G: gcal_location
    const gcalDescription = String(r[7] || '').trim();  // H: gcal_description

    // 表示用（boardは差し替え、イベント自体は常に表示）
    const titleForBoard = boardTitle || gcalTitle;
    const locationForBoard = boardLocation || gcalLocation;

    // メモは合成（board追記 + Google説明）
    // UIで別々に表示したい場合は board_memo と gcal_description を参照
    let memoForBoard = '';
    if (boardMemo && gcalDescription) {
      memoForBoard = boardMemo + '\n――――\n' + gcalDescription;
    } else if (boardMemo) {
      memoForBoard = boardMemo;
    } else if (gcalDescription) {
      memoForBoard = gcalDescription;
    }

    out.push({
      gcal_event_id: String(gcalEventId),
      event_id: String(gcalEventId),
      date: formatISODate_(dd),
      start_time: asTimeStr_(r[2]),   // C: start_time
      end_time: asTimeStr_(r[3]),     // D: end_time
      is_all_day: isTrue_(r[4]),      // E: is_all_day

      // 表示用（board差し替え済み）
      title: titleForBoard,
      location: locationForBoard,
      memo: memoForBoard,

      // 元データ（UI側で必要なら参照）
      gcal_title: gcalTitle,
      gcal_location: gcalLocation,
      gcal_description: gcalDescription,

      // BOARD編集フィールド
      board_title: boardTitle,
      board_location: boardLocation,
      board_memo: boardMemo,
      type: String(r[14] || '').trim(),           // O: type
      display_order: Number(r[15] || 0),          // P: display_order
      pin: isTrue_(r[16]),                        // Q: pin

      gcal_status: String(r[9] || '').trim(),     // J: gcal_status
      _row: EVENTS_DATA_START_ROW + i
    });
  }

  // 並び順: date → pin（優先）→ all-day優先 → display_order → start_time
  out.sort((a, b) => {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    if (a.pin !== b.pin) return a.pin ? -1 : 1;
    // All-day events sort before timed events
    if (a.is_all_day !== b.is_all_day) return a.is_all_day ? -1 : 1;
    if ((a.display_order || 0) !== (b.display_order || 0)) {
      return (a.display_order || 0) - (b.display_order || 0);
    }
    const at = a.start_time || '99:99';
    const bt = b.start_time || '99:99';
    return at < bt ? -1 : 1;
  });

  return out;
}

// =============================================================================
// Googleカレンダー → Sheet 1方向同期（gcal_列のみ更新）
// =============================================================================

/**
 * GCal → Sheet 同期（読み取り専用）
 * - gcal_列のみ更新
 * - board_列は絶対に上書きしない
 */
function syncFromGcalToSheet() {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    console.log('Could not obtain lock for GCal sync');
    return { ok: false, error: 'LOCK_TIMEOUT' };
  }

  try {
    const syncSettings = getGcalSyncSettings_();
    const enabled = syncSettings[GCAL_SETTINGS_KEYS.SYNC_ENABLED];
    const calendarId = syncSettings[GCAL_SETTINGS_KEYS.CALENDAR_ID];

    if (enabled !== true && enabled !== 'TRUE' && enabled !== 'true') {
      console.log('GCal sync disabled');
      return { ok: false, error: 'SYNC_DISABLED' };
    }
    if (!calendarId) {
      console.log('GCal calendar ID not set');
      return { ok: false, error: 'CALENDAR_ID_NOT_SET' };
    }

    const settings = getSettings_();
    const tz = settings.tz;
    const marginDays = Number(syncSettings[GCAL_SETTINGS_KEYS.RANGE_MARGIN_DAYS] || 7);

    // 同期範囲: 基準月-margin ～ 基準月+2ヶ月+margin
    const [y, m] = settings.baseMonth.split('-').map(Number);
    const rangeStart = new Date(y, m - 1, 1 - marginDays);
    const rangeEnd = new Date(y, m + 1, marginDays);

    const timeMin = rangeStart.toISOString();
    const timeMax = rangeEnd.toISOString();

    console.log('GCal sync range:', timeMin, '-', timeMax);

    // Calendar API で取得（読み取りのみ）
    const items = [];
    let pageToken = null;
    const syncToken = syncSettings[GCAL_SETTINGS_KEYS.SYNC_TOKEN] || '';

    try {
      do {
        const options = syncToken
          ? { syncToken: syncToken, showDeleted: true, maxResults: 2500 }
          : { singleEvents: true, showDeleted: true, timeMin: timeMin, timeMax: timeMax, maxResults: 2500 };
        if (pageToken) options.pageToken = pageToken;

        const res = Calendar.Events.list(calendarId, options);
        (res.items || []).forEach(it => items.push(it));
        pageToken = res.nextPageToken;
        if (res.nextSyncToken) {
          setGcalSyncSetting_(GCAL_SETTINGS_KEYS.SYNC_TOKEN, res.nextSyncToken);
        }
      } while (pageToken);
    } catch (err) {
      // syncToken失効時はクリアして再取得
      console.log('SyncToken expired, clearing and retrying...');
      setGcalSyncSetting_(GCAL_SETTINGS_KEYS.SYNC_TOKEN, '');
      pageToken = null;
      do {
        const options = {
          singleEvents: true,
          showDeleted: true,
          timeMin: timeMin,
          timeMax: timeMax,
          maxResults: 2500,
        };
        if (pageToken) options.pageToken = pageToken;

        const res = Calendar.Events.list(calendarId, options);
        (res.items || []).forEach(it => items.push(it));
        pageToken = res.nextPageToken;
        if (res.nextSyncToken) {
          setGcalSyncSetting_(GCAL_SETTINGS_KEYS.SYNC_TOKEN, res.nextSyncToken);
        }
      } while (pageToken);
    }

    console.log('GCal events fetched:', items.length);

    if (items.length === 0) {
      setGcalSyncSetting_(GCAL_SETTINGS_KEYS.LAST_SYNC_AT, new Date().toISOString());
      return { ok: true, imported: 0 };
    }

    // Sheet側のインデックス作成
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_EVENTS);
    const lastRow = sh.getLastRow();

    const idxByGcalId = new Map();  // gcal_event_id -> row number

    if (lastRow >= EVENTS_DATA_START_ROW) {
      const ids = sh.getRange(EVENTS_DATA_START_ROW, 1, lastRow - EVENTS_DATA_START_ROW + 1, 1).getValues();
      ids.forEach((row, i) => {
        const gid = String(row[0] || '').trim();
        if (gid) {
          idxByGcalId.set(gid, EVENTS_DATA_START_ROW + i);
        }
      });
    }

    const importAt = new Date().toISOString();
    let updated = 0;
    let created = 0;

    // 各イベントを処理（gcal_列のみ更新）
    items.forEach(ev => {
      const gid = String(ev.id);
      const patch = gcalEventToPatch_(ev, tz, importAt);

      const existingRow = idxByGcalId.get(gid);

      if (existingRow) {
        // 既存行の gcal_updated と比較
        const sheetUpdated = sh.getRange(existingRow, 9).getValue();  // I: gcal_updated
        const sheetUpdatedTime = sheetUpdated ? new Date(sheetUpdated).getTime() : 0;
        const gcalUpdatedTime = ev.updated ? new Date(ev.updated).getTime() : 0;

        if (gcalUpdatedTime <= sheetUpdatedTime) {
          return;  // Sheetの方が新しい or 同じ → 更新不要
        }

        // gcal_列のみ更新（A〜K列 + U列）
        writeGcalColumnsOnly_(sh, existingRow, patch);
        updated++;
      } else {
        // 新規行（board_列は空で作成）
        const newRow = sh.getLastRow() + 1;
        // C,D列をテキスト形式に設定（時刻の自動変換防止）
        sh.getRange(newRow, 3, 1, 2).setNumberFormat('@');
        sh.getRange(newRow, 1, 1, 21).setValues([[
          patch.gcal_event_id,      // A
          patch.date,               // B
          patch.start_time,         // C
          patch.end_time,           // D
          patch.is_all_day,         // E
          patch.gcal_title,         // F
          patch.gcal_location,      // G
          patch.gcal_description,   // H
          patch.gcal_updated,       // I
          patch.gcal_status,        // J
          patch.gcal_is_deleted,    // K
          '',                       // L: board_title（空）
          '',                       // M: board_location（空）
          '',                       // N: board_memo（空）
          patch.suggested_type || '',  // O: type（自動分類）
          '',                       // P: display_order（空）
          false,                    // Q: pin
          false,                    // R: hidden
          '',                       // S: board_updated_at
          '',                       // T: board_updated_by
          importAt                  // U: gcal_imported_at
        ]]);
        idxByGcalId.set(gid, newRow);
        created++;
      }
    });

    setGcalSyncSetting_(GCAL_SETTINGS_KEYS.LAST_SYNC_AT, importAt);
    console.log('GCal sync completed - updated:', updated, 'created:', created);

    return { ok: true, updated: updated, created: created };

  } catch (e) {
    console.error('syncFromGcalToSheet error:', e);
    return { ok: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * GCalイベントをSheet用patchに変換
 */
function gcalEventToPatch_(ev, tz, importAt) {
  const isCancelled = (ev.status === 'cancelled');

  const patch = {
    gcal_event_id: String(ev.id),
    gcal_title: ev.summary || '',
    gcal_location: ev.location || '',
    gcal_description: (ev.description || '').trim(),
    gcal_updated: ev.updated || '',
    gcal_status: ev.status || '',
    gcal_is_deleted: isCancelled,
    gcal_imported_at: importAt,
  };

  // 日付/時刻
  if (ev.start && ev.start.date) {
    patch.is_all_day = true;
    patch.date = ev.start.date;
    patch.start_time = '';
    patch.end_time = '';
  } else if (ev.start && ev.start.dateTime) {
    patch.is_all_day = false;
    const sd = new Date(ev.start.dateTime);
    patch.date = Utilities.formatDate(sd, tz, 'yyyy-MM-dd');
    patch.start_time = Utilities.formatDate(sd, tz, 'HH:mm');
    if (ev.end && ev.end.dateTime) {
      const ed = new Date(ev.end.dateTime);
      patch.end_time = Utilities.formatDate(ed, tz, 'HH:mm');
    } else {
      patch.end_time = '';
    }
  }

  // 自動分類
  patch.suggested_type = classifyEventType_(patch.gcal_title, patch.gcal_description);

  return patch;
}

/**
 * gcal_列のみ更新（board_列は触らない）
 */
function writeGcalColumnsOnly_(sh, rowNum, patch) {
  // C,D列をテキスト形式に設定（時刻の自動変換防止）
  sh.getRange(rowNum, 3, 1, 2).setNumberFormat('@');
  // A〜K列を一括書き込み
  sh.getRange(rowNum, 1, 1, 11).setValues([[
    patch.gcal_event_id,
    patch.date || '',
    patch.start_time || '',
    patch.end_time || '',
    patch.is_all_day || false,
    patch.gcal_title || '',
    patch.gcal_location || '',
    patch.gcal_description || '',
    patch.gcal_updated || '',
    patch.gcal_status || '',
    patch.gcal_is_deleted || false
  ]]);
  // U(21): gcal_imported_at
  sh.getRange(rowNum, 21).setValue(patch.gcal_imported_at || '');
}

// =============================================================================
// BOARD編集（board_列のみ更新 - gcal_列は触らない）
// =============================================================================

/**
 * BOARD編集保存
 * @param {Object} payload - { gcal_event_id, board_title?, board_location?, board_memo?, type?, display_order?, pin? }
 * ※ hidden は無効（全件表示のため）
 */
function saveBoardOverride(payload) {
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

    if (!payload || !payload.gcal_event_id) {
      throw new Error('INVALID: gcal_event_id が必要です');
    }

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_EVENTS);

    if (!sh) {
      throw new Error('SHEET_NOT_FOUND: ' + SHEET_EVENTS);
    }

    // gcal_event_id で行を検索
    const gid = String(payload.gcal_event_id).trim();
    const lastRow = sh.getLastRow();

    if (lastRow < EVENTS_DATA_START_ROW) {
      throw new Error('NOT_FOUND: Eventsが空です');
    }

    const ids = sh.getRange(EVENTS_DATA_START_ROW, 1, lastRow - EVENTS_DATA_START_ROW + 1, 1).getValues();
    let rowNum = null;

    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === gid) {
        rowNum = EVENTS_DATA_START_ROW + i;
        break;
      }
    }

    if (!rowNum) {
      throw new Error('NOT_FOUND: 対象イベントがありません');
    }

    // 編集可能期間チェック
    const dateVal = sh.getRange(rowNum, 2).getValue();  // B: date
    let dateStr = '';
    if (dateVal instanceof Date) {
      dateStr = Utilities.formatDate(dateVal, settings.tz, 'yyyy-MM-dd');
    } else {
      dateStr = String(dateVal).substring(0, 10);
    }
    assertCanEditDate_(settings, dateStr);

    const tz = settings.tz;
    const now = new Date();

    // board_列のみ更新（L〜R列 + S, T列）
    if (payload.board_title !== undefined) {
      sh.getRange(rowNum, 12).setValue(payload.board_title);  // L: board_title
    }
    if (payload.board_location !== undefined) {
      sh.getRange(rowNum, 13).setValue(payload.board_location);  // M: board_location
    }
    if (payload.board_memo !== undefined) {
      sh.getRange(rowNum, 14).setValue(payload.board_memo);  // N: board_memo
    }
    if (payload.type !== undefined) {
      sh.getRange(rowNum, 15).setValue(payload.type);  // O: type
    }
    if (payload.display_order !== undefined) {
      sh.getRange(rowNum, 16).setValue(payload.display_order);  // P: display_order
    }
    if (payload.pin !== undefined) {
      sh.getRange(rowNum, 17).setValue(isTrue_(payload.pin));  // Q: pin
    }
    // ※ hidden は無効化（全件表示のため、R列は使用しない）

    // 更新日時・更新者
    sh.getRange(rowNum, 19).setValue(Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'));  // S: board_updated_at
    sh.getRange(rowNum, 20).setValue(user.email);  // T: board_updated_by

    return { ok: true, gcal_event_id: gid };

  } catch (e) {
    console.error('saveBoardOverride error:', e);
    return { ok: false, error: e.message || 'SAVE_FAILED' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 編集可能期間チェック
 */
function assertCanEditDate_(settings, isoDate) {
  const r = getEditableRange_(settings);
  const d = new Date(isoDate + 'T00:00:00');
  const from = new Date(r.fromISO + 'T00:00:00');
  const to = new Date(r.toISO + 'T23:59:59');

  if (d < from || d > to) {
    throw new Error('OUT_OF_EDITABLE_RANGE: 当月＋翌月のみ編集可能です');
  }
}

// =============================================================================
// 同期トリガー管理
// =============================================================================

/**
 * 同期トリガーをインストール（1分おき）
 */
function installGcalSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncFromGcalToSheet') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('syncFromGcalToSheet')
    .timeBased()
    .everyMinutes(1)
    .create();

  console.log('GCal sync trigger installed (every 1 minute)');
  return { ok: true, message: 'Trigger installed' };
}

/**
 * 同期トリガーを削除
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
 * 手動同期（UIの再読込ボタン用）
 */
function manualSyncNow() {
  return syncFromGcalToSheet();
}

/**
 * GCal同期の初期設定
 */
function setupGcalSync(calendarId) {
  if (!calendarId) {
    return { ok: false, error: 'Calendar ID is required' };
  }

  setGcalSyncSetting_(GCAL_SETTINGS_KEYS.CALENDAR_ID, calendarId);
  setGcalSyncSetting_(GCAL_SETTINGS_KEYS.SYNC_ENABLED, 'TRUE');
  setGcalSyncSetting_(GCAL_SETTINGS_KEYS.RANGE_MARGIN_DAYS, 7);

  console.log('GCal sync configured for calendar:', calendarId);
  return { ok: true, calendarId: calendarId };
}

// =============================================================================
// ユーティリティ
// =============================================================================

function formatISODate_(d) {
  const settings = getSettings_();
  return Utilities.formatDate(d, settings.tz, 'yyyy-MM-dd');
}

function asTimeStr_(v) {
  if (!v) return '';
  if (typeof v === 'string') return v;
  if (v instanceof Date) {
    const settings = getSettings_();
    return Utilities.formatDate(v, settings.tz, 'HH:mm');
  }
  return String(v);
}

function nextMonthKeyServer_(baseYYYYMM) {
  const [y, m] = baseYYYYMM.split('-').map(Number);
  const d = new Date(y, m - 1, 1);
  d.setMonth(d.getMonth() + 1);
  const settings = getSettings_();
  return Utilities.formatDate(d, settings.tz, 'yyyy-MM');
}

// =============================================================================
// Notes（月単位の備考）- 既存機能維持
// =============================================================================

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

function getNotesMap_(fromMonthYYYYMM, toMonthYYYYMM) {
  const sh = ensureNotesSheet_();
  const lastRow = sh.getLastRow();
  const map = {};

  if (lastRow < 2) return map;

  const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();

  for (const [mk, notes] of values) {
    if (!mk) continue;
    const key = String(mk).trim();
    if (key >= fromMonthYYYYMM && key <= toMonthYYYYMM) {
      map[key] = String(notes || '');
    }
  }

  return map;
}

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

    if (!(user.role === 'editor' || user.role === 'admin')) {
      throw new Error('FORBIDDEN');
    }

    if (!/^\d{4}-\d{2}$/.test(monthKey)) {
      throw new Error('INVALID_MONTH');
    }

    const base = settings.baseMonth;
    const next = nextMonthKeyServer_(base);
    if (!(monthKey === base || monthKey === next)) {
      throw new Error('OUT_OF_EDITABLE_RANGE');
    }

    const sh = ensureNotesSheet_();
    const lastRow = sh.getLastRow();
    const tz = settings.tz;
    const now = new Date();

    if (lastRow >= 2) {
      const keys = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) {
        if (String(keys[i][0]).trim() === monthKey) {
          sh.getRange(2 + i, 2).setValue(String(text || ''));
          sh.getRange(2 + i, 3).setValue(Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'));
          sh.getRange(2 + i, 4).setValue(user.email);
          return { ok: true };
        }
      }
    }

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
// DaySettings（出勤日・休日設定）- 既存機能維持
// =============================================================================

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

function getDaySettingsMap_(fromISO, toISO) {
  const sh = ensureDaySettingsSheet_();
  const lastRow = sh.getLastRow();
  const map = {};

  if (lastRow < 2) return map;

  const settings = getSettings_();
  const tz = settings.tz;
  const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();

  for (const [dateVal, type] of values) {
    if (!dateVal) continue;
    let key;
    if (dateVal instanceof Date) {
      key = Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
    } else {
      key = String(dateVal).trim();
    }
    const typeStr = String(type || '').trim();
    if (typeStr && key.match(/^\d{4}-\d{2}-\d{2}$/)) {
      map[key] = typeStr;
    }
  }

  return map;
}

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

    if (!(user.role === 'editor' || user.role === 'admin')) {
      throw new Error('FORBIDDEN');
    }

    assertCanEditDate_(settings, dateISO);

    const sh = ensureDaySettingsSheet_();
    const lastRow = sh.getLastRow();
    const tz = settings.tz;
    const now = new Date();

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
      if (existingRow) {
        sh.deleteRow(existingRow);
      }
      return { ok: true, action: 'removed' };
    }

    if (existingRow) {
      sh.getRange(existingRow, 2).setValue(type);
      sh.getRange(existingRow, 3).setValue(memo || '');
      sh.getRange(existingRow, 4).setValue(Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'));
      sh.getRange(existingRow, 5).setValue(user.email);
      return { ok: true, action: 'updated' };
    }

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
// 診断・ヘルスチェック・ユーティリティ
// =============================================================================

/**
 * 03_Eventsシートのデータをすべてクリア（ヘッダー行は残す）
 */
function clearAllEvents() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_EVENTS);

  if (!sh) {
    return { ok: false, error: 'SHEET_NOT_FOUND' };
  }

  const lastRow = sh.getLastRow();
  if (lastRow < EVENTS_DATA_START_ROW) {
    return { ok: true, cleared: 0, message: 'データなし' };
  }

  const numRows = lastRow - EVENTS_DATA_START_ROW + 1;
  sh.deleteRows(EVENTS_DATA_START_ROW, numRows);

  // syncTokenもクリア（次回フル同期のため）
  setGcalSyncSetting_(GCAL_SETTINGS_KEYS.SYNC_TOKEN, '');

  console.log('Cleared', numRows, 'rows from Events');
  return { ok: true, cleared: numRows };
}

/**
 * BOARDリセット＆再同期（クリア→同期を一括実行）
 */
function resetAndResync() {
  console.log('=== Starting reset and resync ===');

  // Step 1: クリア
  const clearResult = clearAllEvents();
  console.log('Clear result:', JSON.stringify(clearResult));

  // Step 2: 再同期
  const syncResult = syncFromGcalToSheet();
  console.log('Sync result:', JSON.stringify(syncResult));

  return {
    ok: true,
    cleared: clearResult.cleared || 0,
    synced: (syncResult.created || 0) + (syncResult.updated || 0)
  };
}

function diagnoseGcalSync() {
  const syncSettings = getGcalSyncSettings_();
  const calendarId = syncSettings[GCAL_SETTINGS_KEYS.CALENDAR_ID];
  const enabled = syncSettings[GCAL_SETTINGS_KEYS.SYNC_ENABLED];

  const result = {
    settings: {
      calendarId: calendarId,
      enabled: enabled,
      lastSyncAt: syncSettings[GCAL_SETTINGS_KEYS.LAST_SYNC_AT],
    },
    issues: []
  };

  if (!enabled || (enabled !== true && enabled !== 'TRUE' && enabled !== 'true')) {
    result.issues.push('同期が無効です（gcal_sync_enabledをTRUEに設定）');
  }

  if (!calendarId) {
    result.issues.push('カレンダーIDが未設定です（gcal_calendar_idを設定）');
  } else {
    try {
      const settings = getSettings_();
      const [y, m] = settings.baseMonth.split('-').map(Number);
      const timeMin = new Date(y, m - 1, 1).toISOString();
      const timeMax = new Date(y, m, 0).toISOString();

      Calendar.Events.list(calendarId, {
        timeMin: timeMin,
        timeMax: timeMax,
        maxResults: 1
      });
      result.calendarAccess = true;
    } catch (err) {
      result.calendarAccess = false;
      result.issues.push('カレンダーにアクセスできません: ' + err.message);
    }
  }

  result.ok = result.issues.length === 0;
  console.log('Diagnosis:', JSON.stringify(result, null, 2));
  return result;
}
