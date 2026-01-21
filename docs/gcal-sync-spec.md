# Googleカレンダー双方向同期 設計仕様書

## 1. ゴール定義

### 同期方向
- **A: Sheet → GCal**: Webアプリで保存時に即時反映（作成/更新/削除）
- **B: GCal → Sheet**: 時間トリガーで1〜5分おきに反映（作成/更新/削除）

### 対象範囲
- 基準月の1日 〜 翌月末 + 7日（安全マージン）
- `sync_range_start` / `sync_range_end` で設定

### 書き込み先カレンダー
- 固定（社長カレンダーID or 会社共有カレンダーID）
- 複数カレンダー対応は後回し

---

## 2. データ設計（最重要）

### 2.1 Eventsシートに追加する列

| 列名 | 型 | 説明 |
|------|-----|------|
| `gcal_event_id` | 文字列 | GoogleカレンダーのイベントID |
| `gcal_calendar_id` | 文字列 | 書き込み先カレンダーID |
| `last_modified_at` | ISO日時 | 最終更新日時 |
| `last_modified_by` | 文字列 | 更新者メール |
| `last_sync_at` | ISO日時 | 最後に同期した日時 |
| `sync_source` | 文字列 | 最終更新元（`WEBAPP` / `GCAL` / `SYSTEM`） |
| `sync_hash` | 文字列 | 内容ハッシュ（SHA-256） |
| `is_all_day` | boolean | 終日予定かどうか |
| `is_deleted` | boolean | 削除フラグ（論理削除） |

### 2.2 Settingsシートに追加するキー

| キー | 値の例 | 説明 |
|------|--------|------|
| `gcal_calendar_id` | `primary` or `xxx@group.calendar.google.com` | 同期先カレンダーID |
| `sync_range_start` | `2026-01-01` | 同期開始日 |
| `sync_range_end` | `2026-02-28` | 同期終了日 |
| `gcal_sync_token` | （自動設定） | Calendar API差分取得用トークン |
| `gcal_last_sync_at` | （自動設定） | 最後の同期日時 |
| `loop_guard_seconds` | `30` | ループ防止用のスキップ秒数 |

### 2.3 Googleカレンダー側のメタ情報

`extendedProperties.private` に保存：

```javascript
{
  "lw_event_id": "Sheet側のid",
  "lw_origin": "LW",
  "lw_hash": "同期時のハッシュ"
}
```

---

## 3. ループ防止設計（必須）

### 問題
「Sheet更新 → GCal更新 → そのGCal更新を検知 → Sheet更新…」の無限ループ

### 解決策

1. **sync_source を記録**
   - Sheet→GCal: `sync_source = 'WEBAPP'`
   - GCal→Sheet: `sync_source = 'GCAL'`

2. **last_sync_at で時間ガード**
   ```javascript
   if (now - last_sync_at < 30秒) {
     // スキップ
   }
   ```

3. **sync_hash 一致チェック**
   - 内容が同じなら更新しない

---

## 4. 衝突ルール（Last Write Wins）

```
同じ予定（同一ID）に対して：
  Sheet側 last_modified_at > GCal側 updated
    → Sheetを正としてGCal更新

  逆
    → GCalを正としてSheet更新
```

---

## 5. フィールドマッピング

### Sheet → GCal

| Sheet | GCal |
|-------|------|
| `title` | `summary` |
| `location` | `location` |
| `memo` | `description` |
| `date` + `start_time` | `start.dateTime` |
| `date` + `end_time` | `end.dateTime` |
| `is_all_day=true` | `start.date` / `end.date` |
| `type` | `extendedProperties.private.lw_type` or `colorId` |
| `status` | `status` (`confirmed` / `tentative`) |

### GCal → Sheet

| GCal | Sheet |
|------|-------|
| `summary` | `title` |
| `location` | `location` |
| `description` | `memo` |
| `start/end` | `date`, `start_time`, `end_time`, `is_all_day` |
| `updated` | `last_modified_at` |
| `creator.email` | `last_modified_by` |
| `status='cancelled'` | `is_deleted=true` |

---

## 6. 実装コード

### 6.0 事前設定（必須）

1. Apps Script → サービス → **Google Calendar API** を有効化
2. GCPプロジェクトでも Calendar API が有効であることを確認
3. Settingsシートに `gcal_calendar_id` 等を設定

### 6.1 定数・ユーティリティ

```javascript
/** ====== CONFIG ====== */
const SHEET_NAMES = {
  SETTINGS: '01_Settings',
  USERS: '02_Users',
  EVENTS: '03_Events',
};

const SETTINGS_KEYS = {
  CALENDAR_ID: 'gcal_calendar_id',
  RANGE_START: 'sync_range_start',
  RANGE_END: 'sync_range_end',
  LAST_SYNC_TOKEN: 'gcal_sync_token',
  LAST_SYNC_AT: 'gcal_last_sync_at',
  LOOP_GUARD_SECONDS: 'loop_guard_seconds',
};

const SYNC_SOURCE = {
  WEBAPP: 'WEBAPP',
  GCAL: 'GCAL',
  SYSTEM: 'SYSTEM',
};

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(name) { return ss_().getSheetByName(name); }
function nowIso_() { return new Date().toISOString(); }

function toIsoDate_(d) {
  const dd = new Date(d);
  const y = dd.getFullYear();
  const m = String(dd.getMonth()+1).padStart(2,'0');
  const day = String(dd.getDate()).padStart(2,'0');
  return `${y}-${m}-${day}`;
}

function getSettingsMap_() {
  const sh = sh_(SHEET_NAMES.SETTINGS);
  const values = sh.getRange(2,1, sh.getLastRow()-1, 2).getValues();
  const map = {};
  values.forEach(([k,v]) => { if (k) map[String(k).trim()] = v; });
  return map;
}

function setSetting_(key, value) {
  const sh = sh_(SHEET_NAMES.SETTINGS);
  const last = sh.getLastRow();
  if (last < 2) return;
  const values = sh.getRange(2,1, last-1, 2).getValues();
  for (let i=0;i<values.length;i++){
    if (String(values[i][0]).trim() === key){
      sh.getRange(2+i, 2).setValue(value);
      return;
    }
  }
  sh.getRange(last+1, 1, 1, 2).setValues([[key, value]]);
}

function hashEventPayload_(payload) {
  const s = JSON.stringify(payload);
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function headerIndexMap_(sheet) {
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, idx)=> { if (h) map[String(h).trim()] = idx; });
  return map;
}
```

### 6.2 Sheet→GCal リソース変換

```javascript
function buildGcalEventResource_(e, calendarId) {
  const resource = {
    summary: e.title || '',
    location: e.location || '',
    description: e.memo || '',
    extendedProperties: {
      private: {
        lw_event_id: String(e.id),
        lw_origin: 'LW',
        lw_hash: String(e.sync_hash || ''),
      }
    }
  };

  if (e.is_all_day) {
    // 終日：end.date は翌日が仕様
    const startDate = String(e.date);
    const endDateObj = new Date(startDate + 'T00:00:00');
    endDateObj.setDate(endDateObj.getDate() + 1);
    resource.start = { date: startDate };
    resource.end = { date: toIsoDate_(endDateObj) };
  } else {
    const tz = Session.getScriptTimeZone();
    const start = `${e.date}T${e.start_time || '00:00'}:00`;
    const end = e.end_time ? `${e.date}T${e.end_time}:00` : start;
    resource.start = { dateTime: start, timeZone: tz };
    resource.end = { dateTime: end, timeZone: tz };
  }

  return resource;
}

function buildSyncHashFromEvent_(e) {
  const payload = {
    date: e.date,
    start_time: e.start_time,
    end_time: e.end_time,
    is_all_day: !!e.is_all_day,
    title: e.title || '',
    location: e.location || '',
    memo: e.memo || '',
    type: e.type || '',
    status: e.status || '',
    is_deleted: !!e.is_deleted,
  };
  return hashEventPayload_(payload);
}
```

### 6.3 Webアプリ保存時（Sheet→GCal即時同期）

```javascript
/**
 * WebApp: 予定保存（Sheetへ保存 → GCalへ反映）
 */
function saveEventWithGcalSync(payload) {
  const user = Session.getActiveUser().getEmail() || '';
  const settings = getSettingsMap_();
  const calendarId = settings[SETTINGS_KEYS.CALENDAR_ID] || 'primary';

  // 1) Sheetへ upsert（既存のcreateEvent/updateEventを流用可）
  // ...省略（既存実装を使用）...

  // 2) GCalへ反映
  const e = getEventById_(payload.id).data;
  const result = syncOneEventToGcal_(e, calendarId);

  // 3) gcal_event_id を書き戻し
  updateEventRow_(rowNum, {
    gcal_event_id: result.gcal_event_id,
    last_sync_at: nowIso_(),
  });

  return { ok: true, id: payload.id, gcal_event_id: result.gcal_event_id };
}

function syncOneEventToGcal_(e, calendarId) {
  // 削除フラグなら削除
  if (e.is_deleted) {
    if (e.gcal_event_id) {
      try {
        Calendar.Events.remove(calendarId, e.gcal_event_id);
      } catch (err) { /* 既に無い等は許容 */ }
    }
    return { gcal_event_id: e.gcal_event_id || '' };
  }

  const resource = buildGcalEventResource_(e, calendarId);

  if (!e.gcal_event_id) {
    // 新規作成
    const created = Calendar.Events.insert(resource, calendarId);
    return { gcal_event_id: created.id };
  } else {
    // 更新
    const patched = Calendar.Events.patch(resource, calendarId, e.gcal_event_id);
    return { gcal_event_id: patched.id };
  }
}
```

### 6.4 GCal→Sheet 定期取り込み

```javascript
/**
 * 時間トリガー用：GCalの変更をSheetへ反映
 * 推奨：1〜5分おき
 */
function syncFromGcalToSheet() {
  const settings = getSettingsMap_();
  const calendarId = settings[SETTINGS_KEYS.CALENDAR_ID] || 'primary';
  const rangeStart = settings[SETTINGS_KEYS.RANGE_START];
  const rangeEnd = settings[SETTINGS_KEYS.RANGE_END];
  const loopGuard = Number(settings[SETTINGS_KEYS.LOOP_GUARD_SECONDS] || 30);

  if (!rangeStart || !rangeEnd) {
    throw new Error('sync_range_start / sync_range_end が未設定');
  }

  const timeMin = new Date(rangeStart + 'T00:00:00Z').toISOString();
  const endObj = new Date(rangeEnd + 'T00:00:00Z');
  endObj.setDate(endObj.getDate()+1);
  const timeMax = endObj.toISOString();

  // Calendar API で取得
  const collected = [];
  let pageToken = null;
  const syncToken = settings[SETTINGS_KEYS.LAST_SYNC_TOKEN] || '';

  try {
    do {
      const res = Calendar.Events.list(calendarId, {
        singleEvents: true,
        showDeleted: true,
        timeMin,
        timeMax,
        maxResults: 2500,
        pageToken: pageToken || undefined,
        syncToken: syncToken || undefined,
      });
      (res.items || []).forEach(it => collected.push(it));
      pageToken = res.nextPageToken;
      if (res.nextSyncToken) {
        setSetting_(SETTINGS_KEYS.LAST_SYNC_TOKEN, res.nextSyncToken);
      }
    } while (pageToken);
  } catch (err) {
    // syncToken失効時はクリアして再取得
    setSetting_(SETTINGS_KEYS.LAST_SYNC_TOKEN, '');
    // ...再取得ロジック...
  }

  // Sheet側のインデックス作成
  const sh = eventsSheet_();
  const map = headerIndexMap_(sh);
  // ...gcal_event_id / id でマッピング...

  // 各イベントを処理
  collected.forEach(ev => {
    // ループガード
    const lastSyncAt = settings[SETTINGS_KEYS.LAST_SYNC_AT];
    if (lastSyncAt) {
      const dt = (Date.now() - new Date(lastSyncAt).getTime()) / 1000;
      if (dt < loopGuard) return;
    }

    // 紐付け：extendedProperties.private.lw_event_id 優先
    const priv = ev.extendedProperties?.private || {};
    const lwId = priv.lw_event_id || '';

    // rowNum を特定...

    // 衝突チェック（Last Write Wins）
    const sheetLast = /* Sheet側のlast_modified_at */;
    const gcalUpdated = ev.updated ? new Date(ev.updated) : null;
    if (sheetLast && gcalUpdated && sheetLast > gcalUpdated) {
      return; // Sheetが新しい → スキップ
    }

    // Sheetへ反映
    const patch = gcalEventToSheetPatch_(ev, calendarId);
    // ...updateEventRow_ or insertEventRow_...
  });

  setSetting_(SETTINGS_KEYS.LAST_SYNC_AT, nowIso_());
}

function gcalEventToSheetPatch_(ev, calendarId) {
  const isDeleted = (ev.status === 'cancelled');
  const patch = {
    gcal_event_id: ev.id,
    gcal_calendar_id: calendarId,
    title: ev.summary || '',
    location: ev.location || '',
    memo: ev.description || '',
    is_deleted: isDeleted,
    last_modified_at: ev.updated || nowIso_(),
    last_modified_by: ev.creator?.email || '',
  };

  // 日付/時刻
  if (ev.start?.date) {
    patch.is_all_day = true;
    patch.date = ev.start.date;
    patch.start_time = '';
    patch.end_time = '';
  } else if (ev.start?.dateTime) {
    patch.is_all_day = false;
    const d = new Date(ev.start.dateTime);
    patch.date = toIsoDate_(d);
    patch.start_time = Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm');
    if (ev.end?.dateTime) {
      const e = new Date(ev.end.dateTime);
      patch.end_time = Utilities.formatDate(e, Session.getScriptTimeZone(), 'HH:mm');
    }
  }

  patch.sync_hash = buildSyncHashFromEvent_(patch);
  return patch;
}
```

### 6.5 トリガー設定

```javascript
function installGcalSyncTrigger() {
  // 既存トリガー削除
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncFromGcalToSheet') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 1分おきに実行
  ScriptApp.newTrigger('syncFromGcalToSheet')
    .timeBased()
    .everyMinutes(1)
    .create();
}
```

---

## 7. 実装チェックリスト

### 事前準備
- [ ] Advanced Google Services で Calendar API を有効化
- [ ] Settingsシートに `gcal_calendar_id` を設定
- [ ] Settingsシートに `sync_range_start` / `sync_range_end` を設定
- [ ] Settingsシートに `loop_guard_seconds` を設定（推奨: 30）

### Eventsシート
- [ ] `gcal_event_id` 列を追加
- [ ] `gcal_calendar_id` 列を追加
- [ ] `last_modified_at` 列を追加
- [ ] `last_modified_by` 列を追加
- [ ] `last_sync_at` 列を追加
- [ ] `sync_source` 列を追加
- [ ] `sync_hash` 列を追加
- [ ] `is_all_day` 列を追加
- [ ] `is_deleted` 列を追加

### バックエンド
- [ ] ユーティリティ関数を追加
- [ ] `buildGcalEventResource_()` を実装
- [ ] `buildSyncHashFromEvent_()` を実装
- [ ] `syncOneEventToGcal_()` を実装
- [ ] `saveEventWithGcalSync()` を実装
- [ ] `syncFromGcalToSheet()` を実装
- [ ] `gcalEventToSheetPatch_()` を実装
- [ ] `installGcalSyncTrigger()` を実行

### 動作確認
- [ ] Webアプリで予定保存 → GCalに作成される
- [ ] Webアプリで予定更新 → GCalが更新される
- [ ] GCalで予定変更 → 1分以内にSheetに反映
- [ ] GCalで削除 → Sheetで `is_deleted=true`
- [ ] Sheetで削除 → GCalのイベントが消える

---

## 8. よくある間違いと対策

### 8.1 Calendar API が有効になっていない

**症状**: `Calendar is not defined` エラー

**対策**: Apps Script エディタ → サービス → Calendar API を追加

### 8.2 終日予定のend.dateが1日ずれる

**症状**: 1日の終日予定が2日間になる

**対策**: `end.date` は「終了日の翌日」を指定（Calendar API仕様）
```javascript
endDateObj.setDate(endDateObj.getDate() + 1);
resource.end = { date: toIsoDate_(endDateObj) };
```

### 8.3 ループが止まらない

**症状**: 同期が無限に繰り返される

**対策**:
1. `sync_source` を必ず設定
2. `last_sync_at` で30秒以内はスキップ
3. `sync_hash` が一致したらスキップ

### 8.4 syncTokenが失効

**症状**: `410 Gone` エラー

**対策**: try-catchで検知し、トークンをクリアして全件再取得
```javascript
catch (err) {
  setSetting_(SETTINGS_KEYS.LAST_SYNC_TOKEN, '');
  // 再取得...
}
```

### 8.5 extendedPropertiesが保存されない

**症状**: GCal側にメタ情報がない

**対策**: `insert` だけでなく、新規取り込み時にも `patch` でメタを保存
```javascript
Calendar.Events.patch({
  extendedProperties: { private: { lw_event_id: id, lw_origin: 'LW' } }
}, calendarId, evId);
```

---

## 9. 段階的実装ガイド

### Step A: 最小動作（Sheet→GCal のみ）
1. Calendar API有効化
2. `syncOneEventToGcal_()` 実装
3. 保存時にGCal作成・更新
4. `gcal_event_id` をSheetに保存

### Step B: GCal→Sheet 取り込み
1. `syncFromGcalToSheet()` 実装
2. 時間トリガー設定（1分おき）
3. ループ防止（sync_source / last_sync_at）

### Step C: 差分同期・衝突処理
1. syncToken対応
2. Last Write Wins実装
3. 削除同期
4. sync_hash一致スキップ

---

## 10. コピペ用サマリ（エージェント向け）

```
2ヶ月予定表（Sheets/GAS）とGoogleカレンダーを双方向同期する。

【データ設計】
- Sheets側Eventsに gcal_event_id / last_modified_at / last_sync_at / sync_source / sync_hash を追加
- GCal側は Calendar API（Advanced）で extendedProperties.private に lw_event_id 等を保存

【同期フロー】
- Sheet→GCal: 保存時に即時 insert/patch
- GCal→Sheet: 時間トリガーで1〜5分おきに list（timeMin/timeMax、可能ならsyncToken）

【ループ防止】
- sync_source と last_sync_at（30秒以内スキップ）と hash一致スキップ

【衝突ルール】
- Last Write Wins: Sheet last_modified_at と GCal updated の新しい方を正

【対象期間】
- 基準月〜翌月末＋数日

【除外】
- 繰り返し予定は当面対象外
```
