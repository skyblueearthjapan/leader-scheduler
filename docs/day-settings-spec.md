# 日付設定機能（出勤日/休日設定）設計仕様書

## 1. 機能概要

カレンダー上の特定日付に対して「出勤日」または「休日」を設定し、視覚的に区別できる機能。

### ユースケース
- 工場カレンダー：土曜日を出勤日に設定 → 週末の黄色背景を白に変更
- 祝日設定：平日を休日に設定 → 薄ピンク背景で表示

---

## 2. データ構造

### 2.1 スプレッドシート設計

**シート名**: `04_DaySettings`（自動作成）

| 列 | 内容 | 型 | 例 |
|---|---|---|---|
| A | 日付 | 文字列(YYYY-MM-DD) | `2026-01-10` |
| B | 種別 | 文字列 | `WORKDAY` or `HOLIDAY` |
| C | メモ | 文字列 | （任意） |
| D | 更新日時 | 文字列 | `2026-01-04 07:55:00` |
| E | 更新者 | 文字列 | `user@example.com` |

**ヘッダー行（1行目）**: `日付, 種別, メモ, 更新日時, 更新者`

### 2.2 フロントエンド状態管理

```javascript
// STATE オブジェクトに追加
const STATE = {
  // ... 既存のプロパティ
  daySettings: {}  // { 'YYYY-MM-DD': 'WORKDAY' | 'HOLIDAY' }
};
```

**重要**: 値は文字列のみ（オブジェクトではない）

```javascript
// 正しい形式
STATE.daySettings = {
  '2026-01-10': 'WORKDAY',
  '2026-01-12': 'HOLIDAY'
};

// 間違った形式（これはNG）
STATE.daySettings = {
  '2026-01-10': { type: 'WORKDAY', memo: '' }  // NG!
};
```

---

## 3. バックエンド実装（Google Apps Script）

### 3.1 定数定義

```javascript
const SHEET_DAY_SETTINGS = '04_DaySettings';
```

### 3.2 シート自動作成関数

```javascript
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
```

### 3.3 日付設定取得関数

**重要**: 戻り値は `{ 'YYYY-MM-DD': 'WORKDAY' }` 形式（文字列のみ）

```javascript
/**
 * 指定期間の日付設定を取得
 * @param {string} fromISO - 開始日 (YYYY-MM-DD)
 * @param {string} toISO - 終了日 (YYYY-MM-DD)
 * @returns {Object} { 'YYYY-MM-DD': 'WORKDAY'|'HOLIDAY' }
 */
function getDaySettingsMap_(fromISO, toISO) {
  const sh = ensureDaySettingsSheet_();
  const lastRow = sh.getLastRow();
  const map = {};

  if (lastRow < 2) return map;

  const settings = getSettings_();  // タイムゾーン取得用
  const tz = settings.tz;
  const values = sh.getRange(2, 1, lastRow - 1, 2).getValues(); // A:date, B:type

  for (const [dateVal, type] of values) {
    if (!dateVal) continue;

    let key;
    if (dateVal instanceof Date) {
      // Date型の場合はタイムゾーンを考慮して変換
      key = Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
    } else {
      // 文字列の場合はそのまま使用
      key = String(dateVal).trim();
    }

    const typeStr = String(type || '').trim();
    // YYYY-MM-DD形式かつ有効な種別のみ追加
    if (typeStr && key.match(/^\d{4}-\d{2}-\d{2}$/)) {
      map[key] = typeStr;  // 'WORKDAY' or 'HOLIDAY'（文字列のみ！）
    }
  }

  return map;
}
```

### 3.4 日付設定保存関数

```javascript
/**
 * 日付設定を追加/更新/削除
 * @param {string} dateISO - YYYY-MM-DD
 * @param {string|null} type - 'WORKDAY', 'HOLIDAY', または null（解除）
 * @param {string} memo - メモ（任意）
 * @returns {Object} { ok: boolean, action?: string, error?: string }
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

    // 編集可能期間チェック（既存関数を使用）
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
          key = Utilities.formatDate(dates[i][0], tz, 'yyyy-MM-dd');
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
```

### 3.5 Bootstrap関数への統合

```javascript
function getBootstrap() {
  try {
    const settings = getSettings_();
    const user = getUserContext_();
    const masters = getMasters_();
    const range = getEditableRange_(settings);
    const events = listEvents_(range.fromISO, range.toISO);
    const notes = getNotesMap_(/* ... */);

    // 日付設定を追加
    const daySettings = getDaySettingsMap_(range.fromISO, range.toISO);

    return {
      ok: true,
      user,
      settings,
      masters,
      range,
      events,
      notes,
      daySettings  // ← 追加
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
```

---

## 4. フロントエンド実装

### 4.1 HTML（モーダル内ボタン）

**配置位置**: モーダルの下部（スケジュール入力フォームの後）

```html
<!-- 日付設定ボタン（出勤日/休日）- 下部に配置 -->
<div class="daySettingButtons" id="daySettingButtons">
  <button type="button" class="daySettingBtn daySettingBtn--workday" id="btnWorkday">
    <span class="daySettingBtn__icon">🏭</span>
    <span>出勤日に設定</span>
  </button>
  <button type="button" class="daySettingBtn daySettingBtn--holiday" id="btnHoliday">
    <span class="daySettingBtn__icon">🎌</span>
    <span>休日に設定</span>
  </button>
</div>
```

### 4.2 CSS

```css
:root {
  --holiday: #fce7f3;  /* 薄ピンク - 休日設定 */
}

/* 日付設定ボタン */
.daySettingButtons {
  display: flex;
  gap: 8px;
  margin-top: 16px;
  padding-top: 16px;
  border-top: 1px solid var(--line);
}

.daySettingBtn {
  flex: 1;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 6px;
  padding: 10px 12px;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: #fff;
  cursor: pointer;
  font-size: 13px;
  font-weight: 600;
  transition: all 0.15s;
}

.daySettingBtn:hover {
  border-color: #cbd5e1;
}

.daySettingBtn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

/* アクティブ状態 */
.daySettingBtn.active {
  border-width: 2px;
}

.daySettingBtn--workday.active {
  background: #dbeafe;
  border-color: #3b82f6;
  color: #1d4ed8;
}

.daySettingBtn--holiday.active {
  background: #fce7f3;
  border-color: #ec4899;
  color: #be185d;
}

.daySettingBtn__icon {
  font-size: 16px;
}

/* カレンダーセルの日付設定スタイル */
.isHoliday {
  background: var(--holiday) !important;
}

.isHoliday.dayHead,
.isHoliday.wdayHead {
  color: #be185d !important;
}

.isWorkday {
  background: #ffffff !important;
}

.isWorkday.dayHead,
.isWorkday.wdayHead {
  color: var(--text) !important;
}
```

### 4.3 JavaScript

#### 4.3.1 状態管理に追加

```javascript
const STATE = {
  // ... 既存
  daySettings: {}  // { 'YYYY-MM-DD': 'WORKDAY' | 'HOLIDAY' }
};

// モーダル状態
let DAY_SETTING_CHANGED = false;  // 日付設定が変更されたかどうか
```

#### 4.3.2 Boot関数でのデータ読み込み

```javascript
async function boot() {
  const res = await gas('getBootstrap');

  // ... 既存の処理

  STATE.daySettings = res.daySettings || {};  // 追加

  renderBoards();
}
```

#### 4.3.3 ボタン状態更新関数

```javascript
function updateDaySettingButtons(dateISO) {
  const daySetting = STATE.daySettings[dateISO] || null;
  const btnWorkday = el('btnWorkday');
  const btnHoliday = el('btnHoliday');

  // アクティブ状態をリセット
  btnWorkday.classList.remove('active');
  btnHoliday.classList.remove('active');

  // 現在の設定に応じてアクティブ状態を設定
  if (daySetting === 'WORKDAY') {
    btnWorkday.classList.add('active');
  } else if (daySetting === 'HOLIDAY') {
    btnHoliday.classList.add('active');
  }

  // 編集可否に応じて無効化
  const canRole = STATE.user.role === 'editor' || STATE.user.role === 'admin';
  const canDate = isDateEditable(dateISO);
  const canEdit = canRole && canDate;

  btnWorkday.disabled = !canEdit;
  btnHoliday.disabled = !canEdit;
}
```

#### 4.3.4 トグル関数（楽観的UI更新）

```javascript
async function toggleDaySetting(type) {
  const dateISO = el('f_date').value;
  if (!dateISO) return;

  DAY_SETTING_CHANGED = true;

  const currentSetting = STATE.daySettings[dateISO];

  // 同じボタンを押したら解除、違うボタンなら切り替え
  let newSetting = null;
  if (currentSetting === type) {
    newSetting = null;
  } else {
    newSetting = type;
  }

  // 楽観的UI更新
  if (newSetting) {
    STATE.daySettings[dateISO] = newSetting;
  } else {
    delete STATE.daySettings[dateISO];
  }
  updateDaySettingButtons(dateISO);
  renderBoards();

  // バックグラウンドでサーバーに保存
  enqueueSave({
    execute: async () => {
      const res = await gas('setDaySetting', dateISO, newSetting, '');
      if (res && res.ok) {
        showToast(
          newSetting === 'WORKDAY' ? '出勤日に設定しました' :
          newSetting === 'HOLIDAY' ? '休日に設定しました' :
          '設定を解除しました',
          'success', 2000
        );
      } else {
        // 失敗したら元に戻す
        if (currentSetting) {
          STATE.daySettings[dateISO] = currentSetting;
        } else {
          delete STATE.daySettings[dateISO];
        }
        updateDaySettingButtons(dateISO);
        renderBoards();
        showToast('設定の保存に失敗しました', 'error', 4000);
      }
    },
    onError: (e) => {
      // 失敗したら元に戻す
      if (currentSetting) {
        STATE.daySettings[dateISO] = currentSetting;
      } else {
        delete STATE.daySettings[dateISO];
      }
      updateDaySettingButtons(dateISO);
      renderBoards();
      showToast('設定エラー: ' + e.message, 'error', 4000);
    }
  });
}
```

#### 4.3.5 カレンダー描画での適用

```javascript
function renderMonthGrid(container, monthStart) {
  // ... 既存の処理

  for (let day = 1; day <= 31; day++) {
    if (day <= daysInMonth) {
      const dt = new Date(monthStart.getFullYear(), monthStart.getMonth(), day);
      const iso = ymd(dt);  // 'YYYY-MM-DD'形式

      // 日付設定を取得
      const daySetting = STATE.daySettings[iso];

      // 日付設定による上書き（週末クラス設定の後に適用）
      if (daySetting === 'WORKDAY') {
        dayHead.classList.add('isWorkday');
        wdayHead.classList.add('isWorkday');
        cell.classList.add('isWorkday');
      } else if (daySetting === 'HOLIDAY') {
        dayHead.classList.add('isHoliday');
        wdayHead.classList.add('isHoliday');
        cell.classList.add('isHoliday');
      }
    }
  }
}
```

#### 4.3.6 モーダル開閉時の処理

```javascript
function openCreateModal(dateISO) {
  DAY_SETTING_CHANGED = false;  // リセット
  // ... 既存の処理
  updateDaySettingButtons(dateISO);  // 追加
}

function openEditModal(id) {
  DAY_SETTING_CHANGED = false;  // リセット
  // ... 既存の処理
  updateDaySettingButtons(e.date);  // 追加
}
```

#### 4.3.7 保存時のバリデーション調整

```javascript
function onSave() {
  const payload = getPayload();

  // 日付設定のみ変更した場合は、スケジュール保存をスキップ
  const hasScheduleInput = payload.title.trim() || payload.start_time ||
                          payload.end_time || payload.location || payload.memo;
  if (DAY_SETTING_CHANGED && !hasScheduleInput && MODAL_MODE === 'create') {
    openModal(false);
    return;
  }

  // ... 既存のバリデーション
}
```

#### 4.3.8 イベントリスナー登録

```javascript
document.addEventListener('DOMContentLoaded', () => {
  // ... 既存のリスナー

  // 日付変更時にボタン状態を更新
  el('f_date').addEventListener('change', () => {
    const dateISO = el('f_date').value;
    applyEditability(dateISO);
    updateDaySettingButtons(dateISO);  // 追加
  });

  // 日付設定ボタン
  el('btnWorkday').addEventListener('click', () => toggleDaySetting('WORKDAY'));
  el('btnHoliday').addEventListener('click', () => toggleDaySetting('HOLIDAY'));
});
```

---

## 5. 実装チェックリスト

### バックエンド
- [ ] `SHEET_DAY_SETTINGS` 定数を追加
- [ ] `ensureDaySettingsSheet_()` 関数を追加
- [ ] `getDaySettingsMap_(fromISO, toISO)` 関数を追加
  - **重要**: 戻り値は `{ 'YYYY-MM-DD': 'WORKDAY' }` 形式（文字列のみ）
- [ ] `setDaySetting(dateISO, type, memo)` 関数を追加
- [ ] `getBootstrap()` に `daySettings` を追加

### フロントエンド
- [ ] `STATE.daySettings = {}` を追加
- [ ] `DAY_SETTING_CHANGED` フラグを追加
- [ ] HTMLにボタンを追加（モーダル下部）
- [ ] CSSスタイルを追加
- [ ] `updateDaySettingButtons()` 関数を追加
- [ ] `toggleDaySetting()` 関数を追加
- [ ] `boot()` で `daySettings` を読み込み
- [ ] `renderMonthGrid()` で `isWorkday`/`isHoliday` クラスを適用
- [ ] モーダル開閉時に `updateDaySettingButtons()` を呼び出し
- [ ] `onSave()` でバリデーション調整
- [ ] イベントリスナーを登録

---

## 6. よくある間違いと対策

### 6.1 データ形式の不一致（最重要）

**問題**: バックエンドがオブジェクトを返し、フロントエンドが文字列を期待

```javascript
// バックエンド（間違い）
map[key] = { type: 'WORKDAY', memo: '' };

// フロントエンド（期待）
if (daySetting === 'WORKDAY') { ... }  // 常にfalseになる
```

**対策**: `getDaySettingsMap_` は必ず文字列のみを返す

```javascript
map[key] = typeStr;  // 'WORKDAY' or 'HOLIDAY'
```

### 6.2 タイムゾーン問題

**問題**: スプレッドシートの日付がDate型に変換され、タイムゾーンでずれる

**対策**: 読み込み時にタイムゾーンを明示的に指定

```javascript
if (dateVal instanceof Date) {
  key = Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd');
}
```

### 6.3 楽観的UI更新の巻き戻し忘れ

**問題**: サーバー保存失敗時にUIが不整合になる

**対策**: エラー時に必ず元の状態に戻す

```javascript
onError: (e) => {
  if (currentSetting) {
    STATE.daySettings[dateISO] = currentSetting;
  } else {
    delete STATE.daySettings[dateISO];
  }
  updateDaySettingButtons(dateISO);
  renderBoards();
}
```

### 6.4 モーダル開閉時のフラグリセット忘れ

**問題**: `DAY_SETTING_CHANGED` がリセットされず誤動作

**対策**: `openCreateModal` と `openEditModal` の両方でリセット

```javascript
function openCreateModal(dateISO) {
  DAY_SETTING_CHANGED = false;
  // ...
}
```

---

## 7. テスト手順

1. アプリを開き、任意の日付をクリック
2. モーダル下部の「出勤日に設定」ボタンをクリック
3. カレンダー上でその日の背景が白になることを確認
4. 「保存」をクリック（件名未入力でもエラーなしで閉じる）
5. スプレッドシートの `04_DaySettings` シートにデータが保存されていることを確認
6. アプリを再読み込みし、設定が保持されていることを確認
7. 同じ日をクリックし、「出勤日に設定」ボタンがアクティブ状態であることを確認
8. 再度クリックして設定解除できることを確認
