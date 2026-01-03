# A＋C＋B保存方式 実装ガイド

コーディングエージェント向け指示書

---

## 概要

A＋C＋B保存方式には2つのパターンがあります：

| パターン | 用途 | A | C | B |
|----------|------|---|---|---|
| **UX最適化** | 入力フォームの高速化 | ローカルドラフト | 変更キュー | 遅延サーバ保存 |
| **三層保存** | ファイルアップロード | 原本ファイル | 確認用PDF | 正規化DB |

このドキュメントでは**UX最適化パターン**を解説します。

---

## A＋C＋B（入力UX最適化）の定義

| 層 | 名称 | 役割 |
|----|------|------|
| **A** | ローカルドラフト | 入力をブラウザに即時保存（localStorage/IndexedDB） |
| **C** | 変更キュー | 変更差分を溜めて、まとめて送信 |
| **B** | サーバ確定 | 一定条件でCをサーバへ反映（GAS） |

---

## 目的

入力フォームの「保存処理が重くて入力が止まる」問題を解消する。
ユーザーの入力は**常に即時反映（A）し、サーバ保存（B）は裏側でまとめて（C→B）**行う。

---

## 要件（必須）

### 1) 入力は"常に止めない"
- キー入力・選択・チェックなどのイベントで**サーバ保存を同期的に呼ばない**
- 入力後の表示・編集はローカル状態（A）を参照して即時更新

### 2) 自動保存トリガー（Bへ反映する条件）

次のいずれかで C（キュー）をB（サーバ）へ送る：

| トリガー | 説明 |
|----------|------|
| 無操作時間 | 最後の入力から3秒何もなければ送信 |
| 一定間隔 | 30秒ごとに未送信があれば送信 |
| 画面遷移 | モーダル閉じる/タブ切替前 |
| 明示的保存 | 保存ボタン押下で即送信 |

### 3) 変更の溜め込み（C）
- 変更ごとに「差分イベント」をキューに積む
- 無操作になったらキューを1つのパッチに集約して送る
- サーバ送信成功したらキューをクリア

### 4) ローカル永続化（A）
- ブラウザリロード/クラッシュ/回線断でも復旧できるように
- A（ドラフト）とC（キュー）は**必ず永続化する**

### 5) 状態表示（UX）

画面に必ず表示：
- ✓ 保存済み
- ⏳ 保存中…
- ⚠ 未保存（オフライン）
- ✗ 保存失敗（再試行ボタン）

### 6) データ整合性（重複・競合防止）
- B側は**upsert**（同一レコードIDは更新）に統一
- クライアント側はrevision（連番/timestamp）を持つ
- サーバは古いrevisionを弾く

---

## 実装アーキテクチャ

```
┌─────────────────────────────────────────────────────┐
│                    ブラウザ                          │
├─────────────────────────────────────────────────────┤
│  ┌─────────┐    ┌─────────┐    ┌─────────┐        │
│  │   UI    │───→│ A:Draft │───→│ C:Queue │        │
│  │ (Form)  │    │(local)  │    │(local)  │        │
│  └─────────┘    └─────────┘    └────┬────┘        │
│                                      │             │
│                          3秒無操作 / 30秒間隔       │
│                                      ▼             │
│                              ┌─────────────┐      │
│                              │   flush()   │      │
│                              └──────┬──────┘      │
└─────────────────────────────────────┼──────────────┘
                                      │
                                      ▼
┌─────────────────────────────────────────────────────┐
│                  B: サーバ（GAS）                    │
│  ┌─────────────────────────────────────────────┐   │
│  │  api_applyPatch(recordId, patch, revision)  │   │
│  │  → upsert → return {ok, serverRevision}     │   │
│  └─────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────┘
```

---

## データ構造

### A: draftStore（localStorage）

```javascript
// キー: draft:{app}:{recordId}
{
  "draft:schedule:evt-123": {
    date: "2026-01-15",
    start_time: "13:00",
    end_time: "14:00",
    type: "MEET",
    title: "会議",
    location: "本社",
    memo: "",
    revision: 5,
    updatedAt: "2026-01-03T10:30:00Z"
  }
}
```

### C: changeQueue（localStorage）

```javascript
// キー: queue:{app}:{recordId}
{
  "queue:schedule:evt-123": [
    { field: "title", value: "定例会議", at: 1704268200000 },
    { field: "start_time", value: "14:00", at: 1704268205000 }
  ]
}
```

---

## クライアント実装（JavaScript）

### 主要関数

```javascript
// 1. フィールド変更時
onFieldChange(recordId, field, value) {
  // A: ドラフトに即時反映
  updateDraft(recordId, field, value);

  // C: キューに追加
  pushToQueue(recordId, { field, value, at: Date.now() });

  // 無操作3秒でflush
  scheduleFlush(recordId, 3000);
}

// 2. サーバ送信
async flush(recordId) {
  if (isFlushing) return; // 二重送信防止

  const queue = getQueue(recordId);
  if (!queue.length) return;

  // キュー → パッチに集約（同じfieldは最後の値のみ）
  const patch = mergeQueueToPatch(queue);
  const draft = getDraft(recordId);

  setStatus('saving');

  try {
    const result = await gas('applyPatch', recordId, patch, draft.revision);
    if (result.ok) {
      clearQueue(recordId);
      updateDraftRevision(recordId, result.serverRevision);
      setStatus('saved');
    } else {
      setStatus('conflict');
    }
  } catch (e) {
    setStatus('offline');
    scheduleRetry(recordId, 30000);
  }
}

// 3. 起動時の復元
onBoot() {
  const drafts = getAllDrafts();
  const queues = getAllQueues();

  // 未送信キューがあれば通知
  if (hasUnsentQueues(queues)) {
    showNotification('未送信のデータがあります');
    flushAll();
  }
}
```

---

## サーバ実装（GAS）

```javascript
/**
 * パッチを適用（upsert）
 * @param {string} recordId - レコードID（新規はtmp-xxxの場合あり）
 * @param {Object} patch - 変更フィールド
 * @param {number} clientRevision - クライアント側のリビジョン
 */
function applyPatch(recordId, patch, clientRevision) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const sh = getEventsSheet_();

    // 新規（tmp-で始まる）か既存か
    const isNew = recordId.startsWith('tmp-');

    if (isNew) {
      // 新規作成
      const newId = Utilities.getUuid();
      const newRevision = 1;
      // ... 行追加
      return { ok: true, recordId: newId, serverRevision: newRevision };
    } else {
      // 既存更新
      const row = findRowById_(sh, recordId);
      if (!row) return { ok: false, error: 'NOT_FOUND' };

      const serverRevision = getRevision_(sh, row);

      // リビジョンチェック
      if (clientRevision < serverRevision) {
        return { ok: false, error: 'CONFLICT', serverRevision };
      }

      // パッチ適用
      applyPatchToRow_(sh, row, patch);
      const newRevision = serverRevision + 1;
      setRevision_(sh, row, newRevision);

      return { ok: true, serverRevision: newRevision };
    }
  } finally {
    lock.releaseLock();
  }
}
```

---

## 新規レコードの扱い（パターン1: 仮ID方式）

```
1. モーダル開く → tmp-{uuid} を生成
2. 入力 → A/Cに tmp-xxx で保存
3. 初回flush → サーバが正式ID発行
4. A/Cのキーを tmp-xxx → 正式ID に置換
```

メリット：
- 入力開始時にサーバ通信不要
- オフラインでも新規作成開始可能

---

## 状態表示UI

```html
<div class="save-status">
  <span class="status-icon"></span>
  <span class="status-text"></span>
</div>

<style>
.save-status.saved .status-icon::before { content: "✓"; color: #10b981; }
.save-status.saving .status-icon::before { content: "⏳"; }
.save-status.offline .status-icon::before { content: "⚠"; color: #f59e0b; }
.save-status.error .status-icon::before { content: "✗"; color: #ef4444; }
</style>
```

---

## beforeunload対策

ブラウザ制約で「閉じる直前に必ず送信」は保証できない。

**対策：**
- A/Cがローカルに残ることを正とする
- 次回起動時に未送信キューを検知して自動送信
- 送信できない時は「未保存」表示

---

## 受け入れ条件（テスト観点）

- [ ] 入力中に保存待ちでUIが固まらない
- [ ] 3秒無操作で自動的に「保存中→保存済み」になる
- [ ] 回線断で「未保存」になるが、入力は継続できる
- [ ] 回線復帰で自動送信され、Bに反映される
- [ ] リロードしても入力内容が復元される
- [ ] 同じレコードの保存が重複行にならず更新される（upsert）

---

## 横展開時のカスタマイズ箇所

| 項目 | 説明 |
|------|------|
| recordIdの定義 | 何をキーに1レコードとするか |
| patchのフィールド | フォーム項目の一覧 |
| Bの保存先 | スプレッドシートの列定義 |

それ以外（A/C/Bの仕組み、状態表示、flush/リトライ）は共通化可能。

---

*最終更新: 2026-01-03*
