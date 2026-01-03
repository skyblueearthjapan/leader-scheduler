/**
 * A＋C＋B保存方式（三層保存）サービス
 *
 * A: 原本保存（Source Archive）
 * C: 確認用アウトプット（Confirmation Artifact）
 * B: 正規化データ（Normalized Data）- 既存のスプレッドシート
 */

// =============================================================================
// 定数
// =============================================================================

const ACB_SETTINGS_FOLDER_ID_CELL = 'B10';  // 01_Settingsシートの共有ドライブフォルダID

// フォルダ名
const FOLDER_SOURCE_A = '01_Source_A';
const FOLDER_CONFIRM_C = '02_Confirm_C';
const FOLDER_SYSTEM_B = '03_System_B';
const FOLDER_LOG = '99_Log';

// =============================================================================
// フォルダ管理
// =============================================================================

/**
 * ACB用のルートフォルダIDを取得
 */
function getACBRootFolderId_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  if (!sh) return null;

  const folderId = sh.getRange(ACB_SETTINGS_FOLDER_ID_CELL).getValue();
  return folderId ? String(folderId).trim() : null;
}

/**
 * ACBフォルダ構造を初期化（なければ作成）
 * @param {string} rootFolderId - ルートフォルダID
 * @returns {Object} {sourceA, confirmC, systemB, log}
 */
function ensureACBFolders_(rootFolderId) {
  if (!rootFolderId) {
    throw new Error('ACB_ROOT_FOLDER_NOT_SET: 01_SettingsのB10にフォルダIDを設定してください');
  }

  const rootFolder = DriveApp.getFolderById(rootFolderId);

  const folders = {};
  const folderNames = [FOLDER_SOURCE_A, FOLDER_CONFIRM_C, FOLDER_SYSTEM_B, FOLDER_LOG];

  for (const name of folderNames) {
    const existing = rootFolder.getFoldersByName(name);
    if (existing.hasNext()) {
      folders[name] = existing.next();
    } else {
      folders[name] = rootFolder.createFolder(name);
    }
  }

  return {
    sourceA: folders[FOLDER_SOURCE_A],
    confirmC: folders[FOLDER_CONFIRM_C],
    systemB: folders[FOLDER_SYSTEM_B],
    log: folders[FOLDER_LOG]
  };
}

// =============================================================================
// A: 原本保存（Source Archive）
// =============================================================================

/**
 * アップロードされたファイルを原本として保存
 * @param {string} base64Data - Base64エンコードされたファイルデータ
 * @param {string} filename - ファイル名
 * @param {string} mimeType - MIMEタイプ
 * @param {string} subFolder - サブフォルダ名（任意、例：月キー）
 * @returns {Object} {ok, sourceFileId, sourceUrl, error}
 */
function uploadSourceA(base64Data, filename, mimeType, subFolder) {
  try {
    const rootFolderId = getACBRootFolderId_();
    const folders = ensureACBFolders_(rootFolderId);

    // サブフォルダ（月キーなど）があれば作成
    let targetFolder = folders.sourceA;
    if (subFolder) {
      const subFolders = targetFolder.getFoldersByName(subFolder);
      if (subFolders.hasNext()) {
        targetFolder = subFolders.next();
      } else {
        targetFolder = targetFolder.createFolder(subFolder);
      }
    }

    // ファイル保存
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      mimeType,
      filename
    );
    const file = targetFolder.createFile(blob);

    return {
      ok: true,
      sourceFileId: file.getId(),
      sourceUrl: file.getUrl()
    };
  } catch (e) {
    console.error('uploadSourceA error:', e);
    return { ok: false, error: e.message };
  }
}

/**
 * 原本ファイルのURLを取得
 * @param {string} fileId
 * @returns {string|null}
 */
function getSourceAUrl_(fileId) {
  if (!fileId) return null;
  try {
    return DriveApp.getFileById(fileId).getUrl();
  } catch (e) {
    return null;
  }
}

// =============================================================================
// C: 確認用アウトプット（Confirmation Artifact）
// =============================================================================

/**
 * 確認用PDFを生成して保存
 * @param {string} htmlContent - HTML内容
 * @param {string} filename - ファイル名（拡張子なし）
 * @param {string} subFolder - サブフォルダ名（任意）
 * @returns {Object} {ok, confirmFileId, confirmUrl, error}
 */
function buildConfirmC(htmlContent, filename, subFolder) {
  try {
    const rootFolderId = getACBRootFolderId_();
    const folders = ensureACBFolders_(rootFolderId);

    // サブフォルダ
    let targetFolder = folders.confirmC;
    if (subFolder) {
      const subFolders = targetFolder.getFoldersByName(subFolder);
      if (subFolders.hasNext()) {
        targetFolder = subFolders.next();
      } else {
        targetFolder = targetFolder.createFolder(subFolder);
      }
    }

    // HTML → PDF変換
    const blob = HtmlService.createHtmlOutput(htmlContent)
      .getBlob()
      .setName(filename + '.pdf');

    const file = targetFolder.createFile(blob);

    return {
      ok: true,
      confirmFileId: file.getId(),
      confirmUrl: file.getUrl()
    };
  } catch (e) {
    console.error('buildConfirmC error:', e);
    return { ok: false, error: e.message };
  }
}

/**
 * 確認用HTMLファイルを保存
 * @param {string} htmlContent - HTML内容
 * @param {string} filename - ファイル名（拡張子なし）
 * @param {string} subFolder - サブフォルダ名（任意）
 * @returns {Object} {ok, confirmFileId, confirmUrl, error}
 */
function buildConfirmCHtml(htmlContent, filename, subFolder) {
  try {
    const rootFolderId = getACBRootFolderId_();
    const folders = ensureACBFolders_(rootFolderId);

    let targetFolder = folders.confirmC;
    if (subFolder) {
      const subFolders = targetFolder.getFoldersByName(subFolder);
      if (subFolders.hasNext()) {
        targetFolder = subFolders.next();
      } else {
        targetFolder = targetFolder.createFolder(subFolder);
      }
    }

    const blob = Utilities.newBlob(htmlContent, 'text/html', filename + '.html');
    const file = targetFolder.createFile(blob);

    return {
      ok: true,
      confirmFileId: file.getId(),
      confirmUrl: file.getUrl()
    };
  } catch (e) {
    console.error('buildConfirmCHtml error:', e);
    return { ok: false, error: e.message };
  }
}

/**
 * 確認用ファイルのURLを取得
 * @param {string} fileId
 * @returns {string|null}
 */
function getConfirmCUrl_(fileId) {
  if (!fileId) return null;
  try {
    return DriveApp.getFileById(fileId).getUrl();
  } catch (e) {
    return null;
  }
}

// =============================================================================
// B: 正規化データ（既存のスプレッドシート構造を活用）
// =============================================================================

// Bは既存のCode.gsの関数を使用：
// - createEvent() / updateEvent() / deleteEvent()
// - listEvents_() / getEvents()
//
// A/Cとの紐付けが必要な場合は、04_Eventsシートに
// sourceFileId, confirmFileId 列を追加して対応

// =============================================================================
// ログ出力
// =============================================================================

/**
 * インポートログを保存
 * @param {string} action - アクション名
 * @param {Object} details - 詳細情報
 * @returns {Object} {ok, logFileId, error}
 */
function writeLog_(action, details) {
  try {
    const rootFolderId = getACBRootFolderId_();
    if (!rootFolderId) return { ok: false, error: 'NO_ROOT_FOLDER' };

    const folders = ensureACBFolders_(rootFolderId);
    const settings = getSettings_();
    const tz = settings.tz;
    const now = new Date();
    const timestamp = Utilities.formatDate(now, tz, 'yyyy-MM-dd_HH-mm-ss');

    const logContent = JSON.stringify({
      timestamp: Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'),
      action: action,
      user: Session.getActiveUser().getEmail(),
      details: details
    }, null, 2);

    const filename = `${action}_${timestamp}.json`;
    const blob = Utilities.newBlob(logContent, 'application/json', filename);
    const file = folders.log.createFile(blob);

    return { ok: true, logFileId: file.getId() };
  } catch (e) {
    console.error('writeLog_ error:', e);
    return { ok: false, error: e.message };
  }
}

// =============================================================================
// 統合API（フロントエンドから呼び出し）
// =============================================================================

/**
 * ファイルインポート（A→B→Cの一連処理）
 * @param {Object} params
 * @param {string} params.base64Data - ファイルデータ
 * @param {string} params.filename - ファイル名
 * @param {string} params.mimeType - MIMEタイプ
 * @param {string} params.monthKey - 月キー（YYYY-MM）
 * @returns {Object} {ok, sourceFileId, recordIds, confirmFileId, error}
 */
function importWithACB(params) {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    return { ok: false, error: 'LOCK_TIMEOUT' };
  }

  try {
    const user = getUserContext_();
    if (!(user.role === 'editor' || user.role === 'admin')) {
      return { ok: false, error: 'FORBIDDEN' };
    }

    // Step 1: A（原本保存）
    const sourceResult = uploadSourceA(
      params.base64Data,
      params.filename,
      params.mimeType,
      params.monthKey
    );
    if (!sourceResult.ok) {
      return { ok: false, error: 'SOURCE_SAVE_FAILED: ' + sourceResult.error };
    }

    // Step 2: B（正規化）- ファイルタイプに応じてパース
    // ※実際のパース処理はファイルタイプに応じて実装
    // ここでは基本構造のみ提供
    const recordIds = [];
    // const parsedData = parseFile_(params.base64Data, params.mimeType);
    // for (const item of parsedData) {
    //   const result = createEvent(item);
    //   if (result.ok) recordIds.push(result.event_id);
    // }

    // Step 3: C（確認用生成）
    const confirmHtml = generateImportConfirmHtml_(
      params.filename,
      sourceResult.sourceFileId,
      recordIds.length,
      new Date()
    );
    const confirmResult = buildConfirmCHtml(
      confirmHtml,
      `import_${params.monthKey}_${Date.now()}`,
      params.monthKey
    );

    // ログ出力
    writeLog_('IMPORT', {
      filename: params.filename,
      sourceFileId: sourceResult.sourceFileId,
      recordCount: recordIds.length,
      confirmFileId: confirmResult.confirmFileId
    });

    return {
      ok: true,
      sourceFileId: sourceResult.sourceFileId,
      sourceUrl: sourceResult.sourceUrl,
      recordIds: recordIds,
      confirmFileId: confirmResult.ok ? confirmResult.confirmFileId : null,
      confirmUrl: confirmResult.ok ? confirmResult.confirmUrl : null
    };
  } catch (e) {
    console.error('importWithACB error:', e);
    writeLog_('IMPORT_ERROR', { error: e.message });
    return { ok: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * インポート確認用HTMLを生成
 */
function generateImportConfirmHtml_(filename, sourceFileId, recordCount, timestamp) {
  const settings = getSettings_();
  const tz = settings.tz;
  const timeStr = Utilities.formatDate(timestamp, tz, 'yyyy-MM-dd HH:mm:ss');

  return `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <title>インポート確認 - ${filename}</title>
  <style>
    body { font-family: sans-serif; padding: 20px; }
    h1 { color: #1e40af; }
    table { border-collapse: collapse; margin-top: 20px; }
    th, td { border: 1px solid #e5e7eb; padding: 8px 12px; text-align: left; }
    th { background: #f3f4f6; }
    .success { color: #059669; }
    .info { color: #6b7280; font-size: 12px; margin-top: 20px; }
  </style>
</head>
<body>
  <h1>インポート確認レポート</h1>
  <table>
    <tr><th>項目</th><th>値</th></tr>
    <tr><td>ファイル名</td><td>${filename}</td></tr>
    <tr><td>原本ID</td><td>${sourceFileId}</td></tr>
    <tr><td>インポート件数</td><td class="success">${recordCount} 件</td></tr>
    <tr><td>処理日時</td><td>${timeStr}</td></tr>
  </table>
  <p class="info">このファイルは確認用に自動生成されました。</p>
</body>
</html>`;
}

/**
 * ACB設定状態を取得（デバッグ用）
 */
function getACBStatus() {
  try {
    const rootFolderId = getACBRootFolderId_();
    if (!rootFolderId) {
      return {
        ok: false,
        configured: false,
        message: '01_SettingsのB10にDriveフォルダIDを設定してください'
      };
    }

    const folders = ensureACBFolders_(rootFolderId);
    return {
      ok: true,
      configured: true,
      rootFolderId: rootFolderId,
      folders: {
        sourceA: folders.sourceA.getId(),
        confirmC: folders.confirmC.getId(),
        systemB: folders.systemB.getId(),
        log: folders.log.getId()
      }
    };
  } catch (e) {
    return {
      ok: false,
      configured: false,
      error: e.message
    };
  }
}
