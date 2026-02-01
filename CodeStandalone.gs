// ===== スタンドアロン版 - 複数ドキュメント自動処理 =====
// 設定は Config.gs に記載されています
//
// 使い方:
//   1. Config.gs で設定を編集
//   2. connectToDropbox() でDropbox認証
//   3. processDocument0() 等で初回処理（1ドキュメントずつ）
//   4. setupAutoProcessing() でトリガー設定
//   5. 以降は15分おきに自動実行されます
//   6. 停止したい場合は removeAutoProcessingTrigger() を実行

// ===== 自動処理メイン関数 =====

/**
 * 初回セットアップ（トリガー設定のみ）
 * 初回の全件処理は processOneDocumentFull() で1ドキュメントずつ手動実行してください
 */
function setupAutoProcessing() {
  const service = getDropboxService();

  if (!service.hasAccess()) {
    console.log('Dropboxに接続されていません。connectToDropbox() を実行してください。');
    console.log('認証URL: ' + service.getAuthorizationUrl());
    return;
  }

  // 既存のトリガーを削除
  removeAutoProcessingTrigger();

  // 新しいトリガーを設定（15分おき）
  ScriptApp.newTrigger('autoProcessAllDocuments')
    .timeBased()
    .everyMinutes(AUTO_PROCESS_INTERVAL_MINUTES)
    .create();

  console.log(`トリガーを設定しました（${AUTO_PROCESS_INTERVAL_MINUTES}分おき）`);
  console.log('');
  console.log('【次のステップ】');
  console.log('各ドキュメントの初回処理を行ってください：');
  for (let i = 0; i < TARGET_DOCUMENT_IDS.length; i++) {
    console.log(`  processDocument${i}()  // ${TARGET_DOCUMENT_IDS[i].name}`);
  }
}

/**
 * 1つのドキュメントを全件処理（初回用・手動実行）
 * @param {string} docId - ドキュメントID
 */
function processOneDocumentFull(docId) {
  const service = getDropboxService();

  if (!service.hasAccess()) {
    console.log('Dropboxに接続されていません。');
    return;
  }

  const docInfo = TARGET_DOCUMENT_IDS.find(d => d.id === docId);
  const docName = docInfo ? docInfo.name : docId;

  console.log(`処理開始: ${docName}`);
  const startTime = new Date();

  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    const paragraphs = body.getParagraphs();

    // 画像挿入処理（全件）
    const imageResult = processImagesForDocument(paragraphs, service, true);
    console.log(`  画像: 挿入${imageResult.inserted} 更新${imageResult.updated} NoImage${imageResult.noImage} スキップ${imageResult.skipped}`);

    // コメント色付け処理
    const colorResult = applyCommentColorsForParagraphs(paragraphs);
    console.log(`  色付け: グレー${colorResult.gray} 紫${colorResult.purple} 黒字${colorResult.skipped}`);

    const elapsed = (new Date() - startTime) / 1000;
    console.log(`完了（${elapsed.toFixed(1)}秒）`);

  } catch (e) {
    console.log(`エラー: ${e.message}`);
  }
}

// ===== 各ドキュメントの初回処理用関数（手動実行用） =====
// processDocument0(), processDocument1(), ... として自動生成
function processDocument0() { if (TARGET_DOCUMENT_IDS[0]) processOneDocumentFull(TARGET_DOCUMENT_IDS[0].id); }
function processDocument1() { if (TARGET_DOCUMENT_IDS[1]) processOneDocumentFull(TARGET_DOCUMENT_IDS[1].id); }
function processDocument2() { if (TARGET_DOCUMENT_IDS[2]) processOneDocumentFull(TARGET_DOCUMENT_IDS[2].id); }
function processDocument3() { if (TARGET_DOCUMENT_IDS[3]) processOneDocumentFull(TARGET_DOCUMENT_IDS[3].id); }
function processDocument4() { if (TARGET_DOCUMENT_IDS[4]) processOneDocumentFull(TARGET_DOCUMENT_IDS[4].id); }
function processDocument5() { if (TARGET_DOCUMENT_IDS[5]) processOneDocumentFull(TARGET_DOCUMENT_IDS[5].id); }
function processDocument6() { if (TARGET_DOCUMENT_IDS[6]) processOneDocumentFull(TARGET_DOCUMENT_IDS[6].id); }
function processDocument7() { if (TARGET_DOCUMENT_IDS[7]) processOneDocumentFull(TARGET_DOCUMENT_IDS[7].id); }
function processDocument8() { if (TARGET_DOCUMENT_IDS[8]) processOneDocumentFull(TARGET_DOCUMENT_IDS[8].id); }
function processDocument9() { if (TARGET_DOCUMENT_IDS[9]) processOneDocumentFull(TARGET_DOCUMENT_IDS[9].id); }

/**
 * 定期実行される関数（トリガーから呼ばれる）
 * 全ドキュメントを20件バッチで処理
 */
function autoProcessAllDocuments() {
  const service = getDropboxService();

  if (!service.hasAccess()) {
    console.log('Dropbox認証が切れています。再認証が必要です。');
    return;
  }

  console.log('自動処理を開始します...');
  processAllDocumentsBatch();
}

/**
 * 全ドキュメントを20件バッチで処理（定期実行用）
 */
function processAllDocumentsBatch() {
  const service = getDropboxService();
  const startTime = new Date();

  const results = {
    documents: [],
    totalImages: { inserted: 0, updated: 0, noImage: 0, skipped: 0, errors: 0 },
    totalColors: { gray: 0, purple: 0, skipped: 0 }
  };

  for (const docInfo of TARGET_DOCUMENT_IDS) {
    const docResult = {
      name: docInfo.name,
      images: null,
      colors: null,
      error: null
    };

    try {
      console.log(`処理中: ${docInfo.name}`);
      const doc = DocumentApp.openById(docInfo.id);
      const body = doc.getBody();
      const paragraphs = body.getParagraphs();

      // 画像挿入処理（20件バッチ）
      const imageResult = processImagesForDocument(paragraphs, service, false);
      docResult.images = imageResult;
      results.totalImages.inserted += imageResult.inserted;
      results.totalImages.updated += imageResult.updated;
      results.totalImages.noImage += imageResult.noImage;
      results.totalImages.skipped += imageResult.skipped;
      results.totalImages.errors += imageResult.errors.length;

      // コメント色付け処理
      const colorResult = applyCommentColorsForParagraphs(paragraphs);
      docResult.colors = colorResult;
      results.totalColors.gray += colorResult.gray;
      results.totalColors.purple += colorResult.purple;
      results.totalColors.skipped += colorResult.skipped;

      console.log(`  画像: 挿入${imageResult.inserted} 更新${imageResult.updated} NoImage${imageResult.noImage} スキップ${imageResult.skipped}`);
      console.log(`  色付け: グレー${colorResult.gray} 紫${colorResult.purple} 黒字${colorResult.skipped}`);

    } catch (e) {
      docResult.error = e.message;
      console.log(`  エラー: ${e.message}`);
    }

    results.documents.push(docResult);
  }

  const elapsed = (new Date() - startTime) / 1000;
  console.log(`\n処理完了（${elapsed.toFixed(1)}秒）`);
  console.log(`画像合計: 挿入${results.totalImages.inserted} 更新${results.totalImages.updated} NoImage${results.totalImages.noImage} スキップ${results.totalImages.skipped} エラー${results.totalImages.errors}`);
  console.log(`色付け合計: グレー${results.totalColors.gray} 紫${results.totalColors.purple} 黒字${results.totalColors.skipped}`);

  // スプレッドシートにログ保存
  saveLogToSpreadsheet(startTime, elapsed, results);

  return results;
}

/**
 * 1つのドキュメントの画像処理
 */
function processImagesForDocument(paragraphs, service, fullMode) {
  // 必要なフォルダを特定してメタデータを事前取得
  const requiredFolders = scanRequiredFolders(paragraphs, FACE_PATTERN);
  const metadataCache = prefetchMetadata(service, requiredFolders);

  // 処理対象を収集
  const { targets, skipped, unregisteredChars } = collectTargetParagraphs(paragraphs, FACE_PATTERN, metadataCache);

  if (targets.length === 0) {
    return { inserted: 0, updated: 0, noImage: 0, skipped: skipped.length, errors: [], unregisteredChars };
  }

  // 処理実行
  const limit = fullMode ? null : BATCH_SIZE;
  const { insertedCount, updatedCount, noImageCount, errors } = processImageInsertions(targets, service, limit, metadataCache);

  return {
    inserted: insertedCount,
    updated: updatedCount,
    noImage: noImageCount,
    skipped: skipped.length,
    errors: errors,
    unregisteredChars: unregisteredChars
  };
}

/**
 * 段落配列に対してコメント色付けを適用
 */
function applyCommentColorsForParagraphs(paragraphs) {
  let grayCount = 0;
  let purpleCount = 0;
  let skippedCount = 0;

  for (const para of paragraphs) {
    const text = para.getText();

    if (!text.startsWith('//')) {
      continue;
    }

    // スキップキーワードチェック（黒字に設定）
    const betweenText = getTextBetweenSlashAndColon(text);
    if (betweenText !== null) {
      const normalizedBetween = normalizeText(betweenText);
      const shouldSkip = COMMENT_SKIP_KEYWORDS.some(keyword => normalizedBetween.includes(keyword));
      if (shouldSkip) {
        // 黒字に設定
        const textElement = para.editAsText();
        textElement.setForegroundColor(null);  // nullで黒字に戻る
        skippedCount++;
        continue;
      }
    }

    // 色を決定
    let color;
    const hasColon = text.includes(':') || text.includes('：');
    const endsWithKagi = text.endsWith('》');

    if (hasColon && !endsWithKagi) {
      color = COMMENT_COLOR_PURPLE;
      purpleCount++;
    } else {
      color = COMMENT_COLOR_GRAY;
      grayCount++;
    }

    const textElement = para.editAsText();
    textElement.setForegroundColor(color);
  }

  return { gray: grayCount, purple: purpleCount, skipped: skippedCount };
}

// ===== ログ保存 =====

/**
 * 処理結果をスプレッドシートに保存（ドキュメントごとに1行）
 */
function saveLogToSpreadsheet(startTime, elapsed, results) {
  const errors = [];

  // スプレッドシートにログ保存
  if (LOG_SPREADSHEET_ID) {
    try {
      const ss = SpreadsheetApp.openById(LOG_SPREADSHEET_ID);
      let sheet = ss.getSheetByName('実行ログ');

      // シートがなければ作成
      if (!sheet) {
        sheet = ss.insertSheet('実行ログ');
        // ヘッダー行を追加
        sheet.appendRow([
          '実行日時',
          '処理時間(秒)',
          'ドキュメント',
          '画像挿入',
          '画像更新',
          'NoImage',
          '画像スキップ',
          '画像エラー',
          '色グレー',
          '色紫',
          '色黒字',
          'エラー'
        ]);
        sheet.getRange(1, 1, 1, 12).setFontWeight('bold');
      }

      const timestamp = Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

      // ドキュメントごとに1行追加
      for (const doc of results.documents) {
        if (doc.error) {
          // エラーの場合
          sheet.appendRow([
            timestamp,
            elapsed.toFixed(1),
            doc.name,
            '', '', '', '', '',
            '', '', '',
            doc.error
          ]);
          errors.push(`${doc.name}: ${doc.error}`);
        } else {
          // 正常の場合
          sheet.appendRow([
            timestamp,
            elapsed.toFixed(1),
            doc.name,
            doc.images.inserted,
            doc.images.updated,
            doc.images.noImage,
            doc.images.skipped,
            doc.images.errors.length,
            doc.colors.gray,
            doc.colors.purple,
            doc.colors.skipped,
            doc.images.errors.length > 0 ? doc.images.errors.join(', ') : ''
          ]);

          // 画像エラーがあれば記録
          if (doc.images.errors.length > 0) {
            errors.push(`${doc.name}: ${doc.images.errors.join(', ')}`);
          }
        }
      }

      console.log('ログをスプレッドシートに保存しました');

    } catch (e) {
      console.log(`ログ保存エラー: ${e.message}`);
      errors.push(`ログ保存エラー: ${e.message}`);
    }
  }

  // エラーがあればメール送信
  if (errors.length > 0 && ERROR_NOTIFICATION_EMAIL) {
    sendErrorNotification(startTime, errors);
  }
}

/**
 * エラー通知メールを送信
 */
function sendErrorNotification(startTime, errors) {
  try {
    const timestamp = Utilities.formatDate(startTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    const subject = `[顔画像自動処理] エラー検知 (${errors.length}件)`;
    const body = `顔画像自動処理でエラーが発生しました。

実行日時: ${timestamp}
エラー件数: ${errors.length}件

【エラー詳細】
${errors.map((e, i) => `${i + 1}. ${e}`).join('\n')}

---
このメールは自動送信されています。
`;

    MailApp.sendEmail(ERROR_NOTIFICATION_EMAIL, subject, body);
    console.log(`エラー通知メールを送信しました: ${ERROR_NOTIFICATION_EMAIL}`);

  } catch (e) {
    console.log(`メール送信エラー: ${e.message}`);
  }
}

/**
 * メール送信テスト（権限承認用）
 * 初回実行時に権限承認ダイアログが出ます
 */
function testEmailNotification() {
  if (!ERROR_NOTIFICATION_EMAIL) {
    console.log('ERROR_NOTIFICATION_EMAIL が未設定です。メールアドレスを設定してください。');
    return;
  }

  const testErrors = ['テストエラー1: これはテストです', 'テストエラー2: 正常に受信できれば成功'];
  sendErrorNotification(new Date(), testErrors);
  console.log('テストメールを送信しました。受信を確認してください。');
}

// ===== トリガー管理 =====

/**
 * 自動処理トリガーを削除
 */
function removeAutoProcessingTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;

  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'autoProcessAllDocuments') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  }

  console.log(`トリガーを${removed}件削除しました`);
}

/**
 * トリガーの状態を確認
 */
function checkTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();

  console.log(`設定されているトリガー: ${triggers.length}件`);

  for (const trigger of triggers) {
    console.log(`  - ${trigger.getHandlerFunction()} (${trigger.getEventType()})`);
  }

  const autoTrigger = triggers.find(t => t.getHandlerFunction() === 'autoProcessAllDocuments');
  if (autoTrigger) {
    console.log('\n自動処理トリガー: 有効');
  } else {
    console.log('\n自動処理トリガー: 無効（setupAutoProcessing()で設定してください）');
  }
}

/**
 * 手動で今すぐ全ドキュメントをバッチ処理
 */
function manualProcessAllBatch() {
  const service = getDropboxService();

  if (!service.hasAccess()) {
    console.log('Dropboxに接続されていません。connectToDropbox() を実行してください。');
    return;
  }

  processAllDocumentsBatch();
}

// ===== OAuth2サービス =====
function getDropboxService() {
  return OAuth2.createService('dropbox')
    .setAuthorizationBaseUrl('https://www.dropbox.com/oauth2/authorize')
    .setTokenUrl('https://api.dropboxapi.com/oauth2/token')
    .setClientId(DROPBOX_CLIENT_ID)
    .setClientSecret(DROPBOX_CLIENT_SECRET)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setCache(CacheService.getUserCache())
    .setParam('token_access_type', 'offline');
}

function authCallback(request) {
  const service = getDropboxService();
  const authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('認証成功！このタブを閉じてください。');
  } else {
    return HtmlService.createHtmlOutput('認証失敗。もう一度試してください。');
  }
}

/**
 * Dropbox認証URLを取得（スタンドアロン版用）
 */
function connectToDropbox() {
  const service = getDropboxService();

  if (service.hasAccess()) {
    console.log('既にDropboxに接続されています。');
    return;
  }

  const authorizationUrl = service.getAuthorizationUrl();
  console.log('以下のURLにアクセスしてDropboxを認証してください:');
  console.log(authorizationUrl);
}

/**
 * Dropbox接続をリセット
 */
function disconnectDropbox() {
  const service = getDropboxService();
  service.reset();
  console.log('Dropboxとの接続を解除しました。');
}

/**
 * Dropbox接続テスト
 */
function testDropboxConnection() {
  const service = getDropboxService();

  if (!service.hasAccess()) {
    console.log('Dropboxに接続されていません。');
    return;
  }

  try {
    const url = 'https://api.dropboxapi.com/2/users/get_current_account';
    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${service.getAccessToken()}`,
        'Content-Type': 'application/json'
      },
      payload: 'null',
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      console.log(`接続成功！アカウント: ${data.name.display_name}`);
    } else {
      console.log(`接続失敗: ${response.getResponseCode()}`);
      console.log(`詳細: ${response.getContentText()}`);
    }
  } catch (e) {
    console.log(`エラー: ${e.message}`);
  }
}

// ===== ユーティリティ =====
function hasImageAtStart(paragraph) {
  const numChildren = paragraph.getNumChildren();
  if (numChildren === 0) return false;
  const firstChild = paragraph.getChild(0);
  return firstChild.getType() === DocumentApp.ElementType.INLINE_IMAGE;
}

function getImageAtStart(paragraph) {
  const numChildren = paragraph.getNumChildren();
  if (numChildren === 0) return null;
  const firstChild = paragraph.getChild(0);
  if (firstChild.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
    return firstChild.asInlineImage();
  }
  return null;
}

function parseAltDescription(altDesc) {
  if (!altDesc) return { fileName: null, hash: null };
  const parts = altDesc.split(':');
  return {
    fileName: parts[0] || null,
    hash: parts[1] || null
  };
}

function createAltDescription(fileName, hash) {
  return hash ? `${fileName}:${hash}` : fileName;
}

function normalizeText(text) {
  let normalized = text.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  normalized = normalized.toLowerCase();
  return normalized;
}

function getTextBetweenSlashAndColon(text) {
  const afterSlash = text.substring(2);
  const colonIndex1 = afterSlash.indexOf(':');
  const colonIndex2 = afterSlash.indexOf('：');

  let colonIndex = -1;
  if (colonIndex1 === -1) {
    colonIndex = colonIndex2;
  } else if (colonIndex2 === -1) {
    colonIndex = colonIndex1;
  } else {
    colonIndex = Math.min(colonIndex1, colonIndex2);
  }

  if (colonIndex === -1) {
    return null;
  }

  return afterSlash.substring(0, colonIndex);
}

// ===== 画像処理関数 =====
function scanRequiredFolders(paragraphs, pattern) {
  const folders = new Set();

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    const text = para.getText();
    const match = text.match(pattern);

    if (match) {
      const charName = match[1].trim().replace(/[\s　]+/g, '');
      const baseEnglishName = CHARACTER_MAP[charName];
      if (baseEnglishName) {
        const folderPath = `${DROPBOX_FACES_PATH}/Face_${baseEnglishName}`;
        folders.add(folderPath);
      }
    }
  }

  return Array.from(folders);
}

function prefetchMetadata(service, folders) {
  const metadataCache = {};

  for (const folderPath of folders) {
    try {
      metadataCache[folderPath] = getDropboxFolderMetadata(service, folderPath);
    } catch (e) {
      metadataCache[folderPath] = {};
    }
  }

  return metadataCache;
}

function collectTargetParagraphs(paragraphs, pattern, metadataCache) {
  const targets = [];
  const skipped = [];
  const unregisteredChars = new Set();

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    const text = para.getText();
    const match = text.match(pattern);

    if (match) {
      const charName = match[1].trim().replace(/[\s　]+/g, '');
      const suffix = (match[3] || '').trim().replace(/[\s　]+/g, '');
      const effectiveCharName = charName + suffix;

      const baseEnglishName = CHARACTER_MAP[charName];
      const englishName = CHARACTER_MAP[effectiveCharName];

      if (!baseEnglishName) {
        unregisteredChars.add(charName);
        continue;
      }
      if (!englishName) {
        unregisteredChars.add(effectiveCharName);
        continue;
      }

      const folderPath = `${DROPBOX_FACES_PATH}/Face_${baseEnglishName}`;

      const number3 = match[2].padStart(3, '0');
      const number4 = match[2].padStart(4, '0');
      const candidates = [
        { fileName: `Face_${englishName}_${number3}.png`, number: number3 },
        { fileName: `face_${englishName}_${number3}.png`, number: number3 },
        { fileName: `Face_${englishName}_${number4}.png`, number: number4 },
        { fileName: `face_${englishName}_${number4}.png`, number: number4 },
      ];

      const folderMetadata = metadataCache[folderPath] || {};
      let expectedFileName, number;
      const found = candidates.find(c => folderMetadata[c.fileName]);
      if (found) {
        expectedFileName = found.fileName;
        number = found.number;
      } else {
        expectedFileName = candidates[0].fileName;
        number = number3;
      }

      const existingImage = getImageAtStart(para);

      if (existingImage) {
        const altDesc = existingImage.getAltDescription() || '';

        if (altDesc.startsWith('noimage:')) {
          const currentHash = metadataCache[folderPath] ? metadataCache[folderPath][expectedFileName] : null;

          if (currentHash) {
            targets.push({
              paragraph: para,
              charName: effectiveCharName,
              baseEnglishName: baseEnglishName,
              englishName: englishName,
              number: number,
              fileName: expectedFileName,
              existingImage: existingImage,
              action: 'update_noimage'
            });
          } else {
            skipped.push(para);
          }
          continue;
        }

        const { fileName: existingFileName, hash: existingHash } = parseAltDescription(altDesc);

        if (existingFileName === expectedFileName) {
          const currentHash = metadataCache[folderPath] ? metadataCache[folderPath][expectedFileName] : null;

          if (!currentHash || existingHash === currentHash) {
            skipped.push(para);
            continue;
          }

          targets.push({
            paragraph: para,
            charName: effectiveCharName,
            baseEnglishName: baseEnglishName,
            englishName: englishName,
            number: number,
            fileName: expectedFileName,
            existingImage: existingImage,
            action: 'update_hash'
          });
        } else {
          targets.push({
            paragraph: para,
            charName: effectiveCharName,
            baseEnglishName: baseEnglishName,
            englishName: englishName,
            number: number,
            fileName: expectedFileName,
            existingImage: existingImage,
            action: 'update_name'
          });
        }
      } else {
        targets.push({
          paragraph: para,
          charName: effectiveCharName,
          baseEnglishName: baseEnglishName,
          englishName: englishName,
          number: number,
          fileName: expectedFileName,
          existingImage: null,
          action: 'insert'
        });
      }
    }
  }

  return { targets, skipped, unregisteredChars };
}

function processImageInsertions(targets, service, limit, metadataCache) {
  const toProcess = limit ? targets.slice(0, limit) : targets;
  let insertedCount = 0;
  let updatedCount = 0;
  let noImageCount = 0;
  const errors = [];

  if (!metadataCache) {
    metadataCache = {};
  }

  for (const target of toProcess) {
    const folderPath = `${DROPBOX_FACES_PATH}/Face_${target.baseEnglishName}`;
    const fileName = target.fileName;
    const filePath = `${folderPath}/${fileName}`;

    try {
      const { image, hash } = getImageWithHashFromDropbox(service, filePath, metadataCache);

      let imageToInsert = image;
      let altDesc;
      let isNoImage = false;

      if (!image) {
        imageToInsert = getNoImageBlob();
        if (!imageToInsert) {
          errors.push(`画像なし＆noimage取得失敗: ${fileName}`);
          continue;
        }
        isNoImage = true;
        altDesc = `noimage:${fileName}`;
      } else {
        altDesc = createAltDescription(fileName, hash);
      }

      if (target.existingImage) {
        target.existingImage.removeFromParent();
      }

      const insertedImage = target.paragraph.insertInlineImage(0, imageToInsert);

      if (isNoImage) {
        insertedImage.setWidth(NOIMAGE_SIZE);
        insertedImage.setHeight(NOIMAGE_SIZE);
      } else {
        const width = insertedImage.getWidth();
        const height = insertedImage.getHeight();
        insertedImage.setWidth(width / 3);
        insertedImage.setHeight(height / 3);
      }

      insertedImage.setAltDescription(altDesc);

      if (isNoImage) {
        noImageCount++;
      } else if (target.action === 'insert') {
        insertedCount++;
      } else {
        updatedCount++;
      }
    } catch (e) {
      errors.push(`エラー: ${fileName} - ${e.message}`);
    }
  }

  return { insertedCount, updatedCount, noImageCount, errors, processedCount: toProcess.length };
}

// ===== Dropbox API =====
function getDropboxFolderMetadata(service, folderPath) {
  const metadata = {};
  let cursor = null;
  let hasMore = true;

  while (hasMore) {
    let url, payload;

    if (cursor) {
      url = 'https://api.dropboxapi.com/2/files/list_folder/continue';
      payload = JSON.stringify({ cursor: cursor });
    } else {
      url = 'https://api.dropboxapi.com/2/files/list_folder';
      payload = JSON.stringify({
        path: folderPath,
        recursive: false,
        include_media_info: false,
        include_deleted: false
      });
    }

    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${service.getAccessToken()}`,
        'Content-Type': 'application/json'
      },
      payload: payload,
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();

    if (code === 200) {
      const data = JSON.parse(response.getContentText());

      for (const entry of data.entries) {
        if (entry['.tag'] === 'file') {
          const fileName = entry.name;
          metadata[fileName] = entry.content_hash;
        }
      }

      hasMore = data.has_more;
      cursor = data.cursor;
    } else if (code === 409) {
      return {};
    } else {
      throw new Error(`Dropbox metadata error: ${code}`);
    }
  }

  return metadata;
}

function getImageWithHashFromDropbox(service, filePath, metadataCache) {
  const fileName = filePath.split('/').pop();
  const folderPath = filePath.substring(0, filePath.lastIndexOf('/'));

  let hash = null;
  if (metadataCache[folderPath]) {
    hash = metadataCache[folderPath][fileName];
  } else {
    metadataCache[folderPath] = getDropboxFolderMetadata(service, folderPath);
    hash = metadataCache[folderPath][fileName];
  }

  const image = getImageFromDropbox(service, filePath);
  return { image, hash };
}

let noImageCache = null;
function getNoImageBlob() {
  if (noImageCache) return noImageCache;

  try {
    const response = UrlFetchApp.fetch(NOIMAGE_URL, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      noImageCache = response.getBlob();
      return noImageCache;
    }
  } catch (e) {
    // 取得失敗
  }
  return null;
}

function getImageFromDropbox(service, filePath) {
  const url = 'https://content.dropboxapi.com/2/files/download';

  const apiArg = JSON.stringify({ path: filePath });
  const encodedApiArg = apiArg.replace(/[\u007f-\uffff]/g, function(c) {
    return '\\u' + ('0000' + c.charCodeAt(0).toString(16)).slice(-4);
  });

  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${service.getAccessToken()}`,
      'Dropbox-API-Arg': encodedApiArg,
      'Content-Type': 'application/octet-stream'
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();

  if (code === 200) {
    return response.getBlob();
  } else if (code === 409) {
    return null;
  } else {
    const errorBody = response.getContentText();
    throw new Error(`Dropbox API error: ${code} - ${errorBody}`);
  }
}
