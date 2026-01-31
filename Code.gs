// ===== 設定 =====
// Dropbox App Consoleで取得した値を入力
const DROPBOX_CLIENT_ID = 'ここにApp keyを貼り付け';
const DROPBOX_CLIENT_SECRET = 'ここにApp secretを貼り付け';

// Dropbox上の画像フォルダパス
const DROPBOX_FACES_PATH = '/少年期の終り_画像共有/img/faces';

// No Image画像のDropbox共有リンク（画像が見つからない時に使用）
const NOIMAGE_URL = 'https://www.dropbox.com/scl/fi/ny6cm3boatvpe0s5axg5h/noimage.jpg?rlkey=vcg3cjs1ytfgarx059m191l7u&dl=1';
// No Image画像の表示サイズ（通常の顔画像と同じサイズに設定）
const NOIMAGE_SIZE = 48; 

// キャラクター名→英語名の対応表（追加可能）
const CHARACTER_MAP = {
  'ジョバンニ': 'Giovanni',
  'カムパネルラ': 'Campanella',
  'ドク': 'Doc',
  'デイヴ': 'Dave',
  'ケイト': 'Kate',
  'マーク': 'Mark',
  'モーリィ': 'Mollie',
  'ラビ': 'Rabi',
  'Wi': 'Wi',
  'Wi無': 'Wi_nohood',  // 【Wi】34無 → Wi無として解釈
  'Ｗi無': 'Wi_nohood',
  'ユミ': 'Yumi',
  'ケイト黒': 'Kate_BH',
  'ラブ': 'Love',
  'ラブ面': 'Love_fullface',
  // ↓ 新しいキャラクターはここに追加
};

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
    .setParam('token_access_type', 'offline'); // Refresh Token取得のため必須
}

// OAuth認証コールバック
function authCallback(request) {
  const service = getDropboxService();
  const authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('認証成功！このタブを閉じて、ドキュメントに戻ってください。');
  } else {
    return HtmlService.createHtmlOutput('認証失敗。もう一度試してください。');
  }
}

// ===== ユーティリティ =====
// 段落の先頭に画像があるかチェック（重複防止用）
function hasImageAtStart(paragraph) {
  const numChildren = paragraph.getNumChildren();
  if (numChildren === 0) return false;
  const firstChild = paragraph.getChild(0);
  return firstChild.getType() === DocumentApp.ElementType.INLINE_IMAGE;
}

// 段落先頭の画像を取得
function getImageAtStart(paragraph) {
  const numChildren = paragraph.getNumChildren();
  if (numChildren === 0) return null;
  const firstChild = paragraph.getChild(0);
  if (firstChild.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
    return firstChild.asInlineImage();
  }
  return null;
}

// alt textからファイル名とハッシュを解析
function parseAltDescription(altDesc) {
  if (!altDesc) return { fileName: null, hash: null };
  const parts = altDesc.split(':');
  return {
    fileName: parts[0] || null,
    hash: parts[1] || null
  };
}

// ファイル名とハッシュからalt textを生成
function createAltDescription(fileName, hash) {
  return hash ? `${fileName}:${hash}` : fileName;
}

// 第1パス: 必要なフォルダを特定（メタデータ事前取得用）
function scanRequiredFolders(paragraphs, pattern) {
  const folders = new Set();

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    const text = para.getText();
    const match = text.match(pattern);

    if (match) {
      // 【Wi】34無 → charName="Wi", suffix="無"
      // フォルダは元キャラ(Wi)、ファイル名は複合(Wi無→Wi_nohood)
      const charName = match[1];
      const baseEnglishName = CHARACTER_MAP[charName];  // フォルダ用
      if (baseEnglishName) {
        const folderPath = `${DROPBOX_FACES_PATH}/Face_${baseEnglishName}`;
        folders.add(folderPath);
      }
    }
  }

  return Array.from(folders);
}

// メタデータを事前に一括取得
function prefetchMetadata(service, folders) {
  const metadataCache = {};

  for (const folderPath of folders) {
    try {
      metadataCache[folderPath] = getDropboxFolderMetadata(service, folderPath);
    } catch (e) {
      // エラーが出ても続行（該当フォルダは空扱い）
      metadataCache[folderPath] = {};
    }
  }

  return metadataCache;
}

// 処理対象の段落を収集（パターンマッチ＆不一致検出）
// metadataCache: Dropboxのハッシュ情報（オプション、ハッシュ比較する場合に渡す）
function collectTargetParagraphs(paragraphs, pattern, metadataCache) {
  const targets = [];
  const skipped = [];  // スキップした段落（完全一致）
  const unregisteredChars = new Set();

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    const text = para.getText();
    const match = text.match(pattern);

    if (match) {
      // 【Wi】34無 → charName="Wi", suffix="無", effectiveCharName="Wi無"
      // フォルダは元キャラ(Wi→Wi)、ファイル名は複合(Wi無→Wi_nohood)
      const charName = match[1];
      const suffix = match[3] || '';
      const effectiveCharName = charName + suffix;

      // ベース名（フォルダ用）と複合名（ファイル名用）
      const baseEnglishName = CHARACTER_MAP[charName];  // Wi → "Wi"
      const englishName = CHARACTER_MAP[effectiveCharName];  // Wi無 → "Wi_nohood"

      if (!baseEnglishName) {
        unregisteredChars.add(charName);
        continue;
      }
      if (!englishName) {
        unregisteredChars.add(effectiveCharName);
        continue;
      }

      const folderPath = `${DROPBOX_FACES_PATH}/Face_${baseEnglishName}`;

      // 3桁/4桁 × Face/face の組み合わせを試す
      const number3 = match[2].padStart(3, '0');
      const number4 = match[2].padStart(4, '0');
      const candidates = [
        { fileName: `Face_${englishName}_${number3}.png`, number: number3 },
        { fileName: `face_${englishName}_${number3}.png`, number: number3 },
        { fileName: `Face_${englishName}_${number4}.png`, number: number4 },
        { fileName: `face_${englishName}_${number4}.png`, number: number4 },
      ];

      // メタデータキャッシュで存在確認（上から優先）
      const folderMetadata = metadataCache[folderPath] || {};
      let expectedFileName, number;
      const found = candidates.find(c => folderMetadata[c.fileName]);
      if (found) {
        expectedFileName = found.fileName;
        number = found.number;
      } else {
        // どれも見つからない場合はFace_3桁をデフォルトに
        expectedFileName = candidates[0].fileName;
        number = number3;
      }

      const existingImage = getImageAtStart(para);

      if (existingImage) {
        // 既存画像がある場合: alt textをチェック
        const altDesc = existingImage.getAltDescription() || '';

        // noimage画像かどうかチェック
        if (altDesc.startsWith('noimage:')) {
          // noimageの場合: 正しい画像が存在するか確認
          const currentHash = metadataCache[folderPath] ? metadataCache[folderPath][expectedFileName] : null;

          if (currentHash) {
            // 画像が見つかった → 更新対象
            targets.push({
              paragraph: para,
              charName: effectiveCharName,
              baseEnglishName: baseEnglishName,
              englishName: englishName,
              number: number,
              fileName: expectedFileName,  // 実際に見つかったファイル名
              existingImage: existingImage,
              action: 'update_noimage'  // noimage→正しい画像に更新
            });
          } else {
            // まだ画像がない → スキップ（既にnoimageが入っている）
            skipped.push(para);
          }
          continue;
        }

        const { fileName: existingFileName, hash: existingHash } = parseAltDescription(altDesc);

        // ファイル名が一致するかチェック
        if (existingFileName === expectedFileName) {
          // ファイル名一致 → ハッシュもチェック
          const currentHash = metadataCache[folderPath] ? metadataCache[folderPath][expectedFileName] : null;

          if (!currentHash || existingHash === currentHash) {
            // ハッシュ取得できない or ハッシュ一致 → スキップ
            skipped.push(para);
            continue;
          }

          // ハッシュ不一致 → 画像更新
          targets.push({
            paragraph: para,
            charName: effectiveCharName,
            baseEnglishName: baseEnglishName,
            englishName: englishName,
            number: number,
            fileName: expectedFileName,  // 実際に見つかったファイル名
            existingImage: existingImage,
            action: 'update_hash'  // 画像更新
          });
        } else {
          // ファイル名不一致 → 更新対象
          targets.push({
            paragraph: para,
            charName: effectiveCharName,
            baseEnglishName: baseEnglishName,
            englishName: englishName,
            number: number,
            fileName: expectedFileName,  // 実際に見つかったファイル名
            existingImage: existingImage,
            action: 'update_name'  // 番号変更
          });
        }
      } else {
        // 画像なし → 新規挿入
        targets.push({
          paragraph: para,
          charName: effectiveCharName,
          baseEnglishName: baseEnglishName,
          englishName: englishName,
          number: number,
          fileName: expectedFileName,  // 実際に見つかったファイル名
          existingImage: null,
          action: 'insert'
        });
      }
    }
  }

  return { targets, skipped, unregisteredChars };
}

// 結果をHTMLダイアログで表示
function showResultDialog(result) {
  const { inserted, updated, noImage, skippedProcessed, errors, remaining, unregisteredChars, mode } = result;

  let html = `
    <style>
      body { font-family: sans-serif; font-size: 14px; margin: 0; padding: 16px; }
      .section { margin-bottom: 16px; }
      .section-title { font-weight: bold; margin-bottom: 8px; color: #333; }
      .stat { margin: 4px 0; }
      .stat-ok { color: #2e7d32; }
      .stat-update { color: #7b1fa2; }
      .stat-noimage { color: #9e9e9e; }
      .stat-skip { color: #f57c00; }
      .stat-error { color: #c62828; }
      .stat-remain { color: #1565c0; }
      .list { max-height: 150px; overflow-y: auto; background: #f5f5f5; padding: 8px; border-radius: 4px; font-size: 12px; }
      .list-item { margin: 2px 0; }
    </style>
    <div class="section">
      <div class="section-title">実行結果（${mode}）</div>
      <div class="stat stat-ok">✓ 新規挿入: ${inserted}件</div>
      <div class="stat stat-update">🔄 更新: ${updated}件</div>
      <div class="stat stat-noimage">🖼 NoImage挿入: ${noImage || 0}件</div>
      <div class="stat stat-skip">⏭ スキップ（一致）: ${skippedProcessed}件</div>
      <div class="stat stat-error">⚠ スキップ（エラー）: ${errors.length}件</div>
      <div class="stat stat-remain">📋 残り未処理: ${remaining}件</div>
    </div>
  `;

  if (errors.length > 0) {
    html += `
      <div class="section">
        <div class="section-title">エラー詳細</div>
        <div class="list">
          ${errors.map(e => `<div class="list-item">• ${e}</div>`).join('')}
        </div>
      </div>
    `;
  }

  if (unregisteredChars.size > 0) {
    const charList = Array.from(unregisteredChars);
    html += `
      <div class="section">
        <div class="section-title">未登録キャラクター（CHARACTER_MAPに追加してください）</div>
        <div class="list">
          ${charList.map(c => `<div class="list-item">• ${c}</div>`).join('')}
        </div>
      </div>
    `;
  }

  if (remaining > 0) {
    html += `<div style="color:#666; font-size:12px;">※「次の20件を挿入」でさらに処理できます</div>`;
  }

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(350);
  DocumentApp.getUi().showModalDialog(htmlOutput, '顔画像挿入 - 結果');
}

// ===== メイン関数 =====

// バッチサイズ設定
const BATCH_SIZE = 20;

// パターン定義: 【キャラクター名】番号 または 【キャラクター名】番号X（Xは1文字のサフィックス）
// 例: 【Wi】34無 → キャラクター名="Wi無", 番号="34"
const FACE_PATTERN = /【(.+?)】(\d+)(.)?/;

// 画像挿入の共通処理（挿入/更新対応、alt text保存）
function processImageInsertions(targets, service, limit, metadataCache) {
  const toProcess = limit ? targets.slice(0, limit) : targets;
  let insertedCount = 0;
  let updatedCount = 0;
  let noImageCount = 0;
  const errors = [];

  // メタデータキャッシュがなければ作成
  if (!metadataCache) {
    metadataCache = {};
  }

  for (const target of toProcess) {
    // フォルダは元キャラ（baseEnglishName）
    const folderPath = `${DROPBOX_FACES_PATH}/Face_${target.baseEnglishName}`;

    // collectTargetParagraphsで決定済みのファイル名を使用（Face/face、3桁/4桁対応済み）
    const fileName = target.fileName;
    const filePath = `${folderPath}/${fileName}`;

    try {
      // 画像とハッシュを取得
      const { image, hash } = getImageWithHashFromDropbox(service, filePath, metadataCache);

      let imageToInsert = image;
      let altDesc;
      let isNoImage = false;

      if (!image) {
        // 画像が見つからない場合はnoimage画像を使用
        imageToInsert = getNoImageBlob();
        if (!imageToInsert) {
          errors.push(`画像なし＆noimage取得失敗: ${fileName}`);
          continue;
        }
        isNoImage = true;
        altDesc = `noimage:${fileName}`;  // どのファイルがなかったか記録
      } else {
        altDesc = createAltDescription(fileName, hash);
      }

      // 既存画像があれば削除
      if (target.existingImage) {
        target.existingImage.removeFromParent();
      }

      // 新しい画像を挿入
      const insertedImage = target.paragraph.insertInlineImage(0, imageToInsert);

      // 画像サイズを調整
      if (isNoImage) {
        // noimage画像は固定サイズ
        insertedImage.setWidth(NOIMAGE_SIZE);
        insertedImage.setHeight(NOIMAGE_SIZE);
      } else {
        // 通常画像は1/3に縮小
        const width = insertedImage.getWidth();
        const height = insertedImage.getHeight();
        insertedImage.setWidth(width / 3);
        insertedImage.setHeight(height / 3);
      }

      // alt textを保存
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

// 次の20件を挿入（バッチ処理）
function insertNextBatch() {
  const service = getDropboxService();

  if (!service.hasAccess()) {
    DocumentApp.getUi().alert('Dropboxに接続されていません。\n「顔画像挿入」→「Dropboxに接続」を実行してください。');
    return;
  }

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();

  // 必要なフォルダを特定してメタデータを事前取得
  const requiredFolders = scanRequiredFolders(paragraphs, FACE_PATTERN);
  const metadataCache = prefetchMetadata(service, requiredFolders);

  // 処理対象を収集（ファイル名不一致＆ハッシュ不一致も含む）
  const { targets, skipped, unregisteredChars } = collectTargetParagraphs(paragraphs, FACE_PATTERN, metadataCache);

  if (targets.length === 0) {
    DocumentApp.getUi().alert('処理対象がありません。\n（全て処理済み、またはパターンに一致する行がありません）');
    return;
  }

  // バッチサイズ分だけ処理
  const { insertedCount, updatedCount, noImageCount, errors } = processImageInsertions(targets, service, BATCH_SIZE, metadataCache);
  const remaining = Math.max(0, targets.length - BATCH_SIZE);

  showResultDialog({
    inserted: insertedCount,
    updated: updatedCount,
    noImage: noImageCount,
    skippedProcessed: skipped.length,
    errors: errors,
    remaining: remaining,
    unregisteredChars: unregisteredChars,
    mode: `バッチ ${BATCH_SIZE}件`
  });
}

// 全件挿入（不一致検出＆更新付き）
function insertAllImages() {
  const service = getDropboxService();

  if (!service.hasAccess()) {
    DocumentApp.getUi().alert('Dropboxに接続されていません。\n「顔画像挿入」→「Dropboxに接続」を実行してください。');
    return;
  }

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();

  // 必要なフォルダを特定してメタデータを事前取得
  const requiredFolders = scanRequiredFolders(paragraphs, FACE_PATTERN);
  const metadataCache = prefetchMetadata(service, requiredFolders);

  // 処理対象を収集（ファイル名不一致＆ハッシュ不一致も含む）
  const { targets, skipped, unregisteredChars } = collectTargetParagraphs(paragraphs, FACE_PATTERN, metadataCache);

  if (targets.length === 0) {
    DocumentApp.getUi().alert('処理対象がありません。\n（全て処理済み、またはパターンに一致する行がありません）');
    return;
  }

  // 件数が多い場合は警告
  if (targets.length > 50) {
    const ui = DocumentApp.getUi();
    const response = ui.alert(
      '確認',
      `${targets.length}件の画像を処理します。\n件数が多いため、処理に時間がかかる可能性があります。\n\n続行しますか？`,
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      return;
    }
  }

  // 全件処理
  const { insertedCount, updatedCount, noImageCount, errors } = processImageInsertions(targets, service, null, metadataCache);

  showResultDialog({
    inserted: insertedCount,
    updated: updatedCount,
    noImage: noImageCount,
    skippedProcessed: skipped.length,
    errors: errors,
    remaining: 0,
    unregisteredChars: unregisteredChars,
    mode: '全件'
  });
}

// Dropboxフォルダのメタデータを一括取得（content_hash含む、ページネーション対応）
function getDropboxFolderMetadata(service, folderPath) {
  const metadata = {};
  let cursor = null;
  let hasMore = true;

  while (hasMore) {
    let url, payload;

    if (cursor) {
      // 続きを取得
      url = 'https://api.dropboxapi.com/2/files/list_folder/continue';
      payload = JSON.stringify({ cursor: cursor });
    } else {
      // 最初のリクエスト
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
          // ファイル名をキーにしてcontent_hashを保存
          const fileName = entry.name;
          metadata[fileName] = entry.content_hash;
        }
      }

      hasMore = data.has_more;
      cursor = data.cursor;
    } else if (code === 409) {
      // フォルダが見つからない
      return {};
    } else {
      throw new Error(`Dropbox metadata error: ${code}`);
    }
  }

  return metadata;
}

// 画像とハッシュをDropboxから取得
function getImageWithHashFromDropbox(service, filePath, metadataCache) {
  const fileName = filePath.split('/').pop();
  const folderPath = filePath.substring(0, filePath.lastIndexOf('/'));

  // キャッシュからハッシュを取得（なければフォルダメタデータを取得）
  let hash = null;
  if (metadataCache[folderPath]) {
    hash = metadataCache[folderPath][fileName];
  } else {
    metadataCache[folderPath] = getDropboxFolderMetadata(service, folderPath);
    hash = metadataCache[folderPath][fileName];
  }

  // 画像をダウンロード
  const image = getImageFromDropbox(service, filePath);
  return { image, hash };
}

// No Image画像を取得（キャッシュ付き）
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

// Dropboxから画像を取得
function getImageFromDropbox(service, filePath) {
  const url = 'https://content.dropboxapi.com/2/files/download';

  // パスをASCII文字のみにエンコード
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
    // ファイルが見つからない
    return null;
  } else {
    // エラー詳細を取得
    const errorBody = response.getContentText();
    throw new Error(`Dropbox API error: ${code} - ${errorBody}`);
  }
}

// ===== Dropbox接続・認証 =====
function connectToDropbox() {
  const service = getDropboxService();

  if (service.hasAccess()) {
    DocumentApp.getUi().alert('既にDropboxに接続されています。');
    return;
  }

  const authorizationUrl = service.getAuthorizationUrl();
  const htmlOutput = HtmlService.createHtmlOutput(
    `<p>以下のリンクをクリックしてDropboxを認証してください：</p>
     <p><a href="${authorizationUrl}" target="_blank">Dropboxに接続</a></p>
     <p>認証後、このダイアログを閉じて「画像を挿入する」を実行してください。</p>`
  )
    .setWidth(400)
    .setHeight(150);

  DocumentApp.getUi().showModalDialog(htmlOutput, 'Dropbox認証');
}

// 接続解除
function disconnectDropbox() {
  const service = getDropboxService();
  service.reset();
  DocumentApp.getUi().alert('Dropboxとの接続を解除しました。');
}

// 接続テスト
function testDropboxConnection() {
  const service = getDropboxService();

  if (!service.hasAccess()) {
    DocumentApp.getUi().alert('Dropboxに接続されていません。\n「顔画像挿入」→「Dropboxに接続」を実行してください。');
    return;
  }

  try {
    const url = 'https://api.dropboxapi.com/2/users/get_current_account';
    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${service.getAccessToken()}`
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      DocumentApp.getUi().alert(`接続成功！\nアカウント: ${data.name.display_name}`);
    } else {
      DocumentApp.getUi().alert(`接続失敗: ${response.getResponseCode()}`);
    }
  } catch (e) {
    DocumentApp.getUi().alert(`エラー: ${e.message}`);
  }
}

// ===== デバッグ用 =====
// ドキュメント内の画像のalt textを確認
function debugCheckAltText() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();

  const results = [];
  let count = 0;

  for (const para of paragraphs) {
    const text = para.getText();
    const match = text.match(FACE_PATTERN);
    if (match && hasImageAtStart(para)) {
      const image = getImageAtStart(para);
      const altDesc = image.getAltDescription();
      const suffix = match[3] || '';
      results.push(`${match[1]}${match[2]}${suffix}: "${altDesc || '(なし)'}"`);
      count++;
      if (count >= 10) break;  // 最大10件
    }
  }

  if (results.length === 0) {
    DocumentApp.getUi().alert('画像付きの行が見つかりません。');
  } else {
    DocumentApp.getUi().alert(`画像のalt text（最大10件）:\n\n${results.join('\n')}`);
  }
}

// Dropboxのメタデータ（ハッシュ）を確認
function debugCheckDropboxHash() {
  const service = getDropboxService();

  if (!service.hasAccess()) {
    DocumentApp.getUi().alert('Dropboxに接続されていません。');
    return;
  }

  // 最初のキャラクターのフォルダをテスト
  const testFolder = `${DROPBOX_FACES_PATH}/Face_Giovanni`;

  try {
    const metadata = getDropboxFolderMetadata(service, testFolder);
    const allFiles = Object.keys(metadata);
    const files = allFiles.slice(0, 5);  // 最大5件表示

    if (allFiles.length === 0) {
      DocumentApp.getUi().alert(`フォルダ: ${testFolder}\n\nファイルが見つかりません。`);
    } else {
      const results = files.map(f => `${f}: ${metadata[f] ? metadata[f].substring(0, 16) + '...' : '(なし)'}`);
      DocumentApp.getUi().alert(`フォルダ: ${testFolder}\n\n総ファイル数: ${allFiles.length}件\n\nハッシュ（最大5件表示）:\n${results.join('\n')}`);
    }
  } catch (e) {
    DocumentApp.getUi().alert(`エラー: ${e.message}`);
  }
}

// カスタムメニューを追加
function onOpen() {
  DocumentApp.getUi()
    .createMenu('顔画像挿入')
    .addItem('Dropboxに接続', 'connectToDropbox')
    .addItem('接続をテスト', 'testDropboxConnection')
    .addSeparator()
    .addItem('次の20件を挿入', 'insertNextBatch')
    .addItem('全件挿入（重複スキップ）', 'insertAllImages')
    .addSeparator()
    .addItem('接続を解除', 'disconnectDropbox')
    .addSeparator()
    .addItem('[DEBUG] alt text確認', 'debugCheckAltText')
    .addItem('[DEBUG] Dropboxハッシュ確認', 'debugCheckDropboxHash')
    .addToUi();
}
