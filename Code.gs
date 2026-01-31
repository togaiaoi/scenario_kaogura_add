// ===== 設定 =====
// Dropbox App Consoleで取得した値を入力
const DROPBOX_CLIENT_ID = 'ここにApp keyを貼り付け';
const DROPBOX_CLIENT_SECRET = 'ここにApp secretを貼り付け';

// Dropbox上の画像フォルダパス
const DROPBOX_FACES_PATH = '/少年期の終り_画像共有/img/faces';

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

// 処理対象の段落を収集（パターンマッチ＆画像なし）
function collectTargetParagraphs(paragraphs, pattern, skipProcessed) {
  const targets = [];
  const unregisteredChars = new Set();

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    const text = para.getText();
    const match = text.match(pattern);

    if (match) {
      const charName = match[1];
      const number = match[2].padStart(3, '0');
      const englishName = CHARACTER_MAP[charName];

      if (!englishName) {
        unregisteredChars.add(charName);
        continue;
      }

      // 重複チェック
      if (skipProcessed && hasImageAtStart(para)) {
        continue; // 処理済みなのでスキップ
      }

      targets.push({
        paragraph: para,
        charName: charName,
        englishName: englishName,
        number: number
      });
    }
  }

  return { targets, unregisteredChars };
}

// 結果をHTMLダイアログで表示
function showResultDialog(result) {
  const { inserted, skippedProcessed, errors, remaining, unregisteredChars, mode } = result;

  let html = `
    <style>
      body { font-family: sans-serif; font-size: 14px; margin: 0; padding: 16px; }
      .section { margin-bottom: 16px; }
      .section-title { font-weight: bold; margin-bottom: 8px; color: #333; }
      .stat { margin: 4px 0; }
      .stat-ok { color: #2e7d32; }
      .stat-skip { color: #f57c00; }
      .stat-error { color: #c62828; }
      .stat-remain { color: #1565c0; }
      .list { max-height: 150px; overflow-y: auto; background: #f5f5f5; padding: 8px; border-radius: 4px; font-size: 12px; }
      .list-item { margin: 2px 0; }
    </style>
    <div class="section">
      <div class="section-title">実行結果（${mode}）</div>
      <div class="stat stat-ok">✓ 挿入: ${inserted}件</div>
      <div class="stat stat-skip">⏭ スキップ（処理済み）: ${skippedProcessed}件</div>
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
    html += `<div style="color:#666; font-size:12px;">※「次の10件を挿入」でさらに処理できます</div>`;
  }

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(350);
  DocumentApp.getUi().showModalDialog(htmlOutput, '顔画像挿入 - 結果');
}

// ===== メイン関数 =====

// バッチサイズ設定
const BATCH_SIZE = 20;

// 画像挿入の共通処理
function processImageInsertions(targets, service, limit) {
  const toProcess = limit ? targets.slice(0, limit) : targets;
  let insertedCount = 0;
  const errors = [];

  for (const target of toProcess) {
    const fileName = `Face_${target.englishName}_${target.number}.png`;
    const folderPath = `${DROPBOX_FACES_PATH}/Face_${target.englishName}`;
    const filePath = `${folderPath}/${fileName}`;

    try {
      const image = getImageFromDropbox(service, filePath);
      if (image) {
        const insertedImage = target.paragraph.insertInlineImage(0, image);
        // 画像サイズを1/3に縮小
        const width = insertedImage.getWidth();
        const height = insertedImage.getHeight();
        insertedImage.setWidth(width / 3);
        insertedImage.setHeight(height / 3);
        insertedCount++;
      } else {
        errors.push(`画像なし: ${fileName}`);
      }
    } catch (e) {
      errors.push(`エラー: ${fileName} - ${e.message}`);
    }
  }

  return { insertedCount, errors, processedCount: toProcess.length };
}

// 次の10件を挿入（バッチ処理）
function insertNextBatch() {
  const service = getDropboxService();

  if (!service.hasAccess()) {
    DocumentApp.getUi().alert('Dropboxに接続されていません。\n「顔画像挿入」→「Dropboxに接続」を実行してください。');
    return;
  }

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();
  const pattern = /【(.+?)】(\d+)/;

  // 処理済みをスキップして未処理を収集
  const { targets, unregisteredChars } = collectTargetParagraphs(paragraphs, pattern, true);

  if (targets.length === 0) {
    DocumentApp.getUi().alert('処理対象がありません。\n（全て処理済み、またはパターンに一致する行がありません）');
    return;
  }

  // バッチサイズ分だけ処理
  const { insertedCount, errors } = processImageInsertions(targets, service, BATCH_SIZE);
  const remaining = Math.max(0, targets.length - BATCH_SIZE);

  // 処理済み件数をカウント（全段落から再計算）
  const { targets: remainingTargets } = collectTargetParagraphs(paragraphs, pattern, true);
  const skippedProcessed = paragraphs.filter(p => hasImageAtStart(p)).length;

  showResultDialog({
    inserted: insertedCount,
    skippedProcessed: skippedProcessed - insertedCount, // 今回挿入した分を除く
    errors: errors,
    remaining: remaining,
    unregisteredChars: unregisteredChars,
    mode: `バッチ ${BATCH_SIZE}件`
  });
}

// 全件挿入（重複スキップ付き）
function insertAllImages() {
  const service = getDropboxService();

  if (!service.hasAccess()) {
    DocumentApp.getUi().alert('Dropboxに接続されていません。\n「顔画像挿入」→「Dropboxに接続」を実行してください。');
    return;
  }

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();
  const pattern = /【(.+?)】(\d+)/;

  // 処理済みをスキップして未処理を収集
  const { targets, unregisteredChars } = collectTargetParagraphs(paragraphs, pattern, true);

  if (targets.length === 0) {
    DocumentApp.getUi().alert('処理対象がありません。\n（全て処理済み、またはパターンに一致する行がありません）');
    return;
  }

  // 件数が多い場合は警告
  if (targets.length > 50) {
    const ui = DocumentApp.getUi();
    const response = ui.alert(
      '確認',
      `${targets.length}件の画像を挿入します。\n件数が多いため、処理に時間がかかる可能性があります。\n\n続行しますか？`,
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      return;
    }
  }

  // 全件処理
  const { insertedCount, errors } = processImageInsertions(targets, service, null);

  // 処理済み件数をカウント
  const skippedProcessed = paragraphs.filter(p => hasImageAtStart(p)).length - insertedCount;

  showResultDialog({
    inserted: insertedCount,
    skippedProcessed: skippedProcessed,
    errors: errors,
    remaining: 0,
    unregisteredChars: unregisteredChars,
    mode: '全件'
  });
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
    .addToUi();
}
