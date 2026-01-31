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

// ===== メイン関数 =====
function insertFaceImages() {
  const service = getDropboxService();

  if (!service.hasAccess()) {
    DocumentApp.getUi().alert('Dropboxに接続されていません。\n「顔画像挿入」→「Dropboxに接続」を実行してください。');
    return;
  }

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();

  // 【キャラクター名】番号 のパターン
  const pattern = /【(.+?)】(\d+)/;
  let insertedCount = 0;
  let skippedCount = 0;
  const errors = [];
  const unregisteredChars = new Set(); // 未登録キャラクターを重複なく記録

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    const text = para.getText();
    const match = text.match(pattern);

    if (match) {
      const charName = match[1];
      const number = match[2].padStart(3, '0');
      const englishName = CHARACTER_MAP[charName];

      if (englishName) {
        const fileName = `Face_${englishName}_${number}.png`;
        const folderPath = `${DROPBOX_FACES_PATH}/Face_${englishName}`;
        const filePath = `${folderPath}/${fileName}`;

        try {
          const image = getImageFromDropbox(service, filePath);
          if (image) {
            const insertedImage = para.insertInlineImage(0, image);
            // 画像サイズを1/3に縮小
            const width = insertedImage.getWidth();
            const height = insertedImage.getHeight();
            insertedImage.setWidth(width / 3);
            insertedImage.setHeight(height / 3);
            insertedCount++;
          } else {
            errors.push(`画像なし: ${fileName}`);
            skippedCount++;
          }
        } catch (e) {
          errors.push(`エラー: ${fileName} - ${e.message}`);
          skippedCount++;
        }
      } else {
        unregisteredChars.add(charName); // 未登録キャラを記録
        skippedCount++;
      }
    }
  }

  // 結果を表示
  let message = `完了！\n挿入: ${insertedCount}件\nスキップ: ${skippedCount}件`;

  // 未登録キャラクターを表示
  if (unregisteredChars.size > 0) {
    const charList = Array.from(unregisteredChars).join('、');
    message += `\n\n【未登録キャラクター】\n${charList}\n※ CHARACTER_MAP に追加してください`;
  }

  if (errors.length > 0) {
    message += `\n\nエラー詳細:\n${errors.slice(0, 10).join('\n')}`;
    if (errors.length > 10) {
      message += `\n...他 ${errors.length - 10}件`;
    }
  }
  DocumentApp.getUi().alert(message);
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
    .addItem('画像を挿入する', 'insertFaceImages')
    .addSeparator()
    .addItem('接続を解除', 'disconnectDropbox')
    .addToUi();
}
