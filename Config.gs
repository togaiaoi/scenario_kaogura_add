// ===== 設定ファイル =====
// このファイルには設定値のみを記載します
// コードは CodeStandalone.gs にあります

// ===== Dropbox設定 =====
// Dropbox App Consoleで取得した値を入力
const DROPBOX_CLIENT_ID = 'ここにApp keyを貼り付け';
const DROPBOX_CLIENT_SECRET = 'ここにApp secretを貼り付け';

// Dropbox上の画像フォルダパス
const DROPBOX_FACES_PATH = '/少年期の終り_画像共有/img/faces';

// No Image画像のDropbox共有リンク（画像が見つからない時に使用）
const NOIMAGE_URL = 'https://www.dropbox.com/scl/fi/ny6cm3boatvpe0s5axg5h/noimage.jpg?rlkey=vcg3cjs1ytfgarx059m191l7u&dl=1';
// No Image画像の表示サイズ（通常の顔画像と同じサイズに設定）
const NOIMAGE_SIZE = 48;

// ===== キャラクター設定 =====
// キャラクター名→英語名の対応表
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

// ===== 処理対象ドキュメント =====
// ドキュメントURLの /d/ と /edit の間の文字列がドキュメントID
// 例: https://docs.google.com/document/d/XXXXXXXXXXXXXXX/edit → XXXXXXXXXXXXXXX
const TARGET_DOCUMENT_IDS = [
  // { id: 'ドキュメントID', name: 'ドキュメント名（ログ用）' }
  { id: 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx', name: 'シナリオ1' },
  { id: 'yyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyy', name: 'シナリオ2' },
  // ↓ 新しいドキュメントはここに追加
];

// ===== 実行設定 =====
// バッチサイズ（1回の定期実行で処理する画像数/ドキュメント）
const BATCH_SIZE = 20;

// 定期実行の間隔（分）
const AUTO_PROCESS_INTERVAL_MINUTES = 15;

// ===== ログ・通知設定 =====
// ログ保存用スプレッドシートID（空欄ならログ保存しない）
const LOG_SPREADSHEET_ID = '1I-Quc1r46_e6VFPlidIADoiJ7NspfJ5VHwxHe3kt0PQ';

// エラー通知メールの送信先（空欄ならメール送信しない）
const ERROR_NOTIFICATION_EMAIL = 'ono@wss.tokyo';  // 例: 'your-email@example.com'

// ===== コメント行の色付け設定 =====
// パターン定義: 【キャラクター名】番号 または 【キャラクター名】番号X
const FACE_PATTERN = /【(.+?)】(\d+)(.)?/;

// コメント行の色
const COMMENT_COLOR_GRAY = '#999999';    // 通常のコメント行
const COMMENT_COLOR_PURPLE = '#9a03ff';  // 特殊コメント行（コロン含む＆》で終わらない）

// 黒字のまま維持するキーワード（//と：の間にこれらが含まれる場合はスキップ）
// ※アルファベットは半角小文字で記載（全角/半角、大文字/小文字は自動で正規化される）
const COMMENT_SKIP_KEYWORDS = [
  'マップid',
  'チャプター',
  'se',
  'bgm',
  'bgs',
  '設定',
  'todo',
  'r-code',
  'アイテム',
  '通知',
  '入手',
  'メモ',
  '仕様',
  '実績',
  // ↓ 新しいキーワードはここに追加
];
