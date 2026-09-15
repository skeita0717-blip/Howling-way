/****************************************************************
 * Howling way サイト用  スプレッドシート → JSON API
 * ------------------------------------------------------------
 * このコードを Google スプレッドシートの
 * 「拡張機能 > Apps Script」に貼り付けて公開すると、
 * シートの内容をサイトが読み取れる JSON に変換します。
 *
 * ★ 既存の公開Webアプリ(本番の SHEET_API_URL)は、
 *   このファイルの変更後も「デプロイを更新」しない限り
 *   古いバージョンのまま動作し続けます。
 *   管理画面は別デプロイとして新規公開してください。
 ****************************************************************/

function doGet(e) {
  if (e && e.parameter && e.parameter.admin === '1') {
    return renderAdminPage_();
  }
  return getPublicJson_();
}

/**
 * 既存の公開JSON API本体。
 * ロジックは変更前と完全に同一です(News/Shows/Movies/Discoを読み取って返すだけ)。
 */
function getPublicJson_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var result = {
    news:   readSheet(ss, 'News'),
    shows:  readSheet(ss, 'Shows'),
    movies: readSheet(ss, 'Movies'),
    disco:  readSheet(ss, 'Disco')
  };

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 指定したシートを読み取り、オブジェクトの配列に変換する。
 * シート構成: 1行目=説明文 / 2行目=見出し（キー） / 3行目=日本語の説明 / 4行目以降=実データ
 * 「公開」列が FALSE の行はスキップする。
 * ※この関数は変更していません。
 */
function readSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  var values = sheet.getDataRange().getValues();
  if (values.length < 4) return [];

  var headers = values[1];
  var rows = [];

  for (var i = 3; i < values.length; i++) {
    var row = values[i];
    var obj = {};
    var hasContent = false;

    for (var c = 0; c < headers.length; c++) {
      var key = String(headers[c]).trim();
      if (!key) continue;
      var val = row[c];

      // 日付型はそのまま文字に
      if (val instanceof Date) {
        val = Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy.MM.dd');
      }
      obj[key] = (val === '' || val === null) ? '' : String(val);
      if (obj[key]) hasContent = true;
    }

    // 「公開」列が FALSE / 非公開 / no の行は出さない
    var pub = String(obj['公開'] || obj['publish'] || '').toLowerCase();
    if (pub === 'false' || pub === 'no' || pub === '非公開' || pub === '0') continue;

    // 空行はスキップ
    if (!hasContent) continue;

    // News の本文を <p> でラップ（改行を段落に）
    if (sheetName === 'News' && obj.body) {
      obj.body = obj.body
        .split(/\n+/)
        .filter(function(t){ return t.trim(); })
        .map(function(t){ return '<p>' + t.trim() + '</p>'; })
        .join('');
    }

    delete obj['公開'];
    delete obj['publish'];
    rows.push(obj);
  }

  return rows;
}


/****************************************************************
 * ここから管理画面(News追加)用の追加コード
 *
 * 認証方式: Google Identity Services (GIS) によるGoogleログイン
 *   → クライアント側でIDトークン(JWT)取得
 *   → google.script.run でGASへ送信
 *   → GAS側で tokeninfo エンドポイントを使い署名・aud・exp・email_verified を検証
 *   → 検証済みメールアドレスを PropertiesService の許可リストと照合
 *
 * ★ Session.getActiveUser() は「自分として実行」+個人Gmail間では
 *   信頼できないことが公式ドキュメントで確認できたため、今回は使用しません。
 ****************************************************************/

/**
 * 管理画面(HTML)を返す。
 * ログイン前の状態でも開けるが、フォーム操作(addNews)自体はGAS側で
 * IDトークンを検証するまで一切実行されない。
 */
function renderAdminPage_() {
  return HtmlService
    .createHtmlOutputFromFile('AdminForm')
    .setTitle('Howling way 管理画面')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 初回セットアップ用。スクリプトエディタから手動で1回だけ実行してください。
 * (関数選択プルダウンで setupProperties_ を選び、実行ボタンを押す)
 *
 * ・ALLOWED_EMAILS   : 管理画面を利用できるGoogleアカウントのメールアドレス(カンマ区切り、2件)
 * ・OAUTH_CLIENT_ID  : Google Cloud Consoleで発行するOAuthクライアントID
 *                      (AdminForm.html内の GOOGLE_OAUTH_CLIENT_ID と必ず同じ値にすること)
 *
 * 値を変更したい場合は、この関数内の文字列を書き換えて再実行してください。
 */
function setupProperties_() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('ALLOWED_EMAILS', 's.keita0717@gmail.com,client@example.com');
  props.setProperty('OAUTH_CLIENT_ID', 'YOUR_OAUTH_CLIENT_ID.apps.googleusercontent.com');
}

/**
 * GoogleのIDトークン(JWT)をサーバー側で検証し、
 * 検証済みかつ許可リストに含まれるメールアドレスを返す。
 * 検証に失敗した場合は null を返す(例外は投げない。呼び出し元で判定する)。
 *
 * 検証項目:
 *   1. tokeninfoエンドポイントでの署名検証(HTTPステータス200であること)
 *   2. aud が本アプリのOAuthクライアントIDと一致すること
 *   3. exp(有効期限)が過ぎていないこと
 *   4. email_verified が true であること
 *   5. email が PropertiesService の許可リストに含まれること
 *
 * ※ tokeninfoエンドポイントはGoogle公式ドキュメントで
 *   「本番運用では公式クライアントライブラリ推奨、tokeninfoは開発/デバッグ向け
 *    (リクエスト抑制や断続的エラーの可能性がある)」と明記されているが、
 *   本管理画面は少人数・低頻度利用のため許容範囲と判断している。
 *   将来的にアクセス頻度が増える場合は、JWT検証ライブラリへの置き換えを検討すること。
 */
function verifyIdentity_(idToken) {
  if (!idToken || typeof idToken !== 'string') return null;

  var expectedAud = PropertiesService.getScriptProperties().getProperty('OAUTH_CLIENT_ID') || '';
  if (!expectedAud) {
    // OAuthクライアントIDが未設定の場合は、誰も認証できないようにする(安全側に倒す)
    return null;
  }

  var resp;
  try {
    resp = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true }
    );
  } catch (err) {
    return null;
  }

  if (resp.getResponseCode() !== 200) return null;

  var payload;
  try {
    payload = JSON.parse(resp.getContentText());
  } catch (err) {
    return null;
  }
  if (!payload) return null;

  // --- aud検証(必須。トークンが「このアプリ」宛に発行されたものかを確認) ---
  if (payload.aud !== expectedAud) return null;

  // --- iss検証(Googleが発行したものであることの追加確認) ---
  if (payload.iss !== 'accounts.google.com' && payload.iss !== 'https://accounts.google.com') {
    return null;
  }

  // --- exp検証(tokeninfo自体も期限切れなら400を返すが、念のため二重チェック) ---
  var exp = Number(payload.exp);
  if (!exp || exp < Math.floor(Date.now() / 1000)) return null;

  // --- email_verified検証 ---
  if (String(payload.email_verified) !== 'true') return null;

  var email = String(payload.email || '').trim().toLowerCase();
  if (!email) return null;

  // --- 許可リストとの照合 ---
  var raw = PropertiesService.getScriptProperties().getProperty('ALLOWED_EMAILS') || '';
  var allowList = raw.split(',')
    .map(function(s){ return s.trim().toLowerCase(); })
    .filter(function(s){ return !!s; });

  if (allowList.indexOf(email) === -1) return null;

  return email;
}

/**
 * クライアント側からログイン直後に呼び出す軽量チェック。
 * UX向上のためだけの関数であり、これ自体はセキュリティの本丸ではない
 * (本当のチェックは addNews() 内で毎回独立して行われる)。
 */
function checkLogin(idToken) {
  var email = verifyIdentity_(idToken);
  if (!email) {
    return { authorized: false };
  }
  return { authorized: true, email: email };
}

/**
 * シートのヘッダー行(2行目)を読み、「列名 → 列番号(1始まり)」のマップを作る。
 * readSheet() と同じ列名の考え方を書き込み側でも再利用するための共通処理。
 */
function getHeaderMap_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  var map = {};
  headers.forEach(function(h, i) {
    var key = String(h).trim();
    if (key) map[key] = i + 1; // 1始まりの列番号
  });
  return map;
}

/**
 * HTML特殊文字をエスケープする。
 * News本文・タイトル等は既存仕様上 <p> タグでラップされ innerHTML 的に
 * 描画されるため、ここでエスケープしておかないと <script> 等の任意タグが
 * そのまま実行されてしまう(保存型XSS)。改行(\n)はエスケープ対象外なので、
 * readSheet() 側の「改行 → <p> 変換」ロジックへの影響はない。
 */
function escapeHtml_(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * News管理フォームから呼び出される、Newsシートへの1行追加処理。
 * google.script.run 経由で AdminForm.html から呼ばれる。
 *
 * data = { date, cat, title, body, img, publish(boolean), idToken }
 *
 * ★ HTML側の判定を一切信用しない。ここで独立してIDトークンを再検証し、
 *   許可された2人以外は例外が投げられ、Sheetへの書き込みは一切発生しない。
 */
function addNews(data) {
  data = data || {};

  // --- サーバー側の認証・認可チェック(必須。HTML側のチェックだけに頼らない) ---
  var email = verifyIdentity_(data.idToken);
  if (!email) {
    throw new Error('権限がありません。許可されたGoogleアカウントでログインし直してください。');
  }

  var title = String(data.title || '').trim();
  var body  = String(data.body  || '').trim();

  if (!title) throw new Error('タイトルは必須です。');
  if (!body)  throw new Error('本文は必須です。');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('News');
  if (!sheet) throw new Error('Newsシートが見つかりません。');

  var map = getHeaderMap_(sheet);
  var row = new Array(Math.max(sheet.getLastColumn(), 6)).fill('');

  function setCol(headerName, value) {
    var col = map[headerName];
    if (col) row[col - 1] = value;
  }

  // XSS対策: 保存前にHTML特殊文字をエスケープ(改行はそのまま維持)
  setCol('日付',     escapeHtml_(String(data.date || '').trim()));
  setCol('カテゴリ', escapeHtml_(String(data.cat  || '').trim()));
  setCol('タイトル', escapeHtml_(title));
  setCol('本文',     escapeHtml_(body));
  setCol('画像URL',  escapeHtml_(String(data.img  || '').trim()));
  setCol('公開',     data.publish ? 'TRUE' : 'FALSE');

  // マニュアル記載の運用(「一番上の行が最新として表示される」)に合わせ、
  // データ4行目(見出し・説明の直下)に新規行を挿入する。
  // ※ Newsの表示順は配列順そのまま(JS側でのソートなし)のため、末尾追加(appendRow)は不可。
  sheet.insertRowBefore(4);
  sheet.getRange(4, 1, 1, row.length).setValues([row]);

  return { success: true, message: 'Newsを登録しました。' };
}
