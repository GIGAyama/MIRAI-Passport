/**
 * ============================================================
 *  みらいパスポート v3.0.0 — サーバーサイド (Code.gs)
 * ============================================================
 *
 * ■ このファイルの役割
 *   Google Apps Script（GAS）のバックエンド処理を担当します。
 *   - Webアプリの起動（doGet）
 *   - Google スプレッドシートへのデータ読み書き（DB操作）
 *   - Gemini AI API との通信
 *   - みらいコンパス（外部LMS）との連携
 *
 * ■ 設定不要のしくみ
 *   初回起動時に「初期化開始」ボタンを押すだけで、
 *   スプレッドシート（データベース）が自動作成されます。
 *   手動でスプレッドシートを用意する必要はありません。
 *
 * ■ データベース構造（自動作成されるスプレッドシート）
 *   シート1: Worksheets（ワークシート本体）
 *     A列: taskId       — タスク固有ID
 *     B列: unitName     — 単元名
 *     C列: stepTitle    — 活動タイトル
 *     D列: htmlContent  — ワークシートのHTML本文
 *     E列: lastUpdated  — 最終更新日時
 *     F列: jsonSource   — 元の計画データ（JSON文字列）
 *     G列: canvasJson   — 手書きキャンバスデータ（JSON文字列）
 *     H列: rubricHtml   — ルーブリック（評価基準）HTML
 *     I列: isShared     — 共有フラグ
 *
 *   シート2: Responses（児童の回答データ）
 *     A列: responseId   — 回答固有ID
 *     B列: taskId       — 対象タスクID
 *     C列: studentId    — 児童ID
 *     D列: studentName  — 児童名
 *     E列: submittedAt  — 提出日時
 *     F列: canvasImage  — キャンバスのスナップショット画像（Base64）
 *     G列: textContent  — 自己評価テキスト
 *     H列: status       — 状態（draft / submitted / graded）
 *     I列: feedbackText — 先生からのコメント
 *     J列: score        — 点数（将来用）
 *     K列: feedbackJson — 添削データ（将来用）
 *     L列: canvasJson   — 児童の手書きデータ（JSON文字列）
 *     M列: isPublic     — 広場への公開フラグ
 *     N列: reactions    — リアクション一覧（JSON配列文字列）
 *
 *   シート3: ImportQueue（コンパス連携用の一時キュー）
 *     A列: transactionId — トランザクションID
 *     B列: dataJson      — インポートデータ（JSON文字列）
 *     C列: createdAt     — 作成日時
 *
 * ============================================================
 */

/* ========== 定数 ========== */
var APP_NAME = "みらいパスポート";
var DB_NAME  = APP_NAME + "_DB";

/* ---------- ワークシートシートの列番号（1始まり） ---------- */
var WS_COL_TASK_ID      = 1;  // A列
var WS_COL_UNIT_NAME    = 2;  // B列
var WS_COL_STEP_TITLE   = 3;  // C列
var WS_COL_HTML_CONTENT = 4;  // D列
var WS_COL_LAST_UPDATED = 5;  // E列
var WS_COL_JSON_SOURCE  = 6;  // F列
var WS_COL_CANVAS_JSON  = 7;  // G列
var WS_COL_RUBRIC_HTML  = 8;  // H列
var WS_COL_IS_SHARED    = 9;  // I列
var WS_TOTAL_COLS       = 9;  // 列の総数

/* ---------- レスポンスシートの列番号（1始まり） ---------- */
var RS_COL_RESPONSE_ID  = 1;   // A列
var RS_COL_TASK_ID      = 2;   // B列
var RS_COL_STUDENT_ID   = 3;   // C列
var RS_COL_STUDENT_NAME = 4;   // D列
var RS_COL_SUBMITTED_AT = 5;   // E列
var RS_COL_CANVAS_IMAGE = 6;   // F列
var RS_COL_TEXT_CONTENT  = 7;   // G列
var RS_COL_STATUS       = 8;   // H列
var RS_COL_FEEDBACK_TXT = 9;   // I列
var RS_COL_SCORE        = 10;  // J列
var RS_COL_FEEDBACK_JSON= 11;  // K列
var RS_COL_CANVAS_JSON  = 12;  // L列
var RS_COL_IS_PUBLIC    = 13;  // M列
var RS_COL_REACTIONS    = 14;  // N列
var RS_TOTAL_COLS       = 14;  // 列の総数


/* ============================================================
 *  1. エントリーポイント & テンプレート
 * ============================================================ */

/**
 * Webアプリにアクセスがあったとき最初に呼ばれる関数。
 * URLパラメータを読み取り、index.html テンプレートに注入して返す。
 *
 * 使用されるURLパラメータ:
 *   mode        — "teacher"（先生モード）または "student"（児童モード）
 *   taskId      — 開く課題のID
 *   studentId   — 児童のID（コンパス連携時に自動付与）
 *   studentName — 児童の名前（コンパス連携時に自動付与）
 *   importId    — コンパスからの一括インポート用ID
 *
 * @param {Object} e - GASが自動的に渡すイベントオブジェクト
 * @return {HtmlOutput} 組み立てたHTMLページ
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');

  // URLパラメータをテンプレート変数にセット（未指定は空文字）
  template.mode        = e.parameter.mode        || 'teacher';
  template.taskId      = e.parameter.taskId      || '';
  template.studentId   = e.parameter.studentId   || '';
  template.studentName = e.parameter.studentName  || '';
  template.importId    = e.parameter.importId    || '';

  return template.evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://drive.google.com/uc?id=1mVwFtlrJvqEIk-0Gd03BqmG_-0BZiqY5&.png');
}

/**
 * HTML内で <?!= include('ファイル名'); ?> と書くと、
 * 別ファイルの中身をそのまま埋め込める仕組み。
 * css.html / js_core.html / js_student.html / js_teacher.html の読み込みに使う。
 *
 * @param {string} filename - 読み込むHTMLファイル名（拡張子なし）
 * @return {string} ファイルの中身（HTML文字列）
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/* ============================================================
 *  2. 初期セットアップ（データベース構築）
 * ============================================================ */

/**
 * データベース（スプレッドシート）が既に作成済みかどうかを確認する。
 * アプリ起動時にフロントエンドから呼ばれ、
 * 未作成なら「初期化開始」ボタンの画面を表示する。
 *
 * @return {Object} { isSetup: true/false, dbId: スプレッドシートID or null }
 */
function checkSetupStatus() {
  var props = PropertiesService.getScriptProperties();
  var dbId  = props.getProperty('DB_SS_ID');

  if (dbId) {
    try {
      // IDで開けるかテストする（削除されている場合はエラーになる）
      SpreadsheetApp.openById(dbId);
      return { isSetup: true, dbId: dbId };
    } catch (e) {
      // スプレッドシートが消されている場合
      return { isSetup: false, dbId: null };
    }
  }
  return { isSetup: false, dbId: null };
}

/**
 * 初期化処理。データベース用スプレッドシートを作成し、
 * 3つのシート（Worksheets / Responses / ImportQueue）を準備する。
 * 既に同名のスプレッドシートがGoogleドライブに存在する場合は再利用する。
 *
 * ※ フロントエンドの「初期化開始」ボタンから呼ばれる
 *
 * @return {Object} { success: true, url: スプレッドシートのURL }
 * @throws {Error} 初期化に失敗した場合
 */
function performInitialSetup() {
  var props = PropertiesService.getScriptProperties();
  var ss    = null;

  try {
    // 同名のスプレッドシートがあれば再利用、なければ新規作成
    var files = DriveApp.getFilesByName(DB_NAME);
    if (files.hasNext()) {
      ss = SpreadsheetApp.openById(files.next().getId());
    } else {
      ss = SpreadsheetApp.create(DB_NAME);
    }

    // 各シートをヘッダー付きで確保（既に存在すればスキップ）
    ensureSheet(ss, 'Worksheets', [
      'taskId', 'unitName', 'stepTitle', 'htmlContent',
      'lastUpdated', 'jsonSource', 'canvasJson', 'rubricHtml', 'isShared'
    ]);
    ensureSheet(ss, 'Responses', [
      'responseId', 'taskId', 'studentId', 'studentName',
      'submittedAt', 'canvasImage', 'textContent', 'status',
      'feedbackText', 'score', 'feedbackJson', 'canvasJson',
      'isPublic', 'reactions'
    ]);
    ensureSheet(ss, 'ImportQueue', [
      'transactionId', 'dataJson', 'createdAt'
    ]);

    // スプレッドシートIDをスクリプトプロパティに保存
    props.setProperty('DB_SS_ID', ss.getId());
    return { success: true, url: ss.getUrl() };

  } catch (e) {
    throw new Error("初期化エラー: " + e.message);
  }
}

/**
 * 指定したスプレッドシートにシートが存在しなければ新規作成し、
 * 1行目にヘッダーを書き込む。既に存在する場合は何もしない。
 *
 * @param {Spreadsheet} ss     - 対象のスプレッドシート
 * @param {string}      name   - シート名
 * @param {string[]}    header - ヘッダー行の配列
 * @return {Sheet} 対象シート
 */
function ensureSheet(ss, name, header) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(header);
  }
  return sheet;
}


/* ============================================================
 *  3. 設定管理（APIキー・教師名・コンパスURL）
 * ============================================================ */

/**
 * ユーザー設定を保存する。
 * → 保存先: PropertiesService の UserProperties（ユーザーごとに独立）
 *
 * @param {string} apiKey      - Gemini AI の APIキー
 * @param {string} teacherName - 先生の名前
 * @param {string} compassUrl  - みらいコンパスの WebアプリURL（任意）
 * @return {boolean} true（常に成功）
 */
function saveUserConfig(apiKey, teacherName, compassUrl) {
  var props = {
    'GEMINI_API_KEY': apiKey,
    'TEACHER_NAME':   teacherName
  };
  // compassUrl が渡された場合のみ保存（前後の空白を除去）
  if (compassUrl !== undefined) {
    props['COMPASS_URL'] = compassUrl.trim();
  }
  PropertiesService.getUserProperties().setProperties(props);
  return true;
}

/**
 * カスタム AI プロンプト（単元計画生成用）を保存する。
 *
 * @param {string} promptText - プロンプト文字列
 * @return {boolean} 成功時 true
 */
function saveCustomAiPrompt(promptText) {
  PropertiesService.getUserProperties().setProperty('CUSTOM_AI_PROMPT', promptText || '');
  return true;
}

/**
 * カスタム AI プロンプト（単元計画生成用）を取得する。
 * 保存済みのものがなければ空文字を返す（フロント側でデフォルト値を使う）。
 *
 * @return {Object} { success: true, prompt: 保存済みプロンプト }
 */
function getCustomAiPrompt() {
  var prompt = PropertiesService.getUserProperties().getProperty('CUSTOM_AI_PROMPT') || '';
  return { success: true, prompt: prompt };
}

/**
 * ユーザー設定を読み込む。
 * UserProperties → ScriptProperties の順に探し、
 * 最初に見つかった値を返す（フォールバック機構）。
 *
 * @return {Object} { apiKey, teacherName, compassUrl }
 */
function getUserConfig() {
  var userProps   = PropertiesService.getUserProperties();
  var scriptProps = PropertiesService.getScriptProperties();

  return {
    apiKey:      userProps.getProperty('GEMINI_API_KEY')
                 || scriptProps.getProperty('GEMINI_API_KEY') || '',
    teacherName: userProps.getProperty('TEACHER_NAME') || '',
    compassUrl:  userProps.getProperty('COMPASS_URL')
                 || scriptProps.getProperty('COMPASS_URL') || ''
  };
}


/* ============================================================
 *  4. データベース操作 — ワークシート（Worksheets シート）
 * ============================================================ */

/**
 * データベース用スプレッドシートを開いて返す。
 * スクリプトプロパティに保存されたIDを使用する。
 *
 * @return {Spreadsheet} データベーススプレッドシート
 * @throws {Error} DB_SS_ID が未設定の場合
 */
function getDbSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty('DB_SS_ID');
  if (!id) throw new Error("DB未接続");
  return SpreadsheetApp.openById(id);
}

/**
 * ワークシートをデータベースに保存する。
 * taskId が既に存在する場合は上書き更新、存在しない場合は新規追加。
 *
 * → 書き込み先: Worksheets シートの各列（上部のDB構造コメント参照）
 *
 * @param {Object} data - 保存するワークシートデータ
 * @param {string} data.taskId      - タスクID（未指定なら自動生成）
 * @param {string} data.unitName    - 単元名
 * @param {string} data.stepTitle   - 活動タイトル
 * @param {string} data.htmlContent - ワークシートHTML本文
 * @param {Object} data.jsonSource  - 元の計画データ
 * @param {Object} data.canvasJson  - キャンバスデータ
 * @param {string} data.rubricHtml  - ルーブリックHTML
 * @param {boolean} data.isShared   - 共有フラグ
 * @return {boolean} true
 */
function saveWorksheetToDB(data) {
  var sheet = getDbSpreadsheet().getSheetByName('Worksheets');
  var now   = new Date();

  // 保存用レコードの組み立て
  var taskId      = String(data.taskId || Utilities.getUuid());
  var unitName    = data.unitName    || "無題";
  var stepTitle   = data.stepTitle   || "無題";
  var htmlContent = data.htmlContent || "";
  var jsonSource  = JSON.stringify(data.jsonSource || {});
  var canvasJson  = data.canvasJson ? JSON.stringify(data.canvasJson) : "";
  var rubricHtml  = data.rubricHtml  || "";
  var isShared    = data.isShared    || false;

  // A列（taskId）で既存行を検索
  var found = sheet.getRange("A:A")
    .createTextFinder(taskId)
    .matchEntireCell(true)
    .findNext();

  if (found) {
    // --- 既存行を更新（変化しうる列だけまとめて書き込み） ---
    var row = found.getRow();
    // D列〜I列を一括更新（6列分）
    sheet.getRange(row, WS_COL_HTML_CONTENT, 1, 6).setValues([[
      htmlContent,  // D列: HTML本文
      now,          // E列: 最終更新日時
      jsonSource,   // F列: JSONソース（更新時も最新を保持）
      canvasJson,   // G列: キャンバスデータ
      rubricHtml,   // H列: ルーブリック
      isShared      // I列: 共有フラグ
    ]]);
  } else {
    // --- 新規行を追加 ---
    sheet.appendRow([
      taskId,       // A列: タスクID
      unitName,     // B列: 単元名
      stepTitle,    // C列: 活動タイトル
      htmlContent,  // D列: HTML本文
      now,          // E列: 最終更新日時
      jsonSource,   // F列: JSONソース
      canvasJson,   // G列: キャンバスデータ
      rubricHtml,   // H列: ルーブリック
      isShared      // I列: 共有フラグ
    ]);
  }
  return true;
}

/**
 * 指定した taskId のワークシートをデータベースから読み込む。
 *
 * → 読み込み元: Worksheets シート
 * → 最適化: 1行分を一括取得（セル単位アクセスを排除）
 *
 * @param {string} taskId - 取得するタスクID
 * @return {Object|null} ワークシートデータ、見つからなければ null
 */
function loadWorksheetFromDB(taskId) {
  var sheet = getDbSpreadsheet().getSheetByName('Worksheets');

  // A列でタスクIDを検索
  var found = sheet.getRange("A:A")
    .createTextFinder(String(taskId))
    .matchEntireCell(true)
    .findNext();

  if (!found) return null;

  // 見つかった行のA列〜I列を一括取得（1回の通信で完了）
  var row    = found.getRow();
  var values = sheet.getRange(row, 1, 1, WS_TOTAL_COLS).getValues()[0];

  return {
    taskId:      values[0],                        // A列
    unitName:    values[1],                        // B列
    stepTitle:   values[2],                        // C列
    htmlContent: values[3],                        // D列
    jsonSource:  safeJSONParse(values[5]),          // F列（JSON文字列→オブジェクト）
    canvasJson:  safeJSONParse(values[6]),          // G列（JSON文字列→オブジェクト）
    rubricHtml:  values[7],                        // H列
    isShared:    values[8]                         // I列
  };
}

/**
 * 複数のタスクIDに該当するワークシートをまとめて取得する。
 * 一括生成・一括印刷で使用する。
 *
 * → 読み込み元: Worksheets シートの全データ
 *
 * @param {string[]} taskIds - 取得したいタスクIDの配列
 * @return {Object[]} ワークシートデータの配列
 */
function getWorksheetsByIds(taskIds) {
  var sheet = getDbSpreadsheet().getSheetByName('Worksheets');
  var data  = sheet.getDataRange().getValues();

  // 1行目（ヘッダー）をスキップし、指定IDに一致する行だけ返す
  return data.slice(1)
    .filter(function(row) {
      return taskIds.includes(String(row[0]));
    })
    .map(function(row) {
      return {
        taskId:      row[0],
        unitName:    row[1],
        stepTitle:   row[2],
        htmlContent: row[3],
        canvasJson:  safeJSONParse(row[6]),
        jsonSource:  safeJSONParse(row[5])
      };
    });
}

/**
 * 保存済みワークシートの履歴（最新30件）を取得する。
 * サイドバーの「履歴」タブに表示するためのデータ。
 *
 * → 読み込み元: Worksheets シートの A列（ID）、C列（タイトル）、E列（更新日時）
 *
 * @return {Object[]} { id, title, timestamp } の配列（新しい順）
 */
function getHistory() {
  var sheet = getDbSpreadsheet().getSheetByName('Worksheets');

  // データが1行（ヘッダー）しかなければ空配列
  if (sheet.getLastRow() < 2) return [];

  // A列〜E列の範囲を一括取得
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();

  return data
    .map(function(r) {
      return {
        id:        r[0],
        title:     r[2] || "無題",
        timestamp: new Date(r[4]).getTime()
      };
    })
    .filter(function(item) { return item.id; })  // IDが空の行を除外
    .sort(function(a, b) { return b.timestamp - a.timestamp; })  // 新しい順
    .slice(0, 30);  // 最大30件
}


/**
 * 児童向け: 配信済みワークシートの一覧（軽量版）を取得する。
 * htmlContent を含まないため高速。
 * 児童サイドバーで単元別にワークシートを表示するために使用。
 *
 * → 読み込み元: Worksheets シートの A列（ID）、B列（単元名）、C列（タイトル）
 *
 * @return {Object[]} { taskId, unitName, stepTitle } の配列
 */
function getStudentWorksheetList() {
  var sheet = getDbSpreadsheet().getSheetByName('Worksheets');
  if (sheet.getLastRow() < 2) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  return data
    .filter(function(row) { return row[0]; })
    .map(function(row) {
      return { taskId: String(row[0]), unitName: String(row[1]), stepTitle: String(row[2]) };
    });
}


/* ============================================================
 *  5. データベース操作 — 児童レスポンス（Responses シート）
 * ============================================================ */

/**
 * 児童の回答データを保存する。
 * 同じ taskId × studentId の組み合わせが既に存在すれば更新、
 * 存在しなければ新規追加する。
 *
 * → 書き込み先: Responses シートの各列
 *
 * @param {Object} data - 回答データ
 * @param {string} data.taskId      - 対象タスクID
 * @param {string} data.studentId   - 児童ID
 * @param {string} data.studentName - 児童名
 * @param {string} data.canvasImage - キャンバスの画像（Base64）
 * @param {string} data.textContent - 自己評価テキスト
 * @param {string} data.status      - 状態（draft / submitted）
 * @param {string} data.canvasJson  - キャンバスJSON
 * @param {boolean} data.isPublic   - 広場への公開フラグ
 * @return {Object} { success: true }
 */
function saveStudentResponse(data) {
  var sheet = getDbSpreadsheet().getSheetByName('Responses');
  var now   = new Date();

  // 同じ taskId × studentId の既存行を探す
  // （TextFinder は単一列しか検索できないため、全データをループで確認）
  var vals = sheet.getDataRange().getValues();
  var existingRow = -1;
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][1]) === String(data.taskId) &&
        String(vals[i][2]) === String(data.studentId)) {
      existingRow = i + 1;  // シート上の行番号（1始まり）
      break;
    }
  }

  // isPublic の初期値は true（未指定なら公開）
  var isPublicVal = (data.isPublic === undefined) ? true : data.isPublic;

  if (existingRow > 0) {
    // --- 既存行を更新 ---
    sheet.getRange(existingRow, RS_COL_STUDENT_NAME).setValue(data.studentName);       // D列: 児童名
    sheet.getRange(existingRow, RS_COL_SUBMITTED_AT).setValue(now);                    // E列: 提出日時
    sheet.getRange(existingRow, RS_COL_CANVAS_IMAGE).setValue(data.canvasImage);       // F列: 画像
    sheet.getRange(existingRow, RS_COL_TEXT_CONTENT).setValue(data.textContent);        // G列: テキスト
    sheet.getRange(existingRow, RS_COL_STATUS).setValue(data.status);                  // H列: 状態
    if (data.canvasJson) {
      sheet.getRange(existingRow, RS_COL_CANVAS_JSON).setValue(data.canvasJson);       // L列: キャンバスJSON
    }
    sheet.getRange(existingRow, RS_COL_IS_PUBLIC).setValue(isPublicVal);                // M列: 公開フラグ
  } else {
    // --- 新規行を追加 ---
    sheet.appendRow([
      Utilities.getUuid(),       // A列: 回答ID（自動生成）
      data.taskId,               // B列: タスクID
      data.studentId,            // C列: 児童ID
      data.studentName,          // D列: 児童名
      now,                       // E列: 提出日時
      data.canvasImage || "",    // F列: 画像
      data.textContent || "",    // G列: テキスト
      data.status || "submitted",// H列: 状態
      "",                        // I列: フィードバック（空）
      "",                        // J列: スコア（空）
      "",                        // K列: フィードバックJSON（空）
      data.canvasJson || "",     // L列: キャンバスJSON
      isPublicVal,               // M列: 公開フラグ
      "[]"                       // N列: リアクション（空配列）
    ]);
  }
  return { success: true };
}

/**
 * 指定タスクの全提出データを取得する。
 * 教師の「提出状況ダッシュボード」に表示するためのデータ。
 *
 * → 読み込み元: Responses シートの全行から taskId が一致するものを抽出
 *
 * @param {string} taskId - 対象タスクID
 * @return {Object[]} 提出データの配列
 */
function getTaskSubmissions(taskId) {
  var sheet  = getDbSpreadsheet().getSheetByName('Responses');
  var values = sheet.getDataRange().getValues();

  var results = [];
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][1]) === String(taskId)) {
      results.push({
        rowIndex:     i + 1,      // シート上の行番号（添削保存時に使用）
        studentId:    values[i][2],
        studentName:  values[i][3],
        submittedAt:  values[i][4],
        canvasImage:  values[i][5],
        status:       values[i][7],
        feedbackText: values[i][8],
        canvasJson:   values[i][11]
      });
    }
  }
  return results;
}

/**
 * 教師用管理画面: 全提出データとワークシート一覧を一括取得する。
 * クライアント側でフィルタリングするため、1回の通信で全データを返す。
 * canvasJson は重いので除外し、canvasImage（サムネイル）のみ返す。
 *
 * → 読み込み元: Responses シート + Worksheets シート
 *
 * @return {Object} { submissions: [...], worksheets: [...] }
 */
function getDashboardData() {
  var ss = getDbSpreadsheet();

  // --- ワークシート一覧（軽量: ID + 単元名 + タイトルのみ） ---
  var wsSheet = ss.getSheetByName('Worksheets');
  var worksheets = [];
  if (wsSheet.getLastRow() >= 2) {
    var wsData = wsSheet.getRange(2, 1, wsSheet.getLastRow() - 1, 3).getValues();
    worksheets = wsData
      .filter(function(r) { return r[0]; })
      .map(function(r) {
        return { taskId: String(r[0]), unitName: String(r[1]), stepTitle: String(r[2]) };
      });
  }

  // --- 全提出データ（canvasJson を除外して軽量化） ---
  var resSheet = ss.getSheetByName('Responses');
  var submissions = [];
  if (resSheet.getLastRow() >= 2) {
    var resData = resSheet.getDataRange().getValues();
    for (var i = 1; i < resData.length; i++) {
      var row = resData[i];
      if (!row[0]) continue;  // responseId が空なら skip
      submissions.push({
        rowIndex:     i + 1,
        responseId:   row[0],
        taskId:       String(row[1]),
        studentId:    row[2],
        studentName:  row[3],
        submittedAt:  row[4] ? new Date(row[4]).getTime() : 0,
        canvasImage:  row[5],
        textContent:  row[6],
        status:       row[7],
        feedbackText: row[8]
      });
    }
  }

  return { submissions: submissions, worksheets: worksheets };
}

/**
 * 教師用管理画面: 指定行の提出データからcanvasJsonを取得する。
 * 添削プレビュー時にのみ呼び出す（一覧では不要なため分離）。
 *
 * @param {number} rowIndex - Responsesシートの行番号
 * @return {Object|null} { canvasJson, htmlContent }
 */
function getSubmissionDetail(rowIndex) {
  var ss = getDbSpreadsheet();
  var resSheet = ss.getSheetByName('Responses');
  var row = resSheet.getRange(rowIndex, 1, 1, RS_TOTAL_COLS).getValues()[0];
  var taskId = String(row[1]);

  // 対応するワークシートのHTMLも取得
  var wsSheet = ss.getSheetByName('Worksheets');
  var htmlContent = '';
  var found = wsSheet.getRange("A:A").createTextFinder(taskId).matchEntireCell(true).findNext();
  if (found) {
    htmlContent = wsSheet.getRange(found.getRow(), WS_COL_HTML_CONTENT).getValue();
  }

  return {
    canvasJson:  safeJSONParse(row[11]),
    htmlContent: htmlContent
  };
}

/**
 * 教師が児童の回答に対するフィードバック（添削）を保存する。
 *
 * → 書き込み先: Responses シートの H列（status）、I列（feedbackText）、L列（canvasJson）
 *
 * @param {Object} data - フィードバックデータ
 * @param {number} data.rowIndex     - 対象のシート行番号
 * @param {string} data.feedbackText - コメント
 * @param {string} data.canvasJson   - 赤ペン添削のキャンバスデータ
 * @return {Object} { success: true }
 */
function saveFeedback(data) {
  var sheet = getDbSpreadsheet().getSheetByName('Responses');

  if (data.rowIndex) {
    sheet.getRange(data.rowIndex, RS_COL_STATUS).setValue("graded");                   // H列: 「添削済」に変更
    sheet.getRange(data.rowIndex, RS_COL_FEEDBACK_TXT).setValue(data.feedbackText);     // I列: コメント
    if (data.canvasJson) {
      sheet.getRange(data.rowIndex, RS_COL_CANVAS_JSON).setValue(data.canvasJson);      // L列: 添削キャンバス
    }
  }
  return { success: true };
}

/**
 * 広場（ギャラリー）に公開されている回答を取得する。
 * 児童の「みんなの広場」画面で、友達の作品を表示するのに使用する。
 *
 * → 読み込み元: Responses シートの全行から条件に一致するものを抽出
 *   条件: taskId が一致 AND (status が submitted または graded) AND isPublic が true
 *
 * @param {string} taskId - 対象タスクID
 * @return {Object[]} 公開回答データの配列
 */
function getSharedResponses(taskId) {
  var sheet  = getDbSpreadsheet().getSheetByName('Responses');
  var values = sheet.getDataRange().getValues();

  var results = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];

    // 公開判定（空文字・true・"true" のいずれかなら公開とみなす）
    var isPublic = (row[12] === "" || row[12] === true || row[12] === "true");

    // taskId が一致、かつ提出済みor添削済み、かつ公開設定
    if (String(row[1]) === String(taskId) &&
        (row[7] === 'submitted' || row[7] === 'graded') &&
        isPublic) {
      results.push({
        responseId:  row[0],
        studentId:   row[2],
        studentName: row[3],
        canvasImage: row[5],
        canvasJson:  row[11],
        reactions:   ensureArray(safeJSONParse(row[13]))
      });
    }
  }
  return results;
}

/**
 * 友達の作品にリアクション（スタンプ・コメント）を送る。
 *
 * → 書き込み先: Responses シートの N列（reactions）
 *   既存のリアクション配列に新しいリアクションを追加して上書き保存する。
 *
 * @param {Object} data - リアクションデータ
 * @param {string} data.targetResponseId - 対象の回答ID
 * @param {Object} data.reaction         - リアクション内容 { type, value, fromName }
 * @return {Object} { success: true/false, reactions: 更新後の配列 }
 */
function savePeerReaction(data) {
  var sheet = getDbSpreadsheet().getSheetByName('Responses');

  // A列（responseId）で対象行を検索
  var finder = sheet.getRange("A:A")
    .createTextFinder(data.targetResponseId)
    .matchEntireCell(true)
    .findNext();

  if (finder) {
    var row  = finder.getRow();
    var cell = sheet.getRange(row, RS_COL_REACTIONS);  // N列

    // 既存のリアクション配列を取得し、新しいリアクションを追加
    var current = ensureArray(safeJSONParse(cell.getValue()));
    data.reaction.timestamp = new Date().getTime();
    current.push(data.reaction);

    // 更新した配列をJSON文字列として保存
    cell.setValue(JSON.stringify(current));
    return { success: true, reactions: current };
  }
  return { success: false };
}

/**
 * 指定した児童の回答データを取得する（最新のもの）。
 * 児童がワークシートを開いたとき、以前の回答を復元するために使う。
 *
 * → 読み込み元: Responses シート（末尾から逆順に検索して最初に見つかったもの）
 *
 * @param {string} taskId    - 対象タスクID
 * @param {string} studentId - 児童ID
 * @return {Object|null} 回答データ、見つからなければ null
 */
function getMyResponse(taskId, studentId) {
  var sheet  = getDbSpreadsheet().getSheetByName('Responses');
  var values = sheet.getDataRange().getValues();

  // 末尾から検索（最新の回答を優先的に返す）
  for (var i = values.length - 1; i >= 1; i--) {
    if (String(values[i][1]) === String(taskId) &&
        String(values[i][2]) === String(studentId)) {
      return {
        responseId:   values[i][0],
        status:       values[i][7],
        feedbackText: values[i][8],
        canvasImage:  values[i][5],
        canvasJson:   values[i][11],
        isPublic:     values[i][12],
        reactions:    ensureArray(safeJSONParse(values[i][13]))
      };
    }
  }
  return null;
}


/* ============================================================
 *  6. AI 連携（Gemini API）
 * ============================================================ */

/**
 * Gemini AI API にプロンプトを送信し、生成されたテキストを返す。
 * ワークシートの自動生成・ルーブリック作成に使用する。
 *
 * → 通信先: Google Generative Language API（Gemini 2.5 Flash）
 *
 * @param {string} prompt - AIに送るプロンプト文
 * @return {string} AIが生成したテキスト
 * @throws {Error} APIキー未設定 / APIエラー / 応答なし の場合
 */
function callGeminiAPI(prompt) {
  var apiKey = getUserConfig().apiKey;
  if (!apiKey) {
    throw new Error("Gemini APIキーが設定されていません。先生モードの設定を確認してください。");
  }

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;

  var payload = {
    contents: [{ parts: [{ text: prompt }] }]
  };

  var options = {
    method:             'post',
    contentType:        'application/json',
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var res  = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(res.getContentText());

  // エラーチェック
  if (json.error) {
    throw new Error("AIエラー: " + json.error.message);
  }
  if (!json.candidates || !json.candidates[0].content) {
    throw new Error("AIから応答が得られませんでした。");
  }

  return json.candidates[0].content.parts[0].text;
}

/**
 * AIコーチング対応のワークシートHTMLを生成する。
 * 一括生成時にサーバーサイドから呼ばれる。
 *
 * @param {Object} data - 生成に必要な情報
 * @param {string} data.unitName    - 単元名
 * @param {string} data.stepTitle   - 活動タイトル
 * @param {string} data.description - 活動内容の説明
 * @return {string} 生成されたHTML文字列
 */
function generateSingleWorksheet(data) {
  var prompt = 'あなたは教育工学と個別最適な学びの専門家です。\n'
    + '児童が自立的に学習を進められるよう、以下の活動内容に基づいた「AIコーチング機能付きHTMLワークシート」の本文を作成してください。\n\n'
    + '【活動内容】\n'
    + '単元: ' + data.unitName + '\n'
    + '活動: ' + data.stepTitle + '\n'
    + '内容: ' + data.description + '\n\n'
    + '【デザインの要件】\n'
    + '1. 小学生が親しみやすい言葉遣い。\n'
    + '2. 以下のセクションを必ず含める：\n'
    + '   - 「今日のめあて」（活動内容から具体化）\n'
    + '   - 「AIヒント・ポイント」（この活動でつまずきやすい点や、考えるコツをAIコーチとして助言）\n'
    + '   - 「考えを書くスペース」（<div class="ws-answer" style="border:1px solid #aaa; border-radius:6px; padding:10px; min-height:3em; background:#fffde7;"></div> を使用。問題文には ws-answer を付けないこと）\n'
    + '   - 「自己評価」（3段階のスタンプ選択など）\n'
    + '3. スタイルはBootstrap 5のクラス（card, p-3, mb-3, bg-lightなど）を活用。\n'
    + '4. HTMLの「body内部」のみを出力すること。余計な解説や```htmlタグは不要。';

  return callGeminiAPI(prompt);
}

/**
 * AIでルーブリック（評価基準表）を生成する。
 *
 * @param {Object} data - 生成に必要な情報
 * @param {string} data.unitName    - 単元名
 * @param {string} data.stepTitle   - 活動タイトル
 * @param {string} data.description - 活動内容の説明
 * @return {string} 生成されたHTMLテーブル文字列
 */
function generateRubricAI(data) {
  var prompt = '教育評価専門家としてルーブリック作成。'
    + '単元:' + data.unitName
    + ',活動:' + data.stepTitle
    + ',内容:' + data.description
    + '。3観点3段階,HTMLテーブル形式(table table-bordered),具体的記述。HTMLのみ。';

  return callGeminiAPI(prompt);
}


/* ============================================================
 *  7. ユーティリティ関数
 * ============================================================ */

/**
 * このWebアプリの公開URLを取得する。
 * 児童への「配布URL」生成に使用する。
 *
 * @return {string} WebアプリのURL
 */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * JSON文字列を安全にパース（解析）する。
 * 不正なJSONの場合はエラーにならず null を返す。
 *
 * @param {string} s - パースするJSON文字列
 * @return {*} パース結果、失敗時は null
 */
function safeJSONParse(s) {
  try {
    return JSON.parse(s);
  } catch (e) {
    return null;
  }
}

/**
 * 値が配列であることを保証する。
 * 配列でなければ空配列を返す。
 *
 * @param {*} val - チェックする値
 * @return {Array} 配列
 */
function ensureArray(val) {
  return Array.isArray(val) ? val : [];
}


/* ============================================================
 *  8. みらいコンパス連携
 * ============================================================ */

/**
 * みらいコンパスからのインポートキューを処理する。
 * コンパスが ImportQueue シートに書き込んだデータを読み取り、
 * Worksheets シートに取り込む。
 *
 * → 読み込み元: ImportQueue シート（処理後に該当行を削除）
 * → 書き込み先: Worksheets シート（handleImportUnitPlan 経由）
 *
 * @param {string} importId - インポートキューのトランザクションID
 * @return {Object} 処理結果 { success, taskIds, message }
 */
function consumeImportQueue(importId) {
  if (!importId) return { success: false, message: "ID未指定" };

  var ss    = getDbSpreadsheet();
  var sheet = ss.getSheetByName('ImportQueue');
  if (!sheet) return { success: false, message: "連携シートなし" };

  var data           = sheet.getDataRange().getValues();
  var foundData      = null;
  var deleteRowIndex = -1;

  // A列（transactionId）で検索
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(importId)) {
      foundData      = safeJSONParse(data[i][1]);  // B列: JSON データ
      deleteRowIndex = i + 1;                       // シート上の行番号
      break;
    }
  }

  if (deleteRowIndex > 0 && foundData) {
    // ワークシートシートにインポート
    var result = handleImportUnitPlan(foundData);
    // 処理済みキューを削除
    sheet.deleteRow(deleteRowIndex);
    return {
      success: true,
      taskIds: result.taskIds,
      tasks: result.tasks || [],
      unitName: foundData.unitName || "",
      grade: foundData.grade || "",
      message: result.message
    };
  }
  return { success: false, message: "データが見つかりません" };
}

/**
 * コンパスから受け取った単元計画データを Worksheets シートに取り込む。
 * 既に同じ taskId が存在する場合は、HTML未生成の場合のみ上書きする。
 *
 * → 書き込み先: Worksheets シート
 *
 * @param {Object} data - インポートデータ
 * @param {string} data.unitName - 単元名
 * @param {Object[]} data.tasks  - タスク配列（各要素に taskId, title を含む）
 * @return {Object} { taskIds: 追加されたID配列, message: 結果メッセージ }
 */
function handleImportUnitPlan(data) {
  var ss    = getDbSpreadsheet();
  var sheet = ss.getSheetByName('Worksheets');

  // 既存データのtaskId→行番号マップを作成
  var existingData = sheet.getDataRange().getValues();
  var idMap = {};
  for (var i = 1; i < existingData.length; i++) {
    idMap[String(existingData[i][0])] = i + 1;
  }

  var now       = new Date();
  var unitName  = data.unitName || "無題の単元";
  var addedTaskIds = [];
  var updates   = [];
  var inserts   = [];

  // 各タスクを処理
  data.tasks.forEach(function(task) {
    var taskId = String(task.taskId);
    addedTaskIds.push(taskId);

    var record = [
      taskId,                    // A列: タスクID
      unitName,                  // B列: 単元名
      task.title || "無題",       // C列: 活動タイトル
      "",                        // D列: HTML本文（未生成）
      now,                       // E列: 更新日時
      JSON.stringify(task),      // F列: 元データJSON
      "",                        // G列: キャンバス（空）
      "",                        // H列: ルーブリック（空）
      false                      // I列: 共有フラグ
    ];

    if (idMap.hasOwnProperty(taskId)) {
      // 既存のタスク: HTML未生成の場合のみ更新
      var existingRow = idMap[taskId];
      if (!existingData[existingRow - 1][3]) {
        updates.push({ row: existingRow, values: record });
      }
    } else {
      // 新規タスク: 追加リストに入れる
      inserts.push(record);
    }
  });

  // 新規タスクを一括挿入（1回の通信で複数行を書き込み）
  if (inserts.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, inserts.length, inserts[0].length)
      .setValues(inserts);
  }

  // 既存タスクの更新
  updates.forEach(function(u) {
    sheet.getRange(u.row, 1, 1, u.values.length).setValues([u.values]);
  });

  return {
    taskIds: addedTaskIds,
    tasks: data.tasks || [],
    message: inserts.length + '件追加、' + updates.length + '件更新'
  };
}

/**
 * みらいコンパスへ状態情報を同期送信する。
 * 児童がワークシートを操作した際に、コンパス側の「LiveStatus」を更新する。
 *
 * → 送信先: ユーザー設定で登録された compassUrl（POSTリクエスト）
 *
 * @param {Object} payload - 送信するデータ
 * @param {string} payload.studentId - 児童ID（必須）
 * @param {string} payload.action    - アクション種別
 * @return {Object} { success: true/false }
 */
function syncToCompass(payload) {
  var config     = getUserConfig();
  var compassUrl = config.compassUrl;

  // 送信条件: URL設定済み & ペイロードあり & 児童ID あり
  if (!compassUrl || !payload || !payload.studentId) {
    return { success: false };
  }

  try {
    var options = {
      method:             'post',
      contentType:        'application/json',
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    };
    UrlFetchApp.fetch(compassUrl, options);
    return { success: true };
  } catch (e) {
    console.error("Sync Error:", e);
    return { success: false, error: e.message };
  }
}
