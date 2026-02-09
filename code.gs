/**
 * みらいパスポート v2.2.0
 * Update: Phase 4 AI Coaching - Enhanced Prompt for Worksheet Generation
 * Update: Refined syncToCompass with Task Title metadata
 */

const APP_NAME = "みらいパスポート";
const DB_NAME = APP_NAME + "_DB";

// ==================================================
// 1. エントリーポイント & 初期化 (既存維持)
// ==================================================

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  template.mode = e.parameter.mode || 'teacher';
  template.taskId = e.parameter.taskId || '';
  template.studentId = e.parameter.studentId || '';     
  template.studentName = e.parameter.studentName || ''; 
  template.importId = e.parameter.importId || ''; 
  
  return template.evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://drive.google.com/uc?id=1mVwFtlrJvqEIk-0Gd03BqmG_-0BZiqY5&.png');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function checkSetupStatus() {
  const props = PropertiesService.getScriptProperties();
  const dbId = props.getProperty('DB_SS_ID');
  if (dbId) {
    try {
      SpreadsheetApp.openById(dbId);
      return { isSetup: true, dbId: dbId };
    } catch (e) {
      return { isSetup: false, dbId: null };
    }
  }
  return { isSetup: false, dbId: null };
}

function performInitialSetup() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  try {
    const files = DriveApp.getFilesByName(DB_NAME);
    if (files.hasNext()) ss = SpreadsheetApp.openById(files.next().getId());
    else ss = SpreadsheetApp.create(DB_NAME);

    ensureSheet(ss, 'Worksheets', ['taskId', 'unitName', 'stepTitle', 'htmlContent', 'lastUpdated', 'jsonSource', 'canvasJson', 'rubricHtml', 'isShared']);
    ensureSheet(ss, 'Responses', ['responseId', 'taskId', 'studentId', 'studentName', 'submittedAt', 'canvasImage', 'textContent', 'status', 'feedbackText', 'score', 'feedbackJson', 'canvasJson', 'isPublic', 'reactions']);
    ensureSheet(ss, 'ImportQueue', ['transactionId', 'dataJson', 'createdAt']);

    props.setProperty('DB_SS_ID', ss.getId());
    return { success: true, url: ss.getUrl() };
  } catch (e) {
    throw new Error("初期化エラー: " + e.message);
  }
}

function ensureSheet(ss, name, header) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(header);
  }
  return sheet;
}

// ==================================================
// 2. 設定管理 (既存維持)
// ==================================================

function saveUserConfig(apiKey, teacherName, compassUrl) {
  const props = { 'GEMINI_API_KEY': apiKey, 'TEACHER_NAME': teacherName };
  if (compassUrl !== undefined) props['COMPASS_URL'] = compassUrl.trim();
  PropertiesService.getUserProperties().setProperties(props);
  return true;
}

function getUserConfig() {
  const userProps = PropertiesService.getUserProperties();
  const scriptProps = PropertiesService.getScriptProperties();
  return { 
    apiKey: userProps.getProperty('GEMINI_API_KEY') || scriptProps.getProperty('GEMINI_API_KEY') || '', 
    teacherName: userProps.getProperty('TEACHER_NAME') || '',
    compassUrl: userProps.getProperty('COMPASS_URL') || scriptProps.getProperty('COMPASS_URL') || ''
  };
}

// ==================================================
// 3. データベース操作 (Core Logic)
// ==================================================

function getDbSpreadsheet() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_SS_ID');
  if (!id) throw new Error("DB未接続");
  return SpreadsheetApp.openById(id);
}

function saveWorksheetToDB(data) {
  const sheet = getDbSpreadsheet().getSheetByName('Worksheets');
  const now = new Date();
  const record = {
    taskId: String(data.taskId || Utilities.getUuid()),
    unitName: data.unitName || "無題",
    stepTitle: data.stepTitle || "無題",
    htmlContent: data.htmlContent || "",
    jsonSource: JSON.stringify(data.jsonSource || {}),
    canvasJson: data.canvasJson ? JSON.stringify(data.canvasJson) : "",
    rubricHtml: data.rubricHtml || "",
    isShared: data.isShared || false
  };
  const found = sheet.getRange("A:A").createTextFinder(record.taskId).matchEntireCell(true).findNext();
  if (found) {
    const r = found.getRow();
    sheet.getRange(r, 4).setValue(record.htmlContent);
    sheet.getRange(r, 5).setValue(now);
    sheet.getRange(r, 7).setValue(record.canvasJson);
    sheet.getRange(r, 8).setValue(record.rubricHtml);
    sheet.getRange(r, 9).setValue(record.isShared);
  } else {
    sheet.appendRow([record.taskId, record.unitName, record.stepTitle, record.htmlContent, now, record.jsonSource, record.canvasJson, record.rubricHtml, record.isShared]);
  }
  return true;
}

function loadWorksheetFromDB(taskId) {
  const sheet = getDbSpreadsheet().getSheetByName('Worksheets');
  const found = sheet.getRange("A:A").createTextFinder(String(taskId)).matchEntireCell(true).findNext();
  if (!found) return null;
  const r = found.getRow();
  return {
    taskId: sheet.getRange(r, 1).getValue(),
    unitName: sheet.getRange(r, 2).getValue(),
    stepTitle: sheet.getRange(r, 3).getValue(),
    htmlContent: sheet.getRange(r, 4).getValue(),
    jsonSource: safeJSONParse(sheet.getRange(r, 6).getValue()),
    canvasJson: safeJSONParse(sheet.getRange(r, 7).getValue()),
    rubricHtml: sheet.getRange(r, 8).getValue(),
    isShared: sheet.getRange(r, 9).getValue()
  };
}

function getWorksheetsByIds(taskIds) {
  const sheet = getDbSpreadsheet().getSheetByName('Worksheets');
  const data = sheet.getDataRange().getValues();
  return data.slice(1)
    .filter(row => taskIds.includes(String(row[0])))
    .map(row => ({
      taskId: row[0], unitName: row[1], stepTitle: row[2], htmlContent: row[3],
      canvasJson: safeJSONParse(row[6]), jsonSource: safeJSONParse(row[5])
    }));
}

function getHistory() {
  const sheet = getDbSpreadsheet().getSheetByName('Worksheets');
  if(sheet.getLastRow()<2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 5).getValues();
  return data
    .map(r=>({id:r[0], title:r[2]||"無題", timestamp:new Date(r[4]).getTime()}))
    .filter(i=>i.id)
    .sort((a,b)=>b.timestamp-a.timestamp)
    .slice(0, 30);
}

// ==================================================
// 4. 児童レスポンス管理
// ==================================================

function saveStudentResponse(data) {
  const sheet = getDbSpreadsheet().getSheetByName('Responses');
  const vals = sheet.getDataRange().getValues();
  let row = -1;
  for(let i=1; i<vals.length; i++){
    if(String(vals[i][1]) === String(data.taskId) && String(vals[i][2]) === String(data.studentId)){
      row = i+1; break;
    }
  }
  const isPublicVal = (data.isPublic === undefined) ? true : data.isPublic;
  if(row > 0) {
    sheet.getRange(row, 4).setValue(data.studentName);
    sheet.getRange(row, 5).setValue(new Date());
    sheet.getRange(row, 6).setValue(data.canvasImage);
    sheet.getRange(row, 7).setValue(data.textContent);
    sheet.getRange(row, 8).setValue(data.status);
    if(data.canvasJson) sheet.getRange(row, 12).setValue(data.canvasJson);
    sheet.getRange(row, 13).setValue(isPublicVal);
  } else {
    sheet.appendRow([Utilities.getUuid(), data.taskId, data.studentId, data.studentName, new Date(), data.canvasImage||"", data.textContent||"", data.status||"submitted", "","", "", data.canvasJson||"", isPublicVal, "[]"]);
  }
  return { success: true };
}

function getTaskSubmissions(taskId) {
  const sheet = getDbSpreadsheet().getSheetByName('Responses');
  const values = sheet.getDataRange().getValues();
  return values.map((r, i) => ({ r, rowIndex: i + 1 }))
    .filter(o => o.rowIndex > 1 && String(o.r[1]) === String(taskId))
    .map(o => ({
      rowIndex: o.rowIndex, studentId: o.r[2], studentName: o.r[3], submittedAt: o.r[4],
      canvasImage: o.r[5], textContent: o.r[6], status: o.r[7], feedbackText: o.r[8], canvasJson: o.r[11]
    }));
}

function saveFeedback(data) {
  const sheet = getDbSpreadsheet().getSheetByName('Responses');
  if(data.rowIndex) {
    sheet.getRange(data.rowIndex, 8).setValue("graded");
    sheet.getRange(data.rowIndex, 9).setValue(data.feedbackText);
    if(data.canvasJson) sheet.getRange(data.rowIndex, 12).setValue(data.canvasJson);
  }
  return { success: true };
}

function getSharedResponses(taskId) {
  const sheet = getDbSpreadsheet().getSheetByName('Responses');
  const values = sheet.getDataRange().getValues();
  return values.map((r, i) => ({ r, rowIndex: i + 1 }))
    .filter(o => {
      const isPublic = (o.r[12] === "" || o.r[12] === true || o.r[12] === "true");
      return String(o.r[1]) === String(taskId) && (o.r[7] === 'submitted' || o.r[7] === 'graded') && isPublic;
    })
    .map(o => ({ 
      responseId: o.r[0], studentId: o.r[2], studentName: o.r[3], 
      canvasImage: o.r[5], canvasJson: o.r[11], reactions: ensureArray(safeJSONParse(o.r[13])) 
    }));
}

function savePeerReaction(data) {
  const sheet = getDbSpreadsheet().getSheetByName('Responses');
  const finder = sheet.getRange("A:A").createTextFinder(data.targetResponseId).matchEntireCell(true).findNext();
  if(finder) {
    const row = finder.getRow();
    const cell = sheet.getRange(row, 14);
    let current = ensureArray(safeJSONParse(cell.getValue()));
    current.push({ ...data.reaction, timestamp: new Date().getTime() });
    cell.setValue(JSON.stringify(current));
    return { success: true, reactions: current };
  }
  return { success: false };
}

function getMyResponse(taskId, studentId) {
  const sheet = getDbSpreadsheet().getSheetByName('Responses');
  const values = sheet.getDataRange().getValues();
  for(let i=values.length-1; i>=1; i--){
    if(String(values[i][1]) === String(taskId) && String(values[i][2]) === String(studentId)) {
      return { 
        responseId: values[i][0], status: values[i][7], feedbackText: values[i][8], 
        canvasImage: values[i][5], canvasJson: values[i][11], isPublic: values[i][12],
        reactions: ensureArray(safeJSONParse(values[i][13]))
      };
    }
  }
  return null;
}

// ==================================================
// 5. AI & Utilities (Phase 4: Enhanced Prompts)
// ==================================================

function callGeminiAPI(prompt) {
  const k = getUserConfig().apiKey; 
  if (!k) throw new Error("Gemini APIキーが設定されていません。先生モードの設定を確認してください。");
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${k}`;
  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
  
  const res = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(res.getContentText());
  if (json.error) throw new Error("AIエラー: " + json.error.message);
  if (!json.candidates || !json.candidates[0].content) throw new Error("AIから応答が得られませんでした。");
  return json.candidates[0].content.parts[0].text;
}

/**
 * [Update Phase 4] AIコーチング対応のワークシート生成
 * 単なる入力欄だけでなく、ヒントや自己評価を動的に追加
 */
function generateSingleWorksheet(data) {
  const prompt = `あなたは教育工学と個別最適な学びの専門家です。
児童が自立的に学習を進められるよう、以下の活動内容に基づいた「AIコーチング機能付きHTMLワークシート」の本文を作成してください。

【活動内容】
単元: ${data.unitName}
活動: ${data.stepTitle}
内容: ${data.description}

【デザインの要件】
1. 小学生が親しみやすい言葉遣い。
2. 以下のセクションを必ず含める：
   - 「今日のめあて」（活動内容から具体化）
   - 「AIヒント・ポイント」（この活動でつまずきやすい点や、考えるコツをAIコーチとして助言）
   - 「考えを書くスペース」（<div contenteditable="true" class="rich-editor"></div> を使用）
   - 「自己評価」（3段階のスタンプ選択など）
3. スタイルはBootstrap 5のクラス（card, p-3, mb-3, bg-lightなど）を活用。
4. HTMLの「body内部」のみを出力すること。余計な解説や\`\`\`htmlタグは不要。`;

  return callGeminiAPI(prompt);
}

function generateRubricAI(data) {
  const prompt = `あなたは教育評価の専門家です。以下の授業活動に対するルーブリック（評価基準表）を作成してください。

【単元名】${data.unitName || ""}
【活動タイトル】${data.stepTitle || ""}
【活動内容】${data.description || ""}

【出力ルール】
1. 3観点（知識・技能、思考・判断・表現、主体的に学習に取り組む態度）
2. 各観点3段階（A:十分満足できる、B:おおむね満足できる、C:努力を要する）
3. 各セルに具体的な児童の姿を簡潔に記述
4. HTMLテーブル形式（class="table table-bordered"）で出力
5. HTMLのみ出力（説明文不要）`;
  return callGeminiAPI(prompt);
}

function getWebAppUrl(){ return ScriptApp.getService().getUrl(); }
function safeJSONParse(s){ try { return JSON.parse(s); } catch (e) { return null; } }
function ensureArray(val) { return Array.isArray(val) ? val : []; }

// ==================================================
// 5b. ワークシート+ルーブリック同時生成 & AI添削 & クラス分析
// ==================================================

function generateWorksheetWithRubric(data) {
  const htmlContent = generateSingleWorksheet(data);
  const rubricHtml = generateRubricAI(data);
  return { htmlContent: htmlContent, rubricHtml: rubricHtml };
}

function generateFeedbackAI(data) {
  const prompt = `あなたは小学校の先生のアシスタントです。以下の情報を元に、児童への添削コメントを作成してください。

【ワークシートのタイトル】
${data.worksheetTitle || ""}

【評価基準（ルーブリック）】
${data.rubricHtml || "（なし）"}

【児童の自己評価】
${data.selfEvaluation || "（なし）"}

【児童名】${data.studentName || ""}

【出力ルール】
1. 児童の良い点を具体的に1〜2点褒める（自己評価の内容も参考にする）
2. 改善点やアドバイスを1点、前向きな言い方で伝える
3. 次の学習への意欲が湧くように励ます
4. 小学生にわかりやすい優しい言葉で、3〜5文程度にまとめる
5. テキストのみ出力（HTML不要）`;
  return callGeminiAPI(prompt);
}

function getClassAnalytics(taskId) {
  const sheet = getDbSpreadsheet().getSheetByName('Responses');
  const values = sheet.getDataRange().getValues();
  const responses = values.filter((r, i) => i > 0 && String(r[1]) === String(taskId));

  const total = responses.length;
  const submitted = responses.filter(r => r[7] === 'submitted').length;
  const graded = responses.filter(r => r[7] === 'graded').length;
  const draft = responses.filter(r => r[7] === 'draft').length;

  // 自己評価パース
  const evalCounts = {};
  ['わかった', '考えた', '進んで'].forEach(label => {
    evalCounts[label] = { '◎': 0, '◯': 0, '△': 0 };
  });
  responses.forEach(r => {
    const text = String(r[6] || "");
    const regex = /\[(.+?): (.+?)\]/g;
    let m;
    while ((m = regex.exec(text)) !== null) {
      const label = m[1].trim();
      const val = m[2].trim();
      if (evalCounts[label] && evalCounts[label][val] !== undefined) {
        evalCounts[label][val]++;
      }
    }
  });

  // 児童一覧
  const students = responses.map(r => ({
    studentId: r[2],
    studentName: r[3],
    status: r[7],
    submittedAt: r[4] ? new Date(r[4]).getTime() : null,
    hasFeedback: !!r[8]
  }));

  return { total, submitted, graded, draft, evalCounts, students };
}

// ==================================================
// 6. Compass Integration (Shared DB Mode & Sync)
// ==================================================

function consumeImportQueue(importId) {
  if (!importId) return { success: false, message: "ID未指定" };
  const ss = getDbSpreadsheet();
  const sheet = ss.getSheetByName('ImportQueue');
  if (!sheet) return { success: false, message: "連携シートなし" };
  const data = sheet.getDataRange().getValues();
  let foundData = null, deleteRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(importId)) {
      foundData = safeJSONParse(data[i][1]);
      deleteRowIndex = i + 1; break;
    }
  }
  if (deleteRowIndex > 0 && foundData) {
    const result = handleImportUnitPlan(foundData);
    sheet.deleteRow(deleteRowIndex);
    return { success: true, ...result };
  }
  return { success: false, message: "データが見つかりません" };
}

function handleImportUnitPlan(data) {
  const ss = getDbSpreadsheet();
  const sheet = ss.getSheetByName('Worksheets');
  const existingData = sheet.getDataRange().getValues();
  const idMap = new Map();
  for (let i = 1; i < existingData.length; i++) idMap.set(String(existingData[i][0]), i + 1);
  const now = new Date(), unitName = data.unitName || "無題の単元", addedTaskIds = [];
  const updates = [], inserts = [];
  data.tasks.forEach(task => {
    const taskId = String(task.taskId);
    addedTaskIds.push(taskId);
    const record = [taskId, unitName, task.title || "無題", "", now, JSON.stringify(task), "", "", false];
    if (idMap.has(taskId)) {
      const row = idMap.get(taskId);
      if (!existingData[row - 1][3]) updates.push({ row: row, values: record });
    } else inserts.push(record);
  });
  if (inserts.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, inserts.length, inserts[0].length).setValues(inserts);
  updates.forEach(u => sheet.getRange(u.row, 1, 1, u.values.length).setValues([u.values]));
  return { taskIds: addedTaskIds, message: `${inserts.length}件追加、${updates.length}件更新` };
}

/**
 * [Update Phase 3/4] コンパスへの状態同期送信
 * コンパス側の「LiveStatus」を更新するためのメタデータを送信
 */
function syncToCompass(payload) {
  const config = getUserConfig();
  const compassUrl = config.compassUrl;
  if (!compassUrl || !payload || !payload.studentId) return { success: false };

  try {
    // 連携用の詳細情報を付与（Passport側で持っている活動タイトルなど）
    const options = {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    UrlFetchApp.fetch(compassUrl, options);
    return { success: true };
  } catch (e) {
    console.error("Sync Error:", e);
    return { success: false, error: e.message };
  }
}
