/**
 * みらいパスポート v2.1.1
 * Update: Launch with Student ID/Name (Phase 2 Integration)
 */

const APP_NAME = "みらいパスポート";
const DB_NAME = APP_NAME + "_DB";

// ==================================================
// 1. エントリーポイント & 初期化
// ==================================================

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  
  // URLパラメータの取得（連携用）
  template.mode = e.parameter.mode || 'teacher'; // teacher | student
  template.taskId = e.parameter.taskId || '';
  
  // Phase 2: 自動ログイン用パラメータ
  template.studentId = e.parameter.studentId || '';     // Compass側で管理しているユニークID
  template.studentName = e.parameter.studentName || ''; // 児童名
  
  return template.evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
    if (files.hasNext()) {
      ss = SpreadsheetApp.openById(files.next().getId());
    } else {
      ss = SpreadsheetApp.create(DB_NAME);
    }

    // DB構造定義
    ensureSheet(ss, 'Worksheets', ['taskId', 'unitName', 'stepTitle', 'htmlContent', 'lastUpdated', 'jsonSource', 'canvasJson', 'rubricHtml', 'isShared']);
    ensureSheet(ss, 'Responses', ['responseId', 'taskId', 'studentId', 'studentName', 'submittedAt', 'canvasImage', 'textContent', 'status', 'feedbackText', 'score', 'feedbackJson', 'canvasJson', 'isPublic', 'reactions']);

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
// 2. 設定管理
// ==================================================

function saveUserConfig(apiKey, teacherName) {
  PropertiesService.getUserProperties().setProperties({ 
    'GEMINI_API_KEY': apiKey, 
    'TEACHER_NAME': teacherName 
  });
  return true;
}

function getUserConfig() {
  const userProps = PropertiesService.getUserProperties();
  const scriptProps = PropertiesService.getScriptProperties();
  return { 
    apiKey: userProps.getProperty('GEMINI_API_KEY') || scriptProps.getProperty('GEMINI_API_KEY') || '', 
    teacherName: userProps.getProperty('TEACHER_NAME') || '' 
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
  const getVal = c => sheet.getRange(r, c).getValue();

  return {
    taskId: getVal(1),
    unitName: getVal(2),
    stepTitle: getVal(3),
    htmlContent: getVal(4),
    jsonSource: safeJSONParse(getVal(6)),
    canvasJson: safeJSONParse(getVal(7)),
    rubricHtml: getVal(8),
    isShared: getVal(9)
  };
}

function getWorksheetsByIds(taskIds) {
  const sheet = getDbSpreadsheet().getSheetByName('Worksheets');
  const data = sheet.getDataRange().getValues();
  const results = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (taskIds.includes(String(row[0]))) { 
      results.push({
        taskId: row[0],
        unitName: row[1],
        stepTitle: row[2],
        htmlContent: row[3],
        canvasJson: safeJSONParse(row[6])
      });
    }
  }
  return results; 
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
    const newRow = [
      Utilities.getUuid(), data.taskId, data.studentId, data.studentName, new Date(), 
      data.canvasImage||"", data.textContent||"", data.status||"submitted", 
      "","", "", data.canvasJson||"", isPublicVal, "[]" 
    ];
    sheet.appendRow(newRow);
  }
  return { success: true };
}

function getTaskSubmissions(taskId) {
  const sheet = getDbSpreadsheet().getSheetByName('Responses');
  const values = sheet.getDataRange().getValues();
  return values
    .map((r, i) => ({ r, rowIndex: i + 1 }))
    .filter(o => o.rowIndex > 1 && String(o.r[1]) === String(taskId))
    .map(o => ({
      rowIndex: o.rowIndex,
      studentId: o.r[2],
      studentName: o.r[3],
      submittedAt: o.r[4],
      canvasImage: o.r[5],
      status: o.r[7],
      feedbackText: o.r[8],
      canvasJson: o.r[11]
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
  
  return values
    .map((r, i) => ({ r, rowIndex: i + 1 }))
    .filter(o => {
      const r = o.r;
      const isPublic = (r[12] === "" || r[12] === true || r[12] === "true");
      return String(r[1]) === String(taskId) && (r[7] === 'submitted' || r[7] === 'graded') && isPublic;
    })
    .map(o => ({ 
      responseId: o.r[0],
      studentId: o.r[2],
      studentName: o.r[3], 
      canvasImage: o.r[5],
      canvasJson: o.r[11],
      reactions: ensureArray(safeJSONParse(o.r[13])) 
    }));
}

function savePeerReaction(data) {
  const sheet = getDbSpreadsheet().getSheetByName('Responses');
  const finder = sheet.getRange("A:A").createTextFinder(data.targetResponseId).matchEntireCell(true).findNext();
  
  if(finder) {
    const row = finder.getRow();
    const cell = sheet.getRange(row, 14);
    let current = safeJSONParse(cell.getValue());
    if(!Array.isArray(current)) current = [];
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
        responseId: values[i][0], 
        status: values[i][7], 
        feedbackText: values[i][8], 
        feedbackJson: values[i][10], 
        canvasImage: values[i][5], 
        canvasJson: values[i][11],
        isPublic: values[i][12],
        reactions: ensureArray(safeJSONParse(values[i][13]))
      };
    }
  }
  return null;
}

// ==================================================
// 5. AI & Utilities
// ==================================================

function callGeminiAPI(prompt) {
  const k = getUserConfig().apiKey; 
  if (!k) throw new Error("APIキー未設定");
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${k}`;
  const res = UrlFetchApp.fetch(url, {method:'post',contentType:'application/json',payload:JSON.stringify({contents:[{parts:[{text:prompt}]}]},),muteHttpExceptions:true});
  const json = JSON.parse(res.getContentText());
  if (json.error) throw new Error(json.error.message);
  return json.candidates[0].content.parts[0].text;
}

function generateRubricAI(data) {
  return callGeminiAPI(`教育評価専門家としてルーブリック作成。単元:${data.unitName},活動:${data.stepTitle},内容:${data.description}。3観点3段階,HTMLテーブル形式(table table-bordered),具体的記述。HTMLのみ。`);
}

function getWebAppUrl(){ return ScriptApp.getService().getUrl(); }
function safeJSONParse(s){ try { return JSON.parse(s); } catch (e) { return null; } }
function ensureArray(val) { return Array.isArray(val) ? val : []; }
