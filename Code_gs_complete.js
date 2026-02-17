// ========== 設定セクション ==========
const SPREADSHEET_ID = '1znspyaI-wj70aBkDfOPPrmYjPSSn3mFELW6VNfOSIbM';
const RESPONSE_SHEET_NAME = '回答データ1';  // 回答データ用（フォーム＋照合アプリ）
const HOSPITAL_LIST_SHEET_NAME = '医療機関リスト';  // 医療機関リスト用
const HOSPITAL_QUESTION_TITLE = '1. 受診中の歯科医院を選んでください。';
const JSON_FOLDER_NAME = 'アンケートJSON';  // JSON保存フォルダ名
const JSON_CREATED_COLUMN_NAME = 'JSON作成済み';  // フラグ列の名前
const JSON_PARENT_FOLDER_ID = '1tAUwyUb9B-WH5LrW45-ox4GZFrPJWMnB';  // 共有ドライブ「3.データシート」フォルダ
const VIEW_PASSWORD = 'touka2026';  // データ閲覧用パスワード
const SCAN_PDF_FOLDER_ID = '1blx2Ia2X9blWcYAkGKPGST8MFN3VgJGL';  // 元PDFフォルダ
const DONE_PDF_FOLDER_ID = '1iFx7ngwNSJo80bElYcqw0dyHQyhm0N0N';  // 入力済みPDFフォルダ
const OCR_RESULTS_FOLDER_NAME = 'OCR照合結果';  // OCR照合結果フォルダ名
const OCR_LOCK_TIMEOUT_MIN = 30;  // ファイルロックのタイムアウト（分）
const SPECIMEN_SPREADSHEET_ID = '1y4z8dZjKuJKOS0HUxmnSSfkkvhRWkQpb--TXGm5amvs';  // 検体PpP値シート
const SPECIMEN_SHEET_NAME = 'シート1';
const TRACKING_SPREADSHEET_ID = '1jr25zPPv2qCHkgXNA6oHV5eoSq0RcxpXEirWq0V5YA8';  // HbA1c追跡シート
const TRACKING_SHEET_NAME = 'シート1';
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
// ========== Web App エンドポイント（JSONP対応） ==========
function doGet(e) {
  const action = e.parameter.action;
  let result;
  try {
    switch (action) {
      case 'getHospitalList':
        result = { success: true, data: getHospitalList() };
        break;
      case 'getFormInfo':
        result = { success: true, data: getFormInfo() };
        break;
      case 'addHospital':
        result = { success: true, data: addHospital(e.parameter.name) };
        break;
      case 'deleteHospital':
        result = { success: true, data: deleteHospital(e.parameter.name) };
        break;
      case 'updateHospital':
        result = { success: true, data: updateHospital(e.parameter.oldName, e.parameter.newName) };
        break;
      case 'addSurveyResponse':
        const jsonStr = decodeURIComponent(e.parameter.data);
        const responseData = JSON.parse(jsonStr);
        result = { success: true, data: addSurveyResponse(responseData) };
        break;
      case 'deleteDataByNo':
        const fromNo = parseInt(e.parameter.fromNo, 10);
        const toNo = parseInt(e.parameter.toNo, 10);
        result = { success: true, data: deleteDataByNo(fromNo, toNo) };
        break;
      case 'getNoList':
        result = { success: true, data: getNoList() };
        break;
      case 'updatePdfFilename':
        const targetNo = parseInt(e.parameter.no, 10);
        const pdfName = decodeURIComponent(e.parameter.pdfFilename);
        result = { success: true, data: updatePdfFilename(targetNo, pdfName) };
        break;
      case 'updateSourceToOCR':
        const sourceNo = parseInt(e.parameter.no, 10);
        result = { success: true, data: updateSourceToOCR(sourceNo) };
        break;
      case 'viewData':
        if (e.parameter.password === VIEW_PASSWORD) {
          result = { success: true, data: getViewData() };
        } else {
          result = { success: false, error: 'パスワードが正しくありません' };
        }
        break;
      case 'getUnprocessedPdfs':
        result = { success: true, data: getUnprocessedPdfs() };
        break;
      case 'getDonePdfIndex':
        result = { success: true, data: getDonePdfIndex() };
        break;
      case 'getJsonFileList':
        result = { success: true, data: getJsonFileList() };
        break;
      case 'renameJsonFiles':
        result = { success: true, data: renameJsonFiles(e.parameter.pattern, e.parameter.replacement || '') };
        break;
      case 'trashJsonFiles':
        result = { success: true, data: trashJsonFilesByPattern(e.parameter.pattern) };
        break;
      case 'addSpecimen':
        result = { success: true, data: addSpecimen(e.parameter) };
        break;
      case 'getSpecimenList':
        result = { success: true, data: getSpecimenList(e.parameter.no || '') };
        break;
      case 'addTracking':
        result = { success: true, data: addTracking(e.parameter) };
        break;
      case 'getTrackingList':
        result = { success: true, data: getTrackingList(e.parameter.no || '') };
        break;
      case 'addTrackingColumn':
        result = { success: true, data: addTrackingColumn(e.parameter.columnName || '') };
        break;
      // ========== PDF紐づけツール（Google Drive連携） ==========
      case 'getDrivePdfList':
        result = { success: true, data: getDrivePdfListWithIndex() };
        break;
      case 'addPdfIndexEntry':
        result = { success: true, data: addPdfIndexEntry(e.parameter) };
        break;
      case 'indexPdfWithGemini':
        result = { success: true, data: indexPdfWithGemini(e.parameter.fileId) };
        break;
      case 'indexAllPdfs':
        result = { success: true, data: indexAllPdfs() };
        break;
      case 'linkPdfFromDrive':
        result = { success: true, data: linkPdfFromDrive(e.parameter) };
        break;
      // ========== OCR照合（共有版） ==========
      case 'getOcrResultList':
        result = { success: true, data: getOcrResultList() };
        break;
      case 'getOcrResultDetail':
        result = { success: true, data: getOcrResultDetail(e.parameter.fileId) };
        break;
      case 'getOcrResultText':
        result = { success: true, data: getOcrResultText(e.parameter.fileId) };
        break;
      case 'getOcrResultImage':
        result = { success: true, data: getOcrResultImage(e.parameter.fileId, e.parameter.questionKey) };
        break;
      case 'lockOcrResult':
        result = { success: true, data: lockOcrResult(e.parameter.fileId, e.parameter.user) };
        break;
      case 'unlockOcrResult':
        result = { success: true, data: unlockOcrResult(e.parameter.fileId) };
        break;
      // ========== 検体番号検索（氏名・生年月日） ==========
      case 'searchByNameBirthdate':
        if (e.parameter.password === VIEW_PASSWORD) {
          result = { success: true, data: searchByNameBirthdate(e.parameter) };
        } else {
          result = { success: false, error: 'パスワードが正しくありません' };
        }
        break;
      default:
        result = { success: false, error: 'Unknown action' };
    }
  } catch (error) {
    result = { success: false, error: error.toString() };
  }
  // JSONP形式で返す（CORS回避）
  const callback = e.parameter.callback;
  const jsonOutput = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + jsonOutput + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } else {
    return ContentService.createTextOutput(jsonOutput)
      .setMimeType(ContentService.MimeType.JSON);
  }
}
// ========== 医療機関リスト取得 ==========
function getHospitalList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(HOSPITAL_LIST_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    return [];
  }
  const data = sheet.getRange(1, 1, lastRow, 1).getValues();
  let list = data.map(row => String(row[0]).trim()).filter(name => name !== '');
  return [...new Set(list)];
}
// ========== フォーム情報取得 ==========
function getFormInfo() {
  const form = FormApp.getActiveForm();
  return {
    formUrl: form.getPublishedUrl(),
    entryId: getHospitalQuestionEntryId()
  };
}
// ========== 医療機関質問のEntry ID取得 ==========
// ※ form.createResponse()でプリフィルURLを生成し、正しいentry.XXXを抽出する
function getHospitalQuestionEntryId() {
  const form = FormApp.getActiveForm();
  // ★ ListItem（プルダウン）を検索 ★
  const items = form.getItems(FormApp.ItemType.LIST);
  for (let i = 0; i < items.length; i++) {
    if (items[i].getTitle().includes('受診中の歯科医院')) {
      try {
        // プリフィルURLを生成してEntry IDを抽出
        const listItem = items[i].asListItem();
        const choices = listItem.getChoices();
        if (choices.length > 0) {
          const response = listItem.createResponse(choices[0].getValue());
          const prefillUrl = form.createResponse()
            .withItemResponse(response)
            .toPrefilledUrl();
          // URLから entry.XXXX を抽出
          const match = prefillUrl.match(/entry\.(\d+)=/);
          if (match) {
            Logger.log('プリフィルURLからEntry ID取得: ' + match[1]);
            return match[1];
          }
        }
      } catch (e) {
        Logger.log('プリフィルURL方式でのEntry ID取得失敗: ' + e.toString());
      }
      // フォールバック: Item IDを返す
      return items[i].getId();
    }
  }
  return null;
}
// ========== 医療機関を追加 ==========
function addHospital(name) {
  if (!name || name.trim() === '') {
    throw new Error('医療機関名が空です');
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(HOSPITAL_LIST_SHEET_NAME);
  const existingList = getHospitalList();
  if (existingList.includes(name.trim())) {
    throw new Error('この医療機関名は既に存在します');
  }
  sheet.appendRow([name.trim()]);
  syncHospitalListOnly();
  return { added: name.trim(), list: getHospitalList() };
}
// ========== 医療機関を削除 ==========
function deleteHospital(name) {
  if (!name) {
    throw new Error('医療機関名が指定されていません');
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(HOSPITAL_LIST_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  for (let i = 1; i <= lastRow; i++) {
    const cellValue = sheet.getRange(i, 1).getValue();
    if (cellValue === name) {
      sheet.deleteRow(i);
      syncHospitalListOnly();
      return { deleted: name, list: getHospitalList() };
    }
  }
  throw new Error('指定された医療機関名が見つかりません');
}
// ========== 医療機関名を更新 ==========
function updateHospital(oldName, newName) {
  if (!oldName || !newName) {
    throw new Error('医療機関名が指定されていません');
  }
  if (newName.trim() === '') {
    throw new Error('新しい医療機関名が空です');
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(HOSPITAL_LIST_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const existingList = getHospitalList();
  if (oldName !== newName.trim() && existingList.includes(newName.trim())) {
    throw new Error('この医療機関名は既に存在します');
  }
  for (let i = 1; i <= lastRow; i++) {
    const cellValue = sheet.getRange(i, 1).getValue();
    if (cellValue === oldName) {
      sheet.getRange(i, 1).setValue(newName.trim());
      syncHospitalListOnly();
      return { oldName: oldName, newName: newName.trim(), list: getHospitalList() };
    }
  }
  throw new Error('指定された医療機関名が見つかりません');
}
// ========== Googleフォームの医療機関リストを同期 ==========
function syncHospitalListOnly() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(HOSPITAL_LIST_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  let hList = [];
  if (lastRow >= 1) {
    hList = sheet.getRange(1, 1, lastRow, 1).getValues()
      .map(r => String(r[0]).trim())
      .filter(n => n !== "");
    hList = [...new Set(hList)];
  }
  const form = FormApp.getActiveForm();
  // ★ ListItem（プルダウン）を検索 ★
  const items = form.getItems(FormApp.ItemType.LIST);
  for (let i = 0; i < items.length; i++) {
    if (items[i].getTitle() === HOSPITAL_QUESTION_TITLE) {
      const q = items[i].asListItem();
      q.setChoiceValues(hList);
      break;
    }
  }
}
// ========== フォーム全体を再構築 ==========
function rebuildFullForm() {
  const form = FormApp.getActiveForm();
  // --- 1. 既存の質問をすべて削除 ---
  const items = form.getItems();
  items.forEach(item => form.deleteItem(item));
  // --- 2. 基本設定 ---
  form.setTitle('糖化アンケート_v9_3')
      .setDescription('生活習慣に関するアンケートにご協力ください。')
      .setCollectEmail(true);
  // --- 3. セクションと質問の作成 ---
  // セクション1: 医療機関選択
  const sec1 = form.addPageBreakItem().setTitle('医療機関名');
  // ★ ListItem（プルダウン）を使用 - プリフィルURLに対応 ★
  const q1 = form.addListItem().setTitle(HOSPITAL_QUESTION_TITLE).setRequired(true);
  // セクション2: 基本情報
  const sec2 = form.addPageBreakItem().setTitle('基本情報');
  // ID番号（Q2）
  const valId = FormApp.createTextValidation()
    .requireTextMatchesPattern('^[a-zA-Z0-9+_\\-]+$')
    .setHelpText('半角英数字・記号（+ _ -）で入力してください。')
    .build();
  const q2 = form.addTextItem().setTitle('2. ID番号').setRequired(true).setValidation(valId);
  // 名前（Q3）- カタカナのみ
  const valKatakana = FormApp.createTextValidation()
    .requireTextMatchesPattern('^[ァ-ヶー　 ]+$')
    .setHelpText('カタカナで入力してください。')
    .build();
  const q3 = form.addTextItem().setTitle('3. 名前（カタカナ）').setRequired(true).setValidation(valKatakana);
  // 生年月日（Q4-Q7）- 4つの質問に分割
  const q4 = form.addListItem()
    .setTitle('4. 生年月日 - 年号')
    .setChoiceValues(['昭和', '平成', '令和'])
    .setRequired(true);
  const valYear = FormApp.createTextValidation()
    .requireNumber()
    .setHelpText('半角数字で入力してください。')
    .build();
  const q5 = form.addTextItem()
    .setTitle('5. - 年（半角数字）')
    .setRequired(true)
    .setValidation(valYear);
  const valMonth = FormApp.createTextValidation()
    .requireNumberBetween(1, 12)
    .setHelpText('1～12の半角数字で入力してください。')
    .build();
  const q6 = form.addTextItem()
    .setTitle('6. - 月（半角数字）')
    .setRequired(true)
    .setValidation(valMonth);
  const valDay = FormApp.createTextValidation()
    .requireNumberBetween(1, 31)
    .setHelpText('1～31の半角数字で入力してください。')
    .build();
  const q7 = form.addTextItem()
    .setTitle('7. - 日（半角数字）')
    .setRequired(true)
    .setValidation(valDay);
  // セクション3: 患者情報（性別・血液型・身長・体重）
  const sec3 = form.addPageBreakItem().setTitle('患者情報');
  // 性別（Q9）
  const q9 = form.addMultipleChoiceItem()
    .setTitle('9. 性別')
    .setChoiceValues(['男性', '女性', 'その他', '回答しない'])
    .setRequired(true);
  // 血液型（Q10）
  const q10 = form.addMultipleChoiceItem()
    .setTitle('10. 血液型')
    .setChoiceValues(['A型', 'B型', 'O型', 'AB型', 'わからない'])
    .setRequired(true);
  // 身長・体重（Q11,12）
  const valNum = FormApp.createTextValidation().requireNumber().build();
  const q11 = form.addTextItem().setTitle('11. 身長 cm').setRequired(true).setValidation(valNum);
  const q12 = form.addTextItem().setTitle('12. 体重 kg').setRequired(true).setValidation(valNum);
  // --- 糖尿病・疾患系セクション ---
  const sec5 = form.addPageBreakItem().setTitle('糖尿病について');
  const q13 = form.addMultipleChoiceItem().setTitle('13. 糖尿病と診断されていますか？').setRequired(true);
  const sec6 = form.addPageBreakItem().setTitle('糖尿病の期間');
  const q14 = form.addMultipleChoiceItem().setTitle('14. 何年前からですか？').setChoiceValues(['3年以内', '10年以内', 'もっと以前', 'わからない']).setRequired(true);
  const sec7 = form.addPageBreakItem().setTitle('脂質異常症について');
  const q15 = form.addMultipleChoiceItem().setTitle('15. 脂質異常症と診断されていますか？').setRequired(true);
  const sec8 = form.addPageBreakItem().setTitle('脂質異常症の期間');
  const q16 = form.addMultipleChoiceItem().setTitle('16. 何年前からですか？').setChoiceValues(['3年以内', '3～10年以内', '10年以上前', 'わからない']).setRequired(true);
  const sec9 = form.addPageBreakItem().setTitle('ご兄弟の糖尿病歴');
  const q17 = form.addMultipleChoiceItem().setTitle('17. ご兄弟に糖尿病歴はありますか？').setRequired(true);
  // ※ Q18（ご兄弟の糖尿病の期間）は削除済み
  const sec11 = form.addPageBreakItem().setTitle('ご両親の糖尿病歴');
  const q19 = form.addMultipleChoiceItem().setTitle('19. ご両親に糖尿病歴はありますか？').setRequired(true);
  // ※ Q20（ご両親の糖尿病の期間）は削除済み
  // --- 生活習慣セクション ---
  const sec13 = form.addPageBreakItem().setTitle('生活習慣について');
  form.addMultipleChoiceItem().setTitle('21. 普段、運動をしてますか？').setChoiceValues(['ほぼ毎日', '週2～3回', '週1回以下', 'しない']).setRequired(true);
  form.addCheckboxItem().setTitle('22. 普段、飲む物は何ですか？').setChoiceValues(['有糖飲料(ジュース、炭酸飲料、スポーツドリンク、加糖コーヒーなど)', '無糖飲料(お茶、水、炭酸水、無糖コーヒーなど)']).setRequired(true);
  form.addMultipleChoiceItem().setTitle('23. 普段、お菓子、スイーツなどは食べますか？').setChoiceValues(['ほぼ毎日', '週2～3回', '週1回以下', '食べない']).setRequired(true);
  const q24 = form.addMultipleChoiceItem().setTitle('24. お酒（ビール、ワイン、焼酎、ウイスキーなど）を習慣的に飲みますか？').setRequired(true);
  // --- 飲酒詳細セクション生成関数 ---
  function createDrinkingSec(title, baseNum) {
    const sec = form.addPageBreakItem().setTitle(title);
    form.addListItem().setTitle(baseNum + '. お酒の種類').setChoiceValues(['ビール', '日本酒', '焼酎', 'チューハイ', 'ワイン', 'ウイスキー', 'ブランデー', '梅酒', '泡盛']).setRequired(true);
    form.addListItem().setTitle((baseNum + 1) + '. 週に何回飲みますか？(回)').setChoiceValues(['1', '2', '3', '4', '5', '6', '7']).setRequired(true);
    form.addListItem().setTitle((baseNum + 2) + '. 1回あたりのサイズ/飲み方は？').setChoiceValues(['350ml缶（缶ビール普通サイズ）', '500ml缶（缶ビール大サイズ）', '750mlビン（ワイン普通サイズ）', '375mlビン（ワインハーフボトル）', 'コップ', '水割り', 'お湯割り', 'ロック', '小ジョッキ', '中ジョッキ', '大ジョッキ']).setRequired(true);
    form.addListItem().setTitle((baseNum + 3) + '. 数量').setChoiceValues(['1', '2', '3', '4', '5', '6']).setRequired(true);
    const qNext = form.addMultipleChoiceItem().setTitle((baseNum + 4) + '. 他にもよく飲むお酒はありますか？').setRequired(true);
    return { section: sec, nextQ: qNext };
  }
  const drink1 = createDrinkingSec('飲酒の詳細【回答1】', 25);
  const drink2 = createDrinkingSec('飲酒の詳細【回答2】', 30);
  const drink3 = createDrinkingSec('飲酒の詳細【回答3】', 35);
  // --- 4. 病院リスト取得と紐付け ---
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(HOSPITAL_LIST_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  let hList = [];
  if (lastRow >= 1) {
    hList = sheet.getRange(1, 1, lastRow, 1).getValues()
      .map(r => String(r[0]).trim())
      .filter(n => n !== "");
    hList = [...new Set(hList)];
  }
  Logger.log("医療機関リスト: " + JSON.stringify(hList));
  if (hList.length > 0) {
    q1.setChoiceValues(hList);
  }
  // --- 5. 条件分岐の設定 ---
  q13.setChoices([q13.createChoice('はい', sec6), q13.createChoice('いいえ', sec7)]);
  q15.setChoices([q15.createChoice('はい', sec8), q15.createChoice('いいえ', sec9)]);
  q17.setChoices([q17.createChoice('はい', sec11), q17.createChoice('いいえ', sec11)]);  // Q18削除のため、はい/いいえ両方sec11へ
  q19.setChoices([q19.createChoice('はい', sec13), q19.createChoice('いいえ', sec13)]);  // Q20削除のため、はい/いいえ両方sec13へ
  sec6.setGoToPage(sec7);
  sec8.setGoToPage(sec9);
  q24.setChoices([
    q24.createChoice('はい（習慣的に飲む）', drink1.section),
    q24.createChoice('いいえ（ほとんど飲まない）', FormApp.PageNavigationType.SUBMIT)
  ]);
  drink1.nextQ.setChoices([
    drink1.nextQ.createChoice('ある', drink2.section),
    drink1.nextQ.createChoice('ない', FormApp.PageNavigationType.SUBMIT)
  ]);
  drink2.nextQ.setChoices([
    drink2.nextQ.createChoice('ある', drink3.section),
    drink2.nextQ.createChoice('ない', FormApp.PageNavigationType.SUBMIT)
  ]);
  drink3.nextQ.setChoices([
    drink3.nextQ.createChoice('ある', FormApp.PageNavigationType.SUBMIT),
    drink3.nextQ.createChoice('ない', FormApp.PageNavigationType.SUBMIT)
  ]);
  Logger.log("再構築完了！（Q18・Q20削除版）");
}
// ========== デバッグ・テスト用関数 ==========
function getFormEntryIds() {
  const form = FormApp.getActiveForm();
  Logger.log('=== フォーム情報 ===');
  Logger.log('公開URL: ' + form.getPublishedUrl());
  Logger.log('');
  Logger.log('=== 全質問のEntry ID ===');
  const items = form.getItems();
  items.forEach(item => {
    Logger.log('Title: "' + item.getTitle() + '" | ID: ' + item.getId() + ' | Type: ' + item.getType());
  });
  const hospitalQ = items.find(item => item.getTitle().includes('受診中の歯科医院'));
  if (hospitalQ) {
    Logger.log('');
    Logger.log('=== 医療機関選択のEntry ID ===');
    Logger.log(hospitalQ.getId());
  }
}
function getFormPublishedUrl() {
  const form = FormApp.getActiveForm();
  Logger.log('公開URL: ' + form.getPublishedUrl());
}
function testGetHospitalList() {
  Logger.log(getHospitalList());
}
// ========== POSTエンドポイント（OCR変換データ受信用） ==========
function doPost(e) {
  let result;
  try {
    let body;
    // フォームPOST（iframe経由）の場合は e.parameter.payload を使用
    if (e.parameter && e.parameter.payload) {
      body = JSON.parse(e.parameter.payload);
    } else {
      body = JSON.parse(e.postData.contents);
    }
    const action = body.action;
    switch (action) {
      case 'addSurveyResponse':
        result = { success: true, data: addSurveyResponse(body.data) };
        break;
      case 'savePdfCopy':
        result = { success: true, data: savePdfToGoogleDrive(body.pdfBase64, body.rowNumber, body.idNumber, body.duplicateCount, body.originalFilename) };
        break;
      // ========== OCR照合（共有版） ==========
      case 'saveOcrResult':
        result = { success: true, data: saveOcrResult(body.filename, body.data, body.overwrite) };
        break;
      case 'saveVerifiedOcrResult':
        result = { success: true, data: saveVerifiedOcrResult(body.fileId, body.data) };
        break;
      case 'saveIndexJson':
        result = { success: true, data: saveIndexJsonToDrive(body.indexData) };
        break;
      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (error) {
    result = { success: false, error: error.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
/**
 * OCR変換済みデータをスプレッドシートに行追加
 * @param {Object} data - スプレッドシートの列ヘッダーをキーとしたデータ
 * @returns {Object} 結果（追加行番号、重複チェック結果）
 */
function addSurveyResponse(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
  // ヘッダー行を取得
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  // 重複チェック（同じID番号があるか確認、ただしブロックはしない）
  const idColumn = headers.indexOf('2. ID番号');
  const newId = data['2. ID番号'];
  const lastRow = sheet.getLastRow();
  let duplicateCount = 0;
  if (idColumn >= 0 && newId) {
    if (lastRow > 1) {
      const existingIds = sheet.getRange(2, idColumn + 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < existingIds.length; i++) {
        if (String(existingIds[i][0]).trim() === String(newId).trim()) {
          duplicateCount++;
        }
      }
    }
  }
  // ヘッダーの順序に合わせてデータ配列を構築
  const newRow = lastRow + 1;
  const rowData = headers.map(header => {
    if (header === 'No.') {
      // 通し番号を自動採番（既存データの最大No. + 1）
      if (lastRow > 1) {
        const noColIndex = headers.indexOf('No.');
        const existingNos = sheet.getRange(2, noColIndex + 1, lastRow - 1, 1).getValues();
        let maxNo = 0;
        for (let i = 0; i < existingNos.length; i++) {
          const num = Number(existingNos[i][0]);
          if (!isNaN(num) && num > maxNo) maxNo = num;
        }
        return String(maxNo + 1);
      }
      return '1';
    }
    if (header === 'タイムスタンプ') {
      return new Date();  // 現在時刻
    }
    if (header === 'メールアドレス') {
      return '';  // OCRからの送信にはメールなし
    }
    if (header === JSON_CREATED_COLUMN_NAME) {
      return '';  // JSON作成済みフラグは空
    }
    return data[header] !== undefined ? data[header] : '';
  });
  // 最終行の次に追加
  sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
  // No.列の表示形式を数値に設定（日付表示防止）
  const noColIdx = headers.indexOf('No.');
  if (noColIdx >= 0) {
    sheet.getRange(newRow, noColIdx + 1).setNumberFormat('@');
  }
  // JSON保存（共有ドライブ）
  try {
    const jsonData = {
      metadata: {
        createdAt: new Date().toISOString(),
        source: 'OCR'
      },
      data: {}
    };
    for (let key in data) {
      jsonData.data[key] = data[key];
    }
    saveJsonToGoogleDrive(jsonData, newRow);
    // JSON作成済みフラグを設定
    const jsonColIndex = headers.indexOf(JSON_CREATED_COLUMN_NAME);
    if (jsonColIndex >= 0) {
      sheet.getRange(newRow, jsonColIndex + 1).setValue('済');
    }
  } catch (jsonError) {
    Logger.log('JSON保存エラー（スプレッドシート追加は成功）: ' + jsonError.toString());
  }
  const resultMsg = duplicateCount > 0
    ? 'データを行 ' + newRow + ' に追加しました。（同一ID ' + newId + ' の ' + (duplicateCount + 1) + '件目）'
    : 'データを行 ' + newRow + ' に追加しました。';
  return {
    status: 'success',
    message: resultMsg,
    row: newRow,
    id: newId || null,
    duplicateCount: duplicateCount
  };
}
// ========== JSON保存機能 ==========
/**
 * フォーム送信時に呼び出されるトリガー関数
 */
function onFormSubmit(e) {
  try {
    const row = e.range.getRow();
    const sheet = e.range.getSheet();
    // No.列の自動採番（全行を検索して最大No.+1を割り当て）
    const allHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const noColIndex = allHeaders.indexOf('No.');
    if (noColIndex >= 0) {
      let maxNo = 0;
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        // 新しい行自身を除く全データ行を検索して最大No.を取得
        const existingNos = sheet.getRange(2, noColIndex + 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < existingNos.length; i++) {
          // 自分自身の行（まだ空のはず）はスキップ
          if (i + 2 === row) continue;
          const num = Number(existingNos[i][0]);
          if (!isNaN(num) && num > maxNo) maxNo = num;
        }
      }
      sheet.getRange(row, noColIndex + 1).setNumberFormat('@').setValue(String(maxNo + 1));
    }
    const jsonColumnIndex = getOrCreateJsonFlagColumn(sheet);
    const flagValue = sheet.getRange(row, jsonColumnIndex).getValue();
    if (flagValue === '済') {
      Logger.log('行 ' + row + ' は既にJSON作成済みのためスキップ');
      return;
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const jsonData = buildJsonFromRow(headers, rowData);
    // No.列の値を直接取得してJSON保存に渡す
    const noValue = noColIndex >= 0 ? sheet.getRange(row, noColIndex + 1).getValue() : null;
    saveJsonToGoogleDrive(jsonData, row, noValue);
    sheet.getRange(row, jsonColumnIndex).setValue('済');
    Logger.log('行 ' + row + ' のJSON作成完了（No.' + noValue + '）');
  } catch (error) {
    Logger.log('onFormSubmit エラー: ' + error.toString());
  }
}
/**
 * JSON作成済みフラグ列のインデックスを取得（なければ作成）
 */
function getOrCreateJsonFlagColumn(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === JSON_CREATED_COLUMN_NAME) {
      return i + 1;
    }
  }
  const newColumnIndex = sheet.getLastColumn() + 1;
  sheet.getRange(1, newColumnIndex).setValue(JSON_CREATED_COLUMN_NAME);
  return newColumnIndex;
}
/**
 * 行データからJSONオブジェクトを構築
 */
function buildJsonFromRow(headers, rowData) {
  const json = {
    metadata: {
      createdAt: new Date().toISOString(),
      source: 'GoogleForm'
    },
    data: {}
  };
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    const value = rowData[i];
    if (header === JSON_CREATED_COLUMN_NAME || header === 'No.') {
      continue;
    }
    if (header === 'タイムスタンプ') {
      json.metadata.submittedAt = value ? new Date(value).toISOString() : null;
      continue;
    }
    if (header === 'メールアドレス') {
      json.metadata.email = value || null;
      continue;
    }
    json.data[header] = value !== '' ? value : null;
  }
  return json;
}
/**
 * JSONファイルをGoogleドライブに保存
 * ファイル名形式: No_タイムスタンプ.json
 * 例: 001_20260217_173511.json
 */
function saveJsonToGoogleDrive(jsonData, rowNumber, noValueDirect) {
  const folder = getOrCreateJsonFolder();
  // No.を取得（直接渡された値を優先、なければスプレッドシートから取得）
  let noStr = '';
  if (noValueDirect) {
    noStr = String(noValueDirect).padStart(3, '0');
  } else {
    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const noColIndex = headers.indexOf('No.');
      if (noColIndex >= 0 && rowNumber >= 2) {
        const noValue = sheet.getRange(rowNumber, noColIndex + 1).getValue();
        if (noValue) noStr = String(noValue).padStart(3, '0');
      }
    } catch (e) {
      // No.取得失敗時はrow番号で代替
    }
  }
  if (!noStr) noStr = String(rowNumber).padStart(3, '0');
  // タイムスタンプを取得（データシートのタイムスタンプ優先、なければ現在時刻）
  let tsDate = null;
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const tsColIndex = headers.indexOf('タイムスタンプ');
    if (tsColIndex >= 0 && rowNumber >= 2) {
      const tsValue = sheet.getRange(rowNumber, tsColIndex + 1).getValue();
      if (tsValue) tsDate = new Date(tsValue);
    }
  } catch (e) {
    // 取得失敗時は現在時刻にフォールバック
  }
  if (!tsDate || isNaN(tsDate.getTime())) tsDate = new Date();
  // タイムスタンプ文字列（YYYYMMDD_HHmmss）
  const tsStr = tsDate.getFullYear()
    + String(tsDate.getMonth() + 1).padStart(2, '0')
    + String(tsDate.getDate()).padStart(2, '0')
    + '_'
    + String(tsDate.getHours()).padStart(2, '0')
    + String(tsDate.getMinutes()).padStart(2, '0')
    + String(tsDate.getSeconds()).padStart(2, '0');
  // ファイル名: No_タイムスタンプ.json
  const baseName = noStr + '_' + tsStr;
  // 同じファイル名が既にあれば (2), (3)... を付与
  const escapedBase = baseName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const allFiles = folder.getFiles();
  let sameCount = 0;
  while (allFiles.hasNext()) {
    const f = allFiles.next();
    const fName = f.getName();
    if (fName === baseName + '.json' || new RegExp('^' + escapedBase + '\\(\\d+\\)\\.json$').test(fName)) {
      sameCount++;
    }
  }
  let fileName;
  if (sameCount > 0) {
    fileName = baseName + '(' + (sameCount + 1) + ').json';
  } else {
    fileName = baseName + '.json';
  }
  const jsonString = JSON.stringify(jsonData, null, 2);
  folder.createFile(fileName, jsonString, MimeType.PLAIN_TEXT);
  Logger.log('JSONファイル作成: ' + fileName);
}
/**
 * PDFコピーをGoogleドライブに保存
 * JSONファイルと同じフォルダに同じ命名規則（拡張子だけ.pdf）で保存
 */
function savePdfToGoogleDrive(pdfBase64, rowNumber, idNumber, duplicateCount, originalFilename) {
  const folder = getOrCreateJsonFolder();
  // No.を取得（スプレッドシートから）
  let noStr = '';
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const noColIndex = headers.indexOf('No.');
    if (noColIndex >= 0 && rowNumber >= 2) {
      const noValue = sheet.getRange(rowNumber, noColIndex + 1).getValue();
      if (noValue) noStr = String(noValue).padStart(3, '0');
    }
  } catch (e) {
    // No.取得失敗時はrow番号で代替
  }
  if (!noStr) noStr = String(rowNumber).padStart(3, '0');
  // 日付（YYYYMMDD）
  const now = new Date();
  const dateStr = now.getFullYear()
    + String(now.getMonth() + 1).padStart(2, '0')
    + String(now.getDate()).padStart(2, '0');
  const id = idNumber || 'noID';
  const baseName = noStr + '_OCR_' + id + '_' + dateStr;
  // 同じIDのPDFファイルが既にあれば (2), (3)... を付与
  const idPattern = '_OCR_' + id + '_';
  const allFiles = folder.getFiles();
  let sameIdPdfCount = 0;
  while (allFiles.hasNext()) {
    const f = allFiles.next();
    const name = f.getName();
    if (name.indexOf(idPattern) >= 0 && name.endsWith('.pdf')) {
      sameIdPdfCount++;
    }
  }
  let fileName;
  if (sameIdPdfCount > 0) {
    fileName = baseName + '(' + (sameIdPdfCount + 1) + ').pdf';
  } else {
    fileName = baseName + '.pdf';
  }
  // Base64デコードしてPDFファイルとして保存
  const pdfBlob = Utilities.newBlob(Utilities.base64Decode(pdfBase64), 'application/pdf', fileName);
  folder.createFile(pdfBlob);
  Logger.log('PDFコピー作成: ' + fileName + ' (元: ' + originalFilename + ')');
  // 元PDFをスキャンPDFフォルダから入力済みPDFフォルダへ移動
  let movedOriginal = false;
  if (originalFilename) {
    try {
      const scanFolder = DriveApp.getFolderById(SCAN_PDF_FOLDER_ID);
      const doneFolder = DriveApp.getFolderById(DONE_PDF_FOLDER_ID);
      const matchFiles = scanFolder.getFilesByName(originalFilename);
      if (matchFiles.hasNext()) {
        const origFile = matchFiles.next();
        origFile.moveTo(doneFolder);
        movedOriginal = true;
        Logger.log('元PDF移動: ' + originalFilename + ' → 入力済みPDFフォルダ');
      }
    } catch (moveErr) {
      Logger.log('元PDF移動エラー（PDFコピーは成功）: ' + moveErr.toString());
    }
  }
  return {
    fileName: fileName,
    originalFilename: originalFilename,
    movedOriginal: movedOriginal
  };
}
/**
 * JSON保存用フォルダを取得または作成
 * 共有ドライブ「3.データシート」フォルダ内に作成
 */
function getOrCreateJsonFolder() {
  const parentFolder = DriveApp.getFolderById(JSON_PARENT_FOLDER_ID);
  const folders = parentFolder.getFoldersByName(JSON_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(JSON_FOLDER_NAME);
}
/**
 * 既存データでJSON未作成のものを一括処理
 */
function processExistingDataWithoutJson() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
  if (!sheet) {
    Logger.log('回答シートが見つかりません: ' + RESPONSE_SHEET_NAME);
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('データがありません');
    return;
  }
  const jsonColumnIndex = getOrCreateJsonFlagColumn(sheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let processedCount = 0;
  let skippedCount = 0;
  for (let row = 2; row <= lastRow; row++) {
    const flagValue = sheet.getRange(row, jsonColumnIndex).getValue();
    if (flagValue === '済') {
      skippedCount++;
      continue;
    }
    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!rowData[0]) {
      continue;
    }
    const jsonData = buildJsonFromRow(headers, rowData);
    saveJsonToGoogleDrive(jsonData, row);
    sheet.getRange(row, jsonColumnIndex).setValue('済');
    processedCount++;
  }
  Logger.log('処理完了: ' + processedCount + '件作成, ' + skippedCount + '件スキップ');
}
/**
 * トリガーをセットアップする関数（初回のみ手動実行）
 */
function setupFormSubmitTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  Logger.log('フォーム送信トリガーを設定しました');
}
// ========== No.列の管理 ==========
/**
 * スプレッドシートのA列に「No.」列を挿入し、既存データに通し番号を振る
 * ※初回のみ手動実行する関数
 */
function insertNoColumn() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
  if (!sheet) {
    Logger.log('回答シートが見つかりません: ' + RESPONSE_SHEET_NAME);
    return;
  }
  // 既にNo.列があるか確認
  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (existingHeaders.indexOf('No.') >= 0) {
    Logger.log('No.列は既に存在します。backfillNoColumn() で欠番を埋めてください。');
    return;
  }
  // A列に列を挿入
  sheet.insertColumnBefore(1);
  // ヘッダーに「No.」を設定
  sheet.getRange(1, 1).setValue('No.');
  // 既存データに通し番号を振る
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const numbers = [];
    for (let i = 1; i <= lastRow - 1; i++) {
      numbers.push([i]);
    }
    const strNumbers = numbers.map(function(n) { return [String(n[0])]; });
    sheet.getRange(2, 1, lastRow - 1, 1).setNumberFormat('@').setValues(strNumbers);
  }
  // 「元PDFファイル名」列がなければ末尾に追加
  const updatedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (updatedHeaders.indexOf('元PDFファイル名') < 0) {
    const pdfColIndex = sheet.getLastColumn() + 1;
    sheet.getRange(1, pdfColIndex).setValue('元PDFファイル名');
    Logger.log('「元PDFファイル名」列を追加しました（列' + pdfColIndex + '）');
  }
  Logger.log('No.列を挿入し、' + (lastRow - 1) + '件に通し番号を付与しました。');
}
/**
 * No.列で番号が空の行に通し番号を振る（バックフィル）
 */
function backfillNoColumn() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const noColIndex = headers.indexOf('No.');
  if (noColIndex < 0) {
    Logger.log('No.列が見つかりません。先に insertNoColumn() を実行してください。');
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('データがありません');
    return;
  }
  const noValues = sheet.getRange(2, noColIndex + 1, lastRow - 1, 1).getValues();
  // 既存の最大No.を取得
  let maxNo = 0;
  for (let i = 0; i < noValues.length; i++) {
    const num = Number(noValues[i][0]);
    if (!isNaN(num) && num > maxNo) maxNo = num;
  }
  // 空のセルに番号を振る
  let filledCount = 0;
  for (let i = 0; i < noValues.length; i++) {
    if (noValues[i][0] === '' || noValues[i][0] === null || noValues[i][0] === undefined) {
      maxNo++;
      sheet.getRange(i + 2, noColIndex + 1).setNumberFormat('@').setValue(String(maxNo));
      filledCount++;
    }
  }
  Logger.log('バックフィル完了: ' + filledCount + '件にNo.を付与（最大No.: ' + maxNo + '）');
}
// ※ resetNoColumn() は削除済み
// リナンバリングはJSONファイル名との整合性が崩れるため禁止
// 欠番はそのまま維持すること

/**
 * 「JSON作成済み」列の右に「元PDFファイル名」列を追加する
 * 既に存在する場合はスキップ
 */
function addPdfFilenameColumn() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
  if (!sheet) {
    Logger.log('回答シートが見つかりません: ' + RESPONSE_SHEET_NAME);
    return;
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('元PDFファイル名') >= 0) {
    Logger.log('「元PDFファイル名」列は既に存在します。');
    return;
  }
  const jsonColIndex = headers.indexOf(JSON_CREATED_COLUMN_NAME);
  if (jsonColIndex >= 0) {
    // 「JSON作成済み」の右隣に列を挿入
    sheet.insertColumnAfter(jsonColIndex + 1);
    sheet.getRange(1, jsonColIndex + 2).setValue('元PDFファイル名');
    Logger.log('「JSON作成済み」の右に「元PDFファイル名」列を追加しました（列' + (jsonColIndex + 2) + '）');
  } else {
    // 「JSON作成済み」がなければ末尾に追加
    const newCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, newCol).setValue('元PDFファイル名');
    Logger.log('末尾に「元PDFファイル名」列を追加しました（列' + newCol + '）');
  }
}

/**
 * 「JSON作成済み」列の左に「抜去位置」列を追加する
 * 既に存在する場合はスキップ
 */
function addExtractionPositionColumn() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
  if (!sheet) {
    Logger.log('回答シートが見つかりません: ' + RESPONSE_SHEET_NAME);
    return;
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('抜去位置') >= 0) {
    Logger.log('「抜去位置」列は既に存在します。');
    return;
  }
  const jsonColIndex = headers.indexOf(JSON_CREATED_COLUMN_NAME);
  if (jsonColIndex >= 0) {
    // 「JSON作成済み」の左に列を挿入
    sheet.insertColumnBefore(jsonColIndex + 1);
    sheet.getRange(1, jsonColIndex + 1).setValue('抜去位置');
    Logger.log('「JSON作成済み」の左に「抜去位置」列を追加しました（列' + (jsonColIndex + 1) + '）');
  } else {
    // 「JSON作成済み」がなければ末尾に追加
    const newCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, newCol).setValue('抜去位置');
    Logger.log('末尾に「抜去位置」列を追加しました（列' + newCol + '）');
  }
}

/**
 * 「3. 名前（カタカナ）」列の右に「QRコード回答」列を追加する
 * 既に存在する場合はスキップ
 */
function addQrAnswerColumn() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
  if (!sheet) {
    Logger.log('回答シートが見つかりません: ' + RESPONSE_SHEET_NAME);
    return;
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('QRコード回答') >= 0) {
    Logger.log('「QRコード回答」列は既に存在します。');
    return;
  }
  const nameColIndex = headers.indexOf('3. 名前（カタカナ）');
  if (nameColIndex >= 0) {
    // 「3. 名前（カタカナ）」の右に列を挿入
    sheet.insertColumnAfter(nameColIndex + 1);
    sheet.getRange(1, nameColIndex + 2).setValue('QRコード回答');
    Logger.log('「3. 名前（カタカナ）」の右に「QRコード回答」列を追加しました（列' + (nameColIndex + 2) + '）');
  } else {
    // 「3. 名前（カタカナ）」がなければ末尾に追加
    const newCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, newCol).setValue('QRコード回答');
    Logger.log('末尾に「QRコード回答」列を追加しました（列' + newCol + '）');
  }
}
/**
 * 指定No.範囲のデータを削除（全シート横断 + JSONファイル）
 * 医療機関リストを除く全シートのNo.列を検索して該当行を削除
 * 使い方: deleteDataByNo(50, 54) → No.50〜54を削除
 */
function deleteDataByNo(fromNo, toNo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  const excludeSheetNames = [HOSPITAL_LIST_SHEET_NAME];
  // --- 1. JSONファイルを削除（ゴミ箱に移動） ---
  const folder = getOrCreateJsonFolder();
  const allFiles = folder.getFiles();
  let deletedFiles = 0;
  while (allFiles.hasNext()) {
    const file = allFiles.next();
    const fileName = file.getName();
    // ファイル名の先頭3桁のNo.を取得（例: "052_OCR_A138_..." → 52）
    const match = fileName.match(/^(\d{3})_/);
    if (match) {
      const fileNo = parseInt(match[1], 10);
      if (fileNo >= fromNo && fileNo <= toNo) {
        file.setTrashed(true);
        Logger.log('JSONファイル削除: ' + fileName);
        deletedFiles++;
      }
    }
  }
  // --- 2. 全シートからスプレッドシートの行を削除 ---
  let deletedRows = 0;
  var sheetDetails = [];
  for (let s = 0; s < sheets.length; s++) {
    const sheet = sheets[s];
    const sheetName = sheet.getName();
    if (excludeSheetNames.indexOf(sheetName) >= 0) continue;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const noColIndex = headers.indexOf('No.');
    if (noColIndex < 0) continue;
    const noValues = sheet.getRange(2, noColIndex + 1, lastRow - 1, 1).getValues();
    let sheetDeletedRows = 0;
    // 下の行から削除して行番号のずれを防ぐ
    for (let i = noValues.length - 1; i >= 0; i--) {
      const no = Number(noValues[i][0]);
      if (no >= fromNo && no <= toNo) {
        sheet.deleteRow(i + 2);
        sheetDeletedRows++;
      }
    }
    if (sheetDeletedRows > 0) {
      sheetDetails.push(sheetName + ': ' + sheetDeletedRows + '行');
      deletedRows += sheetDeletedRows;
    }
  }
  // リナンバリングはしない（欠番OK、JSONファイル名との整合性を維持）
  var detailMsg = sheetDetails.length > 0 ? '（' + sheetDetails.join(', ') + '）' : '';
  Logger.log('削除完了: スプレッドシート ' + deletedRows + '行' + detailMsg + ', JSONファイル ' + deletedFiles + '件');
  return {
    deletedRows: deletedRows,
    deletedFiles: deletedFiles,
    message: 'スプレッドシート ' + deletedRows + '行' + detailMsg + ', JSONファイル ' + deletedFiles + '件を削除しました。'
  };
}
/**
 * スプレッドシートのNo.・ID・医療機関・元PDFファイル名の一覧を取得
 */
function getNoList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const noCol = headers.indexOf('No.');
  const idCol = headers.indexOf('2. ID番号');
  const hospitalCol = headers.indexOf('1. 受診中の歯科医院を選んでください。');
  const pdfCol = headers.indexOf('元PDFファイル名');
  const tsCol = headers.indexOf('タイムスタンプ');
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const list = [];
  for (let i = 0; i < data.length; i++) {
    const no = noCol >= 0 ? data[i][noCol] : '';
    if (!no) continue;  // No.がない行はスキップ
    list.push({
      no: no,
      id: idCol >= 0 ? data[i][idCol] : '',
      hospital: hospitalCol >= 0 ? data[i][hospitalCol] : '',
      pdfFilename: pdfCol >= 0 ? data[i][pdfCol] : '',
      timestamp: tsCol >= 0 ? data[i][tsCol] : '',
      row: i + 2
    });
  }
  return list;
}

// ========== 検体番号検索（氏名・生年月日でNo.を検索） ==========
function searchByNameBirthdate(params) {
  const name = (params.name || '').trim();
  const nengo = (params.nengo || '').trim();
  const year = (params.year || '').trim();
  const month = (params.month || '').trim();
  const day = (params.day || '').trim();

  if (!name && !nengo && !year && !month && !day) {
    return { results: [], count: 0, message: '検索条件を1つ以上入力してください' };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  const excludeSheetNames = [HOSPITAL_LIST_SHEET_NAME];
  const results = [];

  for (let s = 0; s < sheets.length; s++) {
    const sheet = sheets[s];
    const sheetName = sheet.getName();
    if (excludeSheetNames.indexOf(sheetName) >= 0) continue;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (headers.indexOf('No.') < 0) continue;

    const noCol = headers.indexOf('No.');
    const nameCol = headers.indexOf('3. 名前（カタカナ）');
    const nengoCol = headers.indexOf('4. 生年月日 - 年号');
    const yearCol = headers.indexOf('5. - 年（半角数字）');
    const monthCol = headers.indexOf('6. - 月（半角数字）');
    const dayCol = headers.indexOf('7. - 日（半角数字）');
    const idCol = headers.indexOf('2. ID番号');
    const hospCol = headers.indexOf('1. 受診中の歯科医院を選んでください。');

    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    for (let r = 0; r < data.length; r++) {
      const row = data[r];
      const rowNo = noCol >= 0 ? row[noCol] : '';
      if (!rowNo) continue;

      // 名前の部分一致検索（スペース無視）
      if (name) {
        const rowName = nameCol >= 0 ? String(row[nameCol]).replace(/\s+/g, '') : '';
        const searchName = name.replace(/\s+/g, '');
        if (!rowName.includes(searchName)) continue;
      }

      // 生年月日の完全一致チェック（空欄の条件は無視）
      if (nengo && nengoCol >= 0) {
        if (String(row[nengoCol]).trim() !== nengo) continue;
      }
      if (year && yearCol >= 0) {
        if (String(row[yearCol]).trim() !== year) continue;
      }
      if (month && monthCol >= 0) {
        if (String(row[monthCol]).trim() !== month) continue;
      }
      if (day && dayCol >= 0) {
        if (String(row[dayCol]).trim() !== day) continue;
      }

      results.push({
        no: rowNo,
        name: nameCol >= 0 ? row[nameCol] : '',
        nengo: nengoCol >= 0 ? row[nengoCol] : '',
        year: yearCol >= 0 ? row[yearCol] : '',
        month: monthCol >= 0 ? row[monthCol] : '',
        day: dayCol >= 0 ? row[dayCol] : '',
        id: idCol >= 0 ? row[idCol] : '',
        hospital: hospCol >= 0 ? row[hospCol] : ''
      });
    }
  }

  // No.昇順でソート
  results.sort(function(a, b) {
    return (Number(a.no) || 0) - (Number(b.no) || 0);
  });

  return { results: results, count: results.length };
}

/**
 * 指定No.の「元PDFファイル名」列を更新
 */
function updatePdfFilename(targetNo, pdfFilename) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getLastRow();
  const noCol = headers.indexOf('No.');
  const pdfCol = headers.indexOf('元PDFファイル名');
  if (noCol < 0) return { success: false, message: 'No.列が見つかりません' };
  if (pdfCol < 0) return { success: false, message: '元PDFファイル名列が見つかりません' };
  const noValues = sheet.getRange(2, noCol + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < noValues.length; i++) {
    if (Number(noValues[i][0]) === targetNo) {
      sheet.getRange(i + 2, pdfCol + 1).setValue(pdfFilename);
      Logger.log('PDF紐づけ更新: No.' + targetNo + ' → ' + pdfFilename);
      return { success: true, message: 'No.' + targetNo + ' に ' + pdfFilename + ' を紐づけました。', row: i + 2 };
    }
  }
  return { success: false, message: 'No.' + targetNo + ' が見つかりません' };
}
/**
 * ②ルート（フォーム入力+PDF紐づけ）のSourceを "GoogleForm" → "OCR" に変更
 * JSONファイル名の _Form_ → _OCR_ リネーム + JSON内のsource書き換え
 * @param {number} targetNo - 対象のNo.
 * @returns {Object} 結果
 */
function updateSourceToOCR(targetNo, extraData) {
  const folder = getOrCreateJsonFolder();
  const noStr = String(targetNo).padStart(3, '0');
  extraData = extraData || {};
  // 該当No.のJSONファイルを検索（新旧両方の命名規則に対応）
  const allFiles = folder.getFiles();
  let renamedJson = false;
  let oldJsonName = '';
  let newJsonName = '';
  while (allFiles.hasNext()) {
    const file = allFiles.next();
    const name = file.getName();
    // 該当No.のJSONファイルを探す（新形式: 001_20260217_173511.json / 旧形式: 001_Form_xxx.json）
    if (name.startsWith(noStr + '_') && name.endsWith('.json')) {
      oldJsonName = name;
      // JSONファイル内のsource・抜去位置・PDFファイル名を書き換え
      try {
        const content = file.getBlob().getDataAsString();
        const jsonData = JSON.parse(content);
        if (jsonData.metadata) {
          jsonData.metadata.source = 'OCR';
        }
        if (jsonData.data) {
          // 抜去位置を更新
          if (extraData.extractionPosition) {
            jsonData.data['抜去位置'] = extraData.extractionPosition;
          }
          // 元PDFファイル名を更新
          if (extraData.pdfFileName) {
            jsonData.data['元PDFファイル名'] = extraData.pdfFileName;
          }
        }
        file.setContent(JSON.stringify(jsonData, null, 2));
      } catch (parseErr) {
        Logger.log('JSON内容更新エラー: ' + parseErr.toString());
      }
      // 旧形式の場合のみリネーム（_Form_ → _OCR_）
      if (name.indexOf('_Form_') >= 0) {
        newJsonName = name.replace('_Form_', '_OCR_');
        file.setName(newJsonName);
        Logger.log('JSONリネーム: ' + oldJsonName + ' → ' + newJsonName);
      } else {
        newJsonName = name;
        Logger.log('JSON内容更新: ' + name);
      }
      renamedJson = true;
      break;
    }
  }
  return {
    success: renamedJson,
    oldJsonName: oldJsonName,
    newJsonName: newJsonName,
    message: renamedJson
      ? 'No.' + targetNo + ' のSource を OCR に変更しました。(' + oldJsonName + ' → ' + newJsonName + ')'
      : 'No.' + targetNo + ' のJSONファイルが見つかりませんでした。'
  };
}
/**
 * 閲覧用データを取得（パスワード検証済みの場合のみ呼ばれる）
 * スプレッドシート内の全シート（医療機関リスト除く、No.列があるもの）からデータを結合
 * No.昇順でソート。JSON作成済み・元PDFファイル名・メールアドレス列は除外
 * 重複No.がある場合はduplicateNos配列で通知
 */
function getViewData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  const excludeSheetNames = [HOSPITAL_LIST_SHEET_NAME];
  let masterHeaders = null;
  let allFilteredRows = [];
  for (let s = 0; s < sheets.length; s++) {
    const sheet = sheets[s];
    const sheetName = sheet.getName();
    if (excludeSheetNames.indexOf(sheetName) >= 0) continue;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (headers.indexOf('No.') < 0) continue;
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const excludeCols = [
      headers.indexOf(JSON_CREATED_COLUMN_NAME),
      headers.indexOf('元PDFファイル名'),
      headers.indexOf('メールアドレス')
    ].filter(function(i) { return i >= 0; });
    if (!masterHeaders) {
      var filtered = headers.filter(function(_, i) { return excludeCols.indexOf(i) < 0; });
      // No.列の前に「シート名」列を挿入
      var noPos = filtered.indexOf('No.');
      if (noPos >= 0) {
        filtered.splice(noPos, 0, 'シート名');
      } else {
        filtered.unshift('シート名');
      }
      masterHeaders = filtered;
    }
    const noIdx = headers.indexOf('No.');
    for (let r = 0; r < data.length; r++) {
      var row = data[r];
      if (row[noIdx] === '' || row[noIdx] === null || row[noIdx] === undefined) continue;
      var mappedRow = [];
      for (let m = 0; m < masterHeaders.length; m++) {
        if (masterHeaders[m] === 'シート名') {
          mappedRow.push(sheetName);
          continue;
        }
        var colIdx = headers.indexOf(masterHeaders[m]);
        if (colIdx < 0) {
          mappedRow.push('');
        } else {
          var cell = row[colIdx];
          if (cell instanceof Date) {
            mappedRow.push(Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss'));
          } else {
            mappedRow.push(cell);
          }
        }
      }
      allFilteredRows.push(mappedRow);
    }
  }
  if (!masterHeaders) return { headers: [], rows: [], totalCount: 0, duplicateNos: [] };
  // No.列で昇順ソート
  var masterNoIdx = masterHeaders.indexOf('No.');
  if (masterNoIdx >= 0) {
    allFilteredRows.sort(function(a, b) {
      return (Number(a[masterNoIdx]) || 0) - (Number(b[masterNoIdx]) || 0);
    });
  }
  // 重複No.を検出
  var duplicateNos = [];
  if (masterNoIdx >= 0) {
    var noCounts = {};
    for (let i = 0; i < allFilteredRows.length; i++) {
      var no = allFilteredRows[i][masterNoIdx];
      if (no !== '' && no !== null) {
        var noKey = String(no);
        noCounts[noKey] = (noCounts[noKey] || 0) + 1;
      }
    }
    for (var key in noCounts) {
      if (noCounts[key] > 1) {
        duplicateNos.push(key);
      }
    }
  }
  return {
    headers: masterHeaders,
    rows: allFilteredRows,
    totalCount: allFilteredRows.length,
    duplicateNos: duplicateNos
  };
}
/**
 * PDF処理状況を取得
 * 未処理: スキャンPDFフォルダ（SCAN_PDF_FOLDER_ID）内のPDF
 * 処理済み: 入力済みPDFフォルダ（DONE_PDF_FOLDER_ID）内のPDF
 * @returns {Object} 未処理・処理済みの一覧
 */
function getUnprocessedPdfs() {
  // 1. 未処理: スキャンPDFフォルダ内のPDFファイルを取得
  var scanFolder = DriveApp.getFolderById(SCAN_PDF_FOLDER_ID);
  var scanFiles = scanFolder.getFiles();
  var unprocessed = [];
  while (scanFiles.hasNext()) {
    var file = scanFiles.next();
    var name = file.getName();
    if (name.toLowerCase().endsWith('.pdf')) {
      unprocessed.push({
        name: name,
        date: Utilities.formatDate(file.getDateCreated(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'),
        size: Math.round(file.getSize() / 1024)
      });
    }
  }
  unprocessed.sort(function(a, b) { return b.date.localeCompare(a.date); });

  // 2. 処理済み: 入力済みPDFフォルダ内のPDFファイルを取得
  var doneFolder = DriveApp.getFolderById(DONE_PDF_FOLDER_ID);
  var doneFiles = doneFolder.getFiles();
  var processed = [];
  while (doneFiles.hasNext()) {
    var file2 = doneFiles.next();
    var name2 = file2.getName();
    if (name2.toLowerCase().endsWith('.pdf')) {
      processed.push({
        name: name2,
        date: Utilities.formatDate(file2.getDateCreated(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'),
        size: Math.round(file2.getSize() / 1024)
      });
    }
  }
  processed.sort(function(a, b) { return b.date.localeCompare(a.date); });

  return {
    unprocessedCount: unprocessed.length,
    processedCount: processed.length,
    unprocessed: unprocessed,
    processed: processed
  };
}
/**
 * アンケートJSONフォルダ内のファイル一覧を取得
 * No.、ソース（Form/OCR）、ID、日付、ファイル名を返す
 */
function getJsonFileList() {
  const folder = getOrCreateJsonFolder();
  const files = folder.getFiles();
  const fileList = [];
  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    // ファイル名パターン: 001_Form_211_20260203.json
    const match = name.match(/^(\d{3})_(Form|OCR)_(.+?)_(\d{8})(?:\(\d+\))?\.json$/);
    fileList.push({
      fileName: name,
      no: match ? parseInt(match[1], 10) : null,
      source: match ? match[2] : null,
      id: match ? match[3] : null,
      date: match ? match[4] : null,
      size: file.getSize(),
      lastUpdated: file.getLastUpdated().toISOString()
    });
  }
  // No.昇順でソート
  fileList.sort(function(a, b) {
    if (a.no === null && b.no === null) return 0;
    if (a.no === null) return 1;
    if (b.no === null) return -1;
    return a.no - b.no;
  });
  return {
    totalFiles: fileList.length,
    files: fileList
  };
}
/**
 * JSONフォルダ内のファイル名から指定パターンを除去してリネーム
 * @param {string} pattern - 除去する文字列（例: "(37)"）
 * @param {string} replacement - 置換後の文字列（デフォルト: 空文字）
 */
function renameJsonFiles(pattern, replacement) {
  const folder = getOrCreateJsonFolder();
  const files = folder.getFiles();
  const renamed = [];
  const skipped = [];
  while (files.hasNext()) {
    const file = files.next();
    const oldName = file.getName();
    if (oldName.indexOf(pattern) < 0) continue;
    const newName = oldName.replace(pattern, replacement || '');
    // リネーム先と同名のファイルが既に存在するかチェック
    const existing = folder.getFilesByName(newName);
    if (existing.hasNext()) {
      skipped.push({ oldName: oldName, newName: newName, reason: '同名ファイルが既に存在' });
      continue;
    }
    file.setName(newName);
    renamed.push({ oldName: oldName, newName: newName });
  }
  return {
    renamedCount: renamed.length,
    skippedCount: skipped.length,
    renamed: renamed,
    skipped: skipped
  };
}
/**
 * JSONフォルダ内のファイル名に指定パターンを含むファイルをゴミ箱に移動
 * @param {string} pattern - 検索する文字列（例: "(37)"）
 */
function trashJsonFilesByPattern(pattern) {
  const folder = getOrCreateJsonFolder();
  const files = folder.getFiles();
  const trashed = [];
  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (name.indexOf(pattern) >= 0) {
      file.setTrashed(true);
      trashed.push(name);
    }
  }
  return {
    trashedCount: trashed.length,
    trashed: trashed
  };
}
// ========== 検体PpP値 ==========
/**
 * 検体PpP値データを追加
 */
function addSpecimen(params) {
  const no = params.no;
  if (!no) return { success: false, message: 'No.が指定されていません' };
  const ss = SpreadsheetApp.openById(SPECIMEN_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SPECIMEN_SHEET_NAME);
  // ヘッダーがなければ作成
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, 5).setValues([['No.', '歯の位置', 'PpP値', '分析日', '備考']]);
  }
  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, 5).setValues([[
    Number(no),
    params.position || '',
    params.pppValue ? Number(params.pppValue) : '',
    params.analysisDate || '',
    params.note || ''
  ]]);
  return { success: true, message: 'No.' + no + ' の検体データを登録しました（行' + newRow + '）', row: newRow };
}
/**
 * 検体PpP値データ一覧を取得
 * @param {string} no - 指定No.（空なら全件）
 */
function getSpecimenList(no) {
  const ss = SpreadsheetApp.openById(SPECIMEN_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SPECIMEN_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { totalCount: 0, rows: [] };
  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const rows = [];
  for (let i = 0; i < data.length; i++) {
    if (no && String(data[i][0]) !== String(no)) continue;
    rows.push({
      no: data[i][0],
      position: data[i][1],
      pppValue: data[i][2],
      analysisDate: data[i][3],
      note: data[i][4]
    });
  }
  return { totalCount: rows.length, rows: rows };
}
// ========== HbA1c追跡 ==========
/**
 * HbA1c追跡データを追加（ヘッダー駆動・動的項目対応）
 * 固定列: No., 測定日, 備考
 * 動的列: ヘッダー名をURLパラメータのキーとして値を受け取る
 */
function addTracking(params) {
  const no = params.no;
  if (!no) return { success: false, message: 'No.が指定されていません' };
  const ss = SpreadsheetApp.openById(TRACKING_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TRACKING_SHEET_NAME);
  // ヘッダーがなければデフォルト作成
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, 6).setValues([['No.', '測定日', 'HbA1c(%)', '体重(kg)', '血糖値(mg/dL)', '備考']]);
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = headers.map(function(h) {
    if (h === 'No.') return Number(no);
    if (h === '測定日') return params.date || '';
    if (h === '備考') return params.note || '';
    // 動的項目: ヘッダー名をそのままキーとして使用
    var val = params[h] || '';
    return val !== '' ? Number(val) : '';
  });
  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
  return { success: true, message: 'No.' + no + ' の追跡データを登録しました（行' + newRow + '）', row: newRow };
}
/**
 * HbA1c追跡データ一覧を取得（ヘッダー情報付き・動的列対応）
 * @param {string} no - 指定No.（空なら全件）
 * @return {Object} { totalCount, rows(配列の配列), headers(配列) }
 */
function getTrackingList(no) {
  const ss = SpreadsheetApp.openById(TRACKING_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TRACKING_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { totalCount: 0, rows: [], headers: [] };
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (lastRow < 2) return { totalCount: 0, rows: [], headers: headers };
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const noIdx = headers.indexOf('No.');
  const rows = [];
  for (var i = 0; i < data.length; i++) {
    if (no && String(data[i][noIdx]) !== String(no)) continue;
    rows.push(data[i]);
  }
  return { totalCount: rows.length, rows: rows, headers: headers };
}
/**
 * HbA1c追跡シートに記録項目（列）を追加
 * 備考列の直前に挿入する
 * @param {string} columnName - ヘッダー名（例: "収縮期血圧(mmHg)"）
 */
function addTrackingColumn(columnName) {
  if (!columnName) return { success: false, message: '項目名が指定されていません' };
  const ss = SpreadsheetApp.openById(TRACKING_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TRACKING_SHEET_NAME);
  // ヘッダーがなければデフォルト作成
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, 3).setValues([['No.', '測定日', '備考']]);
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  // 重複チェック
  if (headers.indexOf(columnName) >= 0) {
    return { success: false, message: '「' + columnName + '」は既に存在します' };
  }
  // 備考列の位置を探す
  var noteIdx = headers.indexOf('備考');
  if (noteIdx >= 0) {
    // 備考の前に列を挿入
    sheet.insertColumnBefore(noteIdx + 1);
    sheet.getRange(1, noteIdx + 1).setValue(columnName);
  } else {
    // 備考がなければ末尾に追加
    var newCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, newCol).setValue(columnName);
  }
  return { success: true, message: '「' + columnName + '」を追加しました' };
}
// ========== PDF紐づけツール（Google Drive連携） ==========
/**
 * Google Driveの元PDFフォルダ内のPDFファイル一覧を返す
 * link_pdf.html の②ステップで使用
 */
function getDrivePdfList() {
  try {
    var folder = DriveApp.getFolderById(SCAN_PDF_FOLDER_ID);
    var files = folder.getFilesByType(MimeType.PDF);
    var list = [];
    while (files.hasNext()) {
      var f = files.next();
      list.push({
        fileId: f.getId(),
        fileName: f.getName(),
        fileSize: f.getSize(),
        lastUpdated: Utilities.formatDate(f.getLastUpdated(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm')
      });
    }
    // ファイル名でソート
    list.sort(function(a, b) {
      return a.fileName.localeCompare(b.fileName);
    });
    return list;
  } catch (err) {
    throw new Error('PDF一覧取得エラー: ' + err.message);
  }
}
/**
 * PDFインデックスシートとマージしてPDF一覧を返す
 */
function getDrivePdfListWithIndex() {
  var pdfList = getDrivePdfList();
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var indexSheet = ss.getSheetByName('PDFインデックス');
    if (!indexSheet || indexSheet.getLastRow() < 2) {
      // インデックスシートがない or データなし → 既存リストをそのまま返す
      return pdfList;
    }
    var indexData = indexSheet.getRange(1, 1, indexSheet.getLastRow(), indexSheet.getLastColumn()).getValues();
    var headers = indexData[0];
    var fileIdCol = headers.indexOf('fileId');
    var hospitalCol = headers.indexOf('hospital');
    var patientIdCol = headers.indexOf('patientId');
    var patientNameCol = headers.indexOf('patientName');
    var birthdateCol = headers.indexOf('birthdate');
    // fileIdをキーにしたマップを作成
    var indexMap = {};
    for (var i = 1; i < indexData.length; i++) {
      var row = indexData[i];
      var fId = row[fileIdCol];
      if (fId) {
        indexMap[fId] = {
          hospital: hospitalCol >= 0 ? row[hospitalCol] : '',
          patientId: patientIdCol >= 0 ? row[patientIdCol] : '',
          patientName: patientNameCol >= 0 ? row[patientNameCol] : '',
          birthdate: birthdateCol >= 0 ? row[birthdateCol] : ''
        };
      }
    }
    // マージ
    for (var j = 0; j < pdfList.length; j++) {
      var info = indexMap[pdfList[j].fileId];
      if (info) {
        pdfList[j].hospital = info.hospital || '';
        pdfList[j].patientId = info.patientId || '';
        pdfList[j].patientName = info.patientName || '';
        pdfList[j].birthdate = info.birthdate || '';
        pdfList[j].indexed = true;
      } else {
        pdfList[j].hospital = '';
        pdfList[j].patientId = '';
        pdfList[j].patientName = '';
        pdfList[j].birthdate = '';
        pdfList[j].indexed = false;
      }
    }
    return pdfList;
  } catch (err) {
    Logger.log('インデックスマージエラー: ' + err.toString());
    return pdfList;
  }
}

/**
 * 入力済みPDFフォルダのPDFインデックスデータ＋No.逆引き結果を返す
 * search_sample.html のPDFインデックス検索用
 */
function getDonePdfIndex() {
  var result = [];
  try {
    // ① 入力済みPDFフォルダのPDFファイル一覧取得
    var doneFolder = DriveApp.getFolderById(DONE_PDF_FOLDER_ID);
    var files = doneFolder.getFilesByType(MimeType.PDF);
    var doneFiles = {}; // fileName → fileId
    while (files.hasNext()) {
      var f = files.next();
      doneFiles[f.getName()] = f.getId();
    }

    // ② PDFインデックスシートから fileId→メタデータ のマップ作成
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var indexSheet = ss.getSheetByName('PDFインデックス');
    var indexMap = {}; // fileId → {hospital, patientId, patientName, birthdate}
    if (indexSheet && indexSheet.getLastRow() >= 2) {
      var indexData = indexSheet.getRange(1, 1, indexSheet.getLastRow(), indexSheet.getLastColumn()).getValues();
      var iHeaders = indexData[0];
      var iFileIdCol = iHeaders.indexOf('fileId');
      var iHospitalCol = iHeaders.indexOf('hospital');
      var iPatientIdCol = iHeaders.indexOf('patientId');
      var iPatientNameCol = iHeaders.indexOf('patientName');
      var iBirthdateCol = iHeaders.indexOf('birthdate');
      for (var i = 1; i < indexData.length; i++) {
        var row = indexData[i];
        var fId = iFileIdCol >= 0 ? row[iFileIdCol] : '';
        if (fId) {
          indexMap[fId] = {
            hospital: iHospitalCol >= 0 ? String(row[iHospitalCol] || '') : '',
            patientId: iPatientIdCol >= 0 ? String(row[iPatientIdCol] || '') : '',
            patientName: iPatientNameCol >= 0 ? String(row[iPatientNameCol] || '') : '',
            birthdate: iBirthdateCol >= 0 ? String(row[iBirthdateCol] || '') : ''
          };
        }
      }
    }

    // ③ 回答データ1シートから 元PDFファイル名→No. の逆引きマップ作成
    var sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var noCol = headers.indexOf('No.');
    var pdfCol = headers.indexOf('元PDFファイル名');
    var pdfToNo = {}; // fileName → No.
    if (noCol >= 0 && pdfCol >= 0 && sheet.getLastRow() >= 2) {
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      for (var j = 0; j < data.length; j++) {
        var no = data[j][noCol];
        var pdfName = String(data[j][pdfCol] || '').trim();
        if (no && pdfName) {
          pdfToNo[pdfName] = String(no);
        }
      }
    }

    // ④ マージ: インデックス済みのPDFのみ、No.付きで返す
    for (var fileName in doneFiles) {
      var fileId = doneFiles[fileName];
      var meta = indexMap[fileId];
      if (!meta) continue; // インデックスされていないPDFはスキップ
      var linkedNo = pdfToNo[fileName] || '';
      result.push({
        fileName: fileName,
        fileId: fileId,
        hospital: meta.hospital,
        patientId: meta.patientId,
        patientName: meta.patientName,
        birthdate: meta.birthdate,
        no: linkedNo
      });
    }

    // ファイル名でソート
    result.sort(function(a, b) { return a.fileName.localeCompare(b.fileName); });
  } catch (err) {
    Logger.log('getDonePdfIndex エラー: ' + err.toString());
  }
  return result;
}

/**
 * PDFインデックスにエントリを追加/更新
 * @param {Object} params - { fileId, fileName, hospital, patientId, patientName, birthdate, source }
 */
/**
 * index.jsonをGoogle DriveのスキャンPDFフォルダに保存
 * 既存ファイルがあれば上書き、なければ新規作成
 * @param {Object} indexData - インデックスデータ（オブジェクト）
 * @returns {Object} 結果
 */
function saveIndexJsonToDrive(indexData) {
  var folder = DriveApp.getFolderById(SCAN_PDF_FOLDER_ID);
  var jsonContent = JSON.stringify(indexData, null, 2);
  // 既存のindex.jsonを検索
  var files = folder.getFilesByName('index.json');
  if (files.hasNext()) {
    var file = files.next();
    file.setContent(jsonContent);
    return { action: 'updated', fileId: file.getId() };
  } else {
    var newFile = folder.createFile('index.json', jsonContent, 'application/json');
    return { action: 'created', fileId: newFile.getId() };
  }
}

function addPdfIndexEntry(params) {
  var fileId = params.fileId;
  var fileName = decodeURIComponent(params.fileName || '');
  var hospital = decodeURIComponent(params.hospital || '');
  var patientId = decodeURIComponent(params.patientId || '');
  var patientName = decodeURIComponent(params.patientName || '');
  var birthdate = decodeURIComponent(params.birthdate || '');
  var source = params.source || 'index_script';
  if (!fileId) {
    return { success: false, message: 'fileIdが必要です' };
  }
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('PDFインデックス');
  // シートがなければ作成
  if (!sheet) {
    sheet = ss.insertSheet('PDFインデックス');
    sheet.appendRow(['fileId', 'fileName', 'hospital', 'patientId', 'patientName', 'birthdate', 'indexedAt', 'source']);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }
  // 既存エントリを検索
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var fileIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < fileIds.length; i++) {
      if (fileIds[i][0] === fileId) {
        // 既存行を更新
        sheet.getRange(i + 2, 1, 1, 8).setValues([[
          fileId, fileName, hospital, patientId, patientName, birthdate,
          Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'), source
        ]]);
        return { success: true, action: 'updated' };
      }
    }
  }
  // 新規追加
  sheet.appendRow([
    fileId, fileName, hospital, patientId, patientName, birthdate,
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'), source
  ]);
  return { success: true, action: 'created' };
}
/**
 * Google Drive上のPDFをアンケートNo.に紐づけ、完了後にPDFを移動先フォルダへ移動
 * link_pdf.html の③ステップで使用
 * @param {Object} params - { no, row, fileId, fileName }
 * @returns {Object} 結果
 */
function linkPdfFromDrive(params) {
  var no = parseInt(params.no, 10);
  var fileId = params.fileId;
  var fileName = decodeURIComponent(params.fileName);
  if (!no || !fileId || !fileName) {
    return { success: false, message: 'パラメータが不足しています (no, fileId, fileName)' };
  }
  // 1. スプレッドシートの「元PDFファイル名」列を更新（既存関数を再利用）
  var updateResult = updatePdfFilename(no, fileName);
  if (!updateResult.success) {
    return updateResult;
  }
  // 1.5. 抜去位置が指定されていれば「抜去位置」列を更新
  var extractionMsg = '';
  var extractionPosition = params.extractionPosition ? decodeURIComponent(params.extractionPosition) : '';
  if (extractionPosition) {
    try {
      var extResult = updateExtractionPosition(no, extractionPosition);
      if (extResult && extResult.success) {
        extractionMsg = ' / 抜去位置: ' + extractionPosition;
      } else {
        extractionMsg = ' / （抜去位置: ' + (extResult.message || '更新失敗') + '）';
      }
    } catch (e) {
      extractionMsg = ' / （抜去位置更新: スキップ）';
    }
  }
  // 2. Source → OCR に変更 + 抜去位置・PDFファイル名をJSONに書き込み
  var sourceMsg = '';
  try {
    var sourceResult = updateSourceToOCR(no, {
      extractionPosition: extractionPosition,
      pdfFileName: fileName
    });
    if (sourceResult && sourceResult.success) {
      sourceMsg = ' / Source → OCR に変更';
    } else {
      sourceMsg = ' / （Source変更: 対象の _Form_ JSONが見つかりません）';
    }
  } catch (e) {
    sourceMsg = ' / （Source変更: スキップ）';
  }
  // 3. PDFファイルを入力済みPDFフォルダへ移動
  var file = DriveApp.getFileById(fileId);
  var destFolder = DriveApp.getFolderById(DONE_PDF_FOLDER_ID);
  file.moveTo(destFolder);
  return {
    success: true,
    message: 'No.' + no + ' に「' + fileName + '」を紐づけ、PDFを移動しました' + extractionMsg + sourceMsg
  };
}

/**
 * 抜去位置を更新
 * @param {number} targetNo - 対象のNo.
 * @param {string} extractionPosition - 抜去位置テキスト（例: "左上6"）
 * @returns {Object} 結果
 */
function updateExtractionPosition(targetNo, extractionPosition) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getLastRow();
  const noCol = headers.indexOf('No.');
  var extCol = headers.indexOf('抜去位置');
  if (noCol < 0) return { success: false, message: 'No.列が見つかりません' };
  // 「抜去位置」列がなければ自動作成
  if (extCol < 0) {
    addExtractionPositionColumn();
    // 列を再取得
    const newHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    extCol = newHeaders.indexOf('抜去位置');
    if (extCol < 0) return { success: false, message: '抜去位置列の作成に失敗しました' };
  }
  const noValues = sheet.getRange(2, noCol + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < noValues.length; i++) {
    if (Number(noValues[i][0]) === targetNo) {
      sheet.getRange(i + 2, extCol + 1).setValue(extractionPosition);
      Logger.log('抜去位置更新: No.' + targetNo + ' → ' + extractionPosition);
      return { success: true, message: '抜去位置を更新しました' };
    }
  }
  return { success: false, message: 'No.' + targetNo + ' が見つかりません' };
}
// ========== OCR照合（共有版）==========
/**
 * OCR照合結果フォルダを取得または作成
 * 共有ドライブ「3.データシート」フォルダ内に作成
 */
function getOrCreateOcrFolder() {
  var parentFolder = DriveApp.getFolderById(JSON_PARENT_FOLDER_ID);
  var folders = parentFolder.getFoldersByName(OCR_RESULTS_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(OCR_RESULTS_FOLDER_NAME);
}
/**
 * OCR結果JSONをGoogle Driveに保存（auto_ocr.pyから呼び出し）
 * @param {string} filename - ファイル名（例: "アンケート_001_ocr.json"）
 * @param {Object} data - OCR結果JSONデータ
 * @returns {Object} 結果
 */
function saveOcrResult(filename, data, overwrite) {
  var folder = getOrCreateOcrFolder();
  // 同名ファイルが既にあれば、overwrite指定時は上書き、それ以外はスキップ
  var existing = folder.getFilesByName(filename);
  if (existing.hasNext()) {
    var existFile = existing.next();
    if (overwrite) {
      // 上書き: 既存ファイルの内容を更新
      existFile.setContent(JSON.stringify(data, null, 2));
      existFile.setDescription(JSON.stringify({
        status: 'unverified',
        sourcePdf: data.metadata ? data.metadata.source_pdf : '',
        processedAt: data.metadata ? data.metadata.processed_at : new Date().toISOString()
      }));
      Logger.log('OCR結果上書き: ' + filename);
      return {
        fileId: existFile.getId(),
        fileName: filename,
        message: 'OCR結果を上書き保存しました',
        skipped: false,
        overwritten: true
      };
    }
    return {
      fileId: existFile.getId(),
      fileName: filename,
      message: '既存ファイルが見つかりました（スキップ）',
      skipped: true
    };
  }
  var jsonString = JSON.stringify(data, null, 2);
  var file = folder.createFile(filename, jsonString, MimeType.PLAIN_TEXT);
  // ファイルのDescriptionにステータスメタデータを保存
  file.setDescription(JSON.stringify({
    status: 'unverified',
    sourcePdf: data.metadata ? data.metadata.source_pdf : '',
    processedAt: data.metadata ? data.metadata.processed_at : new Date().toISOString()
  }));
  Logger.log('OCR結果保存: ' + filename);
  return {
    fileId: file.getId(),
    fileName: filename,
    message: 'OCR結果を保存しました',
    skipped: false
  };
}
/**
 * OCR照合結果ファイルの一覧を取得（メタデータのみ、高速）
 * ファイルのDescription欄からステータス情報を読み取る
 * @returns {Object} ファイル一覧
 */
function getOcrResultList() {
  var folder = getOrCreateOcrFolder();
  var files = folder.getFiles();
  var fileList = [];
  var now = new Date();
  while (files.hasNext()) {
    var file = files.next();
    var name = file.getName();
    if (!name.endsWith('.json')) continue;
    // Descriptionからメタデータを取得
    var desc = file.getDescription();
    var meta = {};
    try {
      meta = desc ? JSON.parse(desc) : {};
    } catch (e) {
      meta = {};
    }
    // ロックのタイムアウトチェック（30分）
    if (meta.status === 'in_progress' && meta.lockedAt) {
      var lockTime = new Date(meta.lockedAt);
      var elapsed = (now.getTime() - lockTime.getTime()) / (1000 * 60);
      if (elapsed > OCR_LOCK_TIMEOUT_MIN) {
        // タイムアウト → ロック解除
        meta.status = 'unverified';
        meta.lockedBy = null;
        meta.lockedAt = null;
        file.setDescription(JSON.stringify(meta));
      }
    }
    fileList.push({
      fileId: file.getId(),
      fileName: name,
      sourcePdf: meta.sourcePdf || '',
      processedAt: meta.processedAt || '',
      status: meta.status || 'unverified',
      lockedBy: meta.lockedBy || null,
      verifiedBy: meta.verifiedBy || null,
      verifiedAt: meta.verifiedAt || null,
      fileSize: file.getSize()
    });
  }
  // ファイル名でソート
  fileList.sort(function(a, b) {
    return a.fileName.localeCompare(b.fileName);
  });
  // ステータス別カウント
  var counts = { unverified: 0, in_progress: 0, verified: 0 };
  for (var i = 0; i < fileList.length; i++) {
    var s = fileList[i].status;
    if (counts[s] !== undefined) counts[s]++;
  }
  return {
    totalCount: fileList.length,
    counts: counts,
    files: fileList
  };
}
/**
 * OCR結果JSONの全文を取得（画像データ含む）
 * @param {string} fileId - Google DriveのファイルID
 * @returns {Object} JSON全文
 */
function getOcrResultDetail(fileId) {
  var file = DriveApp.getFileById(fileId);
  var content = file.getBlob().getDataAsString();
  return JSON.parse(content);
}
/**
 * OCR結果JSONからテキストデータのみ取得（画像を除外、JSONP向け軽量版）
 * @param {string} fileId - Google DriveのファイルID
 * @returns {Object} 画像を除いたJSON
 */
function getOcrResultText(fileId) {
  var file = DriveApp.getFileById(fileId);
  var content = file.getBlob().getDataAsString();
  var data = JSON.parse(content);
  // verification_imagesのキー一覧だけ返す（データは除外）
  var imageKeys = [];
  if (data.verification_images) {
    for (var key in data.verification_images) {
      imageKeys.push(key);
    }
    delete data.verification_images;
  }
  data._imageKeys = imageKeys;
  return data;
}
/**
 * OCR結果JSONから指定した質問の画像1枚を取得
 * @param {string} fileId - Google DriveのファイルID
 * @param {string} questionKey - 質問キー（例: "質問1_名前"）
 * @returns {Object} { key, imageData }
 */
function getOcrResultImage(fileId, questionKey) {
  var file = DriveApp.getFileById(fileId);
  var content = file.getBlob().getDataAsString();
  var data = JSON.parse(content);
  var images = data.verification_images || {};
  var decodedKey = decodeURIComponent(questionKey);
  return {
    key: decodedKey,
    imageData: images[decodedKey] || ''
  };
}
/**
 * OCR結果ファイルをロック（排他制御）
 * @param {string} fileId - Google DriveのファイルID
 * @param {string} user - ロックするユーザー名
 * @returns {Object} ロック結果
 */
function lockOcrResult(fileId, user) {
  var file = DriveApp.getFileById(fileId);
  var desc = file.getDescription();
  var meta = {};
  try {
    meta = desc ? JSON.parse(desc) : {};
  } catch (e) {
    meta = {};
  }
  // 既にロック済みかチェック
  if (meta.status === 'in_progress' && meta.lockedBy) {
    var lockTime = new Date(meta.lockedAt);
    var elapsed = (new Date().getTime() - lockTime.getTime()) / (1000 * 60);
    if (elapsed < OCR_LOCK_TIMEOUT_MIN && meta.lockedBy !== user) {
      return {
        locked: false,
        lockedBy: meta.lockedBy,
        message: meta.lockedBy + ' さんが照合中です（' + Math.round(elapsed) + '分前から）'
      };
    }
  }
  // 照合完了済みのファイルもロック可能（再照合用）
  meta.status = 'in_progress';
  meta.lockedBy = user;
  meta.lockedAt = new Date().toISOString();
  file.setDescription(JSON.stringify(meta));
  return {
    locked: true,
    lockedBy: user,
    message: 'ロックしました'
  };
}
/**
 * OCR結果ファイルのロックを解除
 * @param {string} fileId - Google DriveのファイルID
 * @returns {Object} 結果
 */
function unlockOcrResult(fileId) {
  var file = DriveApp.getFileById(fileId);
  var desc = file.getDescription();
  var meta = {};
  try {
    meta = desc ? JSON.parse(desc) : {};
  } catch (e) {
    meta = {};
  }
  if (meta.status !== 'verified') {
    meta.status = 'unverified';
  }
  meta.lockedBy = null;
  meta.lockedAt = null;
  file.setDescription(JSON.stringify(meta));
  return { message: 'ロックを解除しました' };
}
/**
 * 照合完了データを保存（verify_ocr.htmlから呼び出し）
 * 1. OCR結果JSONの verification_status を verified に更新
 * 2. スプレッドシートにデータを追加（既存addSurveyResponse再利用）
 * 3. 元PDFをSCANフォルダ→DONEフォルダへ移動
 * 4. ファイルDescriptionを更新
 * @param {string} fileId - OCR結果JSONのファイルID
 * @param {Object} data - 照合済みデータ（スプレッドシート列名: 値）
 * @returns {Object} 結果
 */
function saveVerifiedOcrResult(fileId, data) {
  var file = DriveApp.getFileById(fileId);
  // --- 1. スプレッドシートにデータを追加 ---
  var surveyResult = addSurveyResponse(data.surveyData);
  // --- 2. OCR結果JSONのステータスを更新 ---
  var content = file.getBlob().getDataAsString();
  var ocrJson = JSON.parse(content);
  if (ocrJson.metadata) {
    ocrJson.metadata.verification_status = 'verified';
    ocrJson.metadata.verified_by = data.verifiedBy || '';
    ocrJson.metadata.verified_at = new Date().toISOString();
  }
  file.setContent(JSON.stringify(ocrJson, null, 2));
  // --- 3. ファイルDescriptionを更新 ---
  var desc = file.getDescription();
  var meta = {};
  try {
    meta = desc ? JSON.parse(desc) : {};
  } catch (e) {
    meta = {};
  }
  meta.status = 'verified';
  meta.lockedBy = null;
  meta.lockedAt = null;
  meta.verifiedBy = data.verifiedBy || '';
  meta.verifiedAt = new Date().toISOString();
  file.setDescription(JSON.stringify(meta));
  // --- 4. 元PDFの移動（スキャンPDF→入力済みPDF） ---
  var movedPdf = false;
  var sourcePdf = ocrJson.metadata ? ocrJson.metadata.source_pdf : '';
  if (sourcePdf) {
    try {
      var scanFolder = DriveApp.getFolderById(SCAN_PDF_FOLDER_ID);
      var doneFolder = DriveApp.getFolderById(DONE_PDF_FOLDER_ID);
      var pdfFiles = scanFolder.getFilesByName(sourcePdf);
      if (pdfFiles.hasNext()) {
        pdfFiles.next().moveTo(doneFolder);
        movedPdf = true;
        Logger.log('元PDF移動: ' + sourcePdf + ' → 入力済みPDFフォルダ');
      }
    } catch (moveErr) {
      Logger.log('元PDF移動エラー: ' + moveErr.toString());
    }
  }
  Logger.log('照合完了保存: ' + file.getName() + ' (行' + surveyResult.row + ')');
  return {
    message: '照合データを保存しました（行' + surveyResult.row + '）',
    row: surveyResult.row,
    movedPdf: movedPdf,
    surveyResult: surveyResult
  };
}
// ========== PDFインデックス（Gemini OCR） ==========
/**
 * Google Drive上のPDFをGemini APIでOCRし、インデックスに登録
 * @param {string} fileId - Google DriveのファイルID
 * @returns {Object} OCR結果
 */
function indexPdfWithGemini(fileId) {
  if (!fileId) {
    return { success: false, message: 'fileIdが必要です' };
  }
  if (!GEMINI_API_KEY) {
    return { success: false, message: 'GEMINI_API_KEYがスクリプトプロパティに設定されていません' };
  }
  // PDFファイルを取得
  var file;
  try {
    file = DriveApp.getFileById(fileId);
  } catch (e) {
    return { success: false, message: 'ファイルが見つかりません: ' + e.toString() };
  }
  var fileName = file.getName();
  // PDFを画像に変換してBase64エンコード
  var base64Data;
  var mimeType;
  try {
    // Google DriveのサムネイルAPIでPDFを画像化（高解像度）
    var thumbnailUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w2000';
    var token = ScriptApp.getOAuthToken();
    var imgResponse = UrlFetchApp.fetch(thumbnailUrl, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (imgResponse.getResponseCode() === 200) {
      base64Data = Utilities.base64Encode(imgResponse.getContent());
      mimeType = imgResponse.getHeaders()['Content-Type'] || 'image/jpeg';
    } else {
      // フォールバック: PDFをそのまま送信
      var blob = file.getBlob();
      base64Data = Utilities.base64Encode(blob.getBytes());
      mimeType = 'application/pdf';
    }
  } catch (e) {
    // フォールバック: PDFをそのまま送信
    var blob = file.getBlob();
    base64Data = Utilities.base64Encode(blob.getBytes());
    mimeType = 'application/pdf';
  }
  // Gemini APIで4フィールドOCR
  var prompt = 'あなたは日本の医療アンケートのOCRアシスタントです。\n' +
    'このPDFの1ページ目から以下の4つの情報を読み取ってください。\n\n' +
    '1. 医療機関名: ページ右上に印刷されている医療機関名（歯科医院名）\n' +
    '2. 患者さんID: 医療機関名の左側にある手書きの数字・英数字\n' +
    '3. 名前: 「質問1」の回答欄に手書きカタカナで書かれた氏名（氏と名）\n' +
    '4. 生年月日: 「質問2」の回答欄にある元号（昭和・平成・令和）の丸囲みと手書きの年月日\n\n' +
    '以下のJSON形式のみで回答してください（説明不要）:\n' +
    '{\n' +
    '  "hospital": "医療機関名",\n' +
    '  "patientId": "患者ID",\n' +
    '  "patientName": "氏 名",\n' +
    '  "birthdate": "○○XX年YY月ZZ日"\n' +
    '}\n\n' +
    '注意:\n' +
    '- 読み取れない場合は空文字""にしてください\n' +
    '- 生年月日は「昭和35年12月16日」のような形式にしてください\n' +
    '- 名前は「テラニシ キョウコ」のように氏と名の間にスペースを入れてください\n' +
    '- JSONのみ出力し、```やマークダウンは不要です';
  var apiUrl = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + GEMINI_API_KEY;
  var payload = {
    contents: [{
      parts: [
        { text: prompt },
        {
          inline_data: {
            mime_type: mimeType,
            data: base64Data
          }
        }
      ]
    }],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 512
    }
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var response = UrlFetchApp.fetch(apiUrl, options);
  var responseCode = response.getResponseCode();
  if (responseCode !== 200) {
    return { success: false, message: 'Gemini API エラー (' + responseCode + '): ' + response.getContentText().substring(0, 200) };
  }
  var geminiResult = JSON.parse(response.getContentText());
  var textContent = '';
  try {
    textContent = geminiResult.candidates[0].content.parts[0].text;
  } catch (e) {
    return { success: false, message: 'Gemini APIレスポンス解析エラー' };
  }
  // JSONを抽出（```json ... ``` を除去、不正JSONの修復）
  var jsonText = textContent.trim();
  if (jsonText.indexOf('```json') >= 0) {
    jsonText = jsonText.split('```json')[1].split('```')[0].trim();
  } else if (jsonText.indexOf('```') >= 0) {
    jsonText = jsonText.split('```')[1].split('```')[0].trim();
  }
  // JSON部分のみ抽出（{ から最後の } まで）
  var firstBrace = jsonText.indexOf('{');
  var lastBrace = jsonText.lastIndexOf('}');
  if (firstBrace >= 0 && lastBrace > firstBrace) {
    jsonText = jsonText.substring(firstBrace, lastBrace + 1);
  }
  // 末尾カンマを除去（JSON仕様違反の修復）
  jsonText = jsonText.replace(/,\s*}/g, '}').replace(/,\s*]/g, ']');
  var ocrData;
  try {
    ocrData = JSON.parse(jsonText);
  } catch (e) {
    // 最終手段: 正規表現で各フィールドを抽出
    ocrData = {};
    var fields = ['hospital', 'patientId', 'patientName', 'birthdate'];
    for (var fi = 0; fi < fields.length; fi++) {
      var re = new RegExp('"' + fields[fi] + '"\\s*:\\s*"([^"]*)"');
      var m = jsonText.match(re);
      if (m) ocrData[fields[fi]] = m[1];
    }
    if (!ocrData.hospital && !ocrData.patientId && !ocrData.patientName) {
      return { success: false, message: 'OCR結果のJSON解析エラー: ' + jsonText.substring(0, 100) };
    }
  }
  // インデックスに保存
  var indexResult = addPdfIndexEntry({
    fileId: fileId,
    fileName: encodeURIComponent(fileName),
    hospital: encodeURIComponent(ocrData.hospital || ''),
    patientId: encodeURIComponent(ocrData.patientId || ''),
    patientName: encodeURIComponent(ocrData.patientName || ''),
    birthdate: encodeURIComponent(ocrData.birthdate || ''),
    source: 'gemini_gas'
  });
  return {
    success: true,
    fileName: fileName,
    ocrData: ocrData,
    indexAction: indexResult.action || 'unknown'
  };
}
/**
 * 未インデックスの全PDFをGemini OCRでインデックス構築（一括処理）
 * ※GASの実行時間制限（6分）に注意。タイムアウト前に中断して途中経過を返す
 * @returns {Object} 処理結果
 */
function indexAllPdfs() {
  if (!GEMINI_API_KEY) {
    return { success: false, message: 'GEMINI_API_KEYがスクリプトプロパティに設定されていません' };
  }
  var pdfList = getDrivePdfListWithIndex();
  var unindexed = [];
  for (var i = 0; i < pdfList.length; i++) {
    if (!pdfList[i].indexed) {
      unindexed.push(pdfList[i]);
    }
  }
  if (unindexed.length === 0) {
    return { success: true, message: '全てのPDFがインデックス済みです', processed: 0, total: pdfList.length };
  }
  var results = [];
  var startTime = new Date().getTime();
  var MAX_RUNTIME_MS = 5 * 60 * 1000; // 5分で安全に打ち切り
  for (var j = 0; j < unindexed.length; j++) {
    // タイムアウトチェック
    if (new Date().getTime() - startTime > MAX_RUNTIME_MS) {
      return {
        success: true,
        message: '実行時間制限のため中断しました。残りは再度実行してください。',
        processed: results.length,
        remaining: unindexed.length - j,
        total: pdfList.length,
        results: results
      };
    }
    var pdf = unindexed[j];
    try {
      var r = indexPdfWithGemini(pdf.fileId);
      results.push({
        fileName: pdf.fileName,
        success: r.success,
        ocrData: r.ocrData || null,
        message: r.message || ''
      });
    } catch (e) {
      results.push({
        fileName: pdf.fileName,
        success: false,
        message: e.toString()
      });
    }
    // Gemini APIレート制限対策（1秒待機）
    Utilities.sleep(1000);
  }
  return {
    success: true,
    message: '全ての未インデックスPDFを処理しました',
    processed: results.length,
    total: pdfList.length,
    results: results
  };
}
