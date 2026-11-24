function authorizeCheck() {
  const folder = DriveApp.getFolderById('17WWNmtUGGBeyIRrJr7Sayiw5prdBxFqg');
  const tempFile = folder.createFile('auth_test.txt', 'テスト', MimeType.PLAIN_TEXT);
  tempFile.setTrashed(true);
  console.log("認証完了：ファイルの作成が許可されました");
}

/**
 * RIKAKEN ASSET CLOUD Ver.3.10 (自動採番・履歴機能強化・表示バグ修正版)
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();
const UPLOAD_FOLDER_ID = '17WWNmtUGGBeyIRrJr7Sayiw5prdBxFqg';

const SHEET_NAMES = {
  MASTER: 'hub_master',
  INVENTORY: '什器管理',
  FACILITIES: '設備管理',
  CONSTRUCTION: '工事履歴',
  VENDORS: '業者一覧',
  DELETED: '消去済み',
  MOVED: '移動履歴'
};

// ▼ 修正：消去済みと移動履歴の列構成を新しいものに更新
const COLUMN_DEF = {
  '什器管理': [
    "会社名","拠点","管理番号","画像データ添付","フロア","部屋、場所","品目","メーカー","型番、シリアル","購入日","購入金額","業者","更新日時","予算対象要否","概算予算","サイズ","備考","関連書類URL","関連書類URL②","関連書類URL③","アラート有無","アラート日付","アラート表示開始日数","アラート繰り返し"
  ],
  '設備管理': [
    "会社名","拠点","管理番号","設備名","設置日","更新目安","フロア","部屋、場所","メーカー名","型番","業者","状態","予算対象要否","概算予算","更新日","備考","関連書類URL","関連書類URL②","関連書類URL③","アラート有無","アラート日付","アラート表示開始日数","アラート繰り返し"
  ],
  '工事履歴': [
    "会社名","拠点","管理番号","工事名","施工日","フロア","部屋、場所","業者","費用","関連書類URL","関連書類URL②","関連書類URL③","備考"
  ],
  '業者一覧': [
    "業者ID","会社名","担当者名","TEL","Mail","対応可能内容","名刺データ添付","更新日時","カテゴリー"
  ],
  '消去済み': [
    "消去日時","会社名","拠点","管理番号","消去理由","詳細データ抜粋"
  ],
  '移動履歴': [
    "移動日時","旧会社・拠点","新会社・拠点","管理番号","品名","移動理由","詳細データ抜粋"
  ]
};

const DATE_COLUMNS = ["購入日", "設置日", "施工日", "更新日", "更新日時", "アラート日付"];

/** hub_master の列（0始まり）。E列に2文字の拠点番号（例: TK, SI）がある場合、画像ルールで採番する */
const MASTER_COL_COUNTER = 3;
const MASTER_COL_HUB_CODE = 4;

const SHEET_ITEM_MASTER = '品目番号マスタ';
const SHEET_ROOM_MASTER = '部屋番号マスタ';

/** 管理番号表（CSV）由来のシード。拠点名 → 拠点コード */
const HUB_CODE_SEED_ = [
  ['HD東京（備品のみ）', 'HT'], ['HD名古屋（備品のみ）', 'HN'], ['HD賢島', 'KZ'],
  ['東京', 'TK'], ['名古屋', 'NG'], ['つくば', 'TB'], ['千葉', 'CB'], ['宇都宮', 'UT'],
  ['神奈川', 'KN'], ['鶴見', 'TR'], ['鎌倉', 'KM'], ['三島', 'MS'], ['静岡', 'SZ'],
  ['多摩', 'TM'], ['あきる野', 'AK'], ['岐阜', 'GF'], ['三重', 'ME'], ['大阪', 'OO'],
  ['四日市倉庫', 'YK'], ['並木', 'NM'], ['片岡', 'KT'], ['セイミ', 'SI'], ['鶴岡', 'ST'], ['南相馬', 'SM']
];

/** 品目番号マスタ シード [ラベル, 番号, 区分] */
const ITEM_CODE_SEED_ = [
  ['執務デスク', '1', '備品'], ['執務イス', '2', '備品'], ['キャビネット（収納棚）', '3', '備品'],
  ['会議室デスク', '4', '備品'], ['会議室イス', '5', '備品'], ['会議室モニター', '6', '備品'],
  ['大型モニター', '7', '備品'], ['ロッカー', '8', '備品'], ['リフ机', '9', '備品'],
  ['リフイス', '10', '備品'], ['袖机', '11', '備品'], ['加湿器、空気清浄機', '12', '備品'],
  ['シュレッダー', '13', '備品'], ['ZOOMカメラ', '14', '備品'], ['消火器', '15', '備品'],
  ['仕分け棚（購買荷物置き）', '16', '備品'], ['棚、台（物置）', '17', '備品'], ['受付用iPad', '18', '備品'],
  ['冷蔵庫', '19', '備品'], ['ＴＶ', '20', '備品'], ['TV', '20', '備品'], ['洗濯機', '21', '備品'],
  ['空調（室外機）', '60', '設備'], ['空調（室内機）', '61', '設備'], ['エレベーター', '62', '設備'],
  ['キュービクル', '63', '設備'], ['便器（ウオシュレット、センサー）', '64', '設備'], ['給水ポンプ', '65', '設備'],
  ['シャッター', '66', '設備'], ['照明機器', '67', '設備'], ['自動ドア', '68', '設備'],
  ['テレキューブ', '69', '設備'], ['防犯カメラ', '70', '設備'], ['電話機（主装置）', '71', '設備'],
  ['機械式駐車場（自社分）', '72', '設備'], ['壁掛けエアコン', '73', '設備'], ['給湯器', '74', '設備'],
  ['工事履歴', '0', '工事']
];

/** 部屋番号マスタ シード [部屋名, 番号] ※番号は半角英字 */
const ROOM_CODE_SEED_ = [
  ['トイレ男子', 'A'], ['トイレ女子', 'B'], ['給湯室', 'C'], ['リフレッシュスペース', 'D'],
  ['執務室', 'E'], ['倉庫①', 'F'], ['倉庫②', 'G'], ['サーバー室', 'H'], ['危険物倉庫', 'I'],
  ['更衣室', 'J'], ['会議室①', 'N'], ['会議室②', 'O'], ['会議室③', 'P'], ['会議室④', 'Q'],
  ['会議室⑤', 'R'], ['会議室⑥', 'S'], ['個室①', 'V'], ['個室②', 'W'],
  ['本部長室', 'X'], ['専務室', 'Y'], ['社長室', 'Z'],
  ['会議室', 'N']
];

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('RIKAKEN ASSET CLOUD')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getInitData() {
  syncSheetStructure_();
  const masterData = getCompanyStructureFromSheet();
  const numbering = getNumberingMasterMaps_();
  return {
    structure: masterData.structure,
    drawingUrls: masterData.drawingUrls,
    hubCodes: masterData.hubCodes,
    roomCodes: numbering.roomCodes,
    itemCodes: numbering.itemCodes,
    itemOptions: numbering.itemOptions,
    roomOptions: numbering.roomOptions,
    columns: COLUMN_DEF,
    dateColumns: DATE_COLUMNS,
    alertLeadDaysOptions: ALERT_LEAD_DAYS_OPTIONS
  };
}

/** スプレッドシートの列構成を COLUMN_DEF に合わせる（冪等） */
function syncSheetStructure_() {
  [SHEET_NAMES.INVENTORY, SHEET_NAMES.FACILITIES, SHEET_NAMES.CONSTRUCTION].forEach(ensureDocUrlColumns_);
  ensureHistoryColumnRemoved_(SHEET_NAMES.MOVED, '元シート');
  ensureHistoryColumnRemoved_(SHEET_NAMES.DELETED, '元シート');
  [SHEET_NAMES.MOVED, SHEET_NAMES.DELETED].forEach(ensureHeaderRow_);
  syncNumberingMasters_();
}

function ensureDocUrlColumns_(sheetName) {
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet) return;
  const expected = COLUMN_DEF[sheetName];
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  const urlIdx = headers.indexOf('関連書類URL');
  if (urlIdx === -1) {
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    return;
  }
  if (!headers.includes('関連書類URL②')) {
    sheet.insertColumnsAfter(urlIdx + 1, 2);
  }
  sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
}

function ensureHistoryColumnRemoved_(sheetName, colName) {
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet || sheet.getLastColumn() < 1) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = headers.findIndex(h => String(h || '').trim() === colName);
  if (idx > -1) sheet.deleteColumn(idx + 1);
}

function ensureHeaderRow_(sheetName) {
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet) return;
  const expected = COLUMN_DEF[sheetName];
  sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
}

/** 全角英字 → 半角大文字 */
function toHalfWidthAlpha_(s) {
  return String(s || '')
    .replace(/[Ａ-Ｚａ-ｚ]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .toUpperCase();
}

/**
 * 番号マスタの初期化のみ（シートを正とする）
 * - 既にデータがあるシートは一切上書きしない
 * - シートが無い／データ行が空のときだけシードで初期作成
 * - hub_master E列も、空のときだけシードで埋める（既存値は維持）
 */
function syncNumberingMasters_() {
  fillEmptyHubCodesFromSeed_();
  ensureMasterSheetFromSeed_(SHEET_ITEM_MASTER, ['ラベル', '番号', '区分'], ITEM_CODE_SEED_);
  ensureMasterSheetFromSeed_(SHEET_ROOM_MASTER, ['部屋名', '番号'], ROOM_CODE_SEED_);
}

function fillEmptyHubCodesFromSeed_() {
  const sheet = SS.getSheetByName(SHEET_NAMES.MASTER);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const seedMap = {};
  HUB_CODE_SEED_.forEach(pair => {
    seedMap[String(pair[0]).trim()] = String(pair[1]).trim().toUpperCase();
  });

  const lastCol = Math.max(sheet.getLastColumn(), MASTER_COL_HUB_CODE + 1);
  if (sheet.getLastColumn() < MASTER_COL_HUB_CODE + 1) {
    sheet.getRange(1, MASTER_COL_HUB_CODE + 1).setValue('拠点番号');
  }

  const data = sheet.getRange(2, 1, lastRow, lastCol).getValues();
  for (let i = 0; i < data.length; i++) {
    const hubName = String(data[i][1] || '').trim();
    if (!hubName) continue;
    const current = data[i][MASTER_COL_HUB_CODE] !== undefined && data[i][MASTER_COL_HUB_CODE] !== null
      ? String(data[i][MASTER_COL_HUB_CODE]).trim()
      : '';
    // シートを正とする：既に入っている拠点コードは上書きしない
    if (current) continue;

    let code = seedMap[hubName] || '';
    if (!code) {
      const hit = Object.keys(seedMap).find(k => hubName === k || hubName.indexOf(k) === 0 || k.indexOf(hubName) === 0);
      if (hit) code = seedMap[hit];
    }
    if (!code) continue;
    sheet.getRange(i + 2, MASTER_COL_HUB_CODE + 1).setValue(code);
  }
}

/** シートが無い／データが空のときだけシードで初期化。既存データは触らない */
function ensureMasterSheetFromSeed_(sheetName, headers, seedRows) {
  let sheet = SS.getSheetByName(sheetName);
  if (!sheet) {
    sheet = SS.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    seedRows.forEach(seed => {
      const label = String(seed[0] || '').trim();
      if (!label) return;
      const rowVals = seed.map((v, i) => {
        if (i === 1 && headers[1] === '番号') {
          const raw = String(v || '').trim();
          return /[A-Za-zＡ-Ｚａ-ｚ]/.test(raw) ? toHalfWidthAlpha_(raw) : raw;
        }
        return String(v || '').trim();
      });
      sheet.appendRow(rowVals);
    });
    return;
  }

  // 既存シート: ヘッダーが空なら整えるが、データ行は一切変更しない
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  // lastRow >= 2 → シートにデータあり＝正。シードでの上書き・追記はしない
  if (lastRow >= 2) return;

  // ヘッダーのみ（データなし）のときだけ初期シードを入れる
  if (lastRow === 1) {
    const headerCell = String(sheet.getRange(1, 1).getValue() || '').trim();
    if (!headerCell) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  seedRows.forEach(seed => {
    const label = String(seed[0] || '').trim();
    if (!label) return;
    const rowVals = seed.map((v, i) => {
      if (i === 1 && headers[1] === '番号') {
        const raw = String(v || '').trim();
        return /[A-Za-zＡ-Ｚａ-ｚ]/.test(raw) ? toHalfWidthAlpha_(raw) : raw;
      }
      return String(v || '').trim();
    });
    sheet.appendRow(rowVals);
  });
}

function getNumberingMasterMaps_() {
  const roomCodes = {};
  const roomOptions = [];
  const roomSheet = SS.getSheetByName(SHEET_ROOM_MASTER);
  if (roomSheet && roomSheet.getLastRow() >= 2) {
    const vals = roomSheet.getRange(2, 1, roomSheet.getLastRow(), 2).getValues();
    vals.forEach(row => {
      const name = String(row[0] || '').trim();
      const code = toHalfWidthAlpha_(String(row[1] || '').trim());
      if (!name || !code) return;
      roomCodes[name] = code;
      roomOptions.push({ name: name, code: code });
    });
  }

  const itemCodes = { '備品': {}, '設備': {}, '工事': {} };
  const itemOptions = { '備品': [], '設備': [], '工事': [] };
  const itemSheet = SS.getSheetByName(SHEET_ITEM_MASTER);
  if (itemSheet && itemSheet.getLastRow() >= 2) {
    const vals = itemSheet.getRange(2, 1, itemSheet.getLastRow(), 3).getValues();
    vals.forEach(row => {
      const label = String(row[0] || '').trim();
      const code = String(row[1] || '').trim();
      let cat = String(row[2] || '').trim();
      if (!label || code === '') return;
      if (cat !== '備品' && cat !== '設備' && cat !== '工事') {
        const n = Number(code);
        cat = (!isNaN(n) && n >= 60) ? '設備' : (code === '0' ? '工事' : '備品');
      }
      itemCodes[cat][label] = code;
      itemOptions[cat].push({ label: label, code: code });
    });
  }

  return { roomCodes, roomOptions, itemCodes, itemOptions };
}

function itemCategoryForSheet_(sheetName) {
  if (sheetName === SHEET_NAMES.FACILITIES) return '設備';
  if (sheetName === SHEET_NAMES.CONSTRUCTION) return '工事';
  return '備品';
}

function findSheetByAssetId(id) {
  const idStr = String(id || '').trim();
  if (!idStr) return '';
  const targets = [SHEET_NAMES.INVENTORY, SHEET_NAMES.FACILITIES, SHEET_NAMES.CONSTRUCTION];
  for (let t = 0; t < targets.length; t++) {
    const sheetName = targets[t];
    const sheet = SS.getSheetByName(sheetName);
    if (!sheet) continue;
    const headers = COLUMN_DEF[sheetName];
    const idIdx = headers.findIndex(h => h.includes('管理番号'));
    if (idIdx < 0) continue;
    const data = sheet.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]).trim() === idStr) return sheetName;
    }
  }
  return '';
}

function uploadFileToDrive(data, fileName, mimeType) {
  try {
    const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
    const blob = Utilities.newBlob(Utilities.base64Decode(data), mimeType, fileName);
    const file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) { }
    return file.getUrl();
  } catch (e) {
    throw new Error("アップロード失敗: " + e.message);
  }
}

const ALERT_LEAD_DAYS_OPTIONS = [30, 90, 120, 150];

function normalizeAlertLeadDays_(val) {
  const n = Number(String(val == null ? '' : val).trim());
  if (ALERT_LEAD_DAYS_OPTIONS.indexOf(n) > -1) return n;
  return 30;
}

function normalizeAlertRepeat_(val) {
  return String(val || '').trim() === '毎年' ? '毎年' : 'なし';
}

function parseAlertDate_(dateVal) {
  if (!dateVal) return null;
  const dStr = String(dateVal).replace(/-/g, '/');
  const d = new Date(dStr);
  if (isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}

function formatAlertDate_(dateObj) {
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

function isAlertInDisplayWindow_(targetDate, leadDays, today) {
  if (!targetDate) return false;
  const start = new Date(targetDate.getTime());
  start.setDate(start.getDate() - leadDays);
  return today >= start;
}

/**
 * filter:
 *  - active … 通知ONかつ表示期間内（TOPバナー用）
 *  - off    … 通知OFFで予定日あり
 *  - all    … 予定日あり（ON/OFF問わず）
 */
function collectAlertRecords_(filter) {
  const mode = filter || 'active';
  const targets = [SHEET_NAMES.INVENTORY, SHEET_NAMES.FACILITIES];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  let alerts = [];

  targets.forEach(sheetName => {
    const sheet = SS.getSheetByName(sheetName);
    if (!sheet) return;
    const headers = COLUMN_DEF[sheetName];
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return;

    const alertFlagIdx = headers.indexOf('アラート有無');
    const alertDateIdx = headers.indexOf('アラート日付');
    const leadIdx = headers.indexOf('アラート表示開始日数');
    const repeatIdx = headers.indexOf('アラート繰り返し');
    const idIdx = headers.findIndex(h => h.includes('管理番号'));
    const companyIdx = headers.indexOf('会社名');
    const hubIdx = headers.indexOf('拠点');
    const nameIdx = sheetName === '設備管理' ? headers.indexOf('設備名') : headers.indexOf('品目');
    if (alertFlagIdx === -1 || alertDateIdx === -1) return;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const flag = String(row[alertFlagIdx] || '').trim();
      const dateRaw = row[alertDateIdx];
      if (!dateRaw) continue;

      const targetDate = parseAlertDate_(dateRaw);
      if (!targetDate) continue;
      const leadDays = leadIdx > -1 ? normalizeAlertLeadDays_(row[leadIdx]) : 30;
      const repeat = repeatIdx > -1 ? normalizeAlertRepeat_(row[repeatIdx]) : 'なし';
      const inWindow = isAlertInDisplayWindow_(targetDate, leadDays, today);
      const isOn = flag === '有';

      if (mode === 'active' && !(isOn && inWindow)) continue;
      if (mode === 'off' && !(flag === '無')) continue;
      if (mode === 'all' && !(isOn || flag === '無')) continue;

      alerts.push({
        sheet: sheetName,
        company: row[companyIdx],
        hub: row[hubIdx],
        id: row[idIdx],
        name: nameIdx > -1 ? row[nameIdx] : '',
        date: formatAlertDate_(targetDate),
        leadDays: leadDays,
        repeat: repeat,
        flag: isOn ? '有' : '無',
        isOverdue: targetDate < today,
        inWindow: inWindow
      });
    }
  });

  alerts.sort((a, b) => {
    if (a.isOverdue !== b.isOverdue) return a.isOverdue ? -1 : 1;
    return String(a.date).localeCompare(String(b.date));
  });
  return alerts;
}

/** TOPバナー用：通知ONかつ表示期間内のみ */
function getGlobalAlerts() {
  return collectAlertRecords_('active');
}

/** アラート管理用 */
function getAlertManagementList(filter) {
  const mode = filter === 'off' || filter === 'all' ? filter : 'active';
  return collectAlertRecords_(mode);
}

function findAssetRowById_(sheetName, id) {
  const sheet = SS.getSheetByName(sheetName);
  if (!sheet) throw new Error('シートが見つかりません: ' + sheetName);
  const headers = COLUMN_DEF[sheetName];
  const idIdx = headers.findIndex(h => h.includes('管理番号'));
  if (idIdx < 0) throw new Error('管理番号列がありません');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]).trim() === String(id).trim()) {
      return { sheet, headers, rowIndex: i + 1, row: data[i], idIdx };
    }
  }
  throw new Error('対象データが見つかりません: ' + id);
}

function setCellByHeader_(found, headerName, value) {
  const col = found.headers.indexOf(headerName);
  if (col < 0) throw new Error('列がありません: ' + headerName);
  found.sheet.getRange(found.rowIndex, col + 1).setValue(value);
}

/** action: deferDays | deferDate | completeYear | setFlag */
function updateAlertRecord(payload) {
  const sheetName = payload.sheetName;
  const id = payload.id;
  const action = payload.action;
  const found = findAssetRowById_(sheetName, id);
  const dateIdx = found.headers.indexOf('アラート日付');
  const flagIdx = found.headers.indexOf('アラート有無');
  if (dateIdx < 0 || flagIdx < 0) throw new Error('アラート列がありません');

  if (action === 'deferDays') {
    const days = Number(payload.days);
    if (![7, 14, 30, 60].includes(days)) throw new Error('延期日数が不正です');
    const current = parseAlertDate_(found.row[dateIdx]) || new Date();
    current.setHours(0, 0, 0, 0);
    current.setDate(current.getDate() + days);
    setCellByHeader_(found, 'アラート日付', formatAlertDate_(current));
    setCellByHeader_(found, 'アラート有無', '有');
  } else if (action === 'deferDate') {
    const next = parseAlertDate_(payload.date);
    if (!next) throw new Error('日付が不正です');
    setCellByHeader_(found, 'アラート日付', formatAlertDate_(next));
    setCellByHeader_(found, 'アラート有無', '有');
  } else if (action === 'completeYear') {
    const current = parseAlertDate_(found.row[dateIdx]);
    if (!current) throw new Error('通知予定日がありません');
    current.setFullYear(current.getFullYear() + 1);
    setCellByHeader_(found, 'アラート日付', formatAlertDate_(current));
    setCellByHeader_(found, 'アラート有無', '有');
  } else if (action === 'setFlag') {
    const flag = payload.flag === '有' ? '有' : '無';
    setCellByHeader_(found, 'アラート有無', flag);
  } else {
    throw new Error('不明な操作です');
  }

  SpreadsheetApp.flush();
  return getAlertManagementList(payload.refreshFilter || 'active');
}

function getCompanyStructureFromSheet() {
  try {
    const sheet = SS.getSheetByName(SHEET_NAMES.MASTER);
    if (!sheet) return { structure: {}, drawingUrls: {}, hubCodes: {} };
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { structure: {}, drawingUrls: {}, hubCodes: {} };

    const lastCol = Math.max(sheet.getLastColumn(), MASTER_COL_HUB_CODE + 1);
    const data = sheet.getRange(2, 1, lastRow, lastCol).getValues();
    const structure = {};
    const drawingUrls = {};
    const hubCodes = {};

    data.forEach(row => {
      const company = row[0] ? String(row[0]).trim() : "";
      const hub = row[1] ? String(row[1]).trim() : "";
      const url = row[2] ? String(row[2]).trim() : "";
      const hubCode = row[MASTER_COL_HUB_CODE] !== undefined && row[MASTER_COL_HUB_CODE] !== null
        ? String(row[MASTER_COL_HUB_CODE]).trim()
        : "";

      if (company && hub) {
        if (!structure[company]) structure[company] = [];
        if (!structure[company].includes(hub)) structure[company].push(hub);

        if (url) {
          drawingUrls[`${company}_${hub}`] = url;
        }
        if (hubCode) {
          hubCodes[`${company}_${hub}`] = hubCode;
        }
      }
    });
    return { structure, drawingUrls, hubCodes };
  } catch (e) { return { structure: {}, drawingUrls: {}, hubCodes: {} }; }
}

function getSheetData(hubName, sheetType) {
  try {
    const sheet = SS.getSheetByName(sheetType);
    if (!sheet) return { headers: [], rows: [] };
    const headers = COLUMN_DEF[sheetType] || [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { headers: headers, rows: [] };
    const data = sheet.getDataRange().getDisplayValues();
    const rawRows = data.slice(1);
    
    // ▼ 修正：消去済みと移動履歴は、シートのデータをそのまま画面に返すように変更
    if (sheetType === SHEET_NAMES.DELETED || sheetType === SHEET_NAMES.MOVED) {
      return { headers: headers, rows: rawRows };
    }

    let filteredRows = [];
    if (sheetType === SHEET_NAMES.VENDORS) {
      filteredRows = rawRows;
    } else if (hubName) {
      const hubTrim = String(hubName).trim();
      filteredRows = rawRows.filter(row => String(row[1]).trim() === hubTrim);
    } else {
      let sortIdx = headers.findIndex(h => h.includes('更新') || h.includes('施工'));
      let sortedRows = rawRows.slice();
      if (sortIdx > -1) {
        sortedRows.sort((a, b) => {
          const dateA = new Date(a[sortIdx].replace(/-/g, '/'));
          const dateB = new Date(b[sortIdx].replace(/-/g, '/'));
          const timeA = isNaN(dateA) ? 0 : dateA.getTime();
          const timeB = isNaN(dateB) ? 0 : dateB.getTime();
          return timeB - timeA; 
        });
      }
      filteredRows = sortedRows.slice(0, 20);
    }
    return { headers: headers, rows: filteredRows };
  } catch (e) {
    throw new Error("データ取得エラー: " + e.message);
  }
}

function normalizeFloorToken_(floorVal) {
  const s = String(floorVal || '').trim();
  if (!s) return '';
  const m = s.match(/(\d+)/);
  if (m) return m[1];
  return cleanAssetIdToken_(s);
}

function normalizeRoomToken_(roomVal) {
  const s = String(roomVal || '').trim();
  if (!s) return '';
  const mapped = lookupRoomCode_(s);
  if (mapped) return mapped;
  const half = toHalfWidthAlpha_(s);
  const ascii = half.match(/[A-Z0-9]/);
  if (ascii) return ascii[0];
  return '';
}

function lookupRoomCode_(roomVal) {
  const raw = String(roomVal || '').trim();
  if (!raw) return '';
  const maps = getNumberingMasterMaps_();
  if (maps.roomCodes[raw]) return maps.roomCodes[raw];

  // 「会議室」→ 会議室① など: 前方一致・部分一致（長いラベル優先）
  let best = '';
  let bestLen = -1;
  Object.keys(maps.roomCodes).forEach(name => {
    if (name === '会議室' && raw !== '会議室' && !/^会議室$/.test(raw)) {
      // 汎用「会議室」は exact 以外では後回し
    }
    if (raw === name || name.indexOf(raw) === 0 || raw.indexOf(name) === 0 || name.indexOf(raw) > -1 || raw.indexOf(name) > -1) {
      if (name.length > bestLen) {
        bestLen = name.length;
        best = maps.roomCodes[name];
      }
    }
  });
  if (best) return best;

  // 入力がすでに A/Ｎ など1文字の部屋記号
  const one = toHalfWidthAlpha_(raw);
  if (/^[A-Z]$/.test(one)) return one;
  return '';
}

/**
 * 品目・設備・工事名から項目番号を決める
 * 1) 先頭数字 2) 品目番号マスタ（完全一致→部分一致） 3) 工事シートは 0
 */
function resolveItemCategoryCode_(form, headers, sheetName) {
  let nameIdx = -1;
  if (sheetName === '設備管理') nameIdx = headers.indexOf('設備名');
  else if (sheetName === '工事履歴') nameIdx = headers.indexOf('工事名');
  else nameIdx = headers.indexOf('品目');
  if (nameIdx < 0) {
    nameIdx = headers.findIndex(h => h === '品目' || h === '設備名' || h === '工事名');
  }

  const raw = nameIdx >= 0 ? String(form['col_' + nameIdx] || '').trim() : '';
  const leadNum = raw.match(/^(\d+)/);
  if (leadNum) return leadNum[1];

  if (sheetName === SHEET_NAMES.CONSTRUCTION) {
    const fromMaster = lookupItemCode_(raw || '工事履歴', '工事');
    return fromMaster || '0';
  }

  const cat = itemCategoryForSheet_(sheetName);
  const roomIdx = headers.indexOf('部屋、場所');
  const roomName = roomIdx > -1 ? String(form['col_' + roomIdx] || '').trim() : '';
  const code = lookupItemCode_(raw, cat, roomName);
  return code || '0';
}

function lookupItemCode_(rawName, category, roomName) {
  const raw = String(rawName || '').trim();
  if (!raw) return '';
  const maps = getNumberingMasterMaps_();
  const dict = maps.itemCodes[category] || {};
  if (dict[raw]) return String(dict[raw]);

  // 部屋が会議室系なら「会議室○○」を優先
  const roomHint = String(roomName || '');
  if (/会議/.test(roomHint) || /会議/.test(raw)) {
    const preferred = Object.keys(dict)
      .filter(k => k.indexOf('会議') > -1 && (k.indexOf(raw) > -1 || raw.indexOf(k.replace(/^会議室/, '')) > -1 || k.indexOf(raw) > -1))
      .sort((a, b) => b.length - a.length);
    // イス → 会議室イス / 机・デスク → 会議室デスク
    const aliases = [];
    if (/イス|椅子|チェア/i.test(raw)) aliases.push('会議室イス');
    if (/机|デスク/i.test(raw)) aliases.push('会議室デスク');
    if (/モニタ/i.test(raw)) aliases.push('会議室モニター');
    for (let i = 0; i < aliases.length; i++) {
      if (dict[aliases[i]]) return String(dict[aliases[i]]);
    }
    if (preferred.length) return String(dict[preferred[0]]);
  }

  // 部分一致（長いラベル優先）
  let bestLabel = '';
  Object.keys(dict).forEach(label => {
    if (label === raw) {
      bestLabel = label;
      return;
    }
    if (label.indexOf(raw) > -1 || raw.indexOf(label) > -1) {
      if (label.length > bestLabel.length) bestLabel = label;
    }
  });
  if (bestLabel) return String(dict[bestLabel]);

  // 別名ゆるめマッチ
  if (/イス|椅子/.test(raw)) {
    const hit = Object.keys(dict).filter(k => /イス|椅子/.test(k)).sort((a, b) => a.length - b.length);
    if (hit.length === 1) return String(dict[hit[0]]);
  }
  return '';
}

/** 管理番号用に空白・ハイフンを除去して大文字化 */
function cleanAssetIdToken_(val) {
  return String(val || '').trim().replace(/[\s\-]/g, '').toUpperCase();
}

function lookupHubCode_(companyName, hubName) {
  const companyNameT = String(companyName || '').trim();
  const hubNameT = String(hubName || '').trim();
  const masterSheet = SS.getSheetByName(SHEET_NAMES.MASTER);
  if (!masterSheet) return '';
  const data = masterSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowCompany = String(data[i][0] || '').trim();
    const rowHub = String(data[i][1] || '').trim();
    if (rowCompany !== companyNameT || rowHub !== hubNameT) continue;
    const hubCodeRaw = data[i][MASTER_COL_HUB_CODE];
    return hubCodeRaw !== undefined && hubCodeRaw !== null ? String(hubCodeRaw).trim() : '';
  }
  return '';
}

/**
 * 入力値から連番付与前のベースを取り出す。
 * 例: TK4N5-001 → TK4N5 / TK4N5001 → TK4N5 / TK4N5 → TK4N5
 */
function extractAssetIdBase_(rawId) {
  const s = String(rawId || '').trim().toUpperCase().replace(/\s+/g, '');
  if (!s) return '';
  const hyphen = s.match(/^(.+)-(\d{1,3})$/);
  if (hyphen && hyphen[1]) return cleanAssetIdToken_(hyphen[1]);
  const cleaned = cleanAssetIdToken_(s);
  const m = cleaned.match(/^(.*?)(\d{3})$/);
  if (m && m[1]) return m[1];
  return cleaned;
}

/**
 * 同じベース（例: TK4N5）に対する次の3桁連番を返す。
 * 既存 ID が TK4N5-001 / TK4N5001 なら次は 2
 */
function nextSerial3ForExactBase_(sheet, idCol1Based, base) {
  const baseU = cleanAssetIdToken_(base);
  if (!baseU) return 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  const values = sheet.getRange(2, idCol1Based, lastRow, 1).getDisplayValues();
  let max = 0;
  for (let i = 0; i < values.length; i++) {
    const raw = String(values[i][0] || '').trim().toUpperCase().replace(/\s+/g, '');
    if (!raw) continue;
    if (raw.indexOf(baseU + '-') === 0) {
      const rest = raw.slice(baseU.length + 1);
      if (/^\d{1,3}$/.test(rest)) {
        const n = Number(rest);
        if (n > max) max = n;
      }
      continue;
    }
    const idU = cleanAssetIdToken_(raw);
    if (idU.indexOf(baseU) === 0) {
      const rest = idU.slice(baseU.length);
      if (/^\d{3}$/.test(rest)) {
        const n = Number(rest);
        if (n > max) max = n;
      }
    }
  }
  return max + 1;
}

/**
 * ベース + '-' + 末尾3桁連番
 * 例: TK4N5 → TK4N5-001
 */
function generateIdFromBaseWithSerial_(sheet, idIndex0, rawBaseId) {
  const base = extractAssetIdBase_(rawBaseId);
  if (!base) {
    throw new Error('管理番号のベース（例: TK4N5）を入力してください。');
  }
  const next = nextSerial3ForExactBase_(sheet, idIndex0 + 1, base);
  const serial3 = ('000' + next).slice(-3);
  return base + '-' + serial3;
}

/**
 * 管理番号が空のとき、拠点コード・フロア・部屋・品目番号からベースを組み立てる。
 * 例: TK + 4 + N + 5 → TK4N5
 */
function buildAssetBaseFromForm_(companyName, hubName, form, headers, sheetName) {
  const hubCode = lookupHubCode_(companyName, hubName);
  const floorIdx = headers.indexOf('フロア');
  const roomIdx = headers.indexOf('部屋、場所');
  const floorPart = floorIdx > -1 ? normalizeFloorToken_(form['col_' + floorIdx]) : '';
  const roomRaw = roomIdx > -1 ? String(form['col_' + roomIdx] || '').trim() : '';
  const roomPart = roomIdx > -1 ? (lookupRoomCode_(roomRaw) || normalizeRoomToken_(roomRaw)) : '';
  const itemPart = cleanAssetIdToken_(resolveItemCategoryCode_(form, headers, sheetName));

  if (!hubCode) {
    throw new Error('拠点コードが hub_master の E列に登録されていないか、会社名・拠点が一致しません。管理番号を手入力するか、マスタを確認してください。');
  }
  if (!floorPart) {
    throw new Error('フロアを入力するか、管理番号にベース（例: TK4N5）を入力してください。');
  }
  if (!roomPart) {
    throw new Error('部屋、場所を入力（または部屋番号マスタに登録）するか、管理番号にベース（例: TK4N5）を入力してください。');
  }

  return cleanAssetIdToken_(hubCode) + floorPart + roomPart + (itemPart || '0');
}

function saveData(form) {
  const sheetName = form.sheetName;
  const sheet = SS.getSheetByName(sheetName);
  const id = form.dataId; 
  const isNew = (!id || id === "");
  const headers = COLUMN_DEF[sheetName];
  const idIndex = headers.findIndex(h => h.includes('管理番号') || h.includes('業者ID'));
  
  const userProvidedId = form['col_' + idIndex] ? form['col_' + idIndex].trim() : "";
  
  let newId = id;
  if (isNew) {
    const companyIdx = headers.indexOf("会社名");
    const hubIdx = headers.indexOf("拠点");
    const isAssetSheet = companyIdx > -1 && hubIdx > -1 && headers.some(h => String(h).includes('管理番号'));

    if (isAssetSheet) {
      const lock = LockService.getScriptLock();
      if (!lock.tryLock(10000)) {
        throw new Error("現在システムが混み合っています。数秒待ってから再度お試しください。");
      }
      try {
        let base = userProvidedId;
        if (!base) {
          const compName = form['col_' + companyIdx];
          const hubName = form['col_' + hubIdx];
          if (!compName || !hubName) {
            throw new Error('会社名と拠点を選択してください。');
          }
          base = buildAssetBaseFromForm_(compName, hubName, form, headers, sheetName);
        }
        newId = generateIdFromBaseWithSerial_(sheet, idIndex, base);
      } finally {
        lock.releaseLock();
      }
    } else if (userProvidedId) {
      newId = userProvidedId;
    } else {
      const prefixMap = {'業者一覧': 'V'};
      const prefix = prefixMap[sheetName] || 'X';
      newId = prefix + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmmss');
    }
  } else {
    newId = userProvidedId || id;
  }

  let rowData = [];
  headers.forEach((h, i) => {
    let val = form['col_' + i] || "";
    if (i === idIndex) {
      rowData.push(newId);
    } else if (h === '更新日' || h === '更新日時') {
      const format = h === '更新日時' ? 'yyyy/MM/dd HH:mm:ss' : 'yyyy/MM/dd';
      rowData.push(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), format));
    } else if (h === 'アラート表示開始日数') {
      rowData.push(String(normalizeAlertLeadDays_(val)));
    } else if (h === 'アラート繰り返し') {
      rowData.push(normalizeAlertRepeat_(val));
    } else {
      if (DATE_COLUMNS.includes(h) && val.includes('-')) {
         val = val.replace(/-/g, '/');
      }
      rowData.push(val);
    }
  });

  if (isNew) {
    sheet.appendRow(rowData);
  } else {
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIndex]) === String(id)) { 
        sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
        found = true;
        break;
      }
    }
    if(!found) sheet.appendRow(rowData);
  }
  
  SpreadsheetApp.flush();
  Utilities.sleep(500);
  const hubName = form['col_1']; 
  return getSheetData(sheetName === '業者一覧' ? null : hubName, sheetName);
}

// ▼ 修正：移動理由・品名を含めるように更新
function moveAsset(form) {
  const sheetName = form.sheetName;
  const id = form.dataId;
  const newCompany = form.newCompany;
  const newHub = form.newHub;
  const moveDate = form.moveDate;
  const reason = form.reason;

  const sheet = SS.getSheetByName(sheetName);
  const moveSheet = SS.getSheetByName(SHEET_NAMES.MOVED);
  const headers = COLUMN_DEF[sheetName];
  
  const idIndex = headers.findIndex(h => h.includes('管理番号'));
  const companyIndex = headers.indexOf('会社名');
  const hubIndex = headers.indexOf('拠点');
  // 品名または設備名を取得するためのインデックス
  const nameIndex = headers.findIndex(h => h === '品目' || h === '設備名');
  const updateDateIndex = headers.findIndex(h => h === '更新日' || h === '更新日時');

  const data = sheet.getDataRange().getValues();
  let rowToMove = null;
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIndex]) === String(id)) {
      rowToMove = [...data[i]]; 
      rowIndex = i + 1;
      
      data[i][companyIndex] = newCompany;
      data[i][hubIndex] = newHub;
      
      if (updateDateIndex > -1) {
         const format = headers[updateDateIndex] === '更新日時' ? 'yyyy/MM/dd HH:mm:ss' : 'yyyy/MM/dd';
         data[i][updateDateIndex] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), format);
      }
      
      sheet.getRange(rowIndex, 1, 1, data[i].length).setValues([data[i]]);
      break;
    }
  }

  if (rowToMove && moveSheet) {
    const oldLocation = `${rowToMove[companyIndex]} / ${rowToMove[hubIndex]}`;
    const newLocation = `${newCompany} / ${newHub}`;
    const itemName = nameIndex > -1 ? rowToMove[nameIndex] : '';
    const details = rowToMove.filter((val, idx) => idx !== companyIndex && idx !== hubIndex && idx !== nameIndex && val !== "").join(' / ');
    
    moveSheet.appendRow([
      moveDate ? moveDate.replace(/-/g, '/') : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd'),
      oldLocation,
      newLocation,
      id,
      itemName,
      reason,
      details
    ]);
  }

  SpreadsheetApp.flush();
  Utilities.sleep(500);
  return getSheetData(form.currentHub, sheetName);
}

// ▼ 修正：引数に reason (消去理由) を追加し記録
function deleteData(sheetName, id, hubName, reason) {
  const sheet = SS.getSheetByName(sheetName);
  const delSheet = SS.getSheetByName(SHEET_NAMES.DELETED);
  const headers = COLUMN_DEF[sheetName];
  const idIndex = headers.findIndex(h => h.includes('管理番号') || h.includes('業者ID'));
  const companyIndex = headers.indexOf('会社名');
  const hubIndex = headers.indexOf('拠点');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIndex]) === String(id)) {
      const row = sheet.getRange(i + 1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (delSheet) {
        const deletedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
        const company = companyIndex > -1 ? row[companyIndex] : '';
        const hub = hubIndex > -1 ? row[hubIndex] : '';
        const details = row.filter((val, idx) => idx !== companyIndex && idx !== hubIndex && idx !== idIndex && val !== "").join(' / ');
        
        delSheet.appendRow([
          deletedAt,
          company,
          hub,
          id,
          reason || '',
          details
        ]);
      }
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      Utilities.sleep(500);
      break;
    }
  }
  return getSheetData(sheetName === '業者一覧' ? null : hubName, sheetName);
}

function searchAll(keyword) {
  if (!keyword || !String(keyword).trim()) return [];
  const keywords = String(keyword).toLowerCase().trim().split(/[\s　,、]+/).filter(kw => kw.length > 0);
  if (!keywords.length) return [];
  const targets = [SHEET_NAMES.INVENTORY, SHEET_NAMES.FACILITIES, SHEET_NAMES.CONSTRUCTION, SHEET_NAMES.VENDORS];
  let results = [];

  targets.forEach(sheetName => {
    const sheet = SS.getSheetByName(sheetName);
    if (!sheet) return;
    const headers = COLUMN_DEF[sheetName];
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return;

    const idIdx = headers.findIndex(h => h.includes("管理番号") || h.includes("業者ID"));
    const companyIdx = headers.indexOf("会社名");
    const hubIdx = headers.indexOf("拠点");
    const nameIdx = headers.findIndex(h => h === "品目" || h === "設備名" || h === "工事名" || h === "担当者名" || h === "会社名");

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const match = keywords.every(kw => row.some(cell => String(cell).toLowerCase().includes(kw)));
      if (match) {
        results.push({
          sheet: sheetName,
          company: companyIdx > -1 ? row[companyIdx] : '-',
          hub: hubIdx > -1 ? row[hubIdx] : '-',
          id: idIdx > -1 ? row[idIdx] : '-',
          name: nameIdx > -1 ? row[nameIdx] : '詳細を開く',
          snippet: row.filter((c, idx) => c && idx !== companyIdx && idx !== hubIdx && !String(c).startsWith('http')).slice(0, 4).join(' / ')
        });
      }
    }
  });
  return results;
}