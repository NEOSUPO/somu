/**
 * 負荷テスト用 ダミーデータ一括生成スクリプト
 * 実行すると、hub_master以外の6シートの2行目以降をクリアし、各2000件のダミーデータを生成します。
 */
function generateLoadTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const RECORD_COUNT = 2000; // 生成する件数

  // 1. hub_masterから会社名・拠点名のペアを取得（A列:会社名, B列:拠点名）
  const masterSheet = ss.getSheetByName('hub_master');
  if (!masterSheet) {
    SpreadsheetApp.getUi().alert('エラー: hub_masterシートが見つかりません。');
    return;
  }
  const masterValues = masterSheet.getRange(2, 1, Math.max(1, masterSheet.getLastRow() - 1), 2).getValues();
  const hubs = masterValues.filter(row => row[0] && row[1]); // 空行を除外
  if (hubs.length === 0) hubs.push(["ダミー会社", "ダミー拠点"]); // 万が一空だった場合の保険

  // --- ダミーデータのバリエーション ---
  const floors = ["1F", "2F", "3F", "4F", "5F", "12F", "地下1F"];
  const rooms = ["執務室", "会議室A", "会議室B", "エントランス", "EVホール", "社長室", "サーバー室", "倉庫"];
  
  // 什器用
  const inventoryItems = ["オフィスデスク", "エルゴノミクスチェア", "ミーティングテーブル", "ホワイトボード", "ノートPC", "モニター", "キャビネット", "パーテーション", "応接ソファ", "プロジェクター"];
  const inventoryMakers = ["オカムラ", "コクヨ", "イトーキ", "Dell", "HP", "EIZO", "プラス", "内田洋行"];
  
  // 設備用（自動ドア・エレベーター含む）
  const facilityItems = ["天井埋込形エアコン", "自動ドア（両開き）", "乗用エレベーター", "貨物エレベーター", "セキュリティゲート", "入退室管理システム", "LED照明器具", "分電盤", "受水槽", "防犯カメラシステム"];
  const facilityMakers = ["ダイキン", "パナソニック", "ナブテスコ（自動ドア）", "寺岡オートドア", "三菱エレベーター", "日立ビルシステム", "クマヒラ", "アマノ", "東芝", "三菱電機"];

  // 工事・役務用
  const constructions = ["フロアカーペット張替", "空調機定期清掃", "自動ドア保守点検", "エレベーター法定点検", "LED照明化工事", "セキュリティシステム更新", "防災設備点検", "レイアウト変更工事", "LAN配線工事", "受水槽清掃"];

  // 業者用
  const vendorNames = ["株式会社オフィスサプライ", "〇〇システムズ", "日本メンテサービス", "ビルマネジメント株式会社", "〇〇エレベーター保守", "オートドアサービス株式会社", "〇〇電気工事", "〇〇建装", "スマートセキュリティ", "〇〇通信"];
  const vendorCategories = ["什器販売", "IT機器", "空調メンテ", "ビルメンテ", "昇降機メンテ", "建具メンテ", "電気工事", "内装工事", "警備・セキュリティ", "通信工事"];

  // ▼ 追加：画像URLのリスト
  const imageUrls = [
    "https://www.rikaken-hd.co.jp/wp-content/uploads/2023/01/96b01ca409e2c5eaf538702f2c74e894.png",
    "https://www.rikaken-hd.co.jp/wp-content/uploads/2023/01/f88b6ef95708fb574cde2681d530b1e7.png",
    "https://www.rikaken-hd.co.jp/wp-content/uploads/2023/01/5156ea3ceae907999225e5f56a293e2d.png",
    "https://www.rikaken-hd.co.jp/wp-content/uploads/2023/01/4015472caf3505a83cdeee0a654cbaa9.png",
    "https://www.rikaken-hd.co.jp/wp-content/uploads/2023/01/844d12f0e6ca9faebb5ae92f0e5e24a0.png",
    "https://www.rikaken-hd.co.jp/wp-content/uploads/2023/01/e1138c51ae64a7ae49a43956fee7e6bf.png",
    "https://www.rikaken-hd.co.jp/wp-content/uploads/2023/01/5c2eea6a10e74c605cca2dcbe6fc021e.png"
  ];

  // ユーティリティ関数
  const getRandom = (arr) => arr[Math.floor(Math.random() * arr.length)];
  const getRandomPrice = (min, max) => (Math.floor(Math.random() * ((max - min) / 1000)) * 1000 + min).toString();
  const getRandomDate = () => {
    const start = new Date(2020, 0, 1).getTime();
    const end = new Date(2025, 11, 31).getTime();
    const date = new Date(start + Math.random() * (end - start));
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  };
  const getFutureDate = () => {
    const start = new Date(2026, 0, 1).getTime();
    const end = new Date(2028, 11, 31).getTime();
    const date = new Date(start + Math.random() * (end - start));
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  };

  // 生成処理のメイン関数
  function generateSheetData(sheetName, columns, generatorFunc) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    // 2行目以降の既存データをクリア（1行目のヘッダーは守る）
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }

    const data = [];
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');

    for (let i = 0; i < RECORD_COUNT; i++) {
      const hub = getRandom(hubs); // ランダムな会社・拠点のペア
      const rowData = generatorFunc(hub, i, now);
      data.push(rowData);
    }

    // まとめて書き込み（高速化のため）
    sheet.getRange(2, 1, RECORD_COUNT, columns.length).setValues(data);
    SpreadsheetApp.flush();
  }

  // --- 1. 什器管理 (2000件) ---
  generateSheetData('什器管理', COLUMN_DEF['什器管理'], (hub, i, now) => {
    // 会社名, 拠点, 管理番号, 画像データ添付, フロア, 部屋、場所, 品目, メーカー, 型番、シリアル, 購入日, 購入金額, 業者, 更新日時, 予算対象要否, 概算予算, サイズ, 備考, 関連書類URL, アラート有無, アラート日付
    return [
      hub[0], hub[1], `A${Date.now().toString().slice(-6)}${i}`, getRandom(imageUrls), getRandom(floors), getRandom(rooms), 
      getRandom(inventoryItems), getRandom(inventoryMakers), `MDL-${Math.floor(Math.random()*10000)}`, 
      getRandomDate(), getRandomPrice(10000, 300000), getRandom(vendorNames), now, 
      getRandom(["要", "否"]), getRandomPrice(10000, 300000), "W1200 x D600", "負荷テスト用ダミーデータ", "", 
      "無", ""
    ];
  });

  // --- 2. 設備管理 (2000件) ---
  generateSheetData('設備管理', COLUMN_DEF['設備管理'], (hub, i, now) => {
    // 会社名, 拠点, 管理番号, 設備名, 設置日, 更新目安, フロア, 部屋、場所, メーカー名, 型番, 業者, 状態, 予算対象要否, 概算予算, 更新日, 備考, 関連書類URL, アラート有無, アラート日付
    return [
      hub[0], hub[1], `C${Date.now().toString().slice(-6)}${i}`, getRandom(facilityItems), getRandomDate(), getFutureDate(), 
      getRandom(floors), getRandom(rooms), getRandom(facilityMakers), `FAC-${Math.floor(Math.random()*10000)}`, 
      getRandom(vendorNames), "稼働中", "要", getRandomPrice(500000, 5000000), now.split(' ')[0], 
      "法定点検・メンテナンス対象", "", "無", ""
    ];
  });

  // --- 3. 工事履歴 (2000件) ---
  generateSheetData('工事履歴', COLUMN_DEF['工事履歴'], (hub, i, now) => {
    // 会社名, 拠点, 管理番号, 工事名, 施工日, フロア, 部屋、場所, 業者, 費用, 関連書類URL, 備考, アラート有無, アラート日付
    return [
      hub[0], hub[1], `B${Date.now().toString().slice(-6)}${i}`, getRandom(constructions), getRandomDate(), 
      getRandom(floors), getRandom(rooms), getRandom(vendorNames), getRandomPrice(50000, 2000000), "", 
      "定期メンテナンス・改修記録", "無", ""
    ];
  });

  // --- 4. 業者一覧 (2000件) ---
  generateSheetData('業者一覧', COLUMN_DEF['業者一覧'], (hub, i, now) => {
    // 業者ID, 会社名, 担当者名, TEL, Mail, 対応可能内容, 名刺データ添付, 更新日時, カテゴリー
    return [
      `V${Date.now().toString().slice(-6)}${i}`, getRandom(vendorNames) + ` (支店${i})`, "山田 太郎", 
      `03-0000-${Math.floor(1000 + Math.random() * 9000)}`, `dummy${i}@example.com`, 
      "自動ドア保守、エレベーター法定点検、その他役務", getRandom(imageUrls), now, getRandom(vendorCategories)
    ];
  });

  // --- 5. 消去済み (2000件) ---
  generateSheetData('消去済み', COLUMN_DEF['消去済み'], (hub, i, now) => {
    // 元シート, 会社名, 拠点, 管理番号, 詳細データ抜粋, 消去日時
    return [
      getRandom(["什器管理", "設備管理", "工事履歴"]), hub[0], hub[1], `DEL-${Math.floor(Math.random()*100000)}`, 
      "テスト消去データ抜粋テキスト", now
    ];
  });

  // --- 6. 移動履歴 (2000件) ---
  generateSheetData('移動履歴', COLUMN_DEF['移動履歴'], (hub, i, now) => {
    // 移動日, 会社名, 旧拠点, 新拠点, 管理番号, 品目, 移動理由, 備考
    return [
      getRandomDate(), hub[0], hub[1], getRandom(hubs)[1], `MOV-${Math.floor(Math.random()*100000)}`, 
      getRandom(inventoryItems), "レイアウト変更・組織改編に伴う移設", "移動テスト記録"
    ];
  });

  SpreadsheetApp.getUi().alert('✅ 負荷テスト用のダミーデータ（全6シート×各2000件）の生成が完了しました！');
}