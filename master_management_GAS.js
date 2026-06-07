// =====================================================
// 得意先マスタ管理システム - Google Apps Script
// =====================================================
// 【セットアップ手順】
// 1. Googleドライブで新規スプレッドシートを作成
// 2. メニュー「拡張機能」→「Apps Script」を開く
// 3. このコードを全て貼り付けて保存（Ctrl+S）
// 4. 関数「setup」を選択して▶実行ボタンを押す
// 5. スプレッドシートにタブが自動生成されます
// =====================================================

// =====================================================
// ■ 設定（T-PLANNERのCSV列番号 0始まり）
// =====================================================
const TPLANNER_COLS = {
  cd:        0,   // 得意先CD
  name:      1,   // 得意先名
  nameKana:  2,   // 得意先名カナ
  zip:       4,   // 郵便番号
  addr1:     5,   // 住所1
  addr2:     6,   // 住所2
  tel:       7,   // TEL1
  fax:       9,   // FAX
  mail:      10,  // MAIL
  lastBill: -1,   // 前回請求締年月日 ※CSVで確認後に設定
  updatedAt:-1,   // 更新日時         ※CSVで確認後に設定
};

// ■ 自動フラグ判定ルール
const RULES = {
  cdSpotThreshold: 9999999, // このCD以上はスポット案件として自動フラグ
  dormantYears:    3,       // 更新なしX年以上 → 🔴休眠候補
  noBillYears:     2,       // 前回請求X年以上前 → 🟡請求停止候補
};

// ■ DB識別プレフィックス
const DB_PREFIX = {
  DB1: 'H-',  // T-PLANNER DB1（廃棄物収集系）
  DB2: 'K-',  // T-PLANNER DB2（一元管理系）
};

// ■ ダミーレコードCD（取込スキップ）
const SKIP_CDS = {
  DB1: ['0000000001'], // DB1の先頭ダミーレコード
  DB2: [],
};

// =====================================================
// ■ 取込タブ列定義
// =====================================================
const IMPORT_HEADERS = [
  '確認',           // A: ✅承認 / ❌除外 / （空=未確認）
  '自動判定',       // B: 🔴休眠候補 / 🟡請求停止候補 / 🔵スポット
  'T-PLANNER CD',   // C
  '得意先名',       // D
  '得意先名カナ',   // E
  '郵便番号',       // F
  '住所',           // G
  'TEL',            // H
  'FAX',            // I
  'MAIL',           // J
  '前回請求締年月日', // K
  '更新日時',       // L
  '取込日',         // M
];

// =====================================================
// ■ 承認済マスタ列定義
// =====================================================
const MASTER_HEADERS = [
  '統合CD',          // A: H-XXXXXXXXXX or K-XXXXXXXXXX
  'DB区分',          // B: DB1 or DB2
  'T-PLANNER元CD',   // C: 0000001000
  '得意先名',        // D
  '得意先名カナ',    // E
  '郵便番号',        // F
  '住所',            // G
  'TEL',             // H
  'FAX',             // I
  'MAIL',            // J
  '配車システム使用', // K: ✅ / -
  'コンテナ管理使用', // L: ✅ / -
  '承認日',          // M
  '備考',            // N
];

// =====================================================
// ■ カスタムメニュー（スプレッドシートを開いたとき自動表示）
// =====================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 マスタ管理')
    .addItem('📥 CSVを取込（DB1）', 'importCsvDb1')
    .addItem('📥 CSVを取込（DB2）', 'importCsvDb2')
    .addSeparator()
    .addItem('🔍 未確認レコードの件数を確認', 'checkUnconfirmed')
    .addSeparator()
    .addItem('✅ 承認済マスタを更新', 'syncApprovedMaster')
    .addSeparator()
    .addItem('📊 統計を表示', 'showStats')
    .addToUi();
}

// =====================================================
// ■ セットアップ（初回1回だけ実行）
// =====================================================
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Tokyo');

  // 各シートを作成
  _setupGuideSheet(ss);
  _setupImportSheet(ss, 'DB1_取込', 'DB1');
  _setupImportSheet(ss, 'DB2_取込', 'DB2');
  _setupApprovedSheet(ss);
  _setupExcludedSheet(ss);

  // Googleが作ったデフォルトシートを削除
  ['シート1', 'Sheet1'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  ss.setActiveSheet(ss.getSheetByName('📖 使い方'));
  SpreadsheetApp.getUi().alert('✅ セットアップ完了！\n「📖 使い方」タブを確認してください。');
}

// =====================================================
// ■ 各シートのセットアップ
// =====================================================
function _setupGuideSheet(ss) {
  let sheet = ss.getSheetByName('📖 使い方');
  if (!sheet) sheet = ss.insertSheet('📖 使い方');
  else sheet.clear();
  sheet.setTabColor('#1e3a5f');

  const data = [
    ['得意先マスタ管理システム - 使い方'],
    [''],
    ['■ 基本的な流れ（初回・毎回共通）'],
    [''],
    ['  ① T-PLANNERからCSVをエクスポート（DB1とDB2の2ファイル）'],
    [''],
    ['  ② 上メニュー「📋 マスタ管理」→「📥 CSVを取込（DB1）」'],
    ['     上メニュー「📋 マスタ管理」→「📥 CSVを取込（DB2）」'],
    [''],
    ['  ③ 「DB1_取込」「DB2_取込」タブで内容を確認・チェック'],
    ['     （初回のみ全件確認 / 2回目以降は新規追加分だけ確認）'],
    [''],
    ['     🔴 赤背景 : 休眠候補（更新日時が' + RULES.dormantYears + '年以上前）'],
    ['     🟡 黄背景 : 請求停止候補（前回請求が' + RULES.noBillYears + '年以上前）'],
    ['     🔵 青背景 : スポット案件（CDが大きい工事案件ごとのマスタ）'],
    ['     🟢 緑背景 : 承認済み'],
    ['     灰背景   : 除外済み'],
    [''],
    ['     → 確認列に「✅承認」または「❌除外」をドロップダウンで選択'],
    [''],
    ['  ④ 上メニュー「📋 マスタ管理」→「✅ 承認済マスタを更新」'],
    [''],
    ['  ⑤ 配車・コンテナシステムで「お客様マスタ更新」ボタンを押す'],
    [''],
    [''],
    ['■ 統合CDの仕組み'],
    [''],
    ['     DB1の得意先 → H-XXXXXXXXXX （例: H-0000001000）'],
    ['     DB2の得意先 → K-XXXXXXXXXX （例: K-0000000001）'],
    [''],
    ['     → 同じCD番号でも「H-」「K-」で区別するため重複しない'],
    ['     → 売上データをT-PLANNERにUPするとき、'],
    ['       「H-」→ DB1へ元CDで送信、「K-」→ DB2へ元CDで送信'],
    [''],
    [''],
    ['■ 急ぎの場合（1件だけ追加）'],
    [''],
    ['     「✅ 承認済マスタ」タブに直接1行追加'],
    ['     統合CD列に手動で H-XXXXXXXXXX または K-XXXXXXXXXX を入力'],
    ['     → ④⑤を実行するだけでシステムに反映されます'],
  ];

  sheet.getRange(1, 1, data.length, 1).setValues(data);
  sheet.getRange(1, 1).setFontSize(16).setFontWeight('bold').setFontColor('#1e3a5f');
  sheet.getRange(3, 1).setFontWeight('bold').setFontColor('#1e3a5f');
  sheet.getRange(27, 1).setFontWeight('bold').setFontColor('#1e3a5f');
  sheet.getRange(37, 1).setFontWeight('bold').setFontColor('#1e3a5f');
  sheet.setColumnWidth(1, 600);
}

function _setupImportSheet(ss, sheetName, dbType) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  else sheet.clear();
  sheet.setTabColor(dbType === 'DB1' ? '#0891b2' : '#7c3aed');

  // ヘッダー
  sheet.getRange(1, 1, 1, IMPORT_HEADERS.length).setValues([IMPORT_HEADERS]);
  const hdr = sheet.getRange(1, 1, 1, IMPORT_HEADERS.length);
  hdr.setBackground('#1e3a5f').setFontColor('#fff').setFontWeight('bold');

  // 列幅
  const widths = [70, 120, 130, 220, 160, 80, 280, 130, 130, 180, 140, 140, 100];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // 先頭ゼロが消えないようテキスト形式に固定（CD / 郵便番号 / TEL / FAX）
  sheet.getRange('C:C').setNumberFormat('@'); // T-PLANNER CD
  sheet.getRange('F:F').setNumberFormat('@'); // 郵便番号
  sheet.getRange('H:I').setNumberFormat('@'); // TEL / FAX

  // 確認列ドロップダウン
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['✅承認', '❌除外'], true).build();
  sheet.getRange(2, 1, 2000, 1).setDataValidation(rule);

  // 条件付き書式（CUSTOM_FORMULAで行全体を色付け）
  const maxRow = 2001;
  const CF = SpreadsheetApp.BooleanCriteria;
  const fullRange = [sheet.getRange(2, 1, maxRow, IMPORT_HEADERS.length)];
  const colARange  = [sheet.getRange(2, 1, maxRow, 1)];
  const rules = [
    // ✅承認 → 行全体を緑
    SpreadsheetApp.newConditionalFormatRule()
      .withCriteria(CF.CUSTOM_FORMULA, ['=$A2="✅承認"'])
      .setBackground('#d1fae5').setRanges(fullRange).build(),
    // ❌除外 → 行全体をグレー
    SpreadsheetApp.newConditionalFormatRule()
      .withCriteria(CF.CUSTOM_FORMULA, ['=$A2="❌除外"'])
      .setBackground('#f1f5f9').setFontColor('#94a3b8').setRanges(fullRange).build(),
    // 休眠候補 → 行全体を薄赤
    SpreadsheetApp.newConditionalFormatRule()
      .withCriteria(CF.CUSTOM_FORMULA, ['=ISNUMBER(SEARCH("休眠",$B2))'])
      .setBackground('#fee2e2').setRanges(fullRange).build(),
    // 請求停止候補 → 行全体を薄黄
    SpreadsheetApp.newConditionalFormatRule()
      .withCriteria(CF.CUSTOM_FORMULA, ['=ISNUMBER(SEARCH("請求停止",$B2))'])
      .setBackground('#fef9c3').setRanges(fullRange).build(),
    // スポット案件 → 行全体を薄青
    SpreadsheetApp.newConditionalFormatRule()
      .withCriteria(CF.CUSTOM_FORMULA, ['=ISNUMBER(SEARCH("スポット",$B2))'])
      .setBackground('#eff6ff').setRanges(fullRange).build(),
    // 未確認（A列が空）→ A列セルを黄色
    SpreadsheetApp.newConditionalFormatRule()
      .whenCellEmpty().setBackground('#fef08a').setRanges(colARange).build(),
  ];
  sheet.setConditionalFormatRules(rules);
  sheet.setFrozenRows(1);
}

function _setupApprovedSheet(ss) {
  let sheet = ss.getSheetByName('✅ 承認済マスタ');
  if (!sheet) sheet = ss.insertSheet('✅ 承認済マスタ');
  else sheet.clear();
  sheet.setTabColor('#16a34a');

  sheet.getRange(1, 1, 1, MASTER_HEADERS.length).setValues([MASTER_HEADERS]);
  const hdr = sheet.getRange(1, 1, 1, MASTER_HEADERS.length);
  hdr.setBackground('#16a34a').setFontColor('#fff').setFontWeight('bold');

  const widths = [150, 65, 130, 220, 160, 80, 280, 130, 130, 180, 100, 100, 100, 200];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // 先頭ゼロが消えないようテキスト形式に固定
  sheet.getRange('A:A').setNumberFormat('@'); // 統合CD
  sheet.getRange('C:C').setNumberFormat('@'); // T-PLANNER元CD
  sheet.getRange('F:F').setNumberFormat('@'); // 郵便番号
  sheet.getRange('H:I').setNumberFormat('@'); // TEL / FAX

  sheet.setFrozenRows(1);
}

function _setupExcludedSheet(ss) {
  let sheet = ss.getSheetByName('❌ 除外リスト');
  if (!sheet) sheet = ss.insertSheet('❌ 除外リスト');
  else sheet.clear();
  sheet.setTabColor('#dc2626');

  const headers = ['DB区分', 'T-PLANNER CD', '得意先名', '自動判定', '除外日'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const hdr = sheet.getRange(1, 1, 1, headers.length);
  hdr.setBackground('#dc2626').setFontColor('#fff').setFontWeight('bold');
  sheet.setFrozenRows(1);
}

// =====================================================
// ■ CSV取込（メニューから呼び出し）
// =====================================================
function importCsvDb1() { _showImportDialog('DB1'); }
function importCsvDb2() { _showImportDialog('DB2'); }

function _showImportDialog(dbType) {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html><body style="font-family:'Meiryo',sans-serif;padding:20px;background:#f8fafc">
    <h3 style="color:#1e3a5f;margin:0 0 8px">📥 T-PLANNER ${dbType} CSV取込</h3>
    <p style="color:#64748b;font-size:12px;margin:0 0 14px">
      T-PLANNERからエクスポートしたCSVファイルを選択してください。<br>
      Shift-JIS形式に対応しています。すでに承認・除外済みのレコードは引き継がれます。
    </p>
    <input type="file" id="file" accept=".csv"
      style="display:block;margin:0 0 12px;padding:8px;border:1.5px solid #cbd5e1;
             border-radius:6px;width:100%;box-sizing:border-box;font-size:13px">
    <button onclick="upload()" id="btn"
      style="background:#0891b2;color:white;border:none;padding:8px 22px;
             border-radius:8px;cursor:pointer;font-size:14px;font-weight:700">
      📥 取込開始
    </button>
    <div id="status" style="margin-top:12px;font-size:13px;color:#475569;min-height:20px"></div>
    <script>
      function upload() {
        const file = document.getElementById('file').files[0];
        if (!file) { document.getElementById('status').textContent = '⚠ ファイルを選択してください'; return; }
        document.getElementById('btn').disabled = true;
        document.getElementById('status').textContent = '⏳ 取込中...';
        const reader = new FileReader();
        reader.onload = function(e) {
          google.script.run
            .withSuccessHandler(function(msg) {
              document.getElementById('status').innerHTML = msg;
              document.getElementById('btn').disabled = false;
            })
            .withFailureHandler(function(err) {
              document.getElementById('status').textContent = '❌ エラー: ' + err.message;
              document.getElementById('btn').disabled = false;
            })
            .processImportedCsv(e.target.result, '${dbType}');
        };
        reader.readAsDataURL(file);
      }
    </script>
    </body></html>
  `).setWidth(520).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, `T-PLANNER ${dbType} CSV取込`);
}

// =====================================================
// ■ CSV取込の実処理（GAS側）
// =====================================================
function processImportedCsv(dataUrl, dbType) {
  try {
    // Base64デコード → Shift-JIS → UTF-8変換
    const base64 = dataUrl.split(',')[1];
    const bytes = Utilities.base64Decode(base64);
    const csvText = Utilities.newBlob(bytes).getDataAsString('Shift_JIS');
    const rows = Utilities.parseCsv(csvText);

    if (rows.length < 2) return '❌ データがありません（ヘッダー行のみ）';

    const sheetName = dbType === 'DB1' ? 'DB1_取込' : 'DB2_取込';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return '❌ シートが見つかりません: ' + sheetName;

    // 既存の確認状況を保持（CDをキーに）
    const existingStatus = {};
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const existingData = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
      existingData.forEach(r => { if (r[2]) existingStatus[String(r[2]).trim()] = r[0]; });
    }

    // CSVヘッダーから列番号を自動検出
    const colIdx = _detectCsvColumns(rows[0]);

    // データ処理
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
    const now = new Date();
    const skipCds = SKIP_CDS[dbType] || [];
    const results = [];
    let newCount = 0, existingCount = 0, skippedCount = 0;

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const cd = (row[colIdx.cd] || '').trim();
      if (!cd) continue;

      // ダミー・スキップCDを除外
      if (skipCds.includes(cd)) { skippedCount++; continue; }
      // 名称が空のレコードをスキップ
      const name = (row[colIdx.name] || '').trim();
      if (!name) { skippedCount++; continue; }

      const nameKana = (row[colIdx.nameKana] || '').trim();
      const zip = (row[colIdx.zip] || '').trim();
      const addr1 = (row[colIdx.addr1] || '').trim();
      const addr2 = (row[colIdx.addr2] || '').trim();
      const addr = [addr1, addr2].filter(Boolean).join(' ');
      const tel = (row[colIdx.tel] || '').trim();
      const fax = (row[colIdx.fax] || '').trim();
      const mail = (row[colIdx.mail] || '').trim();
      const lastBill = colIdx.lastBill >= 0 ? (row[colIdx.lastBill] || '').trim() : '';
      const updatedAt = colIdx.updatedAt >= 0 ? (row[colIdx.updatedAt] || '').trim() : '';

      // 自動判定フラグ
      const autoFlag = _calcAutoFlag(cd, updatedAt, lastBill, now);

      // 既存の確認状況を引き継ぐ
      const status = existingStatus[cd] !== undefined ? existingStatus[cd] : '';
      if (existingStatus[cd] !== undefined) existingCount++;
      else newCount++;

      results.push([status, autoFlag, cd, name, nameKana, zip, addr, tel, fax, mail, lastBill, updatedAt, today]);
    }

    if (results.length === 0) return '❌ 有効なデータが見つかりませんでした';

    // シートに書き込み
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, IMPORT_HEADERS.length).clearContent();

    // 書き込み前に先頭ゼロ保護（既存シートで未設定の場合に備えて）
    sheet.getRange(2, 3, results.length, 1).setNumberFormat('@'); // T-PLANNER CD
    sheet.getRange(2, 6, results.length, 1).setNumberFormat('@'); // 郵便番号
    sheet.getRange(2, 8, results.length, 2).setNumberFormat('@'); // TEL / FAX

    sheet.getRange(2, 1, results.length, IMPORT_HEADERS.length).setValues(results);

    // スプレッドシートをアクティブタブに切替
    ss.setActiveSheet(sheet);

    return `✅ 取込完了！ 合計 ${results.length}件<br>
      （🆕 新規: <b>${newCount}件</b> / 🔄 既存引継: ${existingCount}件 / ⏭ スキップ: ${skippedCount}件）`;

  } catch(e) {
    return '❌ エラーが発生しました: ' + e.message;
  }
}

// =====================================================
// ■ CSVの列番号を自動検出
// =====================================================
function _detectCsvColumns(headerRow) {
  const KEYWORDS = {
    cd:        ['得意先CD', '得意先ｺｰﾄﾞ', '得意先コード'],
    name:      ['得意先名'],
    nameKana:  ['得意先名カナ', '得意先名ｶﾅ'],
    zip:       ['郵便番号', '〒'],
    addr1:     ['住所1', '住所１', '住所'],
    addr2:     ['住所2', '住所２'],
    tel:       ['TEL1', 'TEL'],
    fax:       ['FAX'],
    mail:      ['MAIL', 'メール'],
    lastBill:  ['前回請求締年月日', '前回請求'],
    updatedAt: ['更新日時', '更新日'],
  };

  const result = Object.fromEntries(Object.keys(KEYWORDS).map(k => [k, -1]));

  headerRow.forEach((cell, i) => {
    const c = (cell || '').toString().trim();
    Object.keys(KEYWORDS).forEach(key => {
      if (result[key] === -1 && KEYWORDS[key].some(kw => c.includes(kw))) {
        result[key] = i;
      }
    });
  });

  // 自動検出できなかった項目はデフォルト値で補完
  Object.keys(TPLANNER_COLS).forEach(key => {
    if (result[key] === -1 && TPLANNER_COLS[key] >= 0) result[key] = TPLANNER_COLS[key];
  });

  return result;
}

// =====================================================
// ■ 自動判定フラグ計算
// =====================================================
function _calcAutoFlag(cd, updatedAt, lastBill, now) {
  const cdNum = parseInt(cd, 10);
  if (cdNum > RULES.cdSpotThreshold) return '🔵スポット案件';

  const flags = [];
  if (updatedAt) {
    try {
      const d = new Date(updatedAt.replace(/\s/, 'T'));
      if (!isNaN(d) && (now - d) / (1000*60*60*24*365) > RULES.dormantYears) flags.push('🔴休眠候補');
    } catch(e) {}
  }
  if (lastBill) {
    try {
      const d = new Date(lastBill.replace(/\s/, 'T'));
      if (!isNaN(d) && (now - d) / (1000*60*60*24*365) > RULES.noBillYears) flags.push('🟡請求停止候補');
    } catch(e) {}
  }
  return flags.join(' ') || '';
}

// =====================================================
// ■ 未確認レコード件数の確認
// =====================================================
function checkUnconfirmed() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let msg = '📋 確認待ちレコード\n\n';
  let total = 0;

  [['DB1_取込', 'DB1'], ['DB2_取込', 'DB2']].forEach(([sn, label]) => {
    const sheet = ss.getSheetByName(sn);
    if (!sheet || sheet.getLastRow() < 2) return;
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    let unconf = 0, newReal = 0;
    data.forEach(r => {
      if (!r[0]) { unconf++; total++; if (!String(r[1]).includes('スポット')) newReal++; }
    });
    msg += `【${label}】 未確認: ${unconf}件（うち要確認: ${newReal}件）\n`;
  });

  if (total === 0) {
    SpreadsheetApp.getUi().alert('✅ 未確認のレコードはありません！\n全て承認済みまたは除外済みです。');
  } else {
    msg += `\n合計 ${total}件 の未確認レコードがあります。\n\n各取込タブの「確認」列（黄色）を確認してください。`;
    SpreadsheetApp.getUi().alert(msg);
  }
}

// =====================================================
// ■ 承認済マスタの更新
// =====================================================
function syncApprovedMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('✅ 承認済マスタ');
  const excludedSheet = ss.getSheetByName('❌ 除外リスト');

  // 既存マスタの補足情報を保持（配車/コンテナフラグ・備考）
  const keepData = {};
  const masterLastRow = masterSheet.getLastRow();
  if (masterLastRow > 1) {
    const existing = masterSheet.getRange(2, 1, masterLastRow - 1, MASTER_HEADERS.length).getValues();
    existing.forEach(r => {
      if (r[0]) keepData[r[0]] = { dispFlag: r[10], contFlag: r[11], note: r[13] };
    });
  }

  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  const newMaster = [];
  const newExcluded = [];

  [['DB1_取込', 'DB1'], ['DB2_取込', 'DB2']].forEach(([sn, dbType]) => {
    const sheet = ss.getSheetByName(sn);
    if (!sheet || sheet.getLastRow() < 2) return;
    const prefix = DB_PREFIX[dbType];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, IMPORT_HEADERS.length).getValues();

    data.forEach(r => {
      const status = r[0];
      const cd = (r[2] || '').toString().trim();
      if (!cd) return;
      const unifiedCd = prefix + cd;

      if (status === '✅承認') {
        const keep = keepData[unifiedCd] || {};
        newMaster.push([
          unifiedCd, dbType, cd,
          r[3], r[4], r[5], r[6], r[7], r[8], r[9],
          keep.dispFlag || '',
          keep.contFlag || '',
          today,
          keep.note || '',
        ]);
      } else if (status === '❌除外') {
        newExcluded.push([dbType, cd, r[3], r[1], today]);
      }
    });
  });

  // 承認済マスタを更新
  if (masterLastRow > 1) masterSheet.getRange(2, 1, masterLastRow - 1, MASTER_HEADERS.length).clearContent();
  if (newMaster.length > 0) {
    // 書き込み前に先頭ゼロ保護（既存シートで未設定の場合に備えて）
    masterSheet.getRange(2, 1, newMaster.length, 1).setNumberFormat('@'); // 統合CD
    masterSheet.getRange(2, 3, newMaster.length, 1).setNumberFormat('@'); // T-PLANNER元CD
    masterSheet.getRange(2, 6, newMaster.length, 1).setNumberFormat('@'); // 郵便番号
    masterSheet.getRange(2, 8, newMaster.length, 2).setNumberFormat('@'); // TEL / FAX

    masterSheet.getRange(2, 1, newMaster.length, MASTER_HEADERS.length).setValues(newMaster);
    // 交互の行色
    newMaster.forEach((_, i) => {
      masterSheet.getRange(i + 2, 1, 1, MASTER_HEADERS.length)
        .setBackground(i % 2 === 0 ? '#f0fdf4' : '#ffffff');
    });
  }

  // 除外リストを更新
  const exLastRow = excludedSheet.getLastRow();
  if (exLastRow > 1) excludedSheet.getRange(2, 1, exLastRow - 1, 5).clearContent();
  if (newExcluded.length > 0) {
    excludedSheet.getRange(2, 1, newExcluded.length, 5).setValues(newExcluded);
  }

  ss.setActiveSheet(masterSheet);
  SpreadsheetApp.getUi().alert(
    `✅ 承認済マスタを更新しました！\n\n` +
    `承認済: ${newMaster.length}件\n除外: ${newExcluded.length}件`
  );
}

// =====================================================
// ■ 統計表示
// =====================================================
function showStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let msg = '📊 マスタ統計\n\n';

  [['DB1_取込', 'DB1'], ['DB2_取込', 'DB2']].forEach(([sn, label]) => {
    const sheet = ss.getSheetByName(sn);
    if (!sheet || sheet.getLastRow() < 2) { msg += `【${label}】 データなし\n\n`; return; }
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    let approved = 0, excluded = 0, unconf = 0, spot = 0, dormant = 0, noBill = 0;
    data.forEach(r => {
      if (r[0] === '✅承認') approved++;
      else if (r[0] === '❌除外') excluded++;
      else unconf++;
      const flag = String(r[1] || '');
      if (flag.includes('スポット')) spot++;
      if (flag.includes('休眠')) dormant++;
      if (flag.includes('請求停止')) noBill++;
    });
    msg += `【${label}】 合計: ${data.length}件\n`;
    msg += `  ✅ 承認済: ${approved}件 ❌ 除外: ${excluded}件 ⬜ 未確認: ${unconf}件\n`;
    msg += `  🔴 休眠候補: ${dormant}件 🟡 請求停止候補: ${noBill}件 🔵 スポット: ${spot}件\n\n`;
  });

  const masterSheet = ss.getSheetByName('✅ 承認済マスタ');
  const masterCount = masterSheet ? Math.max(0, masterSheet.getLastRow() - 1) : 0;
  msg += `【承認済マスタ（配車・コンテナ共通）】\n  有効得意先: ${masterCount}件`;

  SpreadsheetApp.getUi().alert(msg);
}
