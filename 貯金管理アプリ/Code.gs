// ===== 設定 =====
const SPREADSHEET_ID = '1-0ZpPtHxDbj2oPjDTKxmeTagnAzjGqYfZOU4_3ir5Fg';
const SHEET_NAME = '貯金推移';

// ===== エントリポイント =====
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('貯金管理アプリ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ===== データ + 構造を一括取得 =====
function getFullData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // スプレッドシートの全データ（Date型を文字列変換）
    const rawData = sheet.getDataRange().getValues();
    const data = rawData.map(row =>
      row.map(cell =>
        cell instanceof Date
          ? Utilities.formatDate(cell, 'Asia/Tokyo', 'yyyy/MM/dd')
          : cell
      )
    );

    // 行の構造（カテゴリ・ラベル・計算式か否か）
    const structure = buildStructure_(sheet, lastRow, lastCol);

    return JSON.stringify({ success: true, data, structure });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

// ===== 内部: シートから行構造を読み取る =====
function buildStructure_(sheet, lastRow, lastCol) {
  if (lastRow < 2) return [];

  const abData = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const formulas = lastCol >= 3
    ? sheet.getRange(2, 3, lastRow - 1, 1).getFormulas()
    : [];

  const structure = [];
  let currentCat = '';

  for (let i = 0; i < abData.length; i++) {
    const catVal   = String(abData[i][0] || '').trim();
    const labelVal = String(abData[i][1] || '').trim();
    if (!labelVal) continue;
    if (catVal) currentCat = catVal;

    const hasFormula = !!(formulas[i] && formulas[i][0] && String(formulas[i][0]).startsWith('='));
    structure.push({
      row:      i + 2,        // 1始まりの行番号
      category: currentCat,
      label:    labelVal,
      isInput:  !hasFormula,  // 手入力行
      isTotal:  hasFormula,   // 計算式行（削除・編集不可）
    });
  }
  return structure;
}

// ===== 項目を追加 =====
function addItem(category, label) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const structure = buildStructure_(sheet, lastRow, lastCol);

    const catItems = structure.filter(r => r.category === category);
    let insertRow;

    if (catItems.length === 0) {
      // 新カテゴリ: 最初の集計サマリー行の直前に挿入
      const firstSummary = structure.find(r =>
        r.isTotal && ['現金総額', '投資総額', '全合計'].includes(r.label)
      );
      insertRow = firstSummary ? firstSummary.row : lastRow + 1;
    } else {
      // 既存カテゴリ: 合計行の直前 or 最終入力行の次
      const totalRow = catItems.find(r => r.isTotal);
      insertRow = totalRow ? totalRow.row : catItems[catItems.length - 1].row + 1;
    }

    sheet.insertRowBefore(insertRow);
    sheet.getRange(insertRow, 1).setValue(category);
    sheet.getRange(insertRow, 2).setValue(label);

    return JSON.stringify({ success: true });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

// ===== 項目を削除 =====
function deleteItem(rowIndex) {
  try {
    if (rowIndex <= 1) throw new Error('ヘッダー行は削除できません');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    sheet.deleteRow(rowIndex);
    return JSON.stringify({ success: true });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

// ===== データを保存 =====
function saveMonthData(monthStr, inputData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    let targetCol = -1;
    let prevCol = lastCol;

    for (let i = 2; i < headerRow.length; i++) {
      const cellVal = headerRow[i];
      let cellDateStr = '';
      if (cellVal instanceof Date) {
        cellDateStr = Utilities.formatDate(cellVal, 'Asia/Tokyo', 'yyyy/MM');
      } else if (typeof cellVal === 'string' && cellVal.length >= 7) {
        cellDateStr = cellVal.substring(0, 7).replace('-', '/');
      }
      if (cellDateStr === monthStr) {
        targetCol = i + 1;
        prevCol = i;
        break;
      }
    }

    // 新しい月: 列を追加して前月の計算式をコピー
    if (targetCol === -1) {
      targetCol = lastCol + 1;
      prevCol = lastCol;

      const parts = monthStr.split('/');
      const dateVal = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, 1);
      const headerCell = sheet.getRange(1, targetCol);
      headerCell.setValue(dateVal);
      headerCell.setNumberFormat('yyyy/MM/dd');

      // 計算式が入っている行を動的に検出してコピー
      if (lastRow >= 2) {
        const prevFormulas = sheet.getRange(2, prevCol, lastRow - 1, 1).getFormulas();
        for (let i = 0; i < prevFormulas.length; i++) {
          if (prevFormulas[i][0] && String(prevFormulas[i][0]).startsWith('=')) {
            sheet.getRange(i + 2, prevCol).copyTo(sheet.getRange(i + 2, targetCol));
          }
        }
      }
    }

    for (const [rowStr, value] of Object.entries(inputData)) {
      const rowIdx = parseInt(rowStr);
      if (!isNaN(rowIdx) && value !== '' && value !== null && value !== undefined) {
        sheet.getRange(rowIdx, targetCol).setValue(Number(value));
      }
    }

    return JSON.stringify({ success: true });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}
