/**
 * 【メイン】メニュー、UI、一括処理の制御
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ダッシュボード管理')
    .addItem('【全体】集計を実行', 'showDialogForFinal')
    .addItem('【拠点】集計を実行', 'showDialogForLocation')
    .addSeparator()
    .addItem('離脱くんを起動', 'showDialogForRidatsu')
    .addToUi();
}

function showDialogForFinal() { showUpdateDialog('final', '【全体】集計期間の選択'); }
function showDialogForLocation() { showUpdateDialog('location', '【拠点】集計期間の選択'); }

/**
 * 離脱くん：年度選択ダイアログ
 */
function showDialogForRidatsu() {
  const t = HtmlService.createTemplateFromFile('Dialog_Ridatsu');
  
  const today = new Date();
  let currentYear = today.getFullYear();
  if (today.getMonth() + 1 < 4) currentYear--; 

  const validYears = [currentYear];

  try {
    const ss = SpreadsheetApp.openByUrl(ZAIS_CONFIG.CONFIG_URL);
    const sheet = ss.getSheetByName('データセット');
    const data = sheet.getDataRange().getValues();
    
    const nextYear = currentYear + 1;
    const nextKey = nextYear + '確定シフト'; 

    const hasNext = data.some(row => 
      String(row[0]).trim() === nextKey && String(row[1]).trim() !== ""
    );

    if (hasNext) validYears.push(nextYear);
  } catch (e) {
    Logger.log("データセット確認エラー: " + e.message);
  }

  t.validYears = validYears;
  t.defaultYear = currentYear;

  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(600).setHeight(550), '離脱くん：対象年度の選択');
}

/**
 * 離脱くん実行プロセッサ
 */
function processRidatsuUpdate(form) {
  const targetYear = form.targetYear;
  try {
    const result = runRidatsuAnalysis(targetYear);
    const reportData = getRidatsuReport(targetYear, result.fullTimeDocs, result.rawData);
    
    if (!reportData) return "分析は完了しましたがレポート生成に失敗しました。";

    const t = HtmlService.createTemplateFromFile('Dialog_Report');
    t.targetYear = reportData.targetYear;
    t.period     = reportData.period;
    t.definition = reportData.definition;
    t.counts     = reportData.counts;
    t.firstExit  = reportData.firstExit;
    t.retention  = reportData.retention;
    t.aces       = reportData.aces;       
    t.settledAll = reportData.settledAll; 
    t.exitList   = reportData.exitList;   
    
    SpreadsheetApp.getUi().showModalDialog(
      t.evaluate().setWidth(1350).setHeight(920), 
      '離脱・定着分析レポート（' + targetYear + '年度）'
    );
    
    return "分析が完了しました。";
  } catch (e) {
    return "中断エラー: " + e.message;
  }
}

/**
 * 🚀 【全体・拠点】集計期間選択ダイアログ
 * データセットの「確定シフト」から未来の最大月を高速スキャンして初期値に設定
 */
function showUpdateDialog(mode, title) {
  const t = HtmlService.createTemplateFromFile('Dialog');
  t.mode = mode;

  const today = new Date();
  const currentY = today.getFullYear();
  const currentM = today.getMonth() + 1;

  let maxY = currentY;
  let maxM = currentM;
  const availableYears = new Set([currentY - 1, currentY, currentY + 1]);

  const updateMax = (y, m) => {
    availableYears.add(y);
    if (y > maxY || (y === maxY && m > maxM)) {
      maxY = y;
      maxM = m;
    }
  };

  try {
    const ssConfig = SpreadsheetApp.openByUrl(ZAIS_CONFIG.CONFIG_URL);
    const rows = ssConfig.getSheetByName('データセット').getDataRange().getValues();

    let currentShiftUrl = "";
    rows.forEach(r => {
      const key = String(r[0]).trim();
      const url = String(r[1]).trim();
      if (!url) return;
      if (key === '確定シフト') currentShiftUrl = url;
      const match = key.match(/^(\d{4})確定シフト$/);
      if (match) availableYears.add(parseInt(match[1], 10));
    });

    if (currentShiftUrl) {
      const ss = SpreadsheetApp.openByUrl(currentShiftUrl);
      const unifiedSheet = ss.getSheetByName("確定シフト");
      if (unifiedSheet) {
        const lastRow = unifiedSheet.getLastRow();
        if (lastRow > 1) {
          const headers = unifiedSheet.getRange(1, 1, 1, unifiedSheet.getLastColumn()).getValues()[0];
          const dateColIdx = headers.findIndex(h => String(h).includes("勤務日") || String(h).includes("日付"));
          if (dateColIdx !== -1) {
            const scanRows = Math.min(lastRow - 1, 1000);
            const startRow = lastRow - scanRows + 1;
            const dateValues = unifiedSheet.getRange(startRow, dateColIdx + 1, scanRows, 1).getValues();
            dateValues.forEach(r => {
              if (r[0] instanceof Date) updateMax(r[0].getFullYear(), r[0].getMonth() + 1);
            });
          }
        }
      }
    }
  } catch (e) {
    Logger.log("設定読込エラー: " + e.message);
  }

  t.years = Array.from(availableYears).sort((a, b) => a - b);
  t.defaultStart = { year: currentY, month: currentM };
  t.defaultEnd   = { year: maxY, month: maxM };

  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(400).setHeight(400), title);
}

/**
 * 🚀 一括更新プロセッサ（共有キャッシュによる爆速ループ）
 */
function processBatchUpdate(form) {
  const mode = form.mode;
  const startY = Number(form.startYear);
  const startM = Number(form.startMonth);
  const endY = Number(form.endYear);
  const endM = Number(form.endMonth);
  
  let logs = [];
  let currentY = startY;
  let currentM = startM;
  let safetyCounter = 0;

  // ループ全体でマスタ・シフト表を共有するインメモリキャッシュ
  const sharedCache = {};

  try {
    while (safetyCounter < 36) {
      if (currentY > endY || (currentY === endY && currentM > endM)) break;
      let resultMsg = "";
      
      if (mode === 'final') {
        resultMsg = `[全体] ${updateDashboardFinal_Fixed(currentY, currentM, sharedCache)}`;
      } else if (mode === 'location') {
        resultMsg = `[拠点] ${updateLocationDashboard(currentY, currentM, sharedCache)}`;
      }
      
      logs.push(`${currentY}年${currentM}月: ${resultMsg}`);
      currentM++;
      if (currentM > 12) { currentM = 1; currentY++; }
      safetyCounter++;
    }
    SpreadsheetApp.flush();
  } catch (e) {
    return logs.join('\n') + `\n中断エラー: ${e.message}`;
  }
  return logs.join('\n') + '\n\n処理が完了しました。';
}