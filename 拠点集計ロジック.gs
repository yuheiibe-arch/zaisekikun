/**
 * 🏆 【拠点集計】updateLocationDashboard (運用ルール完全準拠・最終完成版)
 */
function updateLocationDashboard(targetYear, targetMonth) {
  const CONFIG_URL = 'https://docs.google.com/spreadsheets/d/1CKiJYyZ0iF6cobH_ods2hrI9RmFKt6gFiRXUjD195o0/edit';
  
  targetYear = Number(targetYear);
  targetMonth = Number(targetMonth);

  const fiscalYear = (targetMonth >= 4) ? targetYear : targetYear - 1;
  const SHEET_NAME_OUT = `${fiscalYear}拠点`; 

  Logger.log(`=== 拠点別集計開始: ${targetYear}年${targetMonth}月 ===`);

  const getUrl = (key) => {
    try {
      const ss = SpreadsheetApp.openByUrl(CONFIG_URL);
      const data = ss.getSheetByName('データセット').getDataRange().getValues();
      const row = data.find(r => String(r[0]).trim() === key);
      return row ? String(row[1]).trim() : "";
    } catch(e) { return ""; }
  };

  const normalizeName = (n) => String(n).replace(/[\s ]+/g, "").trim();
  const cleanId = (rawId) => String(rawId).replace(/\.0+$/, '').trim();
  
  const normalizeClinic = (n) => {
    if (!n) return "";
    let s = String(n).trim().replace(/ヶ/g, "ケ");
    if (s.includes('千葉') && s.match(/ニュータウン|NT|ＮＴ/i)) return '千葉ニュータウン中央';
    if (s === '八千代') return '八千代緑が丘';
    return s;
  };
  const normalizeItemKey = (str) => String(str).replace(/[（）\(\)\s \[\]【】]/g, "").trim();
  const formatDateKey = (d) => (d instanceof Date && !isNaN(d)) ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') : null;

  const fullTimeSet = new Set(); 
  const fullTimeNameSet = new Set();
  const regPtSet = new Set();
  const agencyMap = new Map();
  const openDateMap = new Map();
  const retiredMap = new Map();
  const clinicDict = {};

  const urlRegex = getUrl('正規表現');
  if (urlRegex) {
    try {
      const ssReg = SpreadsheetApp.openByUrl(urlRegex);
      const regData = ssReg.getSheets()[0].getDataRange().getValues();
      const rHead = regData[0].map(String);
      const cOfficial = rHead.findIndex(h => h.includes('正規記載'));
      if (cOfficial !== -1) {
        for (let r = 1; r < regData.length; r++) {
          const offName = String(regData[r][cOfficial]).trim();
          if (offName) {
            clinicDict[offName] = offName;
            rHead.forEach((h, i) => { if (h.includes('表記揺れ')) { const v = String(regData[r][i]).trim(); if (v) clinicDict[v] = offName; } });
          }
        }
      }

      const sheetLoc = ssReg.getSheetByName('拠点名');
      if (sheetLoc) {
        const locData = sheetLoc.getDataRange().getValues();
        for (let i = 1; i < locData.length; i++) {
          const name = normalizeClinic(locData[i][0]); 
          const date = locData[i][7]; 
          if (name && date instanceof Date) openDateMap.set(name, date);
        }
      }
    } catch(e) {}
  }

  const urlCost = getUrl('人件費くん');
  if (urlCost) {
    try {
      const ss = SpreadsheetApp.openByUrl(urlCost);
      const sFt = ss.getSheets().find(s => s.getName().includes('常勤マスタ'));
      if (sFt && sFt.getLastRow() > 1) { 
        sFt.getRange(2, 1, sFt.getLastRow()-1, 2).getValues().forEach(r => {
          if (r[0]) fullTimeNameSet.add(normalizeName(r[0])); 
          if (r[1]) fullTimeSet.add(cleanId(r[1]));
        });
      }
      const sReg = ss.getSheets().find(s => s.getName().includes('定期非常勤マスタ'));
      if (sReg && sReg.getLastRow() > 1) {
        sReg.getRange(2, 2, sReg.getLastRow()-1, 1).getValues().forEach(r => { if(r[0]) regPtSet.add(cleanId(r[0])); });
      }

      let urlRetired = getUrl('退職医師リスト') || urlCost;
      const ssRet = SpreadsheetApp.openByUrl(urlRetired);
      const sRet = ssRet.getSheetByName('退職医師') || ssRet.getSheets().find(s => s.getName().includes('退職'));
      if (sRet) {
        const rData = sRet.getDataRange().getValues();
        const rh = rData[0];
        const rcMed = rh.indexOf('医籍番号');
        const rcName = rh.indexOf('医師名');
        const rcType = rh.indexOf('区分');
        const rcJoin = rh.indexOf('入職日');
        const rcLeave = rh.indexOf('退職日');
        
        for (let i=1; i<rData.length; i++) {
          const rr = rData[i];
          const med = cleanId(rr[rcMed]||"");
          const name = normalizeName(rr[rcName]||"");
          const type = String(rr[rcType]||"").trim();
          const join = rr[rcJoin] instanceof Date ? rr[rcJoin] : null;
          const leave = rr[rcLeave] instanceof Date ? rr[rcLeave] : null;
          if ((med||name) && join && leave) {
            const info = { type: type, join: join, leave: leave };
            if(med) retiredMap.set(med, info);
            if(name) retiredMap.set(name, info);
          }
        }
      }
    } catch(e) {}
  }

  const urlAgency = getUrl('紹介会社進捗');
  const AGENCY_CONFIG = {
    'エムスリー': { name: 0, date: 2, startRow: 3 }, 'エムステージ': { name: 0, date: 2, startRow: 3 },
    '民間医局': { name: 0, date: 1, startRow: 3 }, 'MRT': { name: 3, date: 0, startRow: 3 },
    'マイナビ': { name: 0, date: 2, startRow: 3 }, 'メディウェル': { name: 0, date: 3, startRow: 3 },
    'リクルート': { name: 0, date: 3, startRow: 3 }, 'オンコール': { name: 0, date: 3, startRow: 3 },
    'メディカルリソース': { name: 0, date: 2, startRow: 3 }, '総合メディカル': { name: 0, date: 2, startRow: 3 },
    'その他の会社': { name: 0, date: 3, startRow: 3 }
  };
  if (urlAgency) {
    try {
      const ssAgency = SpreadsheetApp.openByUrl(urlAgency);
      Object.keys(AGENCY_CONFIG).forEach(sheetName => {
        const sheet = ssAgency.getSheetByName(sheetName);
        if (!sheet) return;
        const lastRow = sheet.getLastRow();
        const conf = AGENCY_CONFIG[sheetName];
        if (lastRow < conf.startRow) return;
        sheet.getRange(conf.startRow, 1, lastRow - conf.startRow + 1, sheet.getLastColumn()).getValues().forEach(row => {
          const dVal = row[conf.date];
          const nVal = row[conf.name];
          if (dVal instanceof Date && nVal) agencyMap.set(`${normalizeName(nVal)}_${formatDateKey(dVal)}`, sheetName);
        });
      });
    } catch(e) {}
  }

  const ssOut = SpreadsheetApp.openByUrl(CONFIG_URL);
  const sheetOut = ssOut.getSheetByName(SHEET_NAME_OUT);
  if (!sheetOut) return `エラー: シート「${SHEET_NAME_OUT}」が見つかりません。`;

  const rangeData = sheetOut.getRange(2, 1, sheetOut.getLastRow()-1, 3).getValues();
  const rowMap = new Map();
  const targetClinics = new Set();
  let currentClinicName = "";

  rangeData.forEach((r, idx) => {
    if (r[0]) {
      currentClinicName = normalizeClinic(String(r[0]));
      targetClinics.add(currentClinicName);
    }
    if (currentClinicName && r[2]) {
      rowMap.set(`${currentClinicName}_${normalizeItemKey(r[2])}`, idx + 2);
    }
  });
  let writeCol = (targetMonth >= 4) ? targetMonth : targetMonth + 12;

  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth() + 1;
  const monthDiff = (currentYear - targetYear) * 12 + (currentMonth - targetMonth);

  // ========================================================
  // ★ ピンポイント・データソース特定（シミュレーション完全準拠）
  // ========================================================
  let urlShift = "";
  let targetSheetName = "";

  if (monthDiff <= 1) { 
    urlShift = 'https://docs.google.com/spreadsheets/d/1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co/edit';
    targetSheetName = "確定シフト";
  } else {
    if (targetYear === 2026) {
      urlShift = 'https://docs.google.com/spreadsheets/d/10Z8jg4o7Ri9Ggf7u_jauyDDZa6bVnc-iuuuqmkNDafw/edit';
    } else if (targetYear === 2025) {
      urlShift = 'https://docs.google.com/spreadsheets/d/1JhLJuxkp6T5hEkfultUE57pmRXBI_lkN6ROp2-CuhCc/edit';
    } else {
      urlShift = getUrl(`${targetYear}確定シフト`);
    }
    targetSheetName = `${targetYear}/${String(targetMonth).padStart(2, '0')}`;
  }

  if (!urlShift) return `エラー: 対象月(${targetYear}/${targetMonth})のシフトURLが取得できません`;

  let shiftSheet = null;
  try {
    const ssShift = SpreadsheetApp.openByUrl(urlShift);
    shiftSheet = ssShift.getSheetByName(targetSheetName);
  } catch(e) {
    return `ファイル展開エラー: ${e.message}`;
  }

  if (!shiftSheet || shiftSheet.getLastRow() < 2) return `データなし: 「${targetSheetName}」シートが見つからないか空です`;

  const dataRows = shiftSheet.getDataRange().getValues();
  const headers = dataRows[0].map(String);
  const data = dataRows.slice(1);

  const findCol = (heads, candidates) => {
    for (const key of candidates) {
      const idx = heads.findIndex(h => String(h).replace(/[（）\(\)\s ]/g, "").includes(key));
      if (idx > -1) return idx;
    }
    return -1;
  };

  const C = {
    NAME: findCol(headers, ['氏名', '名前']),
    MED_NO: headers.indexOf('医籍番号'),
    DOC_NO: findCol(headers, ['医師No']),
    CLINIC: findCol(headers, ['クリニック名', '勤務地', '拠点']),
    DATE: findCol(headers, ['勤務日', '日付']),
    W_TOT: findCol(headers, ['時給合計']),
    W1: headers.indexOf('時給1'), W2: headers.indexOf('時給2'), W3: headers.indexOf('時給3'), W4: headers.indexOf('時給4'), W_ADD1: headers.indexOf('追加支給額1'),
    START_M: findCol(headers, ['勤務開始時間', '開始時間']), 
    END_M: findCol(headers, ['勤務終了時間', '終了時間']),
    START_B: findCol(headers, ['元のシフト勤務開始時間', '元のシフト)勤務開始時間', '元のシフト）勤務開始時間']),
    END_B: findCol(headers, ['元のシフト勤務終了時間', '元のシフト)勤務終了時間', '元のシフト）勤務終了時間']),
    DEPT: findCol(headers, ['診療科']),
    NOTE: findCol(headers, ['業務内容', '備考'])
  };

  const sMIdx = headers.findIndex(h => String(h).includes('勤務開始時間') && !String(h).includes('元の'));
  if (sMIdx !== -1) C.START_M = sMIdx;
  const eMIdx = headers.findIndex(h => String(h).includes('勤務終了時間') && !String(h).includes('元の'));
  if (eMIdx !== -1) C.END_M = eMIdx;

  const commentCols = headers.map((h, i) => String(h).includes('スタッフコメント') || String(h).includes('コメント') ? i : -1).filter(i => i !== -1);

  const clinicStats = {};
  targetClinics.forEach(c => {
    clinicStats[c] = {
      ft: { c:new Set(), h:0 }, ftOver: { h:0 }, reg: { c:new Set(), h:0 },
      spot: { c:new Set(), h:0 }, spotAg: { c:new Set(), h:0 }, spotDir: { c:new Set(), h:0 }
    };
  });

  data.forEach(row => {
    if (C.DATE === -1) return;
    const dVal = new Date(row[C.DATE]);
    if (isNaN(dVal.getTime())) return;
    
    if (dVal.getFullYear() !== targetYear || (dVal.getMonth() + 1) !== targetMonth) return;

    const rawClinic = String(row[C.CLINIC]||"").trim();
    const officialClinic = clinicDict[rawClinic] || rawClinic;
    let clinic = normalizeClinic(officialClinic);
    if (!clinicStats[clinic] || /有給|欠勤|医師会|院外/.test(clinic)) return;

    if (['亀有', '北葛西'].includes(clinic)) {
      if (!String(row[C.DEPT]||"").includes('小児科')) return;
    }

    const medNo = cleanId(row[C.MED]||"");
    const nameRaw = String(row[C.NAME]||"").trim();
    const nameKey = normalizeName(nameRaw);
    const id = medNo || nameKey;
    if (!id) return;

    const note = (C.NOTE > -1) ? String(row[C.NOTE] || "") : "";
    const parseMoney = (v) => v ? (typeof v === 'number' ? v : Number(String(v).replace(/[^0-9.-]/g, "")) || 0) : 0;
    const wage = parseMoney(row[C.W_TOT]) + parseMoney(row[C.W1]) + parseMoney(row[C.W2]) + parseMoney(row[C.W3]) + parseMoney(row[C.W4]) + parseMoney(row[C.W_ADD1]);

    const toH = (v) => v instanceof Date ? v.getHours() + v.getMinutes()/60 : (typeof v === 'number' ? v * 24 : null);
    let s = toH(row[C.START_M]), e = toH(row[C.END_M]);
    if (s == null || e == null) { s = toH(row[C.START_B]); e = toH(row[C.END_B]); }
    
    if (s === null || e === null || (e - s) < 0.5) return;

    let h = e - s;
    const bo = Math.max(0, Math.min(e, 15) - Math.max(s, 13));
    if (h > bo && s < 14.9) h -= bo;
    if (h < 0) h = 0;

    const comments = commentCols.map(c => String(row[c]||"")).join(" ");
    const isRelief = /所定休出|所定外|特別時給|追加|欠勤対応|急遽|ヘルプ|バックアップ|民間医局|紹介|加算/.test(comments + " " + note);
    const isRegularNote = /毎週(月|火|水|木|金|土|日)曜|第[1-5](月|火|水|木|金|土|日)曜/.test(note);
    
    let isFt = false;
    let isReg = false;

    if ((fullTimeSet.has(medNo) || fullTimeNameSet.has(nameKey)) && wage === 0) isFt = true;
    if (regPtSet.has(medNo)) isReg = true;

    const retiredInfo = retiredMap.get(medNo) || retiredMap.get(nameKey);
    if (retiredInfo) {
      if (dVal > retiredInfo.leave) {
        isFt = false; isReg = false;
      } else if (dVal < retiredInfo.join) {
        isFt = false; isReg = false;
      } else {
        if (retiredInfo.type === '常勤' && wage === 0) isFt = true;
        if (retiredInfo.type.includes('定期')) isReg = true;
      }
    }

    if (isRelief) {
      clinicStats[clinic].spot.c.add(id); clinicStats[clinic].spot.h += h;
      if (/所定休出|所定外/.test(comments + note)) clinicStats[clinic].ftOver.h += h;
      const key = `${nameKey}_${formatDateKey(dVal)}`;
      if (agencyMap.has(key) || /紹介|民間|MRT|エムスリー/.test(comments)) {
        clinicStats[clinic].spotAg.c.add(id); clinicStats[clinic].spotAg.h += h;
      } else {
        clinicStats[clinic].spotDir.c.add(id); clinicStats[clinic].spotDir.h += h;
      }
    } else if (isFt) {
      clinicStats[clinic].ft.c.add(id); clinicStats[clinic].ft.h += h;
    } else if (isRegularNote || isReg) { 
      clinicStats[clinic].reg.c.add(id); clinicStats[clinic].reg.h += h;
    } else if (h > 0) {
      clinicStats[clinic].spot.c.add(id); clinicStats[clinic].spot.h += h;
      const key = `${nameKey}_${formatDateKey(dVal)}`;
      if (agencyMap.has(key) || /紹介|民間|MRT|エムスリー/.test(comments)) {
        clinicStats[clinic].spotAg.c.add(id); clinicStats[clinic].spotAg.h += h;
      } else {
        clinicStats[clinic].spotDir.c.add(id); clinicStats[clinic].spotDir.h += h;
      }
    }
  });

  const updates = [];
  targetClinics.forEach(clinic => {
    const s = clinicStats[clinic];
    const totalH = s.ft.h + s.reg.h + s.spot.h;
    const openDate = openDateMap.get(clinic);
    const isNotOpenYet = (openDate && openDate > new Date(targetYear, targetMonth, 0));
    const rate = (v) => totalH > 0 ? v/totalH : 0;

    const items = [
      { k: '常勤医師稼働時間数', v: s.ft.h },
      { k: '所定休出勤務医師稼働時間数', v: s.ftOver.h },
      { k: '常勤構成比', v: rate(s.ft.h) },
      { k: '定期非常勤医師数', v: s.reg.c.size },
      { k: '定期非常勤医師稼働時間数', v: s.reg.h },
      { k: '定期非常勤構成比', v: rate(s.reg.h) },
      { k: 'スポット医師数', v: s.spot.c.size },
      { k: 'スポット医師稼働時間数', v: s.spot.h },
      { k: '紹介会社医師数内訳', v: s.spotAg.c.size },
      { k: '紹介会社医師稼働時間数内訳', v: s.spotAg.h },
      { k: '紹介会社構成比', v: rate(s.spotAg.h) },
      { k: '直接応募構成比', v: rate(s.spotDir.h) },
      { k: '常勤医師数', v: s.ft.c.size }
    ];

    items.forEach(item => {
      const row = rowMap.get(`${clinic}_${normalizeItemKey(item.k)}`);
      if (row) {
        let val = isNotOpenYet ? "" : item.v;
        if (val !== "" && !item.k.includes('構成比')) val = Math.round(val);
        if (val === "") val = 0;
        updates.push({ r: row, c: writeCol, v: val });
      }
    });
  });

  if (updates.length > 0) {
    updates.forEach(u => sheetOut.getRange(u.r, u.c).setValue(u.v));
    return "完了";
  }
  return "更新なし";
}