/**
 * 🏆 【拠点集計】updateLocationDashboard (2026新マスタ・所定休出ロジック完全準拠版)
 */
function updateLocationDashboard(targetYear, targetMonth) {
  const CONFIG_URL = 'https://docs.google.com/spreadsheets/d/1CKiJYyZ0iF6cobH_ods2hrI9RmFKt6gFiRXUjD195o0/edit';
  // ★ jinkenhikun1 (06_NewDoctorMasterManager.gs / B_Core_Utilities.gs) と同じ新マスタURL
  const NEW_MASTER_URL = 'https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit';
  const AGENCY_HISTORY_URL = 'https://docs.google.com/spreadsheets/d/1Fd8uOCE1SKvLCIPjZZ7sE2rFsQFOhjVjHqoaqs-pXqE/edit';

  targetYear = Number(targetYear);
  targetMonth = Number(targetMonth);

  const fiscalYear = (targetMonth >= 4) ? targetYear : targetYear - 1;
  const SHEET_NAME_OUT = `${fiscalYear}拠点`;

  Logger.log(`=== 拠点別集計開始: ${targetYear}年${targetMonth}月 (出力先: ${SHEET_NAME_OUT}) ===`);

  const getUrl = (key) => {
    try {
      const ss = SpreadsheetApp.openByUrl(CONFIG_URL);
      const data = ss.getSheetByName('データセット').getDataRange().getValues();
      const row = data.find(r => String(r[0]).trim() === key);
      return row ? String(row[1]).trim() : "";
    } catch(e) { return ""; }
  };

  const cleanId = (rawId) => String(rawId || "").replace(/\.0+$/, '').trim();
  const normalizeItemKey = (str) => String(str || "").replace(/[（）\(\)\s \[\]【】]/g, "").trim();
  const formatDateKey = (d) => (d instanceof Date && !isNaN(d)) ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') : null;
  const createHeaderMap = (headerRow) => {
    const map = {};
    headerRow.forEach((val, idx) => {
      if (val) map[String(val).replace(/\n/g, '').trim()] = idx;
    });
    return map;
  };

  // 表記ゆれ吸収用マップ
  const clinicDict = {};
  const nameAliasMap = new Map();
  const openDateMap = new Map();

  const normalizeName = (name) => {
    if (!name) return "";
    const rawTrimmed = String(name).trim();
    if (nameAliasMap.has(rawTrimmed)) return String(nameAliasMap.get(rawTrimmed)).replace(/[\s ]+/g, "");
    const halfSpaceVer = rawTrimmed.replace(/ /g, ' ');
    if (nameAliasMap.has(halfSpaceVer)) return String(nameAliasMap.get(halfSpaceVer)).replace(/[\s ]+/g, "");
    const noSpaceVer = halfSpaceVer.replace(/\s+/g, '');
    if (nameAliasMap.has(noSpaceVer)) return String(nameAliasMap.get(noSpaceVer)).replace(/[\s ]+/g, "");
    return noSpaceVer;
  };

  const normalizeClinic = (n) => {
    if (!n) return "";
    let s = String(n).trim();
    if (clinicDict[s]) s = clinicDict[s];
    // 【亀有内科】のようなカッコや先頭番号を除去して再照合
    let clean = s.replace(/^[0-9.\s]+/, '').replace(/[【】\(（]?(内科|小児科)[\)）]?/g, "").split('(')[0].split('（')[0].trim();
    if (clinicDict[clean]) s = clinicDict[clean];
    else s = clean;

    s = s.replace(/ヶ/g, "ケ");
    if (s.includes('千葉') && s.match(/ニュータウン|NT|ＮＴ/i)) return '千葉ニュータウン中央';
    if (s === '八千代') return '八千代緑が丘';
    if (s === '市川妙典') return '妙典';
    if (s === 'まるひろ南浦和店') return '南浦和';
    if (s === '千里丘') return '千里';
    return s;
  };

  // ========================================================
  // 1. 「正規表現」シートから拠点名・氏名の表記ゆれと開業日を取得
  // ========================================================
  const urlRegex = getUrl('正規表現');
  if (urlRegex) {
    try {
      const ssReg = SpreadsheetApp.openByUrl(urlRegex);
      // 1枚目のシート（旧・表記揺れ定義）
      const regData = ssReg.getSheets()[0].getDataRange().getValues();
      const rHead = regData[0].map(String);
      const cOfficial = rHead.findIndex(h => h.includes('正規記載'));
      if (cOfficial !== -1) {
        for (let r = 1; r < regData.length; r++) {
          const offName = String(regData[r][cOfficial]).trim();
          if (offName) {
            clinicDict[offName] = offName;
            rHead.forEach((h, i) => {
              if (h.includes('表記揺れ')) {
                const v = String(regData[r][i]).trim();
                if (v) clinicDict[v] = offName;
              }
            });
          }
        }
      }

      // 「拠点名」シート（B〜E列の別名とH列の開業日）
      const sheetLoc = ssReg.getSheetByName('拠点名');
      if (sheetLoc) {
        const locData = sheetLoc.getDataRange().getValues();
        for (let i = 1; i < locData.length; i++) {
          const rawOfficial = String(locData[i][0] || "").trim();
          if (!rawOfficial) continue;
          clinicDict[rawOfficial] = rawOfficial;
          for (let j = 1; j <= 4; j++) {
            const alias = String(locData[i][j] || "").trim();
            if (alias) clinicDict[alias] = rawOfficial;
          }
          const normOfficial = normalizeClinic(rawOfficial);
          const date = locData[i][7];
          if (normOfficial && date instanceof Date) openDateMap.set(normOfficial, date);
        }
      }

      // 「氏名」シート（氏名の表記ゆれ）
      const sheetName = ssReg.getSheetByName('氏名');
      if (sheetName) {
        const nameData = sheetName.getDataRange().getValues();
        for (let i = 1; i < nameData.length; i++) {
          const correctName = String(nameData[i][0] || "").trim();
          if (!correctName) continue;
          for (let j = 1; j < nameData[i].length; j++) {
            const alias = String(nameData[i][j] || "").trim();
            if (alias) nameAliasMap.set(alias, correctName);
          }
        }
      }
    } catch(e) {
      Logger.log(`正規表現読込警告: ${e.message}`);
    }
  }

  // ========================================================
  // 2. 常勤・定期非常勤・退職者マスタの読込 (新仕様 + 旧仕様フォールバック)
  // ========================================================
  const fullTimeMap = new Map(); // key(medNo or nameKey) -> { join, leave }
  const regPtMap = new Map();    // key(medNo or nameKey) -> { join, leave }
  const retiredMap = new Map();

  const registerDoctor = (targetMap, medNo, nameKey, joinDate, leaveDate) => {
    const info = {
      join: (joinDate instanceof Date && !isNaN(joinDate.getTime())) ? joinDate : null,
      leave: (leaveDate instanceof Date && !isNaN(leaveDate.getTime())) ? leaveDate : null
    };
    if (medNo) targetMap.set(medNo, info);
    if (nameKey) targetMap.set(nameKey, info);
  };

  const loadMasterFromSs = (ss) => {
    if (!ss) return;
    const sheets = ss.getSheets();

    // [A] 常勤シート（常勤2026年度 / 常勤マスタ 等）
    const ftSheets = sheets.filter(s => s.getName().includes('常勤'));
    ftSheets.forEach(sFt => {
      if (sFt.getLastRow() < 2) return;
      const data = sFt.getDataRange().getValues();
      const hMap = createHeaderMap(data[0]);
      const cName = hMap['医師名'] ?? hMap['氏名'] ?? hMap['名前'] ?? 0;
      const cMed  = hMap['医籍番号'] ?? 1;
      const cJoin = hMap['入職日'] ?? -1;
      const cRet  = hMap['退職日'] ?? -1;

      for (let i = 1; i < data.length; i++) {
        const r = data[i];
        const nameKey = normalizeName(r[cName]);
        const medNo = cleanId(r[cMed]);
        if (!nameKey && !medNo) continue;
        const joinD = cJoin !== -1 ? r[cJoin] : null;
        const retD  = cRet !== -1 ? r[cRet] : null;
        registerDoctor(fullTimeMap, medNo, nameKey, joinD, retD);
      }
    });

    // [B] 定期非常勤シート（定期非常勤2026年度 / 定期非常勤マスタ / 定期勤務医師 等）
    const ptSheets = sheets.filter(s => s.getName().includes('定期'));
    ptSheets.forEach(sPt => {
      if (sPt.getLastRow() < 2) return;
      const data = sPt.getDataRange().getValues();
      const hMap = createHeaderMap(data[0]);
      const cName = hMap['医師名'] ?? hMap['氏名'] ?? hMap['名前'] ?? 0;
      const cMed  = hMap['医籍番号'] ?? 1;
      const cJoin = hMap['入職日'] ?? -1;
      const cRet  = hMap['退職日'] ?? -1;

      for (let i = 1; i < data.length; i++) {
        const r = data[i];
        const nameKey = normalizeName(r[cName]);
        const medNo = cleanId(r[cMed]);
        if (!nameKey && !medNo) continue;
        const joinD = cJoin !== -1 ? r[cJoin] : null;
        const retD  = cRet !== -1 ? r[cRet] : null;
        registerDoctor(regPtMap, medNo, nameKey, joinD, retD);
      }
    });
  };

  // ① 新・外部医師マスタ（1aEjphEv...）から読込
  try {
    loadMasterFromSs(SpreadsheetApp.openByUrl(NEW_MASTER_URL));
  } catch (e) {
    Logger.log(`新医師マスタ読込警告: ${e.message}`);
  }

  // ② データセット内の「定期勤務医師」「人件費くん」からも補完読込
  ['定期勤務医師', '人件費くん'].forEach(key => {
    const url = getUrl(key);
    if (url && url !== NEW_MASTER_URL) {
      try { loadMasterFromSs(SpreadsheetApp.openByUrl(url)); } catch (e) {}
    }
  });

  // ③ 退職医師リストの読込
  const urlRetired = getUrl('退職医師リスト') || getUrl('人件費くん');
  if (urlRetired) {
    try {
      const ssRet = SpreadsheetApp.openByUrl(urlRetired);
      const sRet = ssRet.getSheetByName('退職医師') || ssRet.getSheets().find(s => s.getName().includes('退職'));
      if (sRet && sRet.getLastRow() > 1) {
        const rData = sRet.getDataRange().getValues();
        const rh = createHeaderMap(rData[0]);
        const rcMed = rh['医籍番号'] ?? -1;
        const rcName = rh['医師名'] ?? rh['氏名'] ?? -1;
        const rcType = rh['区分'] ?? -1;
        const rcJoin = rh['入職日'] ?? -1;
        const rcLeave = rh['退職日'] ?? -1;

        for (let i = 1; i < rData.length; i++) {
          const rr = rData[i];
          const med = rcMed !== -1 ? cleanId(rr[rcMed]) : "";
          const name = rcName !== -1 ? normalizeName(rr[rcName]) : "";
          const type = rcType !== -1 ? String(rr[rcType] || "").trim() : "";
          const join = (rcJoin !== -1 && rr[rcJoin] instanceof Date) ? rr[rcJoin] : null;
          const leave = (rcLeave !== -1 && rr[rcLeave] instanceof Date) ? rr[rcLeave] : null;
          if ((med || name) && leave) {
            const info = { type, join, leave };
            if (med) retiredMap.set(med, info);
            if (name) retiredMap.set(name, info);
          }
        }
      }
    } catch (e) {}
  }

  // ========================================================
  // 3. 紹介会社データの読込 (新・紹介会社医師履歴 + 紹介会社進捗)
  // ========================================================
  const agencyDoctorMedIds = new Set();
  const agencyMap = new Map();

  // ① 新・紹介会社医師履歴シート (医籍番号ベース)
  try {
    const agencySs = SpreadsheetApp.openByUrl(AGENCY_HISTORY_URL);
    const agencySheet = agencySs.getSheetByName("紹介会社医師履歴");
    if (agencySheet && agencySheet.getLastRow() > 1) {
      const agencyData = agencySheet.getDataRange().getValues();
      const aHMap = createHeaderMap(agencyData[0]);
      const cMed = aHMap['医籍番号'] ?? 0;
      for (let i = 1; i < agencyData.length; i++) {
        const rawMedId = cleanId(agencyData[i][cMed]);
        if (rawMedId) agencyDoctorMedIds.add(rawMedId);
      }
    }
  } catch (e) {}

  // ② 従来の紹介会社進捗シート (氏名_勤務日ベース)
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

  // ========================================================
  // 4. 出力先シート（YYYY拠点）の行マッピング取得
  // ========================================================
  const ssOut = SpreadsheetApp.openByUrl(CONFIG_URL);
  const sheetOut = ssOut.getSheetByName(SHEET_NAME_OUT);
  if (!sheetOut) return `エラー: シート「${SHEET_NAME_OUT}」が見つかりません。`;

  const rangeData = sheetOut.getRange(2, 1, sheetOut.getLastRow() - 1, 3).getValues();
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

  // ========================================================
  // 5. シフトデータの取得と集計
  // ========================================================
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth() + 1;
  const monthDiff = (currentYear - targetYear) * 12 + (currentMonth - targetMonth);

  let urlShift = "";
  let targetSheetName = "";

  if (monthDiff <= 1) {
    urlShift = getUrl('統合シフト') || 'https://docs.google.com/spreadsheets/d/1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co/edit';
    targetSheetName = "確定シフト";
  } else {
    if (targetYear === 2026) {
      urlShift = getUrl('2026確定シフト') || 'https://docs.google.com/spreadsheets/d/10Z8jg4o7Ri9Ggf7u_jauyDDZa6bVnc-iuuuqmkNDafw/edit';
    } else if (targetYear === 2025) {
      urlShift = getUrl('2025確定シフト') || 'https://docs.google.com/spreadsheets/d/1JhLJuxkp6T5hEkfultUE57pmRXBI_lkN6ROp2-CuhCc/edit';
    } else {
      urlShift = getUrl(`${targetYear}確定シフト`);
    }
    targetSheetName = `${targetYear}/${String(targetMonth).padStart(2, '0')}`;
  }

  if (!urlShift) return `エラー: 対象月(${targetYear}/${targetMonth})のシフトURLが取得できません`;

  let shiftSheet = null;
  try {
    const ssShift = SpreadsheetApp.openByUrl(urlShift);
    shiftSheet = ssShift.getSheetByName(targetSheetName) || ssShift.getSheetByName("確定シフト");
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
      ft: { c: new Set(), h: 0 },
      ftOver: { h: 0 },
      reg: { c: new Set(), h: 0 },
      spot: { c: new Set(), h: 0 },
      spotAg: { c: new Set(), h: 0 },
      spotDir: { c: new Set(), h: 0 }
    };
  });

  // 時間のシリアル値・Date・文字列("09:00")すべてに対応する変換関数
  const toH = (v) => {
    if (v instanceof Date) return v.getHours() + v.getMinutes() / 60;
    if (typeof v === 'number') return v * 24;
    if (typeof v === 'string' && v.includes(':')) {
      const parts = v.trim().split(':').map(Number);
      if (!isNaN(parts[0]) && !isNaN(parts[1])) return parts[0] + parts[1] / 60;
    }
    return null;
  };

  data.forEach(row => {
    if (C.DATE === -1) return;
    const dVal = new Date(row[C.DATE]);
    if (isNaN(dVal.getTime())) return;
    if (dVal.getFullYear() !== targetYear || (dVal.getMonth() + 1) !== targetMonth) return;

    const rawClinic = String(row[C.CLINIC] || "").trim();
    if (/有給|欠勤|医師会|院外|バックアップ|嘱託医/.test(rawClinic)) return;

    const clinic = normalizeClinic(rawClinic);
    if (!clinicStats[clinic]) return;

    // 亀有・北葛西の診療科チェック（内科と書かれていない限り小児科として集計）
    if (['亀有', '北葛西'].includes(clinic)) {
      const deptStr = String(row[C.DEPT] || "");
      const isNaika = rawClinic.includes('内科') || deptStr.includes('内科');
      if (isNaika && !deptStr.includes('小児科')) return;
    }

    // ★修正: C.MED バグを C.MED_NO に修正
    const medNo = (C.MED_NO > -1) ? cleanId(row[C.MED_NO]) : "";
    const nameRaw = (C.NAME > -1) ? String(row[C.NAME] || "").trim() : "";
    const nameKey = normalizeName(nameRaw);
    const id = medNo || nameKey;
    if (!id) return;

    let s = toH(row[C.START_M]), e = toH(row[C.END_M]);
    if (s == null || e == null) { s = toH(row[C.START_B]); e = toH(row[C.END_B]); }
    if (s === null || e === null) return;
    if (e < s && e !== 0) e += 24;
    if ((e - s) < 0.5) return;

    let h = e - s;
    const bo = Math.max(0, Math.min(e, 15) - Math.max(s, 13));
    if (h > bo && s < 14.9) h -= bo;
    if (h <= 0) return;

    const note = (C.NOTE > -1) ? String(row[C.NOTE] || "") : "";
    const comments = commentCols.map(c => String(row[c] || "")).join(" ");
    const combinedText = comments + " " + note;

    // ========================================================
    // ★ jinkenhikun1 準拠の医師区分・所定休出判定ロジック
    // ========================================================
    let isFt = false;
    let isReg = false;

    const ftInfo = fullTimeMap.get(medNo) || fullTimeMap.get(nameKey);
    if (ftInfo) {
      const afterJoin = !ftInfo.join || dVal >= ftInfo.join;
      const beforeLeave = !ftInfo.leave || dVal <= ftInfo.leave;
      if (afterJoin && beforeLeave) isFt = true;
    }

    const ptInfo = regPtMap.get(medNo) || regPtMap.get(nameKey);
    if (!isFt && ptInfo) {
      const afterJoin = !ptInfo.join || dVal >= ptInfo.join;
      const beforeLeave = !ptInfo.leave || dVal <= ptInfo.leave;
      if (afterJoin && beforeLeave) isReg = true;
    }

    // 退職医師リストによる補正
    const retiredInfo = retiredMap.get(medNo) || retiredMap.get(nameKey);
    if (retiredInfo) {
      if (dVal > retiredInfo.leave || (retiredInfo.join && dVal < retiredInfo.join)) {
        isFt = false;
        isReg = false;
      } else {
        if (retiredInfo.type.includes('常勤')) isFt = true;
        else if (retiredInfo.type.includes('定期')) isReg = true;
      }
    }

    // 所定休出・ヘルプ等の判定 (jinkenhikun1: comment.includes("所定休出"))
    const isKyushutsu = /所定休出|所定外/.test(combinedText);
    const isRelief = isKyushutsu || /特別時給|追加|欠勤対応|急遽|ヘルプ|バックアップ|民間医局|紹介|加算/.test(combinedText);
    const isRegularNote = /毎週(月|火|水|木|金|土|日)曜|第[1-5](月|火|水|木|金|土|日)曜/.test(note);

    // 紹介会社判定 (jinkenhikun1: agencyDoctorMedIds または agencyMap または コメント)
    const dateKey = `${nameKey}_${formatDateKey(dVal)}`;
    const isAgency = (medNo && agencyDoctorMedIds.has(medNo)) || agencyMap.has(dateKey) || /紹介|民間|MRT|エムスリー|マイナビ|メディウェル|リクルート|オンコール/.test(comments);

    if (isKyushutsu) {
      // 🌟 所定休出: 所定休出時間に加算し、かつスポット全体にも含める（C_DailyCost_Writer.gs 156行目準拠）
      clinicStats[clinic].ftOver.h += h;
      clinicStats[clinic].spot.c.add(id);
      clinicStats[clinic].spot.h += h;
      if (isAgency && !isFt && !isReg) {
        clinicStats[clinic].spotAg.c.add(id);
        clinicStats[clinic].spotAg.h += h;
      } else {
        clinicStats[clinic].spotDir.c.add(id);
        clinicStats[clinic].spotDir.h += h;
      }
    } else if (isFt) {
      // 🌟 常勤医師の通常勤務（wage === 0 制限を撤廃し確実に集計）
      clinicStats[clinic].ft.c.add(id);
      clinicStats[clinic].ft.h += h;
    } else if (isRelief) {
      // 🌟 非常勤・スポットのヘルプ/特別時給勤務
      clinicStats[clinic].spot.c.add(id);
      clinicStats[clinic].spot.h += h;
      if (isAgency) {
        clinicStats[clinic].spotAg.c.add(id);
        clinicStats[clinic].spotAg.h += h;
      } else {
        clinicStats[clinic].spotDir.c.add(id);
        clinicStats[clinic].spotDir.h += h;
      }
    } else if (isReg || isRegularNote) {
      // 🌟 定期非常勤医師
      clinicStats[clinic].reg.c.add(id);
      clinicStats[clinic].reg.h += h;
    } else {
      // 🌟 通常スポット医師
      clinicStats[clinic].spot.c.add(id);
      clinicStats[clinic].spot.h += h;
      if (isAgency) {
        clinicStats[clinic].spotAg.c.add(id);
        clinicStats[clinic].spotAg.h += h;
      } else {
        clinicStats[clinic].spotDir.c.add(id);
        clinicStats[clinic].spotDir.h += h;
      }
    }
  });

  // ========================================================
  // 6. シートへの書き込み
  // ========================================================
  const updates = [];
  targetClinics.forEach(clinic => {
    const s = clinicStats[clinic];
    const totalH = s.ft.h + s.reg.h + s.spot.h;
    const openDate = openDateMap.get(clinic);
    const isNotOpenYet = (openDate && openDate > new Date(targetYear, targetMonth, 0));
    const rate = (v) => totalH > 0 ? v / totalH : 0;

    const items = [
      { k: '常勤医師数', v: s.ft.c.size },
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
      { k: '直接応募構成比', v: rate(s.spotDir.h) }
    ];

    items.forEach(item => {
      const row = rowMap.get(`${clinic}_${normalizeItemKey(item.k)}`);
      if (row) {
        let val = isNotOpenYet ? 0 : item.v;
        if (!item.k.includes('構成比')) val = Math.round(val);
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