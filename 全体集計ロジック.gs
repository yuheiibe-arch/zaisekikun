/**
 * 🏆 【全体集計】updateDashboardFinal_Fixed (2026新マスタ完全準拠版)
 */
function updateDashboardFinal_Fixed(targetYear, targetMonth) {
  if (!targetYear || !targetMonth) return "年月指定エラー";
  targetYear = Number(targetYear);
  targetMonth = Number(targetMonth);

  const fiscalYear = (targetMonth >= 4) ? targetYear : targetYear - 1;
  const OUTPUT_SHEET_NAME = `${fiscalYear}年度`;

  const CONFIG_URL = 'https://docs.google.com/spreadsheets/d/1CKiJYyZ0iF6cobH_ods2hrI9RmFKt6gFiRXUjD195o0/edit';
  const NEW_MASTER_URL = 'https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit';

  const getUrlFromDataset = (configUrl, key) => {
    try {
      const ss = SpreadsheetApp.openByUrl(configUrl);
      const data = ss.getSheetByName('データセット').getDataRange().getValues();
      const row = data.find(r => String(r[0]).trim() === key);
      return row ? String(row[1]).trim() : "";
    } catch(e) { return ""; }
  };

  const classifyDept = (str) => {
    const s = String(str).replace(/[\s ]+/g, "").trim();
    const parts = s.split("、");
    if (parts.some(p => p === "小児科" || p === "小児科小児科")) return "Ped";
    if (parts.some(p => p === "内科")) return "Int";
    if (s.includes("ワクチン")) return "Vac";
    return "Other";
  };

  const nameAliasMap = new Map();
  const normalizeName = (n) => {
    if (!n) return "";
    const raw = String(n).trim();
    if (nameAliasMap.has(raw)) return String(nameAliasMap.get(raw)).replace(/[\s ]+/g, "");
    const noSpace = raw.replace(/[\s ]+/g, "");
    if (nameAliasMap.has(noSpace)) return String(nameAliasMap.get(noSpace)).replace(/[\s ]+/g, "");
    return noSpace;
  };
  const cleanId = (rawId) => String(rawId || "").replace(/\.0+$/, '').trim();
  const createHeaderMap = (headerRow) => {
    const map = {};
    headerRow.forEach((val, idx) => {
      if (val) map[String(val).replace(/\n/g, '').trim()] = idx;
    });
    return map;
  };

  let targetColIndex = (targetMonth >= 4) ? targetMonth : targetMonth + 12;

  const urlRegex = getUrlFromDataset(CONFIG_URL, '正規表現');
  const urlMaster = getUrlFromDataset(CONFIG_URL, '医師マスタ');
  const urlCost = getUrlFromDataset(CONFIG_URL, '人件費くん');
  const urlHoliday = getUrlFromDataset(CONFIG_URL, '祝日');

  if (!urlRegex || !urlMaster) return "設定エラー: マスタURL不足";

  const validClinics = [];
  const holidays = new Set();
  const doctorInfoMap = new Map();
  const retiredMap = new Map();

  const targetDateStart = new Date(targetYear, targetMonth - 1, 1);
  const targetEndDate = new Date(targetYear, targetMonth, 0);

  // --- 1. マスタ読込 ---
  try {
    const ss = SpreadsheetApp.openByUrl(urlRegex);
    const regData = ss.getSheets()[0].getDataRange().getValues();
    for (let r = 1; r < regData.length; r++) {
      regData[r].forEach(cell => {
        const val = String(cell).trim();
        if (val) validClinics.push(val);
      });
    }
    const sLoc = ss.getSheetByName('拠点名');
    if (sLoc) {
      sLoc.getDataRange().getValues().slice(1).forEach(r => {
        for (let j = 0; j <= 4; j++) {
          const val = String(r[j] || "").trim();
          if (val) validClinics.push(val);
        }
      });
    }
    const sName = ss.getSheetByName('氏名');
    if (sName) {
      sName.getDataRange().getValues().slice(1).forEach(r => {
        const correct = String(r[0] || "").trim();
        if (correct) r.slice(1).forEach(a => { if (a) nameAliasMap.set(String(a).trim(), correct); });
      });
    }
  } catch(e){}

  try {
    const ss = SpreadsheetApp.openByUrl(urlHoliday);
    const s = ss.getSheetByName('祝日') || ss.getSheets()[0];
    s.getDataRange().getValues().forEach(r => {
      if(r[0] instanceof Date) holidays.add(Utilities.formatDate(r[0], 'Asia/Tokyo', 'yyyy-MM-dd'));
    });
  } catch(e){}

  // ★ 新・外部医師マスタ（常勤2026年度 / 定期非常勤2026年度）および既存マスタの読込
  const loadDoctorInfoFromSs = (ss) => {
    if (!ss) return;
    const sheets = ss.getSheets();

    sheets.filter(s => s.getName().includes('常勤')).forEach(sFt => {
      if (sFt.getLastRow() < 2) return;
      const data = sFt.getDataRange().getValues();
      const hMap = createHeaderMap(data[0]);
      const cName = hMap['医師名'] ?? hMap['氏名'] ?? 0;
      const cMed  = hMap['医籍番号'] ?? 1;
      const cJoin = hMap['入職日'] ?? 5;
      const cRet  = hMap['退職日'] ?? -1;

      for (let i = 1; i < data.length; i++) {
        const r = data[i];
        const name = normalizeName(r[cName]);
        const id = cleanId(r[cMed]);
        if (!name && !id) continue;
        const joinDate = (cJoin !== -1 && r[cJoin] instanceof Date) ? r[cJoin] : null;
        const leaveDate = (cRet !== -1 && r[cRet] instanceof Date) ? r[cRet] : null;
        const info = { type: 'FT', join: joinDate, leave: leaveDate };
        if (id) doctorInfoMap.set(id, info);
        if (name) doctorInfoMap.set(name, info);
      }
    });

    sheets.filter(s => s.getName().includes('定期')).forEach(sReg => {
      if (sReg.getLastRow() < 2) return;
      const data = sReg.getDataRange().getValues();
      const hMap = createHeaderMap(data[0]);
      const cName = hMap['医師名'] ?? hMap['氏名'] ?? 0;
      const cMed  = hMap['医籍番号'] ?? 1;
      const cJoin = hMap['入職日'] ?? 4;
      const cRet  = hMap['退職日'] ?? -1;

      for (let i = 1; i < data.length; i++) {
        const r = data[i];
        const name = normalizeName(r[cName]);
        const medNo = cleanId(r[cMed]);
        if (!name && !medNo) continue;
        const joinDate = (cJoin !== -1 && r[cJoin] instanceof Date) ? r[cJoin] : null;
        const leaveDate = (cRet !== -1 && r[cRet] instanceof Date) ? r[cRet] : null;
        const info = { type: 'Reg', join: joinDate, leave: leaveDate };
        if (medNo && !doctorInfoMap.has(medNo)) doctorInfoMap.set(medNo, info);
        if (name && !doctorInfoMap.has(name)) doctorInfoMap.set(name, info);
      }
    });
  };

  try { loadDoctorInfoFromSs(SpreadsheetApp.openByUrl(NEW_MASTER_URL)); } catch(e){}
  ['定期勤務医師', '人件費くん'].forEach(k => {
    const u = getUrlFromDataset(CONFIG_URL, k);
    if (u && u !== NEW_MASTER_URL) {
      try { loadDoctorInfoFromSs(SpreadsheetApp.openByUrl(u)); } catch(e){}
    }
  });

  try {
    let urlRetired = getUrlFromDataset(CONFIG_URL, '退職医師リスト') || urlCost;
    if (urlRetired) {
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
    }
  } catch(e){}

  // --- 2. 在籍統計 (Stock) ---
  const stockStats = {
    total: 0, fullTime: { count: 0, internal: 0, peds: 0 }, regPt: { count: 0, internal: 0, peds: 0 },
    spot: { count: 0, internal: 0, peds: 0, vaccine: 0 }, newReg: { total: 0, internal: 0, peds: 0, vaccine: 0, ft:0, rpt:0, sp:0 },
    routes: { direct: 0, agent: 0, referral: 0, other: 0 }, access: { all: Array(5).fill(0), ft: Array(5).fill(0), rpt: Array(5).fill(0), sp: Array(5).fill(0) }
  };
  const getDept = (str) => {
    const parts = str.split('、');
    return { hasInt: parts.includes('内科'), hasPed: parts.includes('小児科'), isVac: str.includes('ワクチン') && !parts.includes('内科') && !parts.includes('小児科') };
  };
  const getRoute = (str) => {
    if (str.match(/直接|HP|ホームページ/)) return 'direct';
    if (str.match(/紹介会社|エムスリー|民間/)) return 'agent';
    if (str.match(/紹介|知人/)) return 'referral';
    return 'other';
  };
  const getDaysDiff = (lastLogin, baseDate) => {
    if (!lastLogin || isNaN(lastLogin.getTime())) return 9999;
    if (lastLogin > baseDate) return 0;
    return (baseDate.getTime() - lastLogin.getTime()) / (1000 * 3600 * 24);
  };

  try {
    const ss = SpreadsheetApp.openByUrl(urlMaster);
    const sheet = ss.getSheets()[0];
    const data = sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).getValues();
    data.forEach(row => {
      const cDate = row[36] ? new Date(row[36]) : null;
      if (cDate && cDate > targetEndDate) return;

      const medNo = cleanId(row[10]);
      const nameKey = normalizeName(row[0]);
      if (!medNo && !nameKey) return;

      stockStats.total++;
      const deptRaw = String(row[11]);
      const dept = getDept(deptRaw);
      const routeRaw = String(row[20]);
      const lastLogin = row[38] ? new Date(row[38]) : null;

      let type = 'spot';
      let info = doctorInfoMap.get(medNo) || doctorInfoMap.get(nameKey);

      if (info) {
        const isRetired = info.leave && info.leave < targetDateStart;
        if (!isRetired) {
          if (info.type === 'FT') type = 'fullTime';
          else if (info.type === 'Reg') type = 'regPt';
        }
      }

      const retiredInfo = retiredMap.get(medNo) || retiredMap.get(nameKey);
      if (retiredInfo && retiredInfo.leave < targetDateStart) {
        type = 'spot';
      }

      if (type === 'fullTime') {
        stockStats.fullTime.count++;
        if (dept.hasInt) stockStats.fullTime.internal++;
        if (dept.hasPed) stockStats.fullTime.peds++;
      } else if (type === 'regPt') {
        stockStats.regPt.count++;
        if (dept.hasInt) stockStats.regPt.internal++;
        if (dept.hasPed) stockStats.regPt.peds++;
      } else {
        stockStats.spot.count++;
        const cat = classifyDept(deptRaw);
        if (cat === 'Int') stockStats.spot.internal++;
        else if (cat === 'Ped') stockStats.spot.peds++;
        else if (cat === 'Vac') stockStats.spot.vaccine++;
      }

      if (cDate && cDate.getFullYear() === targetYear && cDate.getMonth() === (targetMonth - 1)) {
        stockStats.newReg.total++;
        const cat = classifyDept(deptRaw);
        if (cat === 'Int') stockStats.newReg.internal++;
        else if (cat === 'Ped') stockStats.newReg.peds++;
        else if (cat === 'Vac') stockStats.newReg.vaccine++;

        stockStats.routes[getRoute(routeRaw)]++;
        if (type === 'fullTime') stockStats.newReg.ft++;
        if (type === 'regPt') stockStats.newReg.rpt++;
        if (type === 'spot') stockStats.newReg.sp++;
      }

      const dDays = getDaysDiff(lastLogin, targetEndDate);
      const inc = (arr) => {
        if (dDays <= 30) arr[0]++; else if (dDays <= 90) arr[1]++; else if (dDays <= 180) arr[2]++; else if (dDays <= 365) arr[3]++; else arr[4]++;
      };
      inc(stockStats.access.all);
      if (type === 'fullTime') inc(stockStats.access.ft);
      if (type === 'regPt') inc(stockStats.access.rpt);
      if (type === 'spot') inc(stockStats.access.sp);
    });
  } catch (e) { return `医師マスタ集計エラー: ${e.message}`; }

  // --- 3. シフト集計 (Active) ---
  const workStats = {
    attr: { fullTime: new Set(), regPt: new Set(), spot: new Set() },
    spotForDashboard: new Set(),
    logic: {
      spot: new Set(),
      spotDetails: {
        days: { Mon:new Set(),Tue:new Set(),Wed:new Set(),Thu:new Set(),Fri:new Set(),Sat:new Set(),Sun:new Set(),Hol:new Set() },
        hours: { Mon:0,Tue:0,Wed:0,Thu:0,Fri:0,Sat:0,Sun:0,Hol:0,Total:0 }
      }
    }
  };

  try {
    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth() + 1;
    const monthDiff = (currentYear - targetYear) * 12 + (currentMonth - targetMonth);

    let urlShift = "";
    let targetSheetName = "";

    if (monthDiff <= 1) {
      urlShift = getUrlFromDataset(CONFIG_URL, '統合シフト') || 'https://docs.google.com/spreadsheets/d/1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co/edit';
      targetSheetName = "確定シフト";
    } else {
      if (targetYear === 2026) {
        urlShift = getUrlFromDataset(CONFIG_URL, '2026確定シフト') || 'https://docs.google.com/spreadsheets/d/10Z8jg4o7Ri9Ggf7u_jauyDDZa6bVnc-iuuuqmkNDafw/edit';
      } else if (targetYear === 2025) {
        urlShift = getUrlFromDataset(CONFIG_URL, '2025確定シフト') || 'https://docs.google.com/spreadsheets/d/1JhLJuxkp6T5hEkfultUE57pmRXBI_lkN6ROp2-CuhCc/edit';
      } else {
        urlShift = getUrlFromDataset(CONFIG_URL, `${targetYear}確定シフト`);
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

    const allDataRows = shiftSheet.getDataRange().getValues();
    const headers = allDataRows[0].map(String);
    const data = allDataRows.slice(1);

    if (headers.length > 0) {
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
        START_M: findCol(headers, ['勤務開始時間', '開始時間']),
        END_M: findCol(headers, ['勤務終了時間', '終了時間']),
        START_B: findCol(headers, ['元のシフト勤務開始時間', '元のシフト)勤務開始時間', '元のシフト）勤務開始時間']),
        END_B: findCol(headers, ['元のシフト勤務終了時間', '元のシフト)勤務終了時間', '元のシフト）勤務終了時間']),
        NOTE: findCol(headers, ['業務内容', '備考'])
      };

      const sMIdx = headers.findIndex(h => String(h).includes('勤務開始時間') && !String(h).includes('元の'));
      if (sMIdx !== -1) C.START_M = sMIdx;
      const eMIdx = headers.findIndex(h => String(h).includes('勤務終了時間') && !String(h).includes('元の'));
      if (eMIdx !== -1) C.END_M = eMIdx;

      const COMMENT_IDXS = headers.map((h, i) => (String(h).includes('スタッフコメント') || String(h).includes('コメント')) ? i : -1).filter(i => i !== -1);

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
        const dateVal = new Date(row[C.DATE]);
        if (isNaN(dateVal.getTime())) return;
        if (dateVal.getFullYear() !== targetYear || (dateVal.getMonth() + 1) !== targetMonth) return;

        const clinicRaw = String(row[C.CLINIC] || "").trim();
        if (/有給|欠勤|医師会|院外|バックアップ|嘱託医/.test(clinicRaw)) return;
        const isNormalClinic = validClinics.some(c => clinicRaw.includes(c) || c.includes(clinicRaw) || (clinicRaw.includes("NT") && c.includes("千葉")));
        if (!isNormalClinic) return;

        let s = toH(row[C.START_M]), e = toH(row[C.END_M]);
        if (s == null || e == null) { s = toH(row[C.START_B]); e = toH(row[C.END_B]); }
        if (s === null || e === null) return;
        if (e < s && e !== 0) e += 24;
        if ((e - s) < 0.5) return;

        let h = e - s;
        const breakOverlap = Math.max(0, Math.min(e, 15) - Math.max(s, 13));
        if (h > breakOverlap && s < 15) { h -= breakOverlap; }
        if (h < 0) h = 0;

        const medNo = (C.MED_NO > -1) ? cleanId(row[C.MED_NO]) : "";
        const nameKey = (C.NAME > -1) ? normalizeName(row[C.NAME] || "") : "";
        const id = medNo || nameKey;
        if (!id) return;

        const note = (C.NOTE > -1) ? String(row[C.NOTE] || "") : "";
        const allComments = COMMENT_IDXS.map(idx => String(row[idx] || "")).join(" ");

        const dStr = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy-MM-dd');
        let day = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'EEE');
        if (holidays.has(dStr) && day !== 'Sun') day = 'Hol';
        const isSunOrHol = (day === 'Sun' || day === 'Hol');

        let attrType = 'Spot';
        let doctorInfo = doctorInfoMap.get(medNo) || doctorInfoMap.get(nameKey);

        if (doctorInfo) {
          const isBeforeJoin = doctorInfo.join && dateVal < doctorInfo.join;
          const isAfterLeave = doctorInfo.leave && dateVal > doctorInfo.leave;
          if (!isBeforeJoin && !isAfterLeave) {
            attrType = doctorInfo.type; // 'FT' or 'Reg'
          }
        }

        const retiredInfo = retiredMap.get(medNo) || retiredMap.get(nameKey);
        if (retiredInfo && dateVal > retiredInfo.leave) {
          attrType = 'Spot';
        }

        if (attrType === 'FT') workStats.attr.fullTime.add(id);
        else if (attrType === 'Reg') workStats.attr.regPt.add(id);
        else workStats.attr.spot.add(id);

        let isDashboardSpot = false;
        if (attrType === 'Spot') {
          isDashboardSpot = true;
        } else if (attrType === 'Reg') {
          if (!/(毎週|第|曜)/.test(note)) isDashboardSpot = true;
        }
        if (isDashboardSpot) workStats.spotForDashboard.add(id);

        const isRelief = /所定休出|所定外|特別時給|追加|欠勤対応|急遽|ヘルプ|バックアップ|民間医局|紹介|加算/.test(allComments + " " + note);
        const hasRegularNote = /【.*】.*(毎週|第[1-5]|曜|[月火水木金土日]：)/.test(note);

        let isSpotLogic = false;
        if (isRelief) isSpotLogic = true;
        else if (hasRegularNote) isSpotLogic = false;
        else {
          if (attrType === 'FT') isSpotLogic = isSunOrHol;
          else isSpotLogic = true;
        }

        if (isSpotLogic) {
          workStats.logic.spot.add(id);
          if (workStats.logic.spotDetails.days[day]) {
            workStats.logic.spotDetails.days[day].add(id);
            workStats.logic.spotDetails.hours[day] += h;
            workStats.logic.spotDetails.hours.Total += h;
          }
        }
      });
    }
  } catch(e) { return `シフト集計エラー: ${e.message}`; }

  // --- 4. 書き込み ---
  try {
    const ss = SpreadsheetApp.openByUrl(CONFIG_URL);
    const sheet = ss.getSheetByName(OUTPUT_SHEET_NAME);
    if (!sheet) return `エラー: シート「${OUTPUT_SHEET_NAME}」がありません`;

    const outMap = new Map();
    const rData = sheet.getRange(2,1,76,3).getValues();
    let lA=""; rData.forEach((r,i)=>{ if(r[0]) lA=String(r[0]); if(r[2]) outMap.set(lA+'_'+r[2], i); });
    const fVals = new Array(76).fill([0]);

    const isRatioItem = (itemName) => itemName.includes('率') || itemName.includes('割') || itemName.includes('構成比');
    const setV = (cA, cC, v) => {
      const idx = outMap.get(cA+'_'+cC);
      if(idx!==undefined) fVals[idx] = [typeof v==='number' && !isRatioItem(cC) ? Math.round(v) : v];
    };
    const safeRatio = (numerator, denominator) => denominator > 0 && isFinite(numerator / denominator) ? numerator / denominator : 0;

    const actFullTime = workStats.attr.fullTime.size;
    const actRegPt = workStats.attr.regPt.size;
    const actSpotAttr = workStats.spotForDashboard.size;

    const allActiveIds = new Set([...workStats.attr.fullTime, ...workStats.attr.regPt, ...workStats.spotForDashboard]);
    const actTotal = allActiveIds.size;

    setV('医師合計','在籍数',stockStats.total); setV('医師合計','稼働数',actTotal); setV('医師合計','アクティブ率',safeRatio(actTotal, stockStats.total));
    setV('常勤医師','在籍数',stockStats.fullTime.count); setV('常勤医師','稼働数',actFullTime); setV('常勤医師','稼働割合',safeRatio(actFullTime, stockStats.fullTime.count));
    setV('常勤医師','内科',stockStats.fullTime.internal); setV('常勤医師','小児科',stockStats.fullTime.peds);
    setV('定期非常勤医師','在籍数',stockStats.regPt.count); setV('定期非常勤医師','稼働数',actRegPt); setV('定期非常勤医師','稼働割合',safeRatio(actRegPt, stockStats.regPt.count));
    setV('定期非常勤医師','内科',stockStats.regPt.internal); setV('定期非常勤医師','小児科',stockStats.regPt.peds);

    setV('スポット医師','在籍数',stockStats.spot.count); setV('スポット医師','稼働数',actSpotAttr); setV('スポット医師','稼働割合',safeRatio(actSpotAttr, stockStats.spot.count));
    setV('スポット医師','内科',stockStats.spot.internal); setV('スポット医師','小児科',stockStats.spot.peds); setV('スポット医師','ワクチン専任',stockStats.spot.vaccine);

    const sd = workStats.logic.spotDetails;
    setV('稼働スポット医師数','月間合計',workStats.logic.spot.size);
    ['Mon','Tue','Wed','Thu','Fri','Sat','Sun','Hol'].forEach((k,i)=> setV('稼働スポット医師数',['月','火','水','木','金','土','日','祝'][i], sd.days[k].size));

    setV('稼働スポット医師稼働時間数','月間合計',sd.hours.Total);
    ['Mon','Tue','Wed','Thu','Fri','Sat','Sun','Hol'].forEach((k,i)=> setV('稼働スポット医師稼働時間数',['月','火','水','木','金','土','日','祝'][i], sd.hours[k]));

    setV('DS登録増加数','登録合計',stockStats.newReg.total);
    setV('DS登録増加数','内科',stockStats.newReg.internal); setV('DS登録増加数','小児科',stockStats.newReg.peds); setV('DS登録増加数','ワクチン専任',stockStats.newReg.vaccine);
    setV('DS登録増加数','常勤合計',stockStats.newReg.ft); setV('DS登録増加数','定期非常勤合計',stockStats.newReg.rpt); setV('DS登録増加数','スポット合計',stockStats.newReg.sp);
    setV('DS登録経路','合計',stockStats.newReg.total);
    setV('DS登録経路','直接',stockStats.routes.direct); setV('DS登録経路','紹介会社',stockStats.routes.agent);
    setV('DS登録経路','その他',stockStats.routes.other); setV('DS登録経路','リファラル',stockStats.routes.referral);

    const acL = ['1ヶ月以内','3ヶ月以内','6ヶ月以内','1年以内','1年以上経過'];
    acL.forEach((l,i)=>{
      setV('ドクターサイトアクセス数',l,stockStats.access.all[i]);
      setV('ドクターサイトアクセス数（常勤）',l,stockStats.access.ft[i]);
      setV('ドクターサイトアクセス数（定期非常勤）',l,stockStats.access.rpt[i]);
      setV('ドクターサイトアクセス数（スポット）',l,stockStats.access.sp[i]);
    });

    const tR = sheet.getRange(2, targetColIndex, 76, 1);
    tR.setValues(fVals);
    tR.setHorizontalAlignment("right");
    return "完了";
  } catch(e) { return `書込エラー: ${e.message}`; }
}