/**
 * ====================================================================
 * 🏆 【全体集計】updateDashboardFinal_Fixed (モジュール分割・高速版)
 * ====================================================================
 * 共通基盤 (共通関数.gs) からマスタと対象月シフトを受け取り、全体統計を集計します。
 */
function updateDashboardFinal_Fixed(targetYear, targetMonth, cache = {}) {
  if (!targetYear || !targetMonth) return "年月指定エラー";
  targetYear = Number(targetYear);
  targetMonth = Number(targetMonth);

  const fiscalYear = (targetMonth >= 4) ? targetYear : targetYear - 1;
  const OUTPUT_SHEET_NAME = `${fiscalYear}年度`;
  const targetColIndex = (targetMonth >= 4) ? targetMonth : targetMonth + 12;

  const targetDateStart = new Date(targetYear, targetMonth - 1, 1);
  const targetEndDate = new Date(targetYear, targetMonth, 0);

  // 1. 共通マスタと医師マスタ全行の取得（キャッシュ利用）
  let core;
  try {
    core = getSharedCoreContext(cache, { needDoctorMaster: true });
  } catch (e) {
    return `マスタ読込エラー: ${e.message}`;
  }

  const {
    ssConfig, normNameFn, validClinics, holidays,
    doctorInfoMap, retiredMap, doctorMasterRows
  } = core;

  // 2. 在籍統計 (Stock) の計算
  const stockStats = {
    total: 0, fullTime: { count: 0, internal: 0, peds: 0 }, regPt: { count: 0, internal: 0, peds: 0 },
    spot: { count: 0, internal: 0, peds: 0, vaccine: 0 }, newReg: { total: 0, internal: 0, peds: 0, vaccine: 0, ft: 0, rpt: 0, sp: 0 },
    routes: { direct: 0, agent: 0, referral: 0, other: 0 }, access: { all: Array(5).fill(0), ft: Array(5).fill(0), rpt: Array(5).fill(0), sp: Array(5).fill(0) }
  };

  const classifyDept = (str) => {
    const s = String(str).replace(/[\s ]+/g, "").trim();
    const parts = s.split("、");
    if (parts.some(p => p === "小児科" || p === "小児科小児科")) return "Ped";
    if (parts.some(p => p === "内科")) return "Int";
    if (s.includes("ワクチン")) return "Vac";
    return "Other";
  };
  const getDept = (str) => {
    const parts = str.split('、');
    return { hasInt: parts.includes('内科'), hasPed: parts.includes('小児科') };
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

  doctorMasterRows.forEach(row => {
    const cDate = row[36] ? (row[36] instanceof Date ? row[36] : new Date(row[36])) : null;
    if (cDate && cDate > targetEndDate) return;

    const medNo = cleanId(row[10]);
    const nameKey = normNameFn(row[0]);
    if (!medNo && !nameKey) return;

    stockStats.total++;
    const deptRaw = String(row[11]);
    const dept = getDept(deptRaw);
    const routeRaw = String(row[20]);
    const lastLogin = row[38] ? (row[38] instanceof Date ? row[38] : new Date(row[38])) : null;

    let type = 'spot';
    const info = (medNo && doctorInfoMap.get(medNo)) || (nameKey && doctorInfoMap.get(nameKey));
    if (info && (!info.leave || info.leave >= targetDateStart)) {
      if (info.type === 'FT') type = 'fullTime';
      else if (info.type === 'Reg') type = 'regPt';
    }

    const retiredInfo = (medNo && retiredMap.get(medNo)) || (nameKey && retiredMap.get(nameKey));
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

  // 3. 対象月のシフトデータを取得（過去月 vs 当月以降の自動切り分け）
  const shiftData = getShiftDataForMonth(targetYear, targetMonth, core, cache);
  if (shiftData.error) return shiftData.error;
  const { rows, C } = shiftData;

  // 4. 稼働統計 (Active) の計算
  const workStats = {
    attr: { fullTime: new Set(), regPt: new Set(), spot: new Set() },
    spotForDashboard: new Set(),
    logic: {
      spot: new Set(),
      spotDetails: {
        days: { Mon: new Set(), Tue: new Set(), Wed: new Set(), Thu: new Set(), Fri: new Set(), Sat: new Set(), Sun: new Set(), Hol: new Set() },
        hours: { Mon: 0, Tue: 0, Wed: 0, Thu: 0, Fri: 0, Sat: 0, Sun: 0, Hol: 0, Total: 0 }
      }
    }
  };

  rows.forEach(row => {
    if (C.DATE === -1) return;
    const dateVal = row[C.DATE] instanceof Date ? row[C.DATE] : new Date(row[C.DATE]);
    if (isNaN(dateVal.getTime()) || dateVal.getFullYear() !== targetYear || (dateVal.getMonth() + 1) !== targetMonth) return;

    const clinicRaw = String(row[C.CLINIC] || "").trim();
    if (/有給|欠勤|医師会|院外|バックアップ|嘱託医/.test(clinicRaw)) return;
    const isNormalClinic = validClinics.some(c => clinicRaw.includes(c) || c.includes(clinicRaw) || (clinicRaw.includes("NT") && c.includes("千葉")));
    if (!isNormalClinic) return;

    const h = calcWorkHours(row, C);
    if (h === null) return;

    const medNo = (C.MED_NO > -1) ? cleanId(row[C.MED_NO]) : "";
    const nameKey = (C.NAME > -1) ? normNameFn(row[C.NAME]) : "";
    const id = medNo || nameKey;
    if (!id) return;

    const note = (C.NOTE > -1) ? String(row[C.NOTE] || "") : "";
    const allComments = C.COMMENT_IDXS.map(idx => String(row[idx] || "")).join(" ");

    const dStr = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy-MM-dd');
    let day = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'EEE');
    if (holidays.has(dStr) && day !== 'Sun') day = 'Hol';
    const isSunOrHol = (day === 'Sun' || day === 'Hol');

    let attrType = 'Spot';
    const doctorInfo = (medNo && doctorInfoMap.get(medNo)) || (nameKey && doctorInfoMap.get(nameKey));
    if (doctorInfo) {
      const okJoin = !doctorInfo.join || dateVal >= doctorInfo.join;
      const okLeave = !doctorInfo.leave || dateVal <= doctorInfo.leave;
      if (okJoin && okLeave) attrType = doctorInfo.type;
    }

    const retiredInfo = (medNo && retiredMap.get(medNo)) || (nameKey && retiredMap.get(nameKey));
    if (retiredInfo && dateVal > retiredInfo.leave) {
      attrType = 'Spot';
    }

    if (attrType === 'FT') workStats.attr.fullTime.add(id);
    else if (attrType === 'Reg') workStats.attr.regPt.add(id);
    else workStats.attr.spot.add(id);

    if (attrType === 'Spot' || (attrType === 'Reg' && !/(毎週|第|曜)/.test(note))) {
      workStats.spotForDashboard.add(id);
    }

    const isRelief = /所定休出|所定外|特別時給|追加|欠勤対応|急遽|ヘルプ|バックアップ|民間医局|紹介|加算/.test(allComments + " " + note);
    const hasRegularNote = /【.*】.*(毎週|第[1-5]|曜|[月火水木金土日]：)/.test(note);

    let isSpotLogic = false;
    if (isRelief) isSpotLogic = true;
    else if (hasRegularNote) isSpotLogic = false;
    else if (attrType === 'FT') isSpotLogic = isSunOrHol;
    else isSpotLogic = true;

    if (isSpotLogic) {
      workStats.logic.spot.add(id);
      if (workStats.logic.spotDetails.days[day]) {
        workStats.logic.spotDetails.days[day].add(id);
        workStats.logic.spotDetails.hours[day] += h;
        workStats.logic.spotDetails.hours.Total += h;
      }
    }
  });

  // 5. 出力シート（YYYY年度）への一括書き込み
  try {
    if (!cache.outFinalSheets) cache.outFinalSheets = {};
    if (!cache.outFinalSheets[OUTPUT_SHEET_NAME]) {
      const sheet = ssConfig.getSheetByName(OUTPUT_SHEET_NAME);
      if (!sheet) return `エラー: シート「${OUTPUT_SHEET_NAME}」がありません`;
      const outMap = new Map();
      const rData = sheet.getRange(2, 1, 76, 3).getValues();
      let lA = "";
      rData.forEach((r, i) => {
        if (r[0]) lA = String(r[0]);
        if (r[2]) outMap.set(lA + '_' + r[2], i);
      });
      cache.outFinalSheets[OUTPUT_SHEET_NAME] = { sheet, outMap };
    }

    const { sheet, outMap } = cache.outFinalSheets[OUTPUT_SHEET_NAME];
    const fVals = Array.from({ length: 76 }, () => [0]);

    const isRatioItem = (itemName) => itemName.includes('率') || itemName.includes('割') || itemName.includes('構成比');
    const setV = (cA, cC, v) => {
      const idx = outMap.get(cA + '_' + cC);
      if (idx !== undefined) fVals[idx] = [typeof v === 'number' && !isRatioItem(cC) ? Math.round(v) : v];
    };
    const safeRatio = (num, den) => den > 0 && isFinite(num / den) ? num / den : 0;

    const actFullTime = workStats.attr.fullTime.size;
    const actRegPt = workStats.attr.regPt.size;
    const actSpotAttr = workStats.spotForDashboard.size;
    const actTotal = new Set([...workStats.attr.fullTime, ...workStats.attr.regPt, ...workStats.spotForDashboard]).size;

    setV('医師合計', '在籍数', stockStats.total); setV('医師合計', '稼働数', actTotal); setV('医師合計', 'アクティブ率', safeRatio(actTotal, stockStats.total));
    setV('常勤医師', '在籍数', stockStats.fullTime.count); setV('常勤医師', '稼働数', actFullTime); setV('常勤医師', '稼働割合', safeRatio(actFullTime, stockStats.fullTime.count));
    setV('常勤医師', '内科', stockStats.fullTime.internal); setV('常勤医師', '小児科', stockStats.fullTime.peds);
    setV('定期非常勤医師', '在籍数', stockStats.regPt.count); setV('定期非常勤医師', '稼働数', actRegPt); setV('定期非常勤医師', '稼働割合', safeRatio(actRegPt, stockStats.regPt.count));
    setV('定期非常勤医師', '内科', stockStats.regPt.internal); setV('定期非常勤医師', '小児科', stockStats.regPt.peds);

    setV('スポット医師', '在籍数', stockStats.spot.count); setV('スポット医師', '稼働数', actSpotAttr); setV('スポット医師', '稼働割合', safeRatio(actSpotAttr, stockStats.spot.count));
    setV('スポット医師', '内科', stockStats.spot.internal); setV('スポット医師', '小児科', stockStats.spot.peds); setV('スポット医師', 'ワクチン専任', stockStats.spot.vaccine);

    const sd = workStats.logic.spotDetails;
    setV('稼働スポット医師数', '月間合計', workStats.logic.spot.size);
    ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun', 'Hol'].forEach((k, i) => setV('稼働スポット医師数', ['月', '火', '水', '木', '金', '土', '日', '祝'][i], sd.days[k].size));

    setV('稼働スポット医師稼働時間数', '月間合計', sd.hours.Total);
    ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun', 'Hol'].forEach((k, i) => setV('稼働スポット医師稼働時間数', ['月', '火', '水', '木', '金', '土', '日', '祝'][i], sd.hours[k]));

    setV('DS登録増加数', '登録合計', stockStats.newReg.total);
    setV('DS登録増加数', '内科', stockStats.newReg.internal); setV('DS登録増加数', '小児科', stockStats.newReg.peds); setV('DS登録増加数', 'ワクチン専任', stockStats.newReg.vaccine);
    setV('DS登録増加数', '常勤合計', stockStats.newReg.ft); setV('DS登録増加数', '定期非常勤合計', stockStats.newReg.rpt); setV('DS登録増加数', 'スポット合計', stockStats.newReg.sp);
    setV('DS登録経路', '合計', stockStats.newReg.total);
    setV('DS登録経路', '直接', stockStats.routes.direct); setV('DS登録経路', '紹介会社', stockStats.routes.agent);
    setV('DS登録経路', 'その他', stockStats.routes.other); setV('DS登録経路', 'リファラル', stockStats.routes.referral);

    const acL = ['1ヶ月以内', '3ヶ月以内', '6ヶ月以内', '1年以内', '1年以上経過'];
    acL.forEach((l, i) => {
      setV('ドクターサイトアクセス数', l, stockStats.access.all[i]);
      setV('ドクターサイトアクセス数（常勤）', l, stockStats.access.ft[i]);
      setV('ドクターサイトアクセス数（定期非常勤）', l, stockStats.access.rpt[i]);
      setV('ドクターサイトアクセス数（スポット）', l, stockStats.access.sp[i]);
    });

    const tR = sheet.getRange(2, targetColIndex, 76, 1);
    tR.setValues(fVals);
    tR.setHorizontalAlignment("right");
    return "完了";
  } catch (e) {
    return `書込エラー: ${e.message}`;
  }
}