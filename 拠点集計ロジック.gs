/**
 * ====================================================================
 * 🏆 【拠点集計】updateLocationDashboard (モジュール分割・高速版)
 * ====================================================================
 * 共通基盤 (共通関数.gs) からマスタと対象月シフトを受け取り、拠点別集計を行います。
 */
function updateLocationDashboard(targetYear, targetMonth, cache = {}) {
  targetYear = Number(targetYear);
  targetMonth = Number(targetMonth);

  const fiscalYear = (targetMonth >= 4) ? targetYear : targetYear - 1;
  const SHEET_NAME_OUT = `${fiscalYear}拠点`;

  // 1. 共通マスタと紹介会社情報の取得（キャッシュ利用）
  let core;
  try {
    core = getSharedCoreContext(cache, { needAgency: true });
  } catch (e) {
    return `マスタ読込エラー: ${e.message}`;
  }

  const {
    ssConfig, normNameFn, normClinicFn,
    openDateMap, fullTimeMap, regPtMap, retiredMap,
    agencyDoctorMedIds, agencyMap
  } = core;

  // 2. 出力先シート（YYYY拠点）のレイアウト取得（年度ごとに1回だけキャッシュ）
  if (!cache.outLocationSheets) cache.outLocationSheets = {};
  if (!cache.outLocationSheets[SHEET_NAME_OUT]) {
    const sheetOut = ssConfig.getSheetByName(SHEET_NAME_OUT);
    if (!sheetOut) return `エラー: シート「${SHEET_NAME_OUT}」が見つかりません。`;
    const lastRow = sheetOut.getLastRow();
    const rangeData = sheetOut.getRange(2, 1, lastRow - 1, 3).getValues();
    const rowMap = new Map();
    const targetClinics = new Set();
    let currentClinicName = "";

    rangeData.forEach((r, idx) => {
      if (r[0]) {
        currentClinicName = normClinicFn(String(r[0]));
        targetClinics.add(currentClinicName);
      }
      if (currentClinicName && r[2]) {
        rowMap.set(`${currentClinicName}_${normalizeItemKey(r[2])}`, idx + 2);
      }
    });
    cache.outLocationSheets[SHEET_NAME_OUT] = { sheetOut, rowMap, targetClinics, numRows: lastRow - 1 };
  }

  const { sheetOut, rowMap, targetClinics, numRows } = cache.outLocationSheets[SHEET_NAME_OUT];
  const writeCol = (targetMonth >= 4) ? targetMonth : targetMonth + 12;

  // 3. 対象月のシフトデータを取得（過去月 vs 当月以降の自動切り分け）
  const shiftData = getShiftDataForMonth(targetYear, targetMonth, core, cache);
  if (shiftData.error) return shiftData.error;
  const { rows, C } = shiftData;

  // 4. 拠点別集計コンテナの初期化
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

  // 5. シフト行の走査と分類
  rows.forEach(row => {
    if (C.DATE === -1) return;
    const dVal = row[C.DATE] instanceof Date ? row[C.DATE] : new Date(row[C.DATE]);
    if (isNaN(dVal.getTime()) || dVal.getFullYear() !== targetYear || (dVal.getMonth() + 1) !== targetMonth) return;

    const rawClinic = String(row[C.CLINIC] || "").trim();
    if (/有給|欠勤|医師会|院外|バックアップ|嘱託医/.test(rawClinic)) return;

    const clinic = normClinicFn(rawClinic);
    if (!clinicStats[clinic]) return;

    // 亀有・北葛西の小児科フィルタ
    if (['亀有', '北葛西'].includes(clinic)) {
      const deptStr = String(row[C.DEPT] || "");
      const isNaika = rawClinic.includes('内科') || deptStr.includes('内科');
      if (isNaika && !deptStr.includes('小児科')) return;
    }

    const h = calcWorkHours(row, C);
    if (h === null) return;

    const medNo = (C.MED_NO > -1) ? cleanId(row[C.MED_NO]) : "";
    const nameKey = (C.NAME > -1) ? normNameFn(row[C.NAME]) : "";
    const id = medNo || nameKey;
    if (!id) return;

    const note = (C.NOTE > -1) ? String(row[C.NOTE] || "") : "";
    const comments = C.COMMENT_IDXS.map(c => String(row[c] || "")).join(" ");
    const combinedText = comments + " " + note;

    // 医師区分の判定
    let isFt = false;
    let isReg = false;

    const ftInfo = (medNo && fullTimeMap.get(medNo)) || (nameKey && fullTimeMap.get(nameKey));
    if (ftInfo && (!ftInfo.join || dVal >= ftInfo.join) && (!ftInfo.leave || dVal <= ftInfo.leave)) {
      isFt = true;
    }

    const ptInfo = (medNo && regPtMap.get(medNo)) || (nameKey && regPtMap.get(nameKey));
    if (!isFt && ptInfo && (!ptInfo.join || dVal >= ptInfo.join) && (!ptInfo.leave || dVal <= ptInfo.leave)) {
      isReg = true;
    }

    const retiredInfo = (medNo && retiredMap.get(medNo)) || (nameKey && retiredMap.get(nameKey));
    if (retiredInfo) {
      if (dVal > retiredInfo.leave || (retiredInfo.join && dVal < retiredInfo.join)) {
        isFt = false; isReg = false;
      } else {
        if (retiredInfo.type.includes('常勤')) isFt = true;
        else if (retiredInfo.type.includes('定期')) isReg = true;
      }
    }

    const isKyushutsu = /所定休出|所定外/.test(combinedText);
    const isRelief = isKyushutsu || /特別時給|追加|欠勤対応|急遽|ヘルプ|バックアップ|民間医局|紹介|加算/.test(combinedText);
    const isRegularNote = /毎週(月|火|水|木|金|土|日)曜|第[1-5](月|火|水|木|金|土|日)曜/.test(note);

    const dateKey = `${nameKey}_${formatDateKey(dVal)}`;
    const isAgency = (medNo && agencyDoctorMedIds.has(medNo)) || agencyMap.has(dateKey) || /紹介|民間|MRT|エムスリー|マイナビ|メディウェル|リクルート|オンコール/.test(comments);

    if (isKyushutsu) {
      clinicStats[clinic].ftOver.h += h;
      clinicStats[clinic].spot.c.add(id);
      clinicStats[clinic].spot.h += h;
      if (isAgency && !isFt && !isReg) {
        clinicStats[clinic].spotAg.c.add(id); clinicStats[clinic].spotAg.h += h;
      } else {
        clinicStats[clinic].spotDir.c.add(id); clinicStats[clinic].spotDir.h += h;
      }
    } else if (isFt) {
      clinicStats[clinic].ft.c.add(id);
      clinicStats[clinic].ft.h += h;
    } else if (isRelief) {
      clinicStats[clinic].spot.c.add(id);
      clinicStats[clinic].spot.h += h;
      if (isAgency) {
        clinicStats[clinic].spotAg.c.add(id); clinicStats[clinic].spotAg.h += h;
      } else {
        clinicStats[clinic].spotDir.c.add(id); clinicStats[clinic].spotDir.h += h;
      }
    } else if (isReg || isRegularNote) {
      clinicStats[clinic].reg.c.add(id);
      clinicStats[clinic].reg.h += h;
    } else {
      clinicStats[clinic].spot.c.add(id);
      clinicStats[clinic].spot.h += h;
      if (isAgency) {
        clinicStats[clinic].spotAg.c.add(id); clinicStats[clinic].spotAg.h += h;
      } else {
        clinicStats[clinic].spotDir.c.add(id); clinicStats[clinic].spotDir.h += h;
      }
    }
  });

  // 6. 対象列への一括書き込み（setValues 1回で完了）
  const colRange = sheetOut.getRange(2, writeCol, numRows, 1);
  const colValues = colRange.getValues();
  let updatedCount = 0;

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
        colValues[row - 2][0] = val;
        updatedCount++;
      }
    });
  });

  if (updatedCount > 0) {
    colRange.setValues(colValues);
    return "完了";
  }
  return "更新なし";
}