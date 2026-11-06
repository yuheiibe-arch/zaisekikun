/**
 * 離脱くんメイン分析ロジック
 */
function runRidatsuAnalysis(targetYear) {
  const CONFIG_URL = 'https://docs.google.com/spreadsheets/d/1CKiJYyZ0iF6cobH_ods2hrI9RmFKt6gFiRXUjD195o0/edit';
  const urlChecklist = 'https://docs.google.com/spreadsheets/d/1he6rfVppfeeQF3LEfu8719nFb952L_zwMNtMFzCrShM/edit';
  const urlMaster = 'https://docs.google.com/spreadsheets/d/1n8o8gzSpfyvse_I3Vsrrx1EpEfnIbNdy4DyHXnWD8TM/edit';
  
  const getUrlSafe = (dUrl, name) => {
    try {
      const ss = SpreadsheetApp.openByUrl(dUrl);
      const sheet = ss.getSheetByName('データセット');
      const data = sheet.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim() === name) {
          return String(data[i][1]).trim();
        }
      }
    } catch(e) {}
    return "";
  };

  const urlCost = getUrlSafe(CONFIG_URL, '人件費くん');
  const urlShift = getUrlSafe(CONFIG_URL, targetYear + '確定シフト');
  const urlRegEx = getUrlSafe(CONFIG_URL, '正規表現');

  if (!urlShift) throw new Error(targetYear + '年度の確定シフトURLが見つかりません');

  const today = new Date();
  const twoMonthsAgo = new Date(today);
  twoMonthsAgo.setMonth(twoMonthsAgo.getMonth() - 2);
  
  const renameClinic = (name) => {
    const n = String(name || "").trim();
    if (n === "亀有") return "亀有（小児科）";
    if (n === "北葛西") return "北葛西（内科）";
    return n;
  };

  const clinicToAreaMap = new Map();
  const validClinicSet = new Set(); 
  try {
    const dataReg = SpreadsheetApp.openByUrl(urlRegEx).getSheets()[0].getDataRange().getValues();
    for (let i = 1; i < dataReg.length; i++) {
      let clinicName = String(dataReg[i][0]).trim(); 
      const area = String(dataReg[i][6]).trim();
      if (clinicName) {
        validClinicSet.add(clinicName);
        clinicToAreaMap.set(renameClinic(clinicName), area);
      }
    }
  } catch (e) {}

  const getArea = (name) => clinicToAreaMap.get(renameClinic(name)) || "不明";

  const fullTimeJoinMap = new Map(); 
  const teikiMeds = new Set();
  try {
    const ssCost = SpreadsheetApp.openByUrl(urlCost);
    const sFt = ssCost.getSheets().find(s => s.getName().includes('常勤マスタ'));
    if (sFt) {
      sFt.getRange(2, 2, sFt.getLastRow() - 1, 5).getValues().forEach(r => {
        if (r[0]) fullTimeJoinMap.set(Number(r[0]), (r[4] instanceof Date) ? Utilities.formatDate(r[4], "Asia/Tokyo", "yyyy/MM/dd") : "不明");
      });
    }
    const sTeiki = ssCost.getSheets().find(s => s.getName().includes('定期非常勤マスタ'));
    if (sTeiki) {
      sTeiki.getRange(2, 2, sTeiki.getLastRow() - 1, 1).getValues().forEach(r => { if (r[0]) teikiMeds.add(Number(r[0])); });
    }
  } catch (e) {}

  const masterMap = new Map();
  const dataMaster = SpreadsheetApp.openByUrl(urlMaster).getSheets()[0].getDataRange().getValues();
  dataMaster.shift(); 
  dataMaster.forEach(r => {
    const medNo = Number(r[10]);
    if (!isNaN(medNo)) masterMap.set(medNo, {
      address: String(r[13] || "") + String(r[14] || ""),
      mainClinic: renameClinic(r[17])
    });
  });

  const medNoMap = new Map();
  let fullTimeDocs = []; 
  const dataCheck = SpreadsheetApp.openByUrl(urlChecklist).getSheets()[0].getDataRange().getValues();

  for (let i = 2; i < dataCheck.length; i++) {
    const r = dataCheck[i];
    const medNo = Number(r[0]);
    const firstDate = r[20];
    if (isNaN(medNo) || !(firstDate instanceof Date) || medNoMap.has(medNo)) continue;

    const fiscal = (firstDate.getMonth() + 1 >= 4) ? firstDate.getFullYear() : firstDate.getFullYear() - 1;
    
    if (fiscal == targetYear && firstDate <= today) {
      if (fullTimeJoinMap.has(medNo) && !teikiMeds.has(medNo)) {
        fullTimeDocs.push({ name: String(r[1]), joinDate: fullTimeJoinMap.get(medNo) });
        continue;
      }
      if (teikiMeds.has(medNo)) continue;

      medNoMap.set(medNo, {
        month: Utilities.formatDate(firstDate, "Asia/Tokyo", "yyyy/MM"),
        name: String(r[1]), medNo: medNo, firstDateStr: Utilities.formatDate(firstDate, "Asia/Tokyo", "yyyy/MM"), 
        dept: String(r[9]).trim(), route: String(r[4]).trim(), history: []
      });
    }
  }

  // シフトスキャン（過去分と未来分を分ける）
  const scanShifts = (ssUrl, isPast) => {
    if (!ssUrl) return;
    try {
      const ss = SpreadsheetApp.openByUrl(ssUrl);
      
      let sheets = [];
      if (isPast) {
        // ★ 過去分はそのまま (YYYY/MM のシートを探す)
        sheets = ss.getSheets().filter(s => s.getName().match(/^\d{4}\/\d{2}$/));
      } else {
        // ★ 未来分は「確定シフト」シートを探す
        const futureSheet = ss.getSheetByName("確定シフト");
        if (futureSheet) sheets.push(futureSheet);
      }

      sheets.forEach(s => {
        if (!s) return;
        const rows = s.getDataRange().getValues();
        const header = rows[0];
        const colMed = header.indexOf('医籍番号');
        const colDate = header.findIndex(h => String(h).includes('勤務日') || String(h).includes('日付'));
        let colCli = header.indexOf('クリニック名');
        if (colCli === -1) colCli = header.findIndex(h => String(h).includes('クリニック'));

        for (let i = 1; i < rows.length; i++) {
          const mNo = Number(rows[i][colMed]);
          const cli = renameClinic(rows[i][colCli]);
          const d = rows[i][colDate];
          if (medNoMap.has(mNo) && d instanceof Date && validClinicSet.has(String(rows[i][colCli] || "").trim())) {
            medNoMap.get(mNo).history.push({ date: d, clinic: cli });
          }
        }
      });
    } catch(e) {}
  };

  // 過去分（YYYY/MMシートから）
  scanShifts(urlShift, true);
  // 未来分（確定シフトシートから）
  scanShifts(urlShift, false);

  let rawData = [];
  medNoMap.forEach(doc => {
    const hist = doc.history.sort((a,b) => a.date - b.date);
    if (hist.length === 0) return;

    const uniqueAllDates = [...new Set(hist.map(h => Utilities.formatDate(h.date, "Asia/Tokyo", "yyyy/MM/dd")))].sort();
    const totalCount = uniqueAllDates.length;
    
    const pastShifts = hist.filter(h => h.date <= today);
    const uniquePastDates = [...new Set(pastShifts.map(h => Utilities.formatDate(h.date, "Asia/Tokyo", "yyyy/MM/dd")))].sort();
    const pastCount = uniquePastDates.length;
    
    const debutClinic = hist[0].clinic;
    const lastPastS = pastShifts[pastShifts.length - 1] || { date: doc.firstDateStr, clinic: "実績なし" };
    
    const debutDate = hist[0].date;
    const limit = new Date(debutDate);
    limit.setMonth(limit.getMonth() + 3);
    const shifts3mCount = uniqueAllDates.filter(d => new Date(d) <= limit).length;
    const futureShifts = hist.filter(h => h.date > today);

    let status = "";
    
    if (shifts3mCount >= 3 || futureShifts.length > 0) status = "継続";
    else if (pastCount === 1) status = "初回離脱";
    else status = "離脱";

    if (lastPastS.date instanceof Date && lastPastS.date < twoMonthsAgo) {
      if (status === "継続") {
        status = "離脱"; 
      }
    }

    const cCounts = hist.reduce((acc, c) => { 
      const n = renameClinic(c.clinic);
      acc[n] = (acc[n] || 0) + 1; 
      return acc; 
    }, {});
    
    const top3 = Object.entries(cCounts)
      .filter(entry => entry[1] >= 5) 
      .sort((a,b) => b[1] - a[1])     
      .slice(0, 3)                    
      .map(e => e[0])                 
      .join("、");
    
    const m = masterMap.get(doc.medNo) || {};
    rawData.push({
      name: doc.name, count: totalCount, adopt: doc.firstDateStr, 
      lastD: (lastPastS.date instanceof Date) ? Utilities.formatDate(lastPastS.date, "JST", "yyyy/MM/dd") : lastPastS.date, 
      lastC: lastPastS.clinic, debutC: debutClinic, 
      route: doc.route, emp: m.mainClinic || "不明", addr: m.address || "不明", 
      dept: doc.dept,
      area: getArea(lastPastS.clinic), top3: top3, status: status, rawMonth: doc.month
    });
  });

  return { fullTimeDocs: fullTimeDocs, rawData: rawData };
}

function getRidatsuReport(targetYear, fullTimeDocs, rawData) {
  if (!rawData) return null;

  const cleanAddress = (addr) => {
    if (!addr) return "";
    let temp = addr.replace(/[0-9０-９].*$/, "").trim();
    const match = temp.match(/^(.*?[都道府県])?.*?[市区町村]/);
    return match ? match[0] : temp;
  };

  const getDeptMark = (deptStr) => {
    if (!deptStr) return "";
    if (deptStr.includes("小児科")) return "(小)";
    if (deptStr.includes("内科")) return "(内)";
    return "";
  };

  let report = {
    targetYear: targetYear, 
    period: `令和${targetYear - 2018}年度（2025/04/01 〜 未来分含）`,
    definition: "定着率：3ヶ月以内3回稼働、または未来シフトあり。ただし最終勤務から2ヶ月以上空きは離脱扱い。",
    counts: { fullTimeDocs: fullTimeDocs || [] },
    firstExit: { total: 0, byRoute: [] }, 
    retention: { kanto: [], kansai: [], kStats: { active: 0, total: 0 }, sStats: { active: 0, total: 0 } },
    aces: [], settledAll: [], exitList: []
  };

  const today = new Date();
  const currentMonthStr = Utilities.formatDate(today, "Asia/Tokyo", "yyyy/MM");
  const lastMonthDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const lastMonthStr = Utilities.formatDate(lastMonthDate, "Asia/Tokyo", "yyyy/MM");

  let routeMap = {}, retKMap = {}, retSMap = {};

  rawData.forEach(d => {
    if (d.adopt === currentMonthStr || d.adopt === lastMonthStr) {
      return; 
    }

    if (d.area === "関東") report.retention.kStats.total++; 
    else if (d.area === "関西") report.retention.sStats.total++;
    
    const addrSimple = cleanAddress(d.addr);

    if (d.status === "初回離脱") {
      if (!routeMap[d.route]) routeMap[d.route] = { count: 0, names: [] };
      routeMap[d.route].count++;
      routeMap[d.route].names.push(d.name + getDeptMark(d.dept));
    }
    
    const obj = { 
      name: d.name, count: d.count, month: d.adopt, debut: d.debutC, 
      lastD: d.lastD, 
      route: d.route, address: addrSimple, employer: d.emp, main: d.top3 
    };

    if (d.status === "継続") {
      if (d.area === "関東") report.retention.kStats.active++; 
      else if (d.area === "関西") report.retention.sStats.active++;
      
      let target = (d.area === "関東") ? retKMap : retSMap;
      if (!target[d.debutC]) target[d.debutC] = { count: 0, docs: [] };
      target[d.debutC].count++;
      target[d.debutC].docs.push(`${d.name}(${d.count})`);
      
      report.settledAll.push(obj);
    } else {
      report.exitList.push({ 
        name: d.name, count: d.count, adopt: d.adopt, lastD: d.lastD, 
        lastC: d.lastC, route: d.route, emp: d.emp, addr: addrSimple
      });
    }
  });

  report.aces = [...report.settledAll].sort((a,b) => b.count - a.count).slice(0, 30);
  report.settledAll.sort((a,b) => b.count - a.count);

  report.exitList.sort((a,b) => a.adopt.localeCompare(b.adopt));

  report.firstExit.byRoute = Object.entries(routeMap)
    .sort((a,b) => b[1].count - a[1].count)
    .slice(0, 5)
    .map(e => ({ n: e[0], c: e[1].count, docs: e[1].names.join("、") }));

  const sortM = (m) => Object.entries(m)
    .sort((a,b) => b[1].count - a[1].count)
    .slice(0, 5)
    .map(e => ({ n: e[0], c: e[1].count, d: e[1].docs.join("、") }));
  
  report.retention.kanto = sortM(retKMap);
  report.retention.kansai = sortM(retSMap);

  return report;
}