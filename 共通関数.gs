/**
 * ====================================================================
 * 📦 【共通基盤】マスタ読込・シフト切り分け取得・ユーティリティ関数群
 * ====================================================================
 * 拠点集計・全体集計の両方から呼び出されるデータ取得＆キャッシュエンジンです。
 */

const ZAIS_CONFIG = {
  CONFIG_URL: 'https://docs.google.com/spreadsheets/d/1CKiJYyZ0iF6cobH_ods2hrI9RmFKt6gFiRXUjD195o0/edit',
  NEW_MASTER_URL: 'https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit',
  AGENCY_HISTORY_URL: 'https://docs.google.com/spreadsheets/d/1Fd8uOCE1SKvLCIPjZZ7sE2rFsQFOhjVjHqoaqs-pXqE/edit'
};

// --------------------------------------------------------------------
// 汎用ヘルパー関数（他ファイルとの後方互換性も維持）
// --------------------------------------------------------------------

function getUrlFromDataset(configUrl, keyName) {
  try {
    const ss = SpreadsheetApp.openByUrl(configUrl || ZAIS_CONFIG.CONFIG_URL);
    const sheet = ss.getSheetByName('データセット');
    const data = sheet.getDataRange().getValues();
    const row = data.find(r => String(r[0]).trim() === keyName);
    return row ? String(row[1]).trim() : "";
  } catch (e) {
    Logger.log(`URL取得エラー(${keyName}): ${e.message}`);
    return "";
  }
}

function cleanId(rawId) {
  return String(rawId || "").replace(/\.0+$/, '').trim();
}

function normalizeName(name, aliasMap = null) {
  if (!name) return "";
  const rawTrimmed = String(name).trim();
  if (aliasMap) {
    if (aliasMap.has(rawTrimmed)) return String(aliasMap.get(rawTrimmed)).replace(/[\s ]+/g, "");
    const halfSpace = rawTrimmed.replace(/ /g, ' ');
    if (aliasMap.has(halfSpace)) return String(aliasMap.get(halfSpace)).replace(/[\s ]+/g, "");
    const noSpace = halfSpace.replace(/\s+/g, '');
    if (aliasMap.has(noSpace)) return String(aliasMap.get(noSpace)).replace(/[\s ]+/g, "");
    return noSpace;
  }
  return rawTrimmed.replace(/[\s ]+/g, "");
}

function normalizeClinic(name, clinicDict = null) {
  if (!name) return "";
  let s = String(name).trim();
  if (clinicDict && clinicDict[s]) s = clinicDict[s];

  let clean = s.replace(/^[0-9.\s]+/, '').replace(/[【】\(（]?(内科|小児科)[\)）]?/g, "").split('(')[0].split('（')[0].trim();
  if (clinicDict && clinicDict[clean]) s = clinicDict[clean];
  else s = clean;

  s = s.replace(/ヶ/g, "ケ");
  if (s.includes('千葉') && s.match(/ニュータウン|NT|ＮＴ/i)) return '千葉ニュータウン中央';
  if (s === '八千代') return '八千代緑が丘';
  if (s === '市川妙典') return '妙典';
  if (s === 'まるひろ南浦和店') return '南浦和';
  if (s === '千里丘') return '千里';
  return s;
}

function normalizeItemKey(str) {
  return String(str || "").replace(/[（）\(\)\s \[\]【】]/g, "").trim();
}

function formatDateKey(dateObj) {
  if (!dateObj || !(dateObj instanceof Date) || isNaN(dateObj.getTime())) return null;
  return Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function findColumnIndex(headers, keyword) {
  return headers.findIndex(h => String(h).includes(keyword));
}

function createHeaderMap(headerRow) {
  const map = {};
  headerRow.forEach((val, idx) => {
    if (val) map[String(val).replace(/\n/g, '').trim()] = idx;
  });
  return map;
}

/**
 * 勤務開始・終了時刻から実働時間（13-15時休憩控除後）を計算
 */
function calcWorkHours(row, C) {
  const toH = (v) => {
    if (v instanceof Date) return v.getHours() + v.getMinutes() / 60;
    if (typeof v === 'number') return v * 24;
    if (typeof v === 'string' && v.includes(':')) {
      const p = v.trim().split(':').map(Number);
      if (!isNaN(p[0]) && !isNaN(p[1])) return p[0] + p[1] / 60;
    }
    return null;
  };
  let s = toH(row[C.START_M]), e = toH(row[C.END_M]);
  if (s == null || e == null) { s = toH(row[C.START_B]); e = toH(row[C.END_B]); }
  if (s === null || e === null) return null;
  if (e < s && e !== 0) e += 24;
  if ((e - s) < 0.5) return null;

  let h = e - s;
  const bo = Math.max(0, Math.min(e, 15) - Math.max(s, 13));
  if (h > bo && s < 14.9) h -= bo;
  return h > 0 ? h : null;
}

// --------------------------------------------------------------------
// 🚀 共通マスタ一括ローダー（キャッシュ機能付き）
// --------------------------------------------------------------------

/**
 * 全体集計・拠点集計で使うマスタ一式を読み込んで返す（2回目以降はキャッシュから即返却）
 * @param {Object} cache - 共有キャッシュオブジェクト
 * @param {Object} options - { needDoctorMaster: boolean, needAgency: boolean }
 */
function getSharedCoreContext(cache = {}, options = {}) {
  if (!cache.core) {
    const ssConfig = SpreadsheetApp.openByUrl(ZAIS_CONFIG.CONFIG_URL);
    const datasetValues = ssConfig.getSheetByName('データセット').getDataRange().getValues();
    const getUrl = (key) => {
      const row = datasetValues.find(r => String(r[0]).trim() === key);
      return row ? String(row[1]).trim() : "";
    };

    const clinicDict = {};
    const validClinics = [];
    const nameAliasMap = new Map();
    const openDateMap = new Map();
    const holidays = new Set();
    const doctorInfoMap = new Map(); // id or name -> { type: 'FT'|'Reg', join, leave }
    const fullTimeMap = new Map();   // id or name -> { join, leave }
    const regPtMap = new Map();      // id or name -> { join, leave }
    const retiredMap = new Map();

    // 1. 正規表現（拠点名・氏名・開業日）
    const urlRegex = getUrl('正規表現');
    if (urlRegex) {
      try {
        const ssReg = SpreadsheetApp.openByUrl(urlRegex);
        const regData = ssReg.getSheets()[0].getDataRange().getValues();
        const rHead = regData[0].map(String);
        const cOfficial = rHead.findIndex(h => h.includes('正規記載'));
        for (let r = 1; r < regData.length; r++) {
          regData[r].forEach(cell => {
            const val = String(cell).trim();
            if (val) validClinics.push(val);
          });
          if (cOfficial !== -1) {
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
        const sheetLoc = ssReg.getSheetByName('拠点名');
        if (sheetLoc) {
          const locData = sheetLoc.getDataRange().getValues();
          for (let i = 1; i < locData.length; i++) {
            const rawOfficial = String(locData[i][0] || "").trim();
            if (!rawOfficial) continue;
            clinicDict[rawOfficial] = rawOfficial;
            validClinics.push(rawOfficial);
            for (let j = 1; j <= 4; j++) {
              const alias = String(locData[i][j] || "").trim();
              if (alias) {
                clinicDict[alias] = rawOfficial;
                validClinics.push(alias);
              }
            }
            const normOfficial = normalizeClinic(rawOfficial, clinicDict);
            const date = locData[i][7];
            if (normOfficial && date instanceof Date) openDateMap.set(normOfficial, date);
          }
        }
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
      } catch (e) {
        Logger.log(`正規表現読込エラー: ${e.message}`);
      }
    }

    const normNameFn = (n) => normalizeName(n, nameAliasMap);
    const normClinicFn = (c) => normalizeClinic(c, clinicDict);

    // 2. 祝日マスタ
    const urlHoliday = getUrl('祝日');
    if (urlHoliday) {
      try {
        const ssHol = SpreadsheetApp.openByUrl(urlHoliday);
        const sHol = ssHol.getSheetByName('祝日') || ssHol.getSheets()[0];
        sHol.getDataRange().getValues().forEach(r => {
          if (r[0] instanceof Date) holidays.add(Utilities.formatDate(r[0], 'Asia/Tokyo', 'yyyy-MM-dd'));
        });
      } catch (e) {}
    }

    // 3. 新医師マスタ（常勤・定期非常勤）のピンポイント読込
    try {
      const ssMaster = SpreadsheetApp.openByUrl(ZAIS_CONFIG.NEW_MASTER_URL);
      const targetSheets = [
        { name: '【保存】常勤2025年度', type: 'FT', targetMap: fullTimeMap },
        { name: '【保存】定期非常勤2025年度', type: 'Reg', targetMap: regPtMap },
        { name: '常勤2026年度', type: 'FT', targetMap: fullTimeMap },
        { name: '定期非常勤2026年度', type: 'Reg', targetMap: regPtMap }
      ];
      targetSheets.forEach(cfg => {
        const s = ssMaster.getSheetByName(cfg.name);
        if (!s || s.getLastRow() < 2) return;
        const data = s.getDataRange().getValues();
        const hMap = createHeaderMap(data[0]);
        const cName = hMap['医師名'], cMed = hMap['医籍番号'];
        const cJoin = hMap['入職日'] ?? -1, cRet = hMap['退職日'] ?? -1;
        if (cName === undefined || cMed === undefined) return;

        for (let i = 1; i < data.length; i++) {
          const nameKey = normNameFn(data[i][cName]);
          const medNo = cleanId(data[i][cMed]);
          if (!nameKey && !medNo) continue;
          const info = {
            type: cfg.type,
            join: (cJoin !== -1 && data[i][cJoin] instanceof Date) ? data[i][cJoin] : null,
            leave: (cRet !== -1 && data[i][cRet] instanceof Date) ? data[i][cRet] : null
          };
          if (medNo) {
            cfg.targetMap.set(medNo, info);
            if (cfg.type === 'FT' || !doctorInfoMap.has(medNo)) doctorInfoMap.set(medNo, info);
          }
          if (nameKey) {
            cfg.targetMap.set(nameKey, info);
            if (cfg.type === 'FT' || !doctorInfoMap.has(nameKey)) doctorInfoMap.set(nameKey, info);
          }
        }
      });
    } catch (e) {
      Logger.log(`新医師マスタ読込エラー: ${e.message}`);
    }

    // 4. 退職医師リスト
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
            const name = rcName !== -1 ? normNameFn(rr[rcName]) : "";
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

    cache.core = {
      ssConfig, getUrl, normNameFn, normClinicFn,
      validClinics, openDateMap, holidays,
      doctorInfoMap, fullTimeMap, regPtMap, retiredMap
    };
  }

  // オプション1: 全体集計用の「医師マスタ」全行ロード
  if (options.needDoctorMaster && !cache.core.doctorMasterRows) {
    const urlMaster = cache.core.getUrl('医師マスタ');
    const ss = SpreadsheetApp.openByUrl(urlMaster);
    const sheet = ss.getSheets()[0];
    cache.core.doctorMasterRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  }

  // オプション2: 拠点集計用の「紹介会社履歴・進捗」ロード
  if (options.needAgency && !cache.core.agencyLoaded) {
    const agencyDoctorMedIds = new Set();
    const agencyMap = new Map();
    try {
      const agencySs = SpreadsheetApp.openByUrl(ZAIS_CONFIG.AGENCY_HISTORY_URL);
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

    const urlAgency = cache.core.getUrl('紹介会社進捗');
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
            if (dVal instanceof Date && nVal) agencyMap.set(`${cache.core.normNameFn(nVal)}_${formatDateKey(dVal)}`, sheetName);
          });
        });
      } catch (e) {}
    }
    cache.core.agencyDoctorMedIds = agencyDoctorMedIds;
    cache.core.agencyMap = agencyMap;
    cache.core.agencyLoaded = true;
  }

  return cache.core;
}

// --------------------------------------------------------------------
// 🎯 シフトデータの切り分け取得（過去月 vs 当月・翌月以降）
// --------------------------------------------------------------------

/**
 * 対象年月に応じて、過去月は「YYYY確定シフト(YYYY/MM)」、当月以降は「確定シフト(確定シフト)」から
 * シフトデータと解析済み列インデックスを取得して返す（キャッシュ対応）
 */
function getShiftDataForMonth(targetYear, targetMonth, core, cache = {}) {
  const today = new Date();
  const currentY = today.getFullYear();
  const currentM = today.getMonth() + 1;

  // ★ 切り分けルール: 当月・翌月以降は「確定シフト」、過去月は「YYYY確定シフト」
  const isCurrentOrFuture = (targetYear > currentY) || (targetYear === currentY && targetMonth >= currentM);

  let datasetKey = "";
  let targetSheetName = "";

  if (isCurrentOrFuture) {
    datasetKey = "確定シフト";
    targetSheetName = "確定シフト";
  } else {
    const fiscalYear = (targetMonth >= 4) ? targetYear : targetYear - 1;
    datasetKey = `${fiscalYear}確定シフト`;
    targetSheetName = `${targetYear}/${String(targetMonth).padStart(2, '0')}`;
  }

  const urlShift = core.getUrl(datasetKey);
  if (!urlShift) {
    return { error: `エラー: データセットに「${datasetKey}」のURLが設定されていません` };
  }

  if (!cache.shiftFiles) cache.shiftFiles = {};
  if (!cache.shiftSheets) cache.shiftSheets = {};

  const cacheKey = `${urlShift}__${targetSheetName}`;
  if (!cache.shiftSheets[cacheKey]) {
    try {
      const ssShift = cache.shiftFiles[urlShift] || SpreadsheetApp.openByUrl(urlShift);
      cache.shiftFiles[urlShift] = ssShift;

      const shiftSheet = ssShift.getSheetByName(targetSheetName);
      if (!shiftSheet || shiftSheet.getLastRow() < 2) {
        return { error: `データなし: 「${datasetKey}」内に「${targetSheetName}」シートが見つからないか空です` };
      }

      const rawValues = shiftSheet.getDataRange().getValues();
      const headers = rawValues[0].map(String);
      const rows = rawValues.slice(1);

      const findCol = (candidates) => {
        for (const key of candidates) {
          const idx = headers.findIndex(h => String(h).replace(/[（）\(\)\s ]/g, "").includes(key));
          if (idx > -1) return idx;
        }
        return -1;
      };

      const C = {
        NAME: findCol(['氏名', '名前']),
        MED_NO: headers.indexOf('医籍番号'),
        DOC_NO: findCol(['医師No']),
        CLINIC: findCol(['クリニック名', '勤務地', '拠点']),
        DATE: findCol(['勤務日', '日付']),
        START_M: findCol(['勤務開始時間', '開始時間']),
        END_M: findCol(['勤務終了時間', '終了時間']),
        START_B: findCol(['元のシフト勤務開始時間', '元のシフト)勤務開始時間', '元のシフト）勤務開始時間']),
        END_B: findCol(['元のシフト勤務終了時間', '元のシフト)勤務終了時間', '元のシフト）勤務終了時間']),
        DEPT: findCol(['診療科']),
        NOTE: findCol(['業務内容', '備考']),
        COMMENT_IDXS: headers.map((h, i) => (String(h).includes('スタッフコメント') || String(h).includes('コメント')) ? i : -1).filter(i => i !== -1)
      };

      const sMIdx = headers.findIndex(h => String(h).includes('勤務開始時間') && !String(h).includes('元の'));
      if (sMIdx !== -1) C.START_M = sMIdx;
      const eMIdx = headers.findIndex(h => String(h).includes('勤務終了時間') && !String(h).includes('元の'));
      if (eMIdx !== -1) C.END_M = eMIdx;

      cache.shiftSheets[cacheKey] = { rows, C };
    } catch (e) {
      return { error: `ファイル展開エラー(${datasetKey}): ${e.message}` };
    }
  }

  return cache.shiftSheets[cacheKey];
}