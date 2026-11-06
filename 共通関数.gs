/**
 * 【共通】ユーティリティ関数群
 */

// URL取得ヘルパー
function getUrlFromDataset(configUrl, keyName) {
  try {
    const ss = SpreadsheetApp.openByUrl(configUrl);
    const sheet = ss.getSheetByName('データセット');
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === keyName) {
        return String(data[i][1]).trim();
      }
    }
  } catch (e) {
    Logger.log(`URL取得エラー(${keyName}): ${e.message}`);
  }
  return "";
}

// 氏名の正規化（スペース削除）
function normalizeName(name) {
  if (!name) return "";
  return String(name).replace(/[\s　]+/g, "").trim();
}

// クリニック名の正規化（表記ゆれ吸収）
function normalizeClinic(name) {
  if (!name) return "";
  let n = String(name).trim();
  
  // ヶ・ケ統一
  n = n.replace(/ヶ/g, "ケ");

  // 特定のゆれ吸収
  if (n.includes('千葉') && n.match(/ニュータウン|NT|ＮＴ/i)) return '千葉ニュータウン中央';
  if (n === '八千代') return '八千代緑が丘';
  
  // 阿波座・長吉長原などはそのまま返す（シフト表とマスタで文字が完全一致していればOK）
  return n;
}

// 日付の正規化（yyyy-MM-dd）
function formatDateKey(dateObj) {
  if (!dateObj || !(dateObj instanceof Date) || isNaN(dateObj.getTime())) return null;
  return Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy-MM-dd');
}

// 列インデックス検索（あいまい検索）
function findColumnIndex(headers, keyword) {
  return headers.findIndex(h => String(h).includes(keyword));
}