/**
 * ⏰ 自動実行用トリガー関数
 * UIを使わずに、当月と翌月のデータを更新します。
 * 年度またぎ（例: 3月実行時の翌月4月）も updateLocationDashboard 側で自動判定されます。
 */
function triggerUpdateForCurrentAndNext() {
  const today = new Date();
  
  // 1. 当月
  const y1 = today.getFullYear();
  const m1 = today.getMonth() + 1;
  
  // 2. 翌月 (12月の次は翌年1月)
  const nextMonthDate = new Date(today.getFullYear(), today.getMonth() + 1, 1);
  const y2 = nextMonthDate.getFullYear();
  const m2 = nextMonthDate.getMonth() + 1;

  Logger.log(`▼ 自動更新トリガー開始: [当月] ${y1}/${m1}, [翌月] ${y2}/${m2}`);

  // 実行
  const res1 = updateLocationDashboard(y1, m1);
  Logger.log(`当月(${m1}月) 結果: ${res1}`);
  
  const res2 = updateLocationDashboard(y2, m2);
  Logger.log(`翌月(${m2}月) 結果: ${res2}`);
  
  Logger.log("▲ 自動更新トリガー終了");
}