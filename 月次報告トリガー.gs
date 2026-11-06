/**
 * 🚀 【手動トリガー】当月と翌月の全体集計を実行
 * ※名前を変更した「updateDashboardFinal_Fixed」を呼び出すように調整済みです。
 * これで確実に修正版のロジックが動きます。
 */
function triggerUpdateFinal_CurrentAndNext() {
  const today = new Date();

  // --- 1. 当月の計算 ---
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth() + 1;

  // --- 2. 翌月の計算 (12月の次は来年の1月にする処理) ---
  let nextYear = currentYear;
  let nextMonth = currentMonth + 1;
  if (nextMonth > 12) {
    nextMonth = 1;
    nextYear++;
  }

  Logger.log(`=== 自動実行トリガー開始 ===`);

  // 当月実行
  Logger.log(`▶ 当月 (${currentYear}年${currentMonth}月) を実行中...`);
  // ★重要: Fixed版を呼び出します
  const res1 = updateDashboardFinal_Fixed(currentYear, currentMonth);
  Logger.log(`   結果: ${res1}`);

  // 翌月実行
  Logger.log(`▶ 翌月 (${nextYear}年${nextMonth}月) を実行中...`);
  // ★重要: Fixed版を呼び出します
  const res2 = updateDashboardFinal_Fixed(nextYear, nextMonth);
  Logger.log(`   結果: ${res2}`);

  Logger.log(`=== 全処理完了 ===`);
}