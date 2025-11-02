// ★★★ ここをあなたのスプレッドシートIDに置き換えてください ★★★
const SPREADSHEET_ID = '1MWWk7xYLVViQXCAVam2erP80Jtwj3IR1OnE4VcHc3xw';
const SHEET_NAME = '集計'; // データを記録するシート名

/**
 * Webアプリケーションとしてアクセスされたときに実行される関数
 * index.htmlを返します。
 */
function doGet() {
  // HTMLファイルをテンプレートとして読み込み、評価して返す
  const template = HtmlService.createTemplateFromFile('index');
  // スマホ対応のため、viewportタグを追加
  return template.evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * クライアント側（index.html）からデータを受け取り、スプレッドシートに書き込む関数
 * @param {object} formData - HTMLフォームから送られたデータ
 */
function recordKakeibo(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    // --- ★★★ 日付処理を元の正しいロジックに復元 ★★★ ---
    const dateStr = formData.date;
    const nowJST = new Date();
    const selectedDateUTC = new Date(dateStr); // 日付部分のみを持つDateオブジェクト (UTC 00:00:00)

    // 記録する最終的なDateオブジェクトを作成
    const finalDate = new Date();
    
    // 日付部分をフォーム入力で上書き (getFullYear, getMonth, getDate)
    finalDate.setFullYear(selectedDateUTC.getFullYear());
    finalDate.setMonth(selectedDateUTC.getMonth());
    finalDate.setDate(selectedDateUTC.getDate());
    
    // 時刻部分を現在のJST時刻で上書き (getHours, getMinutes, getSeconds)
    finalDate.setHours(nowJST.getHours());
    finalDate.setMinutes(nowJST.getMinutes());
    finalDate.setSeconds(nowJST.getSeconds());
    // ---------------------------------------------------

    // 金額を数値に変換
    const amountValue = Number(formData.amount);
    
    // 記録するデータの配列 (A:日付, B:費目, C:金額, D:メモ)
    const rowData = [
      finalDate,           // A列: 日付 (時刻調整済み)
      formData.item,       // B列: 費目 ('食費' や '収入' などがそのまま入る)
      amountValue,         // C列: 金額
      formData.memo        // D列: メモ
    ];
    
    sheet.appendRow(rowData);

    // 費目が「収入」の場合、メールレポートを送信する
    if (formData.item === '収入') {
      sendMonthlyReport(ss);
      return `収入 ${amountValue.toLocaleString()}円 を記録しました。収支報告メールを送信しました。`;
    }
    
    // 支出の場合
    return `${formData.item} ${amountValue.toLocaleString()}円 を記録しました。`;

  } catch (e) {
    // 構文エラーを解消するため、tryブロックの末尾にcatchブロックを配置
    return "エラーが発生しました: " + e.toString();
  }
}

/**
 * スプレッドシートのデータを読み込み、月別総額と費目別割合を計算する
 * @returns {object} 月の総額、費目別割合、およびエラー情報を含むオブジェクト
 */
function getMonthlySummary() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      // ヘッダー行のみの場合、空のデータを返す
      return { totalExpense: '0', totalIncome: '0', breakdown: {}, error: null };
    }
    
    // データ範囲（A2から最終行までの4列）を取得
    const range = sheet.getRange(2, 1, lastRow - 1, 4);
    const values = range.getValues();
    
    // 現在の月のデータを抽出 (GASサーバーのJST時刻を使用)
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    
    let monthlyTotalExpense = 0;
    let monthlyTotalIncome = 0;
    const categoryTotals = {};
    
    values.forEach(row => {
      const date = row[0]; // A列：日付
      const item = row[1]; // B列：費目
      const amount = row[2]; // C列：金額 (数値)
      
      // 日付が有効なDateオブジェクトであるか確認
      if (date && typeof date.getMonth === 'function') {
        const dataMonth = date.getMonth();
        const dataYear = date.getFullYear();

        // 今月かつ金額が正の値であることを確認 (負の値の入力は考慮しない)
        if (dataMonth === currentMonth && dataYear === currentYear && typeof amount === 'number' && amount > 0) {
          
          if (item === '収入') {
            // 収入として集計
            monthlyTotalIncome += amount;
          } else {
            // 支出として集計（費目が「収入」以外、かつ金額が正の値）
            monthlyTotalExpense += amount;
            
            // 費目ごとの合計を計算（支出のみ）
            const categoryName = item.toString().trim();
            if (categoryName) {
                if (!categoryTotals[categoryName]) {
                    categoryTotals[categoryName] = 0;
                }
                categoryTotals[categoryName] += amount;
            }
          }
        }
      }
    });
    
    const breakdown = {};
    // 費目別割合を計算 (総支出が0の場合は計算しないようにガード)
    if (monthlyTotalExpense > 0) {
        for (const item in categoryTotals) {
          const percentage = (categoryTotals[item] / monthlyTotalExpense) * 100;
          breakdown[item] = {
            amount: categoryTotals[item].toLocaleString(),
            percentage: percentage.toFixed(1)
          };
        }
    }

    return { 
      totalExpense: monthlyTotalExpense.toLocaleString(),
      totalIncome: monthlyTotalIncome.toLocaleString(), 
      breakdown: breakdown, 
      error: null 
    };
    
  } catch (e) {
    return { 
      totalExpense: '0', 
      totalIncome: '0', 
      breakdown: {}, 
      error: e.toString() 
    };
  }
}

// 設定シートの名前
const SETTINGS_SHEET_NAME = '設定'; 

/**
 * 設定シートから指定した設定項目の値を取得する
 * @param {string} key - 設定項目の名前 (例: '収支報告ON/OFF')
 * @returns {string|null} 設定値 (例: 'ON')
 */
function getSettingValue(key) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    
    const values = settingsSheet.getRange('A:B').getValues();

    for (let i = 0; i < values.length; i++) {
      if (values[i][0].toString().trim() === key.trim()) {
        const settingValue = values[i][1].toString().trim();
        return settingValue.toUpperCase(); // 必ず大文字にして返す
      }
    }
    return null;
  } catch (e) {
    Logger.log("設定の読み込み中にエラーが発生しました: " + e);
    return null;
  }
}

/**
 * 月次報告メールを送信する
 */
function sendMonthlyReport(ss) {
    const summary = getMonthlySummary();
    
    if (summary.error) {
        Logger.log("報告失敗: " + summary.error);
        return;
    }
    
    const totalExpense = summary.totalExpense;
    const totalIncome = summary.totalIncome;

    // 数値として計算し直す
    const netBalanceNum = Number(summary.totalIncome.replace(/,/g, '')) - Number(summary.totalExpense.replace(/,/g, ''));
    const netBalance = netBalanceNum.toLocaleString();
    
    let breakdownText = "";
    for (const item in summary.breakdown) {
        const data = summary.breakdown[item];
        breakdownText += `${item}: ${data.amount}円 (${data.percentage}%)\n`;
    }
    
    const recipient = Session.getActiveUser().getEmail(); 
    const subject = `【家計簿】収入が記録されました - 今月の収支速報`;
    const body = `
        収入が記録されましたので、現時点の収支状況をご報告します。

        - 今月の総収入額: ${totalIncome}円
        - 今月の総支出額: ${totalExpense}円
        - 現時点の収支: ${netBalance}円 ${netBalanceNum < 0 ? '(赤字)' : '(黒字)'}

        ---
        費目別内訳（支出のみ）:
        ${breakdownText || '今月は支出がありません。'}
        ---

        詳細はスプレッドシートをご確認ください。
        ${ss.getUrl()}
    `;
    
    MailApp.sendEmail(recipient, subject, body);
    Logger.log("収支報告メールを送信しました。");
}
