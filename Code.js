// ★★★ ここをあなたのスプレッドシートIDに置き換えてください ★★★
const SPREADSHEET_ID = '';
const SHEET_NAME = '集計'; // データを記録するシート名

/**
 * Webアプリケーションとしてアクセスされたときに実行される関数
 * index.htmlを返します。
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * クライアント側（index.html）からデータを受け取り、スプレッドシートに書き込む関数
 */
function recordKakeibo(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    // --- 日付処理 (JSTの時刻を反映させる) ---
    const dateStr = formData.date;
    const nowJST = new Date();
    const selectedDateUTC = new Date(dateStr); 

    const finalDate = new Date();
    finalDate.setFullYear(selectedDateUTC.getFullYear());
    finalDate.setMonth(selectedDateUTC.getMonth());
    finalDate.setDate(selectedDateUTC.getDate());
    finalDate.setHours(nowJST.getHours());
    finalDate.setMinutes(nowJST.getMinutes());
    finalDate.setSeconds(nowJST.getSeconds());
    // ---------------------------------------------------

    const amountValue = Number(formData.amount);
    
    // 記録するデータの配列 (A:日付, B:費目, C:金額, D:メモ)
    const rowData = [
      finalDate,           
      formData.item,       
      amountValue,         
      formData.memo        
    ];
    
    sheet.appendRow(rowData);

    // 費目が「収入」の場合、メールレポートを送信する
    if (formData.item === '収入') {
      sendMonthlyReport(ss);
      return `収入 ${amountValue.toLocaleString()}円 を記録しました。収支報告メールを送信しました。`;
    }
    
    return `${formData.item} ${amountValue.toLocaleString()}円 を記録しました。`;

  } catch (e) {
    return "エラーが発生しました: " + e.toString();
  }
}

/**
 * 指定した年月のデータを集計する
 * @param {number} targetYear - 集計したい年 (例: 2026) / 指定なければ現在
 * @param {number} targetMonth - 集計したい月 (1-12) / 指定なければ現在
 */
function getMonthlySummary(targetYear, targetMonth) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { totalExpense: '0', totalIncome: '0', breakdown: {}, error: null };
    }
    
    // データ範囲を取得
    const range = sheet.getRange(2, 1, lastRow - 1, 4);
    const values = range.getValues();
    
    // 集計対象の年月を決定（引数がなければ現在時刻を使う）
    const now = new Date();
    // 引数で月が渡された場合、JavaScriptの月(0-11)に合わせるために-1する
    const currentYear = targetYear || now.getFullYear();
    const currentMonth = targetMonth ? targetMonth - 1 : now.getMonth();
    
    let monthlyTotalExpense = 0;
    let monthlyTotalIncome = 0;
    const categoryTotals = {};
    
    values.forEach(row => {
      const date = row[0]; 
      const item = row[1]; 
      const amount = row[2]; 
      
      if (date && typeof date.getMonth === 'function') {
        const dataMonth = date.getMonth();
        const dataYear = date.getFullYear();

        // 指定された年月と一致、かつ金額が正の値
        if (dataMonth === currentMonth && dataYear === currentYear && typeof amount === 'number' && amount > 0) {
          
          if (item === '収入') {
            monthlyTotalIncome += amount;
          } else {
            monthlyTotalExpense += amount;
            
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
      year: currentYear,
      month: currentMonth + 1, // 表示用に1-12月に戻す
      totalExpense: monthlyTotalExpense.toLocaleString(),
      totalIncome: monthlyTotalIncome.toLocaleString(), 
      breakdown: breakdown, 
      error: null 
    };
    
  } catch (e) {
    return { error: e.toString() };
  }
}

// 設定シートの名前
const SETTINGS_SHEET_NAME = '設定'; 

function getSettingValue(key) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    const values = settingsSheet.getRange('A:B').getValues();

    for (let i = 0; i < values.length; i++) {
      if (values[i][0].toString().trim() === key.trim()) {
        return values[i][1].toString().trim().toUpperCase();
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}

function sendMonthlyReport(ss) {
    // 引数なしで呼び出すと「今月」の集計になる
    const summary = getMonthlySummary();
    
    if (summary.error) {
        Logger.log("報告失敗: " + summary.error);
        return;
    }
    
    const totalExpense = summary.totalExpense;
    const totalIncome = summary.totalIncome;

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
}