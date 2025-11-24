// --- グローバル設定 ---
const NAVITIME_BASE_URL = "https://www.navitime.co.jp/transfer/searchlist";

const SearchType = {
  Arrival: 0,
  Departure: 1
};

/**
 * [トリガーで毎朝5時に実行]
 * ユーザーリストを参照し、各担当者にスケジュールを送信する
 */
function executeSchedules() {
  
  // 1. IDを使ってスプレッドシートを開く (プロパティから取得)
  const props = PropertiesService.getScriptProperties();
  const SHEET_ID = props.getProperty('SPREADSHEET_ID');
  
  if (!SHEET_ID) {
    Logger.log('エラー: スクリプトプロパティ "SPREADSHEET_ID" が設定されていません。');
    return;
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const scheduleSheet = ss.getSheetByName("schedule");
  const userSheet = ss.getSheetByName("users"); // ★新設シート

  if (!scheduleSheet || !userSheet) {
    Logger.log('エラー: "schedule" または "users" シートが見つかりません。');
    return;
  }

  // 2. ユーザー情報を読み込んで「対応マップ」を作る
  // { "パパ": "U123...", "ママ": "U456..." } の形にする
  const userData = userSheet.getDataRange().getDisplayValues().slice(1);
  const userMap = {};
  userData.forEach(row => {
    const name = row[0];
    const id = row[1];
    if (name && id) {
      userMap[name] = id;
    }
  });

  // 3. 今日の日付情報を取得
  const today = new Date();
  const todayDay = today.getDay(); 
  Logger.log(`--- ${Utilities.formatDate(today, "JST", "yyyy/MM/dd")} の処理を開始 ---`);

  // 4. スケジュールを走査
  const schedules = scheduleSheet.getDataRange().getDisplayValues().slice(1);

  schedules.forEach((row) => {
    const purpose = row[0];
    const departure = row[1];
    const arrival = row[2];
    const via = row[3];
    const enabled = row[4];
    const dayToExecute = row[5];
    const timeToSearch = row[6];
    const targetUserName = row[7]; // ★H列: 送信先ユーザー名

    const isEnabled = (enabled === "ON");
    const isDayMatch = (dayToExecute === '*' || dayToExecute == todayDay);

    // 送信先ユーザー名に対応するIDがあるか確認
    const targetUserId = userMap[targetUserName];

    if (isEnabled && isDayMatch && departure && arrival && targetUserId) {

      // 時刻設定
      const [hour, minute] = timeToSearch.split(':');
      const searchDate = new Date();
      searchDate.setHours(parseInt(hour, 10));
      searchDate.setMinutes(parseInt(minute, 10));
      searchDate.setSeconds(0);

      // URL生成
      const url = createNavitimeUrl(
        departure,
        arrival,
        searchDate,
        SearchType.Departure,
        via
      );

      // 遅延確認用時刻
      const now = new Date();
      const timeString = Utilities.formatDate(now, "JST", "HH:mm:ss");

      const message = `[${purpose}]\n${url}\n\n(送信実行: ${timeString})`;
      
      // ★特定したID宛に送信
      sendLineToUser(message, targetUserId);
      Logger.log(`${targetUserName} さんに送信しました: ${purpose}`);
    } else if (!targetUserId && isEnabled) {
      Logger.log(`スキップ: ユーザー "${targetUserName}" のIDが見つかりません。`);
    }
  });

  Logger.log("--- 処理完了 ---");
}

/**
 * 指定したユーザーIDにLINEメッセージを送信する関数 (引数を変更)
 */
function sendLineToUser(messageText, userId) {
  const TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_TOKEN');

  if (!TOKEN) {
    Logger.log('エラー: LINE_CHANNEL_TOKEN が設定されていません。');
    return;
  }

  const url = "https://api.line.me/v2/bot/message/push";
  
  const payload = {
    "to": userId, // ★引数で受け取ったIDを使う
    "messages": [
      {
        "type": "text",
        "text": messageText
      }
    ]
  };

  const options = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + TOKEN,
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload)
  };

  try {
    UrlFetchApp.fetch(url, options);
  } catch (e) {
    Logger.log(`LINE送信エラー (${userId}宛): ` + e.message);
  }
}

// --- NAVITIME URL生成関数は変更なし (そのまま残してください) ---
function createNavitimeUrl(from, to, date, searchType, through) {
  const params = {
    orvStationName: from,
    dnvStationName: to,
    year: date.getFullYear(),
    month: date.getMonth() + 1,
    day: date.getDate(),
    hour: date.getHours(),
    minute: date.getMinutes(),
    thr1StationName: through,
    basis: searchType
  };
  const queryParts = [];
  for (const [key, value] of Object.entries(params)) {
    if (value !== undefined && value !== null && value !== '') {
      queryParts.push(`${encodeURIComponent(key)}=${encodeURIComponent(value.toString())}`);
    }
  }
  return `${NAVITIME_BASE_URL}?${queryParts.join('&')}`;
}