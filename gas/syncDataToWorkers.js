function syncUserFullData(userId) {
  // 1. ユーザー基本情報を取得 (既存のロジック)
  const userData = getUserInfoFromSheet(userId); 
  // 2. 予約・出席状況を取得 (以前作成した getAllReservationsForUser)
  const reservations = getAllReservationsForUser(userId);

  const fullData = {
    data: userData, // className, upperLimitNumber 等
    myReservedDates: reservations.myReservedDates,
    myAttendedDates: reservations.myAttendedDates
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify({
      type: "USER_FULL_SYNC",
      userId: userId,
      data: fullData
    })
  };
  UrlFetchApp.fetch(WORKERS_URL, options);
}

// ユーザーの全予約データを取得するヘルパー関数
function getAllReservationsForUser(userId) {
  const resSheet = SPREADSHEET.getSheetByName(SHEET_NAME_RESERVATIONS);
  const data = resSheet.getDataRange().getValues();
  const myReservedDates = [];
  const myAttendedDates = [];
  const ssTimezone = SPREADSHEET.getSpreadsheetTimeZone();

  // 1行目から最後までループ
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[RES_COL_USER_ID] !== userId) continue;

    const status = row[RES_COL_STATUS];
    const dateRaw = row[RES_COL_DATE];
    const startTimeRaw = row[RES_COL_START_TIME];
    
    if (!(dateRaw instanceof Date)) continue;

    const dateString = Utilities.formatDate(dateRaw, ssTimezone, 'yyyy-MM-dd');
    const startTimeStr = (startTimeRaw instanceof Date) 
        ? Utilities.formatDate(startTimeRaw, ssTimezone, 'HH:mm') 
        : String(startTimeRaw).trim();
    const fullDateKey = `${dateString} ${startTimeStr}`;

    if (status === '確定') {
       // キャンセル期限などの詳細も必要ならここで取得
       const reservationId = row[RES_COL_RESERVATION_ID];
       const cancellableUntilRaw = row[RES_COL_CANCELLABLE_UNTIL];
       let cancellableUntilString = "";
        if (cancellableUntilRaw instanceof Date) {
            cancellableUntilString = Utilities.formatDate(cancellableUntilRaw, ssTimezone, 'yyyy-MM-dd HH:mm');
        }
       
       let obj = {};
       obj[fullDateKey] = { reservationId: reservationId, cancellableUntil: cancellableUntilString };
       myReservedDates.push(obj);
    } else if (status === '受講済み') {
       myAttendedDates.push(fullDateKey);
    }
  }
  return { myReservedDates, myAttendedDates };
}

/**
 * スプレッドシートの「users」シートから特定ユーザーの基本情報を取得する
 */
function getUserInfoFromSheet(userId) {
  const sheet = SPREADSHEET.getSheetByName(SHEET_NAME_USERS);
  if (!sheet) throw new Error("ユーザーシートが見つかりません。");

  const data = sheet.getDataRange().getValues();
  const header = data[0];
  
  // 各項目の列インデックスを特定（列順が変わっても動くように）
  const col = {
    userId: header.indexOf("ユーザID"),
    displayName: header.indexOf("名前"),
    className: header.indexOf("クラス（一般／おとな美文字）"),
    upperLimit: header.indexOf("稽古回数（初回選択）"),
    limitThis: header.indexOf("今月の稽古回数"),
    limitNext: header.indexOf("来月の稽古回数")
  };

  // ユーザーを検索
  let userRow = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][col.userId] === userId) {
      userRow = data[i];
      break;
    }
  }

  if (!userRow) {
    return null; // 見つからない場合はnullを返し、新規登録フローへ
  }

  // Workersへ送るためのデータ構造を作成
  return {
    userId: userRow[col.userId],
    displayName: userRow[col.displayName],
    className: userRow[col.className],
    upperLimitNumber: userRow[col.upperLimit],
    upperLimitNumberThisMonth: userRow[col.limitThis],
    upperLimitNumberNextMonth: userRow[col.limitNext]
  };
}

/**
 * 予約完了後のWorkers同期を1回で実行する
 */
function syncReservationToWorkers(userId, userData, capacityData) {
  
  const payload = {
    type: "RESERVATION_SYNC", // Workers側で新設したタイプ
    userId: userId,
    userData: userData,      // userInfoFull の内容
    capacityData: capacityData // getCapacityData() の内容
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(WORKERS_URL, options);
    console.log("Workers Sync Response: " + response.getContentText());
  } catch (e) {
    console.error("Workersへの同期に失敗しました: " + e.toString());
  }
}
