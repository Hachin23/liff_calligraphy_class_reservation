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
 * 以下のスプレッドシートから特定ユーザーの基本情報を取得する
 * usersシート
 * userMonthlySubscriptionsシート
 * userTicketsシート
 */
function getUserInfoFromSheet(userId) {
  const userSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USERS);
  if (!userSheet) throw new Error("usersシートが見つかりません。");
  
  const userData = userSheet.getDataRange().getValues();
  const userHeader = userData[0];

  // 各項目の列インデックスを特定（列順が変わっても動くように）
  const userCol = {
    userId: userHeader.indexOf("ユーザID"),
    displayName: userHeader.indexOf("名前"),
    className: userHeader.indexOf("クラス（一般／おとな美文字）")
  };

  // ユーザーを検索
  let userRow = null;
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][userCol.userId] === userId) {
      userRow = userData[i];
      break;
    }
  }

  if (!userRow) {
    return null; // 見つからない場合はnullを返し、新規登録フローへ
  }

  const userMonthlySubscriptionsSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USER_MONTHLY_SUBSCRIPTIONS);
  if (!userMonthlySubscriptionsSheet) throw new Error("userMonthlySubscriptionsシートが見つかりません。");

  const userMonthlySubscriptionsData = userMonthlySubscriptionsSheet.getDataRange().getValues();
  const userMonthlySubscriptionsHeader = userMonthlySubscriptionsData[0];

  const userMonthlySubscriptionsCol = {
    userId: userMonthlySubscriptionsHeader.indexOf("ユーザID"),
    upperLimit: userMonthlySubscriptionsHeader.indexOf("稽古回数（デフォルト）"),
    limitThis: userMonthlySubscriptionsHeader.indexOf("今月の稽古回数"),
    limitNext: userMonthlySubscriptionsHeader.indexOf("来月の稽古回数")
  };

  // ユーザーの月の稽古回数を検索
  let userMonthlySubscriptionsRow = null;
  for (let i = 1; i < userMonthlySubscriptionsData.length; i++) {
    if (userMonthlySubscriptionsData[i][userMonthlySubscriptionsCol.userId] === userId) {
      userMonthlySubscriptionsRow = userMonthlySubscriptionsData[i];
      break;
    }
    // 見つからない場合は、チケット利用のみの生徒（回数を「0」に固定）
  }
  
  const userTicketsSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USER_TICKETS);
  if (!userTicketsSheet) throw new Error("userTicketsシートが見つかりません。");

  const userTicketsData = userTicketsSheet.getDataRange().getValues();
  const userTicketsHeader = userTicketsData[0];

  const userTicketsCol = {
    remainingNumber: userTicketsHeader.indexOf("残数"),
    expirationDate: userTicketsHeader.indexOf("有効期限")
  };
  
  // ユーザーのチケット情報を検索
  const dateStringNow = Utilities.formatDate(new Date(), SPREADSHEET.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  let userTicketsRows = userTicketsData.slice(1).filter(row => row[0] === userId)
  // 有効期限が残っているものだけ取得
  let validUserTicketsRows = userTicketsRows.filter(row => Utilities.formatDate(row[userTicketsCol.expirationDate], SPREADSHEET.getSpreadsheetTimeZone(), 'yyyy-MM-dd') > dateStringNow)
  let remainingNumberTotal = validUserTicketsRows.length === 0 ? 0 : validUserTicketsRows.map(row => row[userTicketsCol.remainingNumber]).reduce((total, num) => total + num, 0);

  // Workersへ送るためのデータ構造を作成
  return {
    userId: userRow[userCol.userId],
    displayName: userRow[userCol.displayName],
    className: userRow[userCol.className],

    upperLimitNumber: userMonthlySubscriptionsRow == null ? 0 : userMonthlySubscriptionsRow[userMonthlySubscriptionsCol.upperLimit],
    upperLimitNumberThisMonth: userMonthlySubscriptionsRow == null ? 0 : userMonthlySubscriptionsRow[userMonthlySubscriptionsCol.limitThis],
    upperLimitNumberNextMonth: userMonthlySubscriptionsRow == null ? 0 : userMonthlySubscriptionsRow[userMonthlySubscriptionsCol.limitNext],

    ticketInfo: {
      dispInfo: validUserTicketsRows.map(row => {
        const obj = {};
        obj.remainingNumber = row[userTicketsCol.remainingNumber];
        obj.expirationDate = Utilities.formatDate(row[userTicketsCol.expirationDate], SPREADSHEET.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
        return obj;
      }),
      remainingNumberTotal: remainingNumberTotal,
      // チケット購入履歴があるか
      purchaseHistory: userTicketsRows.length !== 0 ? true : false
    },
    afterInitialRegistration: false
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
