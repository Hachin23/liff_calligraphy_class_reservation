/**
 * 本番環境移行用：全データを一括でWorkers KVへ同期する。
 */
function runFullMigrationToProduction() {
  Logger.log("--- 本番環境への一括移行を開始します ---");

  try {
    // 1. Config情報の同期
    Logger.log("1/3: Config情報を同期中...");
    syncConfigToWorkers();

    // 2. 残席情報（Capacity）の同期
    Logger.log("2/3: 全クラスの残席情報を同期中...");
    syncCapacityToWorkers();

    // 3. 全ユーザー情報の同期
    Logger.log("3/3: 全ユーザーの予約状況を同期中...");
    migrateAllUsersToWorkers();

    Logger.log("--- 全ての移行プロセスが正常に完了しました ---");
  } catch (e) {
    Logger.log("!!! 移行プロセス中にエラーが発生しました !!!");
    Logger.log("エラー内容: " + e.toString());
  }
}

/**
 * 移行用：全ユーザーの現在の予約状況をWorkers KVへ一括同期する
 */
function migrateAllUsersToWorkers() {
  const usersSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USERS);
  const data = usersSheet.getDataRange().getValues();
  
  // 1行目（ヘッダー）を除いてループ
  for (let i = 1; i < data.length; i++) {
    const userId = data[i][0]; // A列がIDの場合
    if (userId) {
      try {
        syncUserFullData(userId); // 前に作った一括同期関数
        Logger.log("同期完了: " + userId);
        Utilities.sleep(200); // 連続リクエストによる負荷を考慮
      } catch (e) {
        Logger.log("同期失敗: " + userId + " / " + e.message);
      }
    }
  }
  Logger.log("全ユーザーの移行が完了しました");
}

// =========================================================================
// Cloudflareへ転送処理（全データを同期）
// =========================================================================
function syncCapacityToWorkers(userIdForTest) {
  const ssTimezone = SPREADSHEET.getSpreadsheetTimeZone();
  const listSheet = SPREADSHEET.getSheetByName(SHEET_NAME_RESERVATIONS_LIST);
  
  // 1. 全ての残席データを取得（カレンダー表示用）
  // 過去分も含めると重いので、今日以降のデータに絞るのが理想ですが、まずは全件で実装
  const capacityResult = getCapacityForMonthInternalAll(listSheet, ssTimezone);
  
  const payload = {
    type: "CAPACITY_ALL",
    data: capacityResult.capacityData
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload)
  };

  UrlFetchApp.fetch(WORKERS_URL, options);
  console.log("Cloudflareへの同期が完了しました。");
}

// 既存の関数を少し改造して、月指定なしで全取得するヘルパー
function getCapacityForMonthInternalAll(listSheet, ssTimezone) {
  const allListValues = listSheet.getDataRange().getValues();
  if (allListValues.length <= 1) return { success: true, capacityData: {} };
  
  const listValues = allListValues.slice(1);
  const capacityData = {};

  listValues.forEach(row => {
    const dateRaw = row[RES_LIST_COL_DATE];
    if (!(dateRaw instanceof Date)) return;
    
    const dateString = Utilities.formatDate(dateRaw, ssTimezone, 'yyyy-MM-dd');
    const startTimeRaw = row[RES_LIST_COL_START_TIME];
    const startTime = (startTimeRaw instanceof Date) ? Utilities.formatDate(startTimeRaw, ssTimezone, 'HH:mm') : String(startTimeRaw);
    const endTimeRaw = row[RES_LIST_COL_END_TIME];
    const endTime = (endTimeRaw instanceof Date) ? Utilities.formatDate(endTimeRaw, ssTimezone, 'HH:mm') : String(endTimeRaw);

    if (!capacityData[dateString]) capacityData[dateString] = [];
    capacityData[dateString].push({
      lessonId: row[RES_LIST_COL_LESSON_ID],
      startTime: startTime,
      endTime: endTime,
      className: row[RES_LIST_COL_CLASS_NAME],
      remainingCapacity: parseInt(row[RES_LIST_COL_REMAINING_CAPACITY], 10)
    });
  });
  return { success: true, capacityData: capacityData };
}

/**
 * スプレッドシートのConfig情報をWorkers KVへ同期する
 */
function syncConfigToWorkers() {
  const configSheet = SPREADSHEET.getSheetByName(SHEET_NAME_CONFIG);
  
  // スプレッドシートから設定を読み取る (セルの位置は現在のシートに合わせて調整してください)
  // 例: B2にバージョン、B5:B10にクラス名、C5:C10に上限回数がある想定
  const configData = {
    version: configSheet.getRange('B2').getValue(),
    CLASS_INFO: {
      CLASS_NAME: configSheet.getRange(3, 2, 1, configSheet.getLastColumn()-1).getValues()[0].filter(String),
      UPPER_LIMIT_NUMBER: configSheet.getRange(4, 2, 1, configSheet.getLastColumn()-1).getValues()[0].filter(String)
    },
    CALENDAR_INFO: {
      FIRST_DAY_OF_WEEK: configSheet.getRange('B5').getValue()
    }
  };

  const payload = {
    type: "CONFIG_UPDATE",
    data: configData
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload)
  };

  try {
    UrlFetchApp.fetch(WORKERS_URL, options);
    Logger.log("Configの同期に成功しました");
  } catch (e) {
    Logger.log("Configの同期失敗: " + e.toString());
  }
}