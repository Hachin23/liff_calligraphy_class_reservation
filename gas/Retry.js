/**
 * プロセス種別
 */
const PROCESS_TYPE = {
  SYNC_WORKERS: "SYNC_WORKERS"
};

/**
 * リトライステータス
 */
const RETRY_STATUS = {
  WAITING: "WAITING",
  SUCCESS: "SUCCESS",
  FAILED: "FAILED"
};

const RETRY_LIMIT = 3;

/**
 * retryシート 列インデックス定数
 */
const RETRY_COL_RETRY_ID = 0;
const RETRY_COL_PROCESS_TYPE = 1;
const RETRY_COL_ARGS = 2;
const RETRY_COL_RETRY_COUNT = 3;
const RETRY_COL_STATUS = 4;
const RETRY_COL_ERROR_CONTENTS = 5;
const RETRY_COL_REGISTER_DATE = 6;
const RETRY_COL_FINAL_EXECUTE_DATE = 7;

/**
 * 時間主導トリガーから呼び出す再送処理
 */
function retryProcess() {

  if (getConfig().maintenancemode === "ON") {
    Logger.log("メンテナンス中のため実行されませんでした");
    return;
  }

  const lock = LockService.getScriptLock();
  let locked = false;
  
  try {
    lock.waitLock(5000);
    locked = true;

    // スプレッドシート読み込み
    const retrySheet = SPREADSHEET.getSheetByName(SHEET_NAME_RETRY);
    const retrySheetDataWithIndex = retrySheet.getDataRange().getValues().map((row, index) => ({
      rowIndex: index,
      rowData: row
    }))
    .filter(retry => retry.rowIndex > 0)
    .filter(retry => retry.rowData[RETRY_COL_RETRY_COUNT] < RETRY_LIMIT)
    .filter(retry => retry.rowData[RETRY_COL_STATUS] === RETRY_STATUS.WAITING);

    // 対象の再送処理がない場合は、終了
    if (retrySheetDataWithIndex.length === 0) {
      return;
    }

    retrySheetDataWithIndex.forEach(retry => {
      const processType = retry.rowData[RETRY_COL_PROCESS_TYPE];
      const args =  JSON.parse(retry.rowData[RETRY_COL_ARGS]);

      // 処理分岐
      let result = "";
      switch (processType) {
        case PROCESS_TYPE.SYNC_WORKERS:
          result = retrySyncWorkers(args.userId);
          break;
        default:
          result = {
            success: false,
            error: `未対応の処理種別です: ${processType}`
          };
      }
      updateRetryStatus(retrySheet, retry, result.success, result.error);
    });
  
  } catch(e) {
    console.log(`リトライ処理を実行できませんでした: ${e.message}`);
    return;
  } finally {
    if (locked) {
      lock.releaseLock();
    }
  }
}

/**
 * Workers同期失敗データを再同期する
 */
function retrySyncWorkers(userId) {

  try {
    const userData = getUserInfoFromSheet(userId);
    const capacityData = getCapacityData();
    const reservations = getAllReservationsForUser(userId);

    const fullData = {
      data: userData,
      myReservedDates: reservations.myReservedDates,
      myAttendedDates: reservations.myAttendedDates
    };

    syncReservationToWorkers(userId, fullData, capacityData);

    return {
      success: true
    }
  } catch (e) {
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * リトライシート情報を更新
 */
function updateRetryStatus(retrySheet, retryData, success, errorContent = "") {

  const currentCount = retryData.rowData[RETRY_COL_RETRY_COUNT];
  const nextCount = success ? currentCount : currentCount + 1;

  let status;

  if (success) {
    status = RETRY_STATUS.SUCCESS;
  } else if (nextCount >= RETRY_LIMIT) {
    status = RETRY_STATUS.FAILED;
  } else {
    status = RETRY_STATUS.WAITING;
  }

  // リトライ回数
  retrySheet
    .getRange(retryData.rowIndex + 1, RETRY_COL_RETRY_COUNT + 1)
    .setValue(nextCount);
  // ステータス
  retrySheet
    .getRange(retryData.rowIndex + 1, RETRY_COL_STATUS + 1)
    .setValue(status);
  // エラー内容
  retrySheet
  .getRange(retryData.rowIndex + 1, RETRY_COL_ERROR_CONTENTS + 1)
  .setValue(errorContent);
  // 最終実行時間
  retrySheet
  .getRange(retryData.rowIndex + 1, RETRY_COL_FINAL_EXECUTE_DATE + 1)
  .setValue(new Date());
}

/**
 * リトライ処理をretryシートに登録
 */
function retryProcessLog(log) {
  const sheet = SPREADSHEET.getSheetByName(SHEET_NAME_RETRY);

  sheet.appendRow([
    log.retryId,
    log.processType,
    log.args,
    log.retryCount,
    log.status,
    log.errorContents,
    log.registerDate,
    log.finalExecuteDate
  ]);
}
