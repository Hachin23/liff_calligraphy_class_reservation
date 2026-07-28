// ================================
// 定数
// ================================
// 共通で利用するプロパティサービスのインスタンス
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
// スプレッドシートID（指定するスプレッドシートによって変更する）
const SPREADSHEET_ID = getRequiredProperty("SPREADSHEET_ID");
// 毎回スプレッドシートを開くと、レスポンスが悪くなるので定数化
const SPREADSHEET = SpreadsheetApp.openById(SPREADSHEET_ID);
// usersシート
const SHEET_NAME_USERS = 'users';
// reservaionシート
const SHEET_NAME_RESERVATIONS = 'reservations';
// userMonthlySubscriptionsシート
const SHEET_NAME_USER_MONTHLY_SUBSCRIPTIONS = 'userMonthlySubscriptions';
// userTicketsシート
const SHEET_NAME_USER_TICKETS = 'userTickets';
// configシート
const SHEET_NAME_CONFIG = 'config';
// scheduleTemplateシート
const SHEET_NAME_SCHEDULE_TEMPLATE = 'scheduleTemplate';
// reservationsListシート
const SHEET_NAME_RESERVATIONS_LIST = 'reservationsList';
// exceptionsシート
const SHEET_NAME_EXCEPTIONS = 'exceptions';

// ユーザ情報のキャッシュ保持期間（秒）
const CACHE_SECONDS = 300;
// cliend_idの検証用に定義（チャネルID）
const LIFF_CLIENT_ID = getRequiredProperty("LIFF_CLIENT_ID");
// LINEのアクセストークン検証APIのURL
const LINE_VERIFY_URL = "https://api.line.me/oauth2/v2.1/verify";
// LINEのプロフィール取得APIのURL
const LINE_PROFILE_URL = "https://api.line.me/v2/profile";

const WORKERS_URL = getRequiredProperty("WORKERS_URL");

// 予約可能日リスト生成と管理関連の定数
const NUMBER_TO_DAY_NAME = ['日', '月', '火', '水', '木', '金', '土']; // 逆引き用

// 予約シート、ユーザーシートの列インデックス定数
const RES_COL_RESERVATION_ID = 0;
const RES_COL_DATE = 1;
const RES_COL_START_TIME = 2;
const RES_COL_LESSON_ID = 3;
const RES_COL_USER_ID = 4;
const RES_COL_RESERVER_NAME = 5;
const RES_COL_SELECTED_CLASS_NAME = 6;
const RES_COL_DURATION_MINUTES = 7; // 期間（分）
const RES_COL_END_TIME = 8;         // 終了時間 (New)
const RES_COL_CANCELLABLE_UNTIL = 9;
const RES_COL_RESERVATION_DATETIME = 10;
const RES_COL_OPERATION = 11;
const RES_COL_STATUS = 12;
const RES_COL_MEMO = 13;
const RES_COL_CALENDAR_EVENT_ID = 14;

const RES_LIST_COL_DATE = 0;
const RES_LIST_COL_DAY_OF_WEEK = 1;
const RES_LIST_COL_START_TIME = 2;
const RES_LIST_COL_END_TIME = 3;
const RES_LIST_COL_CLASS_NAME = 4;
const RES_LIST_COL_LESSON_ID = 5;
const RES_LIST_COL_CAPACITY = 6;
const RES_LIST_COL_REMAINING_CAPACITY = 7;

// GoogleカレンダーID(連携するカレンダーID)
const ADMIN_CALENDAR_ID = getRequiredProperty("ADMIN_CALENDAR_ID");

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEET_NAME_CONFIG);
  
  if (!configSheet) {
    throw new Error(`${SHEET_NAME_CONFIG} シートが見つかりません。`);
  }

  const classNames = configSheet.getRange('B3:Z3').getValues()[0].filter(n => n);
  const firstDayOfWeek = configSheet.getRange('B4').getValue();

  // CONFIG オブジェクトを作成
  const CONFIG = {
    version: configSheet.getRange('B2').getValue(),
    CLASS_INFO: {
      CLASS_NAME: classNames
    },
    CALENDAR_INFO: {
      FIRST_DAY_OF_WEEK: firstDayOfWeek
    }
  }
  return CONFIG;
}


// ===========================================
// doPost ユーザ情報取得/登録・予約登録・キャンセル
// ===========================================
function doPost(e) {
  const params = e.parameter;
  const mode = params.mode;
  let result;
  try {
    switch (mode) {
      case "verifyAndGetUserInfo":
        result = verifyTokenAndGetUserInfo(params.accessToken);
        return result;
      case "registerUserInfo":
        result = registerUserClassGAS(params);
        break;
      case "cancelReservation":
        verifyToken(params.accessToken);
        result = handleCancelReservation(params);
        break;
      case "getCalendarData":
        result = getCalendarData(params);
        break;
      case "makeReservation":
        verifyToken(params.accessToken);
        result = makeReservation(params);
        break;
      default:
        result = { success: false, message: "Invalid mode" };
        break;
    }
  } catch (error) {
    result = { success: false, message: "GAS Internal Error: " + error.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 予約可能日リストシートから、指定された年月のクラス残席情報を取得・整形して返す。
 * @param {Object} params - { year: number, month: number }
 * @returns {Object} { success: boolean, capacityData: { 'YYYY-MM-DD': [{ lessonId: 1, startTime: 'HH:mm', className: '...', remainingCapacity: N }, ...] } }
 */
function getCalendarData(params) {
    const targetYear = parseInt(params.year, 10);
    const targetMonth = parseInt(params.month, 10); // 1-12
    const userId = params.userId;
    const monthKey = params.monthKey;

    // パラメータチェック
    if (isNaN(targetYear) || isNaN(targetMonth) || !userId) {
        return { success: false, message: 'Invalid or missing parameters.' };
    }
    const listSheet = SPREADSHEET.getSheetByName(SHEET_NAME_RESERVATIONS_LIST);
    if (!listSheet) {
        return { success: false, message: 'ReservationsList sheet not found.' };
    }

    // キャッシュサービスを取得
    const userCache = CacheService.getUserCache();

    //【ユーザーキャッシュ】自身の予約日のキャッシュチェック (10分間)
    const userCacheKey = 'USER_RESERVATIONS_' + userId + "_" + monthKey;
    let myReservedDates = [];
    let myAttendedDates = [];
    
    const cachedUser = userCache.get(userCacheKey);
    if (cachedUser) {
        const userResult = JSON.parse(cachedUser);
        myReservedDates = userResult.myReservedDates || [];
        myAttendedDates = userResult.myAttendedDates || [];
        Logger.log(`ユーザー予約情報をキャッシュから取得しました。`);
    } else {
        // キャッシュミスの場合: スプレッドシートから取得 (既存のロジック)
        const userResult = fetchUserReservationsFromSpreadsheet(userId, targetYear, targetMonth); // 既存の予約取得ロジックを別関数に切り出し
        myReservedDates = userResult.myReservedDates || [];
        myAttendedDates = userResult.myAttendedDates || [];

        // 💡 取得後、10分間キャッシュに保存 (300秒)
        userCache.put(userCacheKey, JSON.stringify(userResult), 300); 
        Logger.log(`ユーザー予約情報をスプレッドシートから取得し、キャッシュしました。`);
    }

    // 3. データを結合して返却
    return { 
        success: true, 
        myReservedDates: myReservedDates, 
        myAttendedDates: myAttendedDates 
    };
}

function fetchUserReservationsFromSpreadsheet(userId, targetYear, targetMonth) {

    const resSheet = SPREADSHEET.getSheetByName(SHEET_NAME_RESERVATIONS);
    const ssTimezone = SPREADSHEET.getSpreadsheetTimeZone();
    
    const myReservedDates = [];
    const myAttendedDates = [];

    if (!resSheet) {
        Logger.log('予約シートが見つかりません。');
        return { myReservedDates, myAttendedDates };
    }
    
    const allReservationData = resSheet.getDataRange().getValues();
    const HEADER_ROW_COUNT = 1; // ヘッダー行をスキップ

    for (let i = HEADER_ROW_COUNT; i < allReservationData.length; i++) {
        const row = allReservationData[i];
        
        // 1. ユーザーIDの一致を確認
        if (row[RES_COL_USER_ID] === userId) {
            const status = row[RES_COL_STATUS];
            const dateRaw = row[RES_COL_DATE];
            const startTimeRaw = row[RES_COL_START_TIME]; // 開始時刻を取得
            const reservationId = row[RES_COL_RESERVATION_ID];
            const cancellableUntilRaw = row[RES_COL_CANCELLABLE_UNTIL];
            
            // 日付データが有効なDateオブジェクトであることを確認
            if (!dateRaw || !(dateRaw instanceof Date)) continue;

            const date = new Date(dateRaw);

            // 対象月でなければスキップ
            if (date.getFullYear() !== targetYear || date.getMonth() + 1 !== targetMonth) continue;

            // 日付を 'YYYY-MM-DD' 形式の文字列に変換（LIFF側のカレンダー描画キーに合わせる）
            const dateString = Utilities.formatDate(dateRaw, ssTimezone, 'yyyy-MM-dd');

            const startTimeStr = (startTimeRaw instanceof Date) 
                ? Utilities.formatDate(startTimeRaw, ssTimezone, 'HH:mm') 
                : String(startTimeRaw).trim();
            
            const fullDateKey = `${dateString} ${startTimeStr}`;

            if (cancellableUntilRaw instanceof Date || typeof cancellableUntilRaw === 'number') {
              cancellableUntil = new Date(cancellableUntilRaw); 
            } else if (typeof cancellableUntilRaw === 'string') {
              cancellableUntil = new Date(cancellableUntilRaw.replace(/\//g, '-'));
            } else {
              Logger.log(`Error: Invalid cancellableUntil value: ${cancellableUntilRaw}`);
              return { success: false, message: "予約データの形式が不正です。" };
            }

            const cancellableUntilString = Utilities.formatDate(cancellableUntilRaw, ssTimezone, 'yyyy-MM-dd HH:mm');
            
            // 2. ステータスに基づきリストに分類
            if (status === '確定') {
              let myReservedDateObj = {};
              myReservedDateObj[fullDateKey] = {
                reservationId: reservationId,
                cancellableUntil: cancellableUntilString
              };
              myReservedDates.push(myReservedDateObj);
            } else if (status === '受講済み') {
                myAttendedDates.push(fullDateKey);
            }
        }
    }
    return { myReservedDates, myAttendedDates };
}

// ====================================
// 予約画面用ロジック (予約実行)
// ====================================

/**
 * 予約を実行し、Reservationsシートに記録し、reservationsListの残席数を減らす。
 * @param {Object} params - { userId, lessonId, date, time, className }
 * @returns {Object} { success: boolean, message: string }
 */
function makeReservation(params) {
  const lock = LockService.getScriptLock();
  lock.tryLock(30000); // 30秒間ロックを試みる

  if (lock.hasLock()) {
    try {    
      const { userId, lessonId, date, time, className} = params;
      
      const resSheet = SPREADSHEET.getSheetByName(SHEET_NAME_RESERVATIONS);
      const listSheet = SPREADSHEET.getSheetByName(SHEET_NAME_RESERVATIONS_LIST);
      const usersSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USERS);

      if (!resSheet || !listSheet || !usersSheet) {
          return { success: false, message: '必要なシートが見つかりません (reservations, reservationsList, users)。' };
      }
      
      // ユーザーの予約シート全体を取得
      const allReservations = resSheet.getDataRange().getValues().slice(1);
      const listData = listSheet.getDataRange().getValues();
      const userRow = usersSheet.getDataRange().getValues().slice(1).find(row => row[0] === userId);

      const ssTimezone = SPREADSHEET.getSpreadsheetTimeZone();

      // 1. ユーザー情報の取得と予約回数上限チェック
      const currentDateTime = new Date();
      // 予約しようとしている日付から、その月の「年-月」キーを生成
      // Dateオブジェクトに渡すために 'YYYY-MM-DD' のハイフンをスラッシュに置換
      const targetDate = new Date(date.replace(/-/g, '/'));
      const targetYearMonth = Utilities.formatDate(targetDate, ssTimezone, 'yyyy-MM');


      if (!userRow) {
          return { success: false, message: 'ユーザー情報が見つかりません。' };
      }
      const userClassName = userRow[2];           // C列: ClassName
      const reserverName = userRow ? userRow[1] : "不明"; // B列: DisplayNameを想定
      const userLimit = parseInt(userRow[3], 10);
      const limitNumberIntThisMonth = parseInt(userRow[4], 10);
      const limitNumberIntNextMonth = parseInt(userRow[5], 10);

      if (userClassName !== className) {
          return { success: false, message: 'この予約クラスは、あなたの登録クラスと異なります。' };
      }

      //予約の重複チェック
      const isDuplicateReservation = allReservations.some(row => {
        if (row[RES_COL_USER_ID] !== userId) return false;

        if (row[RES_COL_STATUS] !== '確定' &&
            row[RES_COL_STATUS] !== '受講済み') {
          return false;
        }

        const reservationDate = Utilities.formatDate(
          row[RES_COL_DATE],
          ssTimezone,
          'yyyy-MM-dd'
        );

        const reservationTimeRaw = row[RES_COL_START_TIME];

        const reservationTime = reservationTimeRaw instanceof Date
            ? Utilities.formatDate(reservationTimeRaw, ssTimezone, "HH:mm")
            : String(reservationTimeRaw).trim();

        return reservationDate === date &&
              reservationTime === time;
      });

      if (isDuplicateReservation) {
        return { success: false, message: 'この日時は既に予約済みです。'};
      }

      const currentReservations = allReservations
          .filter(row => {
              const reservationDateRaw = row[RES_COL_DATE]; // 予約日のデータ (Dateオブジェクト)
              const reservationUserId = row[RES_COL_USER_ID]; // ユーザーIDの列
              const reservationStatus = row[RES_COL_STATUS]; // ステータスの列

              // 1. ユーザーIDが一致しない場合は、除外
              if (reservationUserId !== userId) {
                  return false;
              }

              if (reservationStatus !== '確定' && reservationStatus !== '受講済み') {
                  return false;
              }

              // 2. 予約日が存在し、Dateオブジェクトであることを確認
              if (!reservationDateRaw || !(reservationDateRaw instanceof Date)) {
                  return false; 
              }
              
              // 3. 予約日の「年-月」を抽出し、予約対象月と比較
              const reservationYearMonth = Utilities.formatDate(reservationDateRaw, ssTimezone, 'yyyy-MM');
              
              return reservationYearMonth === targetYearMonth;
          }).length;
      
      const targetMonth = Utilities.formatDate(currentDateTime, ssTimezone, 'MM');
      const reservationMonth = Utilities.formatDate(targetDate, ssTimezone, 'MM');
      const checkUpperLimit = targetMonth === reservationMonth ? limitNumberIntThisMonth : limitNumberIntNextMonth;

      if (currentReservations >= checkUpperLimit) {
          // ユーザーにどの月の上限に達したかを明確に伝える
          const targetMonthDisplay = Utilities.formatDate(targetDate, ssTimezone, 'M月');
          return { success: false, message: `${targetMonthDisplay}分の予約上限回数（${checkUpperLimit}回）に達しています。既に${currentReservations}回予約済みです。` };
      }

      // 2. 予約枠リストの残席チェックと更新
      const dateStr = date; // YYYY-MM-DD
      const startTimeStr = time; // HH:mm
      const targetKey = `${lessonId}_${dateStr}_${startTimeStr}`; 

      let targetListRowIndex = -1; // 1-basedのシート行番号
      let currentRemainingCapacity = 0;
      let endTimeStrList = '';

      // メモリ内で該当行を検索
      for (let i = 1; i < listData.length; i++) { 
          const listRow = listData[i];
          
          // 予約枠のキーを再構成
          const rowDateStr = Utilities.formatDate(new Date(listRow[RES_LIST_COL_DATE]), ssTimezone, 'yyyy-MM-dd');
          const rowStartTimeRaw = listRow[RES_LIST_COL_START_TIME];
          let rowStartTime = (rowStartTimeRaw instanceof Date) 
              ? Utilities.formatDate(rowStartTimeRaw, ssTimezone, 'HH:mm') 
              : String(rowStartTimeRaw).trim();

          const rowKey = `${listRow[RES_LIST_COL_LESSON_ID]}_${rowDateStr}_${rowStartTime}`; 
          
          if (rowKey === targetKey) {
              targetListRowIndex = i + 1; // 1-based
              currentRemainingCapacity = parseInt(listRow[RES_LIST_COL_REMAINING_CAPACITY], 10);

              const endTimeRaw = listRow[RES_LIST_COL_END_TIME];
              endTimeStrList = (endTimeRaw instanceof Date) 
                  ? Utilities.formatDate(endTimeRaw, ssTimezone, 'HH:mm') 
                  : String(endTimeRaw).trim();

              break;
          }
      }

      if (targetListRowIndex === -1) {
          return { success: false, message: '指定された予約枠が見つかりません。' };
      }
      
      // 3. 残席確認
      if (currentRemainingCapacity <= 0) {
          return { success: false, message: '残席がありません。' };
      }

      // 当日予約の場合、開始時刻が現在時刻を過ぎていないかチェック
      const targetDateStr = date.replace(/-/g, '/'); // Dateコンストラクタ用に 'YYYY/MM/DD' に変換

      // 予約開始日時 (YYYY/MM/DD HH:mm) のDateオブジェクトを作成
      const reservationEndDateTime = new Date(`${targetDateStr} ${endTimeStrList}`);

      // スプレッドシートのタイムゾーンを考慮して現在時刻を取得
      // GASの実行環境の現在時刻と、スプレッドシートのタイムゾーンに合わせて調整した予約時刻を比較
      // 注意: new Date() はGASのサーバー時刻を使用するため、厳密にはタイムゾーン調整が必要です。
      // ここではシンプルに、予約開始時刻が現在時刻よりも過去でないかを確認します。
      if (reservationEndDateTime.getTime() <= currentDateTime.getTime()) {
        return { success: false, message: 'この予約枠はすでに終了しています。他の時間帯をお選びください。' };
      }

      // 4. 残席数の更新 (decrement)
      const newRemainingCapacity = currentRemainingCapacity - 1;
      listSheet.getRange(targetListRowIndex, RES_LIST_COL_REMAINING_CAPACITY + 1).setValue(newRemainingCapacity); 

      // 5. 予約レコードの作成と書き込み (reservationsシート)
      const reservationId = Utilities.getUuid(); 
      const now = new Date();

      const startDateTimeStr = `${date.replace(/-/g, '/')} ${time}`; 
      const startDateTime = new Date(startDateTimeStr);

      const endDateTimeStr = `${date.replace(/-/g, '/')} ${endTimeStrList}`;
      const endDateTime = new Date(endDateTimeStr);
      
      const durationMinutes = Math.round((endDateTime.getTime() - startDateTime.getTime()) / (1000 * 60));

      const cancellableUntil = new Date(startDateTime.getTime() - (24 * 60 * 60 * 1000));
      const cancellableUntilStr = Utilities.formatDate(cancellableUntil, ssTimezone, 'yyyy/MM/dd HH:mm');

      const newRow = [];
      newRow[RES_COL_RESERVATION_ID] = reservationId;
      newRow[RES_COL_DATE] = new Date(date.replace(/-/g, '/'));
      newRow[RES_COL_START_TIME] = time;
      newRow[RES_COL_LESSON_ID] = lessonId;
      newRow[RES_COL_USER_ID] = userId;
      newRow[RES_COL_RESERVER_NAME] = reserverName;
      newRow[RES_COL_SELECTED_CLASS_NAME] = className;
      newRow[RES_COL_DURATION_MINUTES] = durationMinutes; // 計算値
      newRow[RES_COL_END_TIME] = endTimeStrList;         // 予約枠リストから取得した終了時間
      newRow[RES_COL_CANCELLABLE_UNTIL] = cancellableUntilStr;
      newRow[RES_COL_RESERVATION_DATETIME] = now;
      newRow[RES_COL_OPERATION] = '確定';
      newRow[RES_COL_STATUS] = '確定';
      newRow[RES_COL_MEMO] = '';
      newRow[RES_COL_CALENDAR_EVENT_ID] = ''; // 処理が重いので、バッチ処理でカレンダー連携で対応するので、ここは空で登録。

      resSheet.appendRow(newRow);
      // 予約一覧のキャッシュを削除 (既存)
      deleteReservationsCache(userId, targetYearMonth);
      // Cloudflareへの同期処理実行
      // syncCapacityToWorkers();
      // syncUserFullData(userId);

      const reservations = getAllReservationsForUser(userId);
      const userInfoFull = {
        data: {
          userId: userId,
          displayName: reserverName,
          className: userClassName,
          upperLimitNumber: userLimit,
          upperLimitNumberThisMonth: limitNumberIntThisMonth,
          upperLimitNumberNextMonth: limitNumberIntNextMonth
        },
        myReservedDates: reservations.myReservedDates,
        myAttendedDates: reservations.myAttendedDates
      };

      const capacityData = getCapacityData();
      syncReservationToWorkers(userId, userInfoFull, capacityData);

      // 最終的なメッセージを組み立てて返す
      return { 
        success: true,
        message: `予約が正常に完了しました。`,
        reservationDateTime: `${date.replace(/-/g, '/')} ${time}～`,
        cancellableUntil: cancellableUntilStr,
        userInfo: userInfoFull,
        capacityData: capacityData
      };

    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "現在、他の処理が実行中です。しばらくしてから再度お試しください。" };
  }
}

/**
 * 特定のユーザーの予約リストキャッシュを削除する
 * @param {string} userId - LINE User ID
 */
function deleteReservationsCache(userId, monthKey) {
    const userCache = CacheService.getUserCache();
    const userCacheKey = "USER_RESERVATIONS_" + userId + "_" + monthKey;
    // ユーザキャッシュを削除
    userCache.remove(userCacheKey);

    const capacityCache = CacheService.getPublicCache();
    const capacityCacheKey = 'CAPACITY_' + monthKey;
    capacityCache.remove(capacityCacheKey);
}

// =========================================================================
// キャンセル処理
// =========================================================================
/**
 * 予約をキャンセル済みに更新する
 * @param {object} params - リクエストパラメータ
 */
function handleCancelReservation(params) {
  const lock = LockService.getScriptLock();
  lock.tryLock(30000);

  if (lock.hasLock()) {
    try {

      const ssTimezone = SPREADSHEET.getSpreadsheetTimeZone();
      const { userId, reservationId, admin = false } = params;

      if (!userId || !reservationId) {
        return { success: false, message: "必須パラメータが不足しています。" };
      }

      const resSheet = SPREADSHEET.getSheetByName(SHEET_NAME_RESERVATIONS);
      const listSheet = SPREADSHEET.getSheetByName(SHEET_NAME_RESERVATIONS_LIST);

      if (!resSheet || !listSheet) {
        return { success: false, message: `シート '${SHEET_NAME_RESERVATIONS}' が見つかりません。` };
      }

      // 1. 予約レコードの高速検索とデータ取得 ---
      const allReservationData = resSheet.getDataRange().getValues();
      const listData = listSheet.getDataRange().getValues(); // シート全体を読み込み

      const HEADER_ROW_COUNT = 1;
      let row = -1; // 予約シートの行番号 (1-based)
      let values = null;
      
      // 予約ID (A列 = インデックス0) でメモリ内を検索
      for (let i = HEADER_ROW_COUNT; i < allReservationData.length; i++) {
          if (allReservationData[i][0] === reservationId) {
              row = i + 1; // 1-basedのシート行番号
              values = allReservationData[i]; // 該当行の全データ (0-based array)
              break;
          }
      }

      if (row === -1) {
          return { success: false, message: "指定された予約が見つかりません。" };
      }

      // 予約シートのインデックス (0-based)
      const COL_DATE_IDX = RES_COL_DATE;
      const COL_START_TIME_IDX = RES_COL_START_TIME;
      const COL_LESSON_ID_IDX = RES_COL_LESSON_ID;
      const COL_USER_ID_IDX = RES_COL_USER_ID;
      const COL_CANCELLABLE_UNTIL_IDX = RES_COL_CANCELLABLE_UNTIL;
      const COL_STATUS_IDX = RES_COL_STATUS;
      const COL_OPERATION_IDX = RES_COL_OPERATION;

      // 1. ユーザーIDの確認
      if (values[COL_USER_ID_IDX] !== userId) {
          return { success: false, message: "この予約のキャンセル権限がありません。" };
      }

      // 2. ステータスの確認
      if (values[COL_STATUS_IDX] !== '確定') {
          return { success: false, message: "この予約は既にキャンセル済みか、確定していません。" };
      }

      // 3. キャンセル期限の確認 (I列) と安全なパース
      let cancellableUntilRaw = values[COL_CANCELLABLE_UNTIL_IDX];
      let cancellableUntil;

      if (cancellableUntilRaw instanceof Date || typeof cancellableUntilRaw === 'number') {
          cancellableUntil = new Date(cancellableUntilRaw); 
      } else if (typeof cancellableUntilRaw === 'string') {
          cancellableUntil = new Date(cancellableUntilRaw.replace(/\//g, '-'));
      } else {
          Logger.log(`Error: Invalid cancellableUntil value: ${cancellableUntilRaw}`);
          return { success: false, message: "予約データの形式が不正です。" };
      }

      const now = new Date();
      if (!admin && now.getTime() >= cancellableUntil.getTime()) {
          return { success: false, message: "キャンセル期限が過ぎています。" };
      }
      
      // --- 4. キャンセル実行 (reservationsシートの更新: K, L, M列をバッチ更新) ---
      const cancelDateTimeStr = Utilities.formatDate(now, ssTimezone, 'yyyy/MM/dd HH:mm:ss');

      resSheet.getRange(row, COL_OPERATION_IDX + 1, 1, 3).setValues([[
          'キャンセル',           // K: 操作
          'キャンセル済み',       // L: ステータス
          cancelDateTimeStr       // M: メモ (キャンセル日時として記録)
      ]]);

      // --- 5. 残席リスト (reservationsList) の残席数回復 (改善策 B + 定員制御) ---
      
      const date = values[COL_DATE_IDX];
      const startTime = values[COL_START_TIME_IDX];
      const lessonId = values[COL_LESSON_ID_IDX];
      
      // ターゲットキーの生成
      const dateStr = Utilities.formatDate(date, ssTimezone, 'yyyy-MM-dd');
      const startTimeStr = (startTime instanceof Date) 
          ? Utilities.formatDate(startTime, ssTimezone, 'HH:mm') 
          : String(startTime).trim();
      const targetKey = `${lessonId}_${dateStr}_${startTimeStr}`; 

      let listUpdated = false;

      // メモリ内で該当行を検索
      for (let i = 1; i < listData.length; i++) { 
          const listRow = listData[i];
          
          // reservationsListの行データからキーを再構成
          const rowDateStr = Utilities.formatDate(new Date(listRow[RES_LIST_COL_DATE]), ssTimezone, 'yyyy-MM-dd');
          const rowStartTimeRaw = listRow[RES_LIST_COL_START_TIME];
          let rowStartTime = (rowStartTimeRaw instanceof Date) 
              ? Utilities.formatDate(rowStartTimeRaw, ssTimezone, 'HH:mm') 
              : String(rowStartTimeRaw).trim();

          const rowKey = `${listRow[RES_LIST_COL_LESSON_ID]}_${rowDateStr}_${rowStartTime}`; // listRow[5] is LessonID
          
          if (rowKey === targetKey) {
              const sheetRowNumber = i + 1; 
              
              const capacity = parseInt(listRow[RES_LIST_COL_CAPACITY], 10);        // G列: 定員 (Index 6)
              const remainingCapacity = parseInt(listRow[RES_LIST_COL_REMAINING_CAPACITY], 10); // H列: 残席数 (Index 7)

              let newRemainingCapacity = remainingCapacity + 1;

              // 【定員制御】新しい残席数が定員を超えないように制御する
              if (newRemainingCapacity > capacity) {
                  newRemainingCapacity = capacity;
                  Logger.log(`Warning: Remaining capacity for ${targetKey} exceeded capacity (${capacity}). Capped at capacity.`);
              }

              // 残席数を更新 (H列はシート上で8番目)
              listSheet.getRange(sheetRowNumber, RES_LIST_COL_REMAINING_CAPACITY + 1).setValue(newRemainingCapacity); 
              
              listUpdated = true;
              break;
          }
      }
      
      if (!listUpdated) {
          Logger.log(`Warning: Failed to update reservationsList for key: ${targetKey}`);
      }

      // キャッシュ削除
      const monthKey = Utilities.formatDate(date, ssTimezone, 'yyyy-MM');
      deleteReservationsCache(userId, monthKey);
      const cancelDateTime = `${dateStr.replace(/-/g, '/')} ${startTimeStr}～`;
      // Cloudflareへの同期処理実行
      // syncCapacityToWorkers();
      // syncUserFullData(userId);

      const usersSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USERS);
      const userRow = usersSheet.getDataRange().getValues().slice(1).find(row => row[0] === userId);
      if (!userRow) {
          return { success: false, message: 'ユーザー情報が見つかりません。' };
      }
      const userClassName = userRow[2];           // C列: ClassName
      const reserverName = userRow ? userRow[1] : "不明"; // B列: DisplayNameを想定
      const userLimit = parseInt(userRow[3], 10);
      const limitNumberIntThisMonth = parseInt(userRow[4], 10);
      const limitNumberIntNextMonth = parseInt(userRow[5], 10);
      
      const reservations = getAllReservationsForUser(userId);
      const userInfoFull = {
        data: {
          userId: userId,
          displayName: reserverName,
          className: userClassName,
          upperLimitNumber: userLimit,
          upperLimitNumberThisMonth: limitNumberIntThisMonth,
          upperLimitNumberNextMonth: limitNumberIntNextMonth
        },
        myReservedDates: reservations.myReservedDates,
        myAttendedDates: reservations.myAttendedDates
      };

      const capacityData = getCapacityData();
      syncReservationToWorkers(userId, userInfoFull, capacityData);

      return { 
        success: true,
        message: "予約をキャンセルしました。",
        cancelDateTime: cancelDateTime,
        userInfo: userInfoFull,
        capacityData: capacityData
      };
    } catch (e) {
      return { success: false, message: `キャンセル処理中にエラーが発生しました。${e.toString()}` };
    } finally {
      lock.releaseLock(); // 忘れずにロックを解除
    }
  } else {
    return { success: false, message: "現在、システムが混み合っています。しばらくしてから再度お試しください。" };
  }
}

function verifyToken(accessToken) {
  // 1. トークンの検証とユーザーIDの取得
  const verifyResponse = UrlFetchApp.fetch(LINE_VERIFY_URL + "?access_token=" + accessToken);
  const verifyData = JSON.parse(verifyResponse.getContentText());
  
  // 検証に失敗した場合 (例: トークンが無効・期限切れ)
  if (verifyData.error || !verifyData.client_id) {
      throw new Error("Token verification failed or expired.");
  }

  // client_idの厳格な検証
  if (verifyData.client_id !== LIFF_CLIENT_ID) {
      // トークンが別のアプリのものである場合、不正アクセスと見なす
      throw new Error("Token client_id mismatch. Invalid application source.");
  }
  return true;
}

