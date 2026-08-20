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
// retryシート
const SHEET_NAME_RETRY = 'retry';

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

const RESERVATION_STATUS = {
  CONFIRMED: "確定",
  CANCELED: "キャンセル"
};

// reservationシート 列インデックス定数
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
const RES_COL_USAGE_TYPE = 15;
const RES_COL_TICKET_ID = 16;

// reservationListシート　列インデックス定数
const RES_LIST_COL_DATE = 0;
const RES_LIST_COL_DAY_OF_WEEK = 1;
const RES_LIST_COL_START_TIME = 2;
const RES_LIST_COL_END_TIME = 3;
const RES_LIST_COL_CLASS_NAME = 4;
const RES_LIST_COL_LESSON_ID = 5;
const RES_LIST_COL_CAPACITY = 6;
const RES_LIST_COL_REMAINING_CAPACITY = 7;

// userTicketsシート 列インデックス定数
const USER_TICKETS_COL_TICKET_ID = 0;
const USER_TICKETS_COL_USER_ID = 1;
const USER_TICKETS_COL_USER_NAME = 2;
const USER_TICKETS_COL_PURCHASE_NUM = 3;
const USER_TICKETS_COL_LESSON_FEE = 4;
const USER_TICKETS_COL_REMAINING_NUM = 5;
const USER_TICKETS_COL_PURCHASE_DATE = 6;
const USER_TICKETS_COL_EXPIRE_DATE = 7;
const USER_TICKETS_COL_STATUS = 8;
const USER_TICKETS_COL_REGISTER_DATE = 9;
const USER_TICKETS_COL_UPDATE_DATE = 10;
const USER_TICKETS_COL_REMARKS = 11;

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
        verifyToken(params.accessToken);
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
          const usageType = row[RES_COL_USAGE_TYPE];
            
            // 2. ステータスに基づきリストに分類
            if (status === '確定') {
              let myReservedDateObj = {};
              myReservedDateObj[fullDateKey] = {
                reservationId: reservationId,
                cancellableUntil: cancellableUntilString,
                usageType: usageType
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
      const { userId, lessonId, date, time, className } = params;
      
      const resSheet = SPREADSHEET.getSheetByName(SHEET_NAME_RESERVATIONS);
      const listSheet = SPREADSHEET.getSheetByName(SHEET_NAME_RESERVATIONS_LIST);
      const usersSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USERS);
      const userMonthlySubscriptionsSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USER_MONTHLY_SUBSCRIPTIONS);
      const userTicketsSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USER_TICKETS);

      if (!resSheet || !listSheet || !usersSheet || !userTicketsSheet || !userMonthlySubscriptionsSheet) {
        return { success: false, message: '必要なシートが見つかりません (reservations, reservationsList, users, userTickets, userMonthlySubscriptionsSheet)。' };
      }
      
      // ユーザーの予約シート全体を取得
      const allReservations = resSheet.getDataRange().getValues().slice(1);
      const listData = listSheet.getDataRange().getValues();
      const userRow = usersSheet.getDataRange().getValues().slice(1).find(row => row[0] === userId);
      const userMontlyRow = userMonthlySubscriptionsSheet.getDataRange().getValues().slice(1).find(row => row[0] === userId);
      const userTicketsRows = userTicketsSheet.getDataRange().getValues();
      const userTicketsData = userTicketsRows.slice(1).filter(ticket => ticket[USER_TICKETS_COL_USER_ID] === userId);

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
      let userLimit = 0;
      let limitNumberIntThisMonth = 0;
      let limitNumberIntNextMonth = 0;
      
      if (userMontlyRow) {
        userLimit = parseInt(userMontlyRow[2], 10);
        limitNumberIntThisMonth = parseInt(userMontlyRow[3], 10);
        limitNumberIntNextMonth = parseInt(userMontlyRow[4], 10);
      }


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
        return { success: false, message: 'この日時は既に予約済みです。' };
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
      
      const now = new Date();
      const dateString = date;
      const userTicketsRowsWithIndex = userTicketsRows.map((row, index) => ({
        rowIndex: index,
        rowData: row
      }));
      const validUserTicketsRows = userTicketsRowsWithIndex
        .filter(ticket =>
        ticket.rowData[USER_TICKETS_COL_USER_ID] === userId &&
          Utilities.formatDate(
            ticket.rowData[USER_TICKETS_COL_EXPIRE_DATE], SPREADSHEET.getSpreadsheetTimeZone(), 'yyyy-MM-dd') >= dateString
      ).sort((a, b) => {
        // 有効期限が近い順
        const expireDiff =
        a.rowData[USER_TICKETS_COL_EXPIRE_DATE] -
        b.rowData[USER_TICKETS_COL_EXPIRE_DATE];

        if (expireDiff !== 0) {
          return expireDiff;
        }
        // 同じ有効期限なら登録順（シートの行番号）
        return a.rowIndex - b.rowIndex;
      });
      let remainingNumberTotal = validUserTicketsRows.length === 0 ? 0 : validUserTicketsRows.map(ticket => ticket.rowData[USER_TICKETS_COL_REMAINING_NUM]).reduce((total, num) => total + num, 0);
      // 消費するチケットを特定（条件：残数が1以上のチケットを消費）
      let targetTicket = validUserTicketsRows.find(ticket => ticket.rowData[USER_TICKETS_COL_REMAINING_NUM] > 0);

      const checkUpperLimit = targetMonth === reservationMonth ? limitNumberIntThisMonth : limitNumberIntNextMonth;
      if (userMontlyRow) {
        const checkRemainingNumber = remainingNumberTotal;
        const checkTotalNumber = checkUpperLimit + checkRemainingNumber;
        if (currentReservations >= checkTotalNumber) {
          // ユーザーにどの月の上限に達したかを明確に伝える
          const targetMonthDisplay = Utilities.formatDate(targetDate, ssTimezone, 'M月');
          return { success: false, message: `${targetMonthDisplay}分の予約上限回数（${checkTotalNumber}回）に達しています。既に${currentReservations}回予約済みです。` };
        }
      } else {
        if (!targetTicket) {
          return { success: false, message: '利用可能なチケットがありません。' };
        }
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

      // targetTicket.rowData（配列データ）とシートの値の両方を更新
      let usageType = '';
      if (userMontlyRow && currentReservations < checkUpperLimit) {
        // 月謝枠を利用
        usageType = 'monthly';
      } else {
        // チケット利用
        usageType = 'ticket';

        if (!targetTicket) {
          return {
            success: false,
            message: '利用可能なチケットがありません。'
          };
        }
        remainingNumberTotal--;
        targetTicket.rowData[USER_TICKETS_COL_REMAINING_NUM]--;
        const newRemainingNum = targetTicket.rowData[USER_TICKETS_COL_REMAINING_NUM];
        userTicketsSheet.getRange(targetTicket.rowIndex + 1, USER_TICKETS_COL_REMAINING_NUM + 1).setValue(newRemainingNum);
      }

      // 残席数の更新 (decrement)
      const newRemainingCapacity = currentRemainingCapacity - 1;
      listSheet.getRange(targetListRowIndex, RES_LIST_COL_REMAINING_CAPACITY + 1).setValue(newRemainingCapacity);

      // 5. 予約レコードの作成と書き込み (reservationsシート)
      const reservationId = Utilities.getUuid();

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
      newRow[RES_COL_USAGE_TYPE] = usageType === "ticket" ? "チケット" : "月謝";
      newRow[RES_COL_TICKET_ID] = usageType === 'ticket' ? targetTicket.rowData[USER_TICKETS_COL_TICKET_ID] : '';

      resSheet.appendRow(newRow);
      // 予約一覧のキャッシュを削除 (既存)
      deleteReservationsCache(userId, targetYearMonth);

      const reservations = getAllReservationsForUser(userId);
      const userInfoFull = {
        data: {
          userId: userId,
          displayName: reserverName,
          className: userClassName,
          upperLimitNumber: userLimit,
          upperLimitNumberThisMonth: limitNumberIntThisMonth,
          upperLimitNumberNextMonth: limitNumberIntNextMonth,
          ticketInfo: {
            dispInfo: validUserTicketsRows.map(ticket => ({
              remainingNumber: ticket.rowData[USER_TICKETS_COL_REMAINING_NUM],
              expirationDate: Utilities.formatDate(ticket.rowData[USER_TICKETS_COL_EXPIRE_DATE], ssTimezone, 'yyyy-MM-dd'),
              purchaseNumber: ticket.rowData[USER_TICKETS_COL_PURCHASE_NUM]
            })),
            remainingNumberTotal: remainingNumberTotal,
            // チケット購入履歴があるか
            purchaseHistory: userTicketsData.length !== 0
          }
        },
        myReservedDates: reservations.myReservedDates,
        myAttendedDates: reservations.myAttendedDates
      };

      const capacityData = getCapacityData();
      try {
        syncReservationToWorkers(userId, userInfoFull, capacityData);
      } catch (e) {
        retryProcessLog({
          retryId: Utilities.getUuid(),
          processType: PROCESS_TYPE.SYNC_WORKERS,
          args: JSON.stringify({ userId: userId}),
          retryCount: 0,
          status: RETRY_STATUS.WAITING,
          errorContents: e.message,
          registerDate: new Date(),
          finalExecuteDate: "",
        })
        Logger.log(`Workers同期エラー: ${e.stack || e}`);
      }
      // 最終的なメッセージを組み立てて返す
      return {
        success: true,
        message: `予約が正常に完了しました。`,
        reservationDateTime: `${date.replace(/-/g, '/')} ${time}～`,
        cancellableUntil: cancellableUntilStr,
        userInfo: userInfoFull,
        capacityData: capacityData
      };
    } catch(e) {
      Logger.log(e);
      return {
        success: false,
        message: '予約処理中にエラーが発生しました。'
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
      const { userId, reservationId, admin = false, attendedCancel = false } = params;
      const allowAttendedCancel = admin && attendedCancel;

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

      const usersSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USERS);
      const userRow = usersSheet.getDataRange().getValues().slice(1).find(row => row[0] === userId);
      if (!userRow) {
          return { success: false, message: 'ユーザー情報が見つかりません。' };
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
      if (!admin && values[COL_STATUS_IDX] !== '確定') {
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
      
      // チケット管理シートからユーザーのチケット情報を取得      
      const userTicketsSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USER_TICKETS);
      const userTicketsRows = userTicketsSheet.getDataRange().getValues();
      const userTicketsData = userTicketsRows.slice(1).filter(ticket => ticket[USER_TICKETS_COL_USER_ID] === userId);
      
      const userTicketsRowsWithIndex = userTicketsRows.map((row, index) => ({
        rowIndex: index,
        rowData: row
      }));
      
      const reservationsRowsWithIndex = allReservationData.map((row, index) => ({
        rowIndex: index,
        rowData: row
      }));
      
      const date = values[COL_DATE_IDX];
      const dateStr = Utilities.formatDate(date, ssTimezone, 'yyyy-MM-dd');
      // キャンセルでは残数「0」も検索対象
      const validUserTicketsRows = userTicketsRowsWithIndex
        .filter(ticket => ticket.rowData[USER_TICKETS_COL_USER_ID] === userId
          // 有効期限がキャンセル日以降のチケット
          && Utilities.formatDate(ticket.rowData[USER_TICKETS_COL_EXPIRE_DATE], SPREADSHEET.getSpreadsheetTimeZone(), 'yyyy-MM-dd') >= dateStr
          // 一度も消費されていないチケットは対象外
          && ticket.rowData[USER_TICKETS_COL_PURCHASE_NUM] !== ticket.rowData[USER_TICKETS_COL_REMAINING_NUM]
      ).sort((ticket1, ticket2) => {
        // 有効期限が近い順
        const expireDiff =
        ticket1.rowData[USER_TICKETS_COL_EXPIRE_DATE] -
        ticket2.rowData[USER_TICKETS_COL_EXPIRE_DATE];

        if (expireDiff !== 0) {
          return expireDiff;
        }
        // 同じ有効期限なら登録順（シートの行番号）
        return ticket1.rowIndex - ticket2.rowIndex;
      })
      .map(ticket => {
        const ticketId = ticket.rowData[USER_TICKETS_COL_TICKET_ID];

        const reservations = reservationsRowsWithIndex.filter(reservation =>
          reservation.rowData[RES_COL_TICKET_ID] === ticketId &&
          (
            reservation.rowData[RES_COL_STATUS] === "確定" ||
            (allowAttendedCancel && reservation.rowData[RES_COL_STATUS] === "受講済み")
          )
        )
        .sort((reservation1, reservation2) => {
        // 予約日が新しい順
        const date1 = Utilities.formatDate(reservation1.rowData[RES_COL_DATE], ssTimezone, 'yyyy-MM-dd');
        const date2 = Utilities.formatDate(reservation2.rowData[RES_COL_DATE], ssTimezone, 'yyyy-MM-dd');
        if (date1 !== date2) {
          return date2.localeCompare(date1);
        }

        // 同じ日なら開始時刻が遅い順
        const time1 = reservation1.rowData[RES_COL_START_TIME] instanceof Date
          ? Utilities.formatDate(
              reservation1.rowData[RES_COL_START_TIME],
              ssTimezone,
              'HH:mm'
            )
          : String(reservation1.rowData[RES_COL_START_TIME]).trim();

        const time2 = reservation2.rowData[RES_COL_START_TIME] instanceof Date
          ? Utilities.formatDate(
            reservation2.rowData[RES_COL_START_TIME],
            ssTimezone,
              'HH:mm'
            )
          : String(reservation2.rowData[RES_COL_START_TIME]).trim();

        return time2.localeCompare(time1);
        });

        if (reservations.length === 0) {
          throw new Error(`チケットに紐づく予約が見つかりません。ticketId: ${ticketId}`);
        }

        return {
          ticket: ticket,
          reservations: reservations
        };
      });
      
      // 月稽古管理シートからユーザーの稽古回数を取得
      const userMonthlySubscriptionsSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USER_MONTHLY_SUBSCRIPTIONS);
      const userMontlyRow = userMonthlySubscriptionsSheet.getDataRange().getValues().slice(1).find(row => row[0] === userId);
      
      let hasMonthly = userMontlyRow !== undefined;
      let hasTickets = validUserTicketsRows.length !== 0;
      const reservationTicketId = values[RES_COL_TICKET_ID];
      
      if (hasTickets) {
        let targetTicketIndex = null;
        if (reservationTicketId) {
          targetTicketIndex = validUserTicketsRows.findIndex(cancelInfo => cancelInfo.ticket.rowData[USER_TICKETS_COL_TICKET_ID] === reservationTicketId);
          if (targetTicketIndex === -1) {
            throw new Error(`キャンセル対象のチケットが見つかりません。ticketId: ${reservationTicketId}`);
          }
        }
        if (hasMonthly) {
          //チケット + 月謝の処理
          cancelMonthlyAndTicketConsumption(userTicketsSheet, resSheet, userMontlyRow, validUserTicketsRows, targetTicketIndex, {rowIndex: row, rowData: values})
        } else {
          cancelTicketConsumption(userTicketsSheet, resSheet, validUserTicketsRows, targetTicketIndex);
        }
      }

      // キャンセル実行 (reservationsシートの更新: K, L, M列をバッチ更新) ---
      const cancelDateTimeStr = Utilities.formatDate(now, ssTimezone, 'yyyy/MM/dd HH:mm:ss');

      resSheet.getRange(row, COL_OPERATION_IDX + 1, 1, 3).setValues([[
          'キャンセル',           // K: 操作
          'キャンセル済み',       // L: ステータス
          cancelDateTimeStr       // M: メモ (キャンセル日時として記録)
      ]]);

      // 残席リスト (reservationsList) の残席数回復
      const startTime = values[COL_START_TIME_IDX];
      const lessonId = values[COL_LESSON_ID_IDX];
      
      // ターゲットキーの生成
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

      const userClassName = userRow[2];           // C列: ClassName
      const reserverName = userRow ? userRow[1] : "不明"; // B列: DisplayNameを想定
      let userLimit = 0;
      let limitNumberIntThisMonth = 0;
      let limitNumberIntNextMonth = 0;
      if (userMontlyRow) {
        userLimit = parseInt(userMontlyRow[2], 10);
        limitNumberIntThisMonth = parseInt(userMontlyRow[3], 10);
        limitNumberIntNextMonth = parseInt(userMontlyRow[4], 10);
      }
      
      const reservations = getAllReservationsForUser(userId);
      const availableUserTickets = userTicketsRows.slice(1)
        .filter(row =>
          row[USER_TICKETS_COL_USER_ID] === userId &&
          Utilities.formatDate(row[USER_TICKETS_COL_EXPIRE_DATE], ssTimezone, 'yyyy-MM-dd') >= dateStr
       )
        .sort((ticket1, ticket2) => {
          // 有効期限が近い順
          const expireDiff =
          ticket1[USER_TICKETS_COL_EXPIRE_DATE] -
          ticket2[USER_TICKETS_COL_EXPIRE_DATE];

          if (expireDiff !== 0) {
            return expireDiff;
          }
          // 同じ有効期限なら登録順（シートの行番号）
          return ticket1[USER_TICKETS_COL_REGISTER_DATE] - ticket2[USER_TICKETS_COL_REGISTER_DATE];
      });
      let remainingNumberTotal = availableUserTickets.length === 0 ? 0 : availableUserTickets.map(row => row[USER_TICKETS_COL_REMAINING_NUM]).reduce((total, num) => total + num, 0);
      const userInfoFull = {
        data: {
          userId: userId,
          displayName: reserverName,
          className: userClassName,
          upperLimitNumber: userLimit,
          upperLimitNumberThisMonth: limitNumberIntThisMonth,
          upperLimitNumberNextMonth: limitNumberIntNextMonth,
          ticketInfo: {
            dispInfo: availableUserTickets.map(row => ({
              remainingNumber: row[USER_TICKETS_COL_REMAINING_NUM],
              expirationDate: Utilities.formatDate(row[USER_TICKETS_COL_EXPIRE_DATE], ssTimezone, 'yyyy-MM-dd'),
              purchaseNumber: row[USER_TICKETS_COL_PURCHASE_NUM]
            })),
            remainingNumberTotal: remainingNumberTotal,
            // チケット購入履歴があるか
            purchaseHistory: userTicketsData.length !== 0
          }
        },
        myReservedDates: reservations.myReservedDates,
        myAttendedDates: reservations.myAttendedDates
      };

      const capacityData = getCapacityData();

      try {
        syncReservationToWorkers(userId, userInfoFull, capacityData);
      } catch (e) {
        retryProcessLog({
          retryId: Utilities.getUuid(),
          processType: PROCESS_TYPE.SYNC_WORKERS,
          args: JSON.stringify({ userId: userId}),
          retryCount: 0,
          status: RETRY_STATUS.WAITING,
          errorContents: e.message,
          registerDate: new Date(),
          finalExecuteDate: "",
        })
        Logger.log(`Workers同期エラー: ${e.stack || e}`);
      }

      if (admin) {
        sendLineMessage(userId,
`教室側で下記のご予約をキャンセルいたしました。

・予約日時：${dateStr} ${startTimeStr}～

ご確認のほど、よろしくお願いいたします。`
        );
      }

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
  const verifyResponse = UrlFetchApp.fetch(
    LINE_VERIFY_URL + "?access_token=" + encodeURIComponent(accessToken),
    {
      muteHttpExceptions: true
    }
  );
  const verifyData = JSON.parse(verifyResponse.getContentText());
  
  // 検証に失敗した場合 (例: トークンが無効・期限切れ)
  if (verifyData.error || !verifyData.client_id) {
      throw new Error("TOKEN_EXPIRED");
  }

  // client_idの厳格な検証
  if (verifyData.client_id !== LIFF_CLIENT_ID) {
      // トークンが別のアプリのものである場合、不正アクセスと見なす
      throw new Error("TOKEN_INVALID");
  }
  return true;
}

/**
 * チケット消費分のキャンセル処理
 *
 * キャンセルによって空いたチケットの消費枠を、
 * 後続チケットの予約から有効期限内のものを探して連鎖的にスライドする。
 *
 * 例：
 *
 * A：7/30期限
 *   └ 7/2予約 ← キャンセル
 *
 * B：8/10期限
 *   └ 7/20予約
 *
 * C：8/30期限
 *   └ 7/25予約
 *
 * ↓
 *
 * A ← Bの7/20予約
 * B ← Cの7/25予約
 * C → 残数+1
 *
 * 有効期限を超える予約はスライドせず、
 * その予約は元のチケットに残す。
 */
function cancelTicketConsumption(
  userTicketsSheet,
  reservationsSheet,
  validUserTicketsRows,
  targetTicketIndex,
  skipFinalIncrementTicketId = null
) {
  const ssTimezone = SPREADSHEET.getSpreadsheetTimeZone();

  // ============================================================
  // 現在の「移動先」となるチケット
  //
  // 最初はキャンセル対象チケット。
  // B → A とスライドした後は、Bが次の移動先になる。
  // ============================================================
  let currentTargetIndex = targetTicketIndex;
  let currentTargetTicketData = validUserTicketsRows[currentTargetIndex];
  let currentTargetTicket = currentTargetTicketData.ticket;

  // ============================================================
  // 現在の移動先チケットの有効期限
  // ============================================================
  let currentTargetExpireDateStr = Utilities.formatDate(
    currentTargetTicket.rowData[USER_TICKETS_COL_EXPIRE_DATE],
    ssTimezone,
    'yyyy-MM-dd'
  );

  // ============================================================
  // 後続チケットを順番に確認
  // ============================================================
  for (let sourceTicketIndex = currentTargetIndex + 1; sourceTicketIndex < validUserTicketsRows.length; sourceTicketIndex++) {
    
    const sourceTicketData = validUserTicketsRows[sourceTicketIndex];
    const sourceTicket = sourceTicketData.ticket;
    const reservations = sourceTicketData.reservations;

    // 予約がないチケットはスキップ
    if (!reservations || reservations.length === 0) {
      continue;
    }

    // ==========================================================
    // 後続チケットの予約から、
    // 現在の移動先チケットの有効期限内に収まる予約を探す
    //
    // 例：
    // B：
    //   8/2  → Aの期限外
    //   7/20 → Aの期限内
    //
    // 8/2で終了せず、7/20まで確認する。
    // ==========================================================
    let slideReservation = null;

    for (let reservationIndex = 0; reservationIndex < reservations.length; reservationIndex++) {
      const reservation = reservations[reservationIndex];
      const reservationDateStr = Utilities.formatDate(
        reservation.rowData[RES_COL_DATE],
        ssTimezone,
        'yyyy-MM-dd'
      );

      if (reservationDateStr <= currentTargetExpireDateStr) {
        slideReservation = reservation;
        break;
      }
    }

    // ==========================================================
    // スライド可能な予約がなければ、
    // この後続チケットから先を確認する
    //
    // 重要：
    // ここで終了しない。
    // さらに後ろのチケットにスライド可能な予約が
    // 存在する可能性があるため。
    // ==========================================================
    if (!slideReservation) {
      continue;
    }

    const reservationDateStr = Utilities.formatDate(
      slideReservation.rowData[RES_COL_DATE],
      ssTimezone,
      'yyyy-MM-dd'
    );

    // ==========================================================
    // 予約を現在の移動先チケットへスライド
    // ==========================================================
    reservationsSheet
      .getRange(
        slideReservation.rowIndex + 1,
        RES_COL_TICKET_ID + 1
      )
      .setValue(
        currentTargetTicket.rowData[
          USER_TICKETS_COL_TICKET_ID
        ]
      );

    Logger.log(
      `チケットをスライドしました。` +
      ` targetTicketId=${currentTargetTicket.rowData[USER_TICKETS_COL_TICKET_ID]},` +
      ` sourceTicketId=${sourceTicket.rowData[USER_TICKETS_COL_TICKET_ID]},` +
      ` reservationDate=${reservationDateStr},` +
      ` targetTicketExpireDate=${currentTargetExpireDateStr}`
    );

    // ==========================================================
    // 今回予約を移動した「元チケット」が、
    // 次のスライド先になる
    //
    // 例：
    // A ← B
    // ↓
    // 次は B ← C
    // ==========================================================
    currentTargetIndex = sourceTicketIndex;
    currentTargetTicketData = sourceTicketData;
    currentTargetTicket = sourceTicket;

    currentTargetExpireDateStr = Utilities.formatDate(
      currentTargetTicket.rowData[
        USER_TICKETS_COL_EXPIRE_DATE
      ],
      ssTimezone,
      'yyyy-MM-dd'
    );

    // ==========================================================
    // 次のチケットから探し直す
    // ==========================================================
    // forループの次の処理へ
  }

  // ============================================================
  // 最後までスライドした結果、
  // 最後の移動元チケットに1枠空きができる
  //
  // 例：
  // A ← B ← C
  //
  // 最終的にCの残数を+1する。
  //
  // 途中で一切スライドできなかった場合は、
  // Aの残数を+1する。
  // ============================================================

  const finalTicketId = currentTargetTicket.rowData[USER_TICKETS_COL_TICKET_ID];

  if (finalTicketId !== skipFinalIncrementTicketId) {

    const remainingNum =
      currentTargetTicket.rowData[USER_TICKETS_COL_REMAINING_NUM];

    currentTargetTicket.rowData[USER_TICKETS_COL_REMAINING_NUM] =
      remainingNum + 1;
    
    userTicketsSheet
      .getRange(
        currentTargetTicket.rowIndex + 1,
        USER_TICKETS_COL_REMAINING_NUM + 1
      )
      .setValue(remainingNum + 1);

    Logger.log(
      `チケットスライド処理が完了しました。` +
      ` 最終的に残数を1戻したticketId=${finalTicketId}`
    );

  } else {

    Logger.log(
      `チケット残数+1をスキップしました。` +
      ` ticketId=${finalTicketId}`
    );

  }

}

function slideMonthlyAndTicket(
  userTicketsSheet,
  resSheet, 
  userMontlyRow,
  validUserTicketsRows,
  targetReservationData
) {
  
  const ssTimezone = SPREADSHEET.getSpreadsheetTimeZone();

  // ============================================================
  // キャンセル対象の予約月
  // ============================================================
  const selectedCancelMonth = Utilities.formatDate(targetReservationData.rowData[RES_COL_DATE], ssTimezone, "yyyy-MM");

  // ============================================================
  // 月謝上限
  // ============================================================
  const now = new Date();

  const currentMonth = Utilities.formatDate(now, ssTimezone, "yyyy-MM");
  const isThisMonth = currentMonth === selectedCancelMonth;

  const upperLimit = isThisMonth ? userMontlyRow[3] : userMontlyRow[4]
  const userId = userMontlyRow[0];
  
  // ============================================================
  // キャンセル後の対象月の予約を取得
  // ============================================================
  const resSheetData = resSheet.getDataRange().getValues().slice(1);
  
  const targetMonthUserReservation = resSheetData.filter(
    reservation => Utilities.formatDate(reservation[RES_COL_DATE], ssTimezone, "yyyy-MM") === selectedCancelMonth
      && reservation[RES_COL_USER_ID] === userId
      && reservation[RES_COL_STATUS] !== "キャンセル済み"
  );

  // ============================================================
  // 月謝枠を超えていない場合はチケット消費なし
  // ============================================================
  if (targetMonthUserReservation.length <= upperLimit) {
    return;
  }
  // ============================================================
  // 対象月のチケット予約を取得
  //
  // validUserTicketsRows は有効期限順なので、
  // 「有効期限が遠いチケット」を後ろから探す。
  // ============================================================
  for (let ticketIndex = validUserTicketsRows.length - 1; ticketIndex >= 0; ticketIndex--) {
    const ticketData = validUserTicketsRows[ticketIndex];
    const ticket = ticketData.ticket;
    const reservations = ticketData.reservations || [];

    // ------------------------------------------------------------
    // 対象月の確定予約だけを取得
    // ------------------------------------------------------------
    const targetMonthReservations = reservations.filter(
      reservation =>
        reservation.rowData[RES_COL_STATUS] === "確定" &&
        Utilities.formatDate(
          reservation.rowData[RES_COL_DATE],
          ssTimezone,
          "yyyy-MM"
        ) === selectedCancelMonth
    );

    if (targetMonthReservations.length === 0) {
      continue;
    }

    // ------------------------------------------------------------
    // 複数予約がある場合は、予約日が遅いものを月謝へ戻す
    // ------------------------------------------------------------
    targetMonthReservations.sort((a, b) => {
      const date1 = Utilities.formatDate(
        a.rowData[RES_COL_DATE],
        ssTimezone,
        'yyyy-MM-dd'
      );

      const date2 = Utilities.formatDate(
        b.rowData[RES_COL_DATE],
        ssTimezone,
        'yyyy-MM-dd'
      );

      // 予約日が新しい順
      if (date1 !== date2) {
        return date2.localeCompare(date1);
      }

      // 同じ日なら開始時刻が遅い順
      const time1 = a.rowData[RES_COL_START_TIME] instanceof Date
        ? Utilities.formatDate(
            a.rowData[RES_COL_START_TIME],
            ssTimezone,
            'HH:mm'
          )
        : String(a.rowData[RES_COL_START_TIME]).trim();

      const time2 = b.rowData[RES_COL_START_TIME] instanceof Date
        ? Utilities.formatDate(
            b.rowData[RES_COL_START_TIME],
            ssTimezone,
            'HH:mm'
          )
        : String(b.rowData[RES_COL_START_TIME]).trim();

      return time2.localeCompare(time1);
    });

    const targetReservation = targetMonthReservations[0];

    // ============================================================
    // チケット予約 → 月謝予約へ変更
    // ============================================================
    resSheet
      .getRange(
        targetReservation.rowIndex + 1,
        RES_COL_USAGE_TYPE + 1
      )
      .setValue("月謝");

    // 月謝になったのでチケットIDを削除
    resSheet
      .getRange(
        targetReservation.rowIndex + 1,
        RES_COL_TICKET_ID + 1
      )
      .setValue("");

    // ============================================================
    // 元のチケットに1枠戻す
    // ============================================================
    const remainingNum =
      Number(
        ticket.rowData[USER_TICKETS_COL_REMAINING_NUM]
      );

    ticket.rowData[USER_TICKETS_COL_REMAINING_NUM] =
      remainingNum + 1;

    userTicketsSheet
      .getRange(
        ticket.rowIndex + 1,
        USER_TICKETS_COL_REMAINING_NUM + 1
      )
      .setValue(remainingNum + 1);

    // ============================================================
    // メモリ上の予約からも移動した予約を削除
    //
    // この後 cancelTicketConsumption() を呼ぶため、
    // 「まだこのチケットに紐づいている予約」だけを残す。
    // ============================================================
    const movedReservationIndex =
      ticketData.reservations.findIndex(
        reservation =>
          reservation.rowIndex === targetReservation.rowIndex
      );

    if (movedReservationIndex !== -1) {
      ticketData.reservations.splice(
        movedReservationIndex,
        1
      );
    }

    Logger.log(
      `月謝枠へチケット予約をスライドしました。` +
      ` ticketId=${ticket.rowData[USER_TICKETS_COL_TICKET_ID]},` +
      ` reservationRow=${targetReservation.rowIndex + 1}`
    );

    // ============================================================
    // まだこのチケットに予約が残っている場合
    //
    // → 空いた1枠について後続チケットからスライド
    // ============================================================
    if (ticketData.reservations.length > 0) {
      cancelTicketConsumption(
        userTicketsSheet,
        resSheet,
        validUserTicketsRows,
        ticketIndex,
        ticket.rowData[USER_TICKETS_COL_TICKET_ID]
      );
    }
    return;
  }

  // ============================================================
  // 月謝枠を超えているのに対象チケット予約が見つからない
  // ============================================================
  throw new Error(
    `月謝へ戻すチケット予約が見つかりません。` +
    ` userId=${userId}, month=${selectedCancelMonth}`
  );
}

function cancelMonthlyAndTicketConsumption(
  userTicketsSheet,
  resSheet,
  userMontlyRow,
  validUserTicketsRows,
  targetTicketIndex,
  // {rowIndex: rowIndex, rowData: reservationData}
  targetReservationData
) {
  const targetCancelUsageType = targetReservationData.rowData[RES_COL_USAGE_TYPE];

  if (targetCancelUsageType === "月謝") {
    slideMonthlyAndTicket(userTicketsSheet, resSheet, userMontlyRow, validUserTicketsRows, targetReservationData);
  } else {
    cancelTicketConsumption(userTicketsSheet, resSheet, validUserTicketsRows, targetTicketIndex);
  }
}



