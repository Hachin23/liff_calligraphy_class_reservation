// ====================================
// 定数
// ====================================

// 予約可能日作成する月の範囲（2 = 来月まで、3 = 再来月まで）
const GENERATE_RESERVATIONS_LIST_MONTH = 3 //　2か月だと作成する月によっては、作成されない授業が発生する可能性があるので、念のため3か月分

// ====================================
// GAS カスタムメニューとフォーム表示関数
// ====================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('レッスン管理')
    .addItem('代理予約', 'openAdminBookingForm')
    .addItem('代理キャンセル', 'openAdminCancleForm')
    .addItem('チケット登録', 'openAdminTicketsPurchaseForm')
    .addItem('稽古回数登録', 'openAdminMonthlySubscriptionSettingForm')
    .addSeparator() // 区切り線を追加
    .addItem('例外日の追加・設定', 'openExceptionForm')
    .addItem('予約可能日リストを生成', 'executeReservationListGeneration')
    .addToUi();
}

// ====================================
// GASの実行承認を実行するための関数
// ====================================
function initializeAuth() {
  SpreadsheetApp.getActiveSpreadsheet(); // SpreadsheetAppへのアクセスを要求
  CalendarApp.getCalendars();           // CalendarAppへのアクセスを要求
  UrlFetchApp.fetch('https://google.com'); // 外部リクエストを要求
  SpreadsheetApp.getUi(); //uiを呼び出し
}

/**
 * 例外日の入力フォーム（サイドバー）を開く
 */
function openExceptionForm() {
  // ExceptionForm.html ファイルを事前に作成してください
  const html = HtmlService.createHtmlOutputFromFile('ExceptionForm') 
      .setTitle('例外日の設定')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 代理予約フォームを開く
 */
function openAdminBookingForm() {
  // AdminBookingForm.html ファイルを事前に作成してください
  const html = HtmlService.createHtmlOutputFromFile('AdminBookingForm') 
      .setTitle('代理予約')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 代理キャンセルフォームを開く
 */
function openAdminCancleForm() {
  // AdminCancleForm.html ファイルを事前に作成してください
  const html = HtmlService.createHtmlOutputFromFile('AdminCancleForm') 
      .setTitle('代理キャンセル')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * チケット登録フォームを開く
 */
function openAdminTicketsPurchaseForm() {
  // AdminTicketsPurchaseForm.html ファイルを事前に作成してください
  const html = HtmlService.createHtmlOutputFromFile('AdminTicketsPurchaseForm') 
      .setTitle('チケット登録')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 稽古回数登録フォームを開く
 */
function openAdminMonthlySubscriptionSettingForm() {
  // AdminMonthlySubscriptionSettingForm.html ファイルを事前に作成してください
  const html = HtmlService.createHtmlOutputFromFile('AdminMonthlySubscriptionSettingForm') 
      .setTitle('稽古回数登録')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ====================================
// 代理予約・キャンセル共通処理
// ====================================

/**
 * 生徒名簿をJSON形式で取得（UI用）
 */
function getStudentListForUI() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users');
  const data = sheet.getDataRange().getValues();
  // ヘッダーを除き、[名前, LINE ID, クラス名] を抽出
  return data.slice(1).map(row => ({ lineId: row[0], name: row[1], className: row[2]}));
}

// ====================================
// 代理予約処理
// ====================================

/**
 * 予約可能リスト
 */
function getReservationListForUI() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('reservationsList');
  const reservationListData = sheet.getDataRange().getValues();
  const timezone = Session.getScriptTimeZone();
  // ヘッダーを除外して reduce で集約
  return reservationListData.slice(1).reduce((acc, row) => {
    const [rawDate, , startTime, , className, lessonId, , remainingSeats] = row;
    if (!rawDate) return acc; // 空行対策
    
    // 日付を "yyyy-MM-dd" 形式の文字列キーに変換
    const dateKey = Utilities.formatDate(new Date(rawDate), timezone, "yyyy-MM-dd");

    if (!acc[dateKey]) {
      acc[dateKey] = [];
    }

    // 管理者が判断しやすい情報をセットにして格納
    acc[dateKey].push({
      lessonId: lessonId,
      startDate: startTime ? Utilities.formatDate(new Date(startTime), timezone, "HH:mm") : "",
      className: className,
      remainingSeats: remainingSeats
    });

    return acc;
  }, {});
}

/**
 * 実際の予約処理（管理者による代理実行）
 */
function adminProxyBooking(userId, lessonId, date, time, className) {
  try {
    // デバッグ・確認用メッセージ
    const result = makeReservation({ userId: userId, lessonId: lessonId, date: date, time: time, className: className });
    const msg = `【登録完了】\n生徒ID: ${userId}\n日時: ${date} ${time}\nクラス: ${className}\n(LessonID: ${lessonId})\nメッセージ: ${result.message}`;
    console.log(msg);
    return msg;
  } catch (e) {
    return "エラーが発生しました: " + e.toString();
  }
}

// ====================================
// 代理キャンセル処理
// ====================================

/**
 * ユーザの予約済みリストを取得
 */
function getReservationsForUI() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('reservations');
  const reservationData = sheet.getDataRange().getValues();
  const timezone = Session.getScriptTimeZone();

  // ヘッダーを除外して reduce で集約
  return reservationData.slice(1).reduce((acc, row) => {
    const [reservationId, rawDate, startTime, , userId, , className, , , , , , status, , ] = row;
    if (status !== '確定') return acc; //キャンセルは除外
    if (!reservationId) return acc; // 空行対策
    
    // 日付を "yyyy-MM-dd" 形式の文字列キーに変換
    const dateKey = Utilities.formatDate(new Date(rawDate), timezone, "yyyy-MM-dd");

    if (!acc[dateKey]) {
      acc[dateKey] = [];
    }

    // 管理者が判断しやすい情報をセットにして格納
    acc[dateKey].push({
      reservationId: reservationId,
      startDate: startTime ? Utilities.formatDate(new Date(startTime), timezone, "HH:mm") : "",
      userId: userId,
      className: className
    });

    return acc;
  }, {});
}

/**
 * 実際のキャンセル処理（管理者による代理実行）
 */
function adminProxyCancel(userId, reservationId, attendedCancel) {
  try {
    // デバッグ・確認用メッセージ
    result = handleCancelReservation({ userId: userId, reservationId: reservationId, admin: true,  attendedCancel: attendedCancel});
    const msg = `${result.message}\n生徒ID: ${userId}\n予約ID: ${reservationId}`;
    console.log(msg);
    return msg;
  } catch (e) {
    return "エラーが発生しました: " + e.toString();
  }
}

/**
 * 指定したユーザーの予約情報を取得
 */
function getUserReservations(userId, attendedCancel) {
  const reservationSheet = SPREADSHEET.getSheetByName(SHEET_NAME_RESERVATIONS);
  const userReservationData = reservationSheet.getDataRange().getValues().
    filter(reservation => 
      reservation[RES_COL_USER_ID] === userId && 
      (reservation[RES_COL_STATUS] === RESERVATION_STATUS.CONFIRMED || (attendedCancel && reservation[RES_COL_STATUS] === RESERVATION_STATUS.ATTENDED))
    );
  // 管理画面表示用の予約情報を返す
  const result = userReservationData.map(reservation => (
    {
      reservationId: reservation[RES_COL_RESERVATION_ID],
      date: Utilities.formatDate(reservation[RES_COL_DATE], SPREADSHEET.getSpreadsheetTimeZone(), 'yyyy-MM-dd'),
      startTime: Utilities.formatDate(reservation[RES_COL_START_TIME], SPREADSHEET.getSpreadsheetTimeZone(), 'HH:mm'),
      className: reservation[RES_COL_SELECTED_CLASS_NAME],
      status: reservation[RES_COL_STATUS] === "受講済み" ? reservation[RES_COL_STATUS] : "予約中"
    }));
  Logger.log(result);
  return result;
}

// =============================================
// チケット回数と値段をスクリプトプロパティから取得
// =============================================
function getTicketPrices() {

  const value = PropertiesService
    .getScriptProperties()
    .getProperty("TICKET_PRICES");

  if (!value) return [];

  return value.split(",").map(item => {
    const [count, price] = item.split(":");
    return {
      count: Number(count),
      price: Number(price)
    };
  }).sort((a, b) => a.count - b.count);
}

function getTicketPriceMap() {
  const value = PropertiesService
    .getScriptProperties()
    .getProperty("TICKET_PRICES");

  return value.split(",").reduce((obj, item) => {
    const [count, price] = item.split(":");
    obj[count] = Number(price);
    return obj;
  }, {});
}

// =========================
// チケット登録処理（管理者画面からの呼出し）
// =========================
function saveTicket(data) {

  const lock = LockService.getScriptLock();
  lock.tryLock(30000); // 30秒間ロックを試みる

  if (lock.hasLock()) {
    try {

      const userTicketSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USER_TICKETS);
      const reservationsSheet = SPREADSHEET.getSheetByName(SHEET_NAME_RESERVATIONS);
      if (!reservationsSheet || !userTicketSheet) {
        throw new Error(`${SHEET_NAME_RESERVATIONS} ${SHEET_NAME_USER_TICKETS} のいずれかのシートが見つかりません。`) 
      }

      const ticketId = Utilities.getUuid();
      const priceMap = getTicketPriceMap();
      const price = priceMap[data.count];
    
      if (!price) {
        throw new Error("料金設定が存在しません。");
      }
    
      const purchaseDate = new Date(data.purchaseDate);
      const expireDate = new Date(purchaseDate);
      const months = data.count <= 3 ? 3 : 6;
      expireDate.setMonth(expireDate.getMonth() + months + 1);
      expireDate.setDate(0);
      const expireDateStr = Utilities.formatDate(
      expireDate,
      Session.getScriptTimeZone(),
      "yyyy/MM/dd"
      );
      const purchaseDateStr = Utilities.formatDate(
      purchaseDate,
      Session.getScriptTimeZone(),
      "yyyy/MM/dd"
      );
    
      const student = getStudentByLineId(data.userId);
      const now = new Date();
      const ssTimezone = SPREADSHEET.getSpreadsheetTimeZone();
      const dateString = Utilities.formatDate(now, ssTimezone, "yyyy-MM-dd");
      
      const userTicketRows = userTicketSheet.getDataRange().getValues()
        .filter(ticket =>
          ticket[USER_TICKETS_COL_USER_ID] === data.userId &&
          Utilities.formatDate(ticket[USER_TICKETS_COL_EXPIRE_DATE], ssTimezone, 'yyyy-MM-dd') >= dateString
        );

      // チケット情報を登録
      userTicketSheet.appendRow([
        ticketId,
        data.userId,
        student.name,
        data.count,
        price,
        data.count,
        purchaseDateStr,
        expireDateStr,
        "有効",
        now,
        now,
        data.memo
      ]);
      
      if (userTicketRows.length > 0) {
        // 未消費を含めた有効期限が有効なチケットを取得
        const ticketReservationData = getTicketReservations(
          reservationsSheet,
          userTicketSheet,
          data.userId,
        );
        // 現在のチケット予約の紐付けを一旦解除
        clearTicketReservations(
          reservationsSheet,
          userTicketSheet,
          ticketReservationData
        );

        // 有効期限の近いチケットから再割り当て
        reassignTicketReservations(
          reservationsSheet,
          userTicketSheet,
          ticketReservationData
        );
      }

      syncUserFullData(data.userId);

    } finally {
      lock.releaseLock();
    }
  } else {
    throw new Error("現在、他の処理が実行中です。しばらくしてから再度お試しください。");
  }
}

function getStudentByLineId(lineId) {
  const students = getStudentListForUI();
  const student = students.find(s => s.lineId === lineId);
  if (!student) {
    throw new Error("生徒が見つかりません。");
  }
  return student;
}

// ======================
// 月の稽古回数を登録する処理
// ======================
function saveMonthlyLesson(data) {
  
  const lock = LockService.getScriptLock();
  lock.tryLock(30000); // 30秒間ロックを試みる

  if (lock.hasLock()) {
    try {
      const userMonthlySubscriptionsSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USER_MONTHLY_SUBSCRIPTIONS);
      const reservationsSheet = SPREADSHEET.getSheetByName(SHEET_NAME_RESERVATIONS);
      const userTicketSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USER_TICKETS);
      if (!userMonthlySubscriptionsSheet || !reservationsSheet || !userTicketSheet) {
        throw new Error(`${SHEET_NAME_RESERVATIONS} ${SHEET_NAME_USER_MONTHLY_SUBSCRIPTIONS} ${SHEET_NAME_USER_TICKETS} のいずれかのシートが見つかりません。`) 
      }
      if (!data.userId) {
        throw new Error("生徒を選択してください");
      }
      const exists = userMonthlySubscriptionsSheet.getDataRange().getValues().slice(1).some(row => row[0] === data.userId);
      if(exists){
        throw new Error("登録済みの生徒です");
      }
  
      // 月謝管理シートに追記
      userMonthlySubscriptionsSheet.appendRow([
        data.userId,
        data.userName,
        data.defaultCount,
        data.currentCount,
        data.nextCount
      ]);
      
      // チケット予約のスライド処理
      const monthlyData = getMonthlyTicketReservations(reservationsSheet, userTicketSheet, data.userId);
      // 今月
      convertTicketReservationsToMonthly(reservationsSheet, userTicketSheet, monthlyData.thisMonth, data.currentCount);
      // 来月
      convertTicketReservationsToMonthly(reservationsSheet, userTicketSheet, monthlyData.nextMonth, data.nextCount);

      // WorkersKVのユーザ情報を更新
      syncUserFullData(data.userId);
    } finally {
      lock.releaseLock();
    }
  } else {
    throw new Error("現在、他の処理が実行中です。しばらくしてから再度お試しください。");
  }
}

function convertTicketReservationsToMonthly(
  reservationsSheet,
  userTicketSheet,
  monthlyReservations,
  monthlyCount
) {
  if (monthlyCount <= 0) {
    return;
  }

  let reservationCnt = 0;

  for (const reservationInfo of monthlyReservations) {
    for (const reservation of reservationInfo.reservations) {

      // 利用種別の変更（チケット → 月謝）
      reservationsSheet.getRange(reservation.rowIndex + 1, RES_COL_USAGE_TYPE + 1).setValue("月謝");
      // 月謝になったのでチケットIDを削除
      reservationsSheet.getRange(reservation.rowIndex + 1, RES_COL_TICKET_ID + 1).setValue("");
      // チケットの残数を1戻す
      const remainingNum = Number(reservationInfo.ticket.rowData[USER_TICKETS_COL_REMAINING_NUM]);
      reservationInfo.ticket.rowData[USER_TICKETS_COL_REMAINING_NUM] = remainingNum + 1;
      userTicketSheet.getRange(reservationInfo.ticket.rowIndex + 1, USER_TICKETS_COL_REMAINING_NUM + 1).setValue(remainingNum + 1);

      ++reservationCnt;
      if (reservationCnt === monthlyCount) {
        break;
      }
    }
    if (reservationCnt === monthlyCount) {
      break;
    }
  }
}

function convertMonthlyReservationsToTicket(
  reservationsSheet,
  userTicketSheet,
  userId,
  monthlyReservations,
  monthlyCount
) {
  const ssTimezone = SPREADSHEET.getSpreadsheetTimeZone();
  const dateString = Utilities.formatDate(new Date(), ssTimezone, "yyyy-MM-dd");
  const userTicketsRows = userTicketSheet.getDataRange().getValues();
  const userTicketsRowsWithIndex = userTicketsRows.map((row, index) => ({
    rowIndex: index,
    rowData: row
  }));
  const validUserTicketsRows = userTicketsRowsWithIndex
    .filter(ticket =>
    ticket.rowData[USER_TICKETS_COL_USER_ID] === userId &&
    Utilities.formatDate(ticket.rowData[USER_TICKETS_COL_EXPIRE_DATE], ssTimezone, 'yyyy-MM-dd') >= dateString &&
    ticket.rowData[USER_TICKETS_COL_REMAINING_NUM] > 0
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

  const availableTicketCount = validUserTicketsRows
    .reduce((total, ticket) => total + Number(ticket.rowData[USER_TICKETS_COL_REMAINING_NUM]), 0);

  if (availableTicketCount < monthlyCount) {
    throw new Error(
      `チケットの残数が不足しています。必要枚数：${monthlyCount}枚、利用可能枚数：${availableTicketCount}枚`
    );
  }

  for (let i = 0; i < monthlyCount; i++) {
    const reservation = monthlyReservations[i];
    const targetTicket = validUserTicketsRows.find(ticket => ticket.rowData[USER_TICKETS_COL_REMAINING_NUM] > 0);
    // 月謝 → チケット
    reservationsSheet
      .getRange(reservation.rowIndex + 1, RES_COL_USAGE_TYPE + 1)
      .setValue("チケット");
    // チケットIDを設定
    reservationsSheet
      .getRange(reservation.rowIndex + 1, RES_COL_TICKET_ID + 1)
      .setValue(targetTicket.rowData[USER_TICKETS_COL_TICKET_ID]);
    // チケット残数を1減らす
    const remainingNum = Number(targetTicket.rowData[USER_TICKETS_COL_REMAINING_NUM]) - 1;
    targetTicket.rowData[USER_TICKETS_COL_REMAINING_NUM] = remainingNum;
    userTicketSheet
      .getRange(targetTicket.rowIndex + 1, USER_TICKETS_COL_REMAINING_NUM + 1)
      .setValue(remainingNum);
  }
}

// ====================================
// 例外日データ書き込みロジック
// ====================================

/**
 * フォームから受け取った例外日データをExceptionsシートに追記する
 */
function addExceptionEntry(dateStr, type, targetId, note) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_EXCEPTIONS);
  if (!sheet) {
    return `エラー: シート '${SHEET_NAME_EXCEPTIONS}' が見つかりません。`;
  }
  
  const dateStringJp = dateStr.replace(/-/g, '/');
  const lastRow = sheet.getLastRow();
  const nextId = 'E' + (lastRow + 1).toString().padStart(3, '0');
  const targetValue = (type === '特定ID休講' || type === '臨時開講') ? targetId : ''; 

  // Exceptionsシートの構成: | A: ID | B: 日付 | C: 例外タイプ | D: 対象ID/備考 | E: 備考 |
  const rowData = [
    nextId, 
    dateStringJp, 
    type, 
    targetValue, 
    note
  ];

  sheet.appendRow(rowData);
  
  return `${type} (${dateStr}) の例外設定が完了しました。`;
}

/**
 * UIオブジェクトを安全に取得します。
 * UIが存在しないコンテキストで実行された場合は null を返します。
 * @returns {GoogleAppsScript.Base.Ui | null}
 */
function safelyGetUi() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    // ログに記録しますが、エラーはスローしません
    // Logger.log("UI環境外で実行されました。"); 
    return null;
  }
}

/**
 * 予約リスト生成を実行し、確実に完了シグナルを送るためのラッパー関数。
 */
function executeReservationListGeneration(event) {
  
  // 💡 UIが存在する場合、通知用に取得
  const ui = SpreadsheetApp.getUi();
  const uiExists = ui && typeof ui.alert === 'function';

  try {
    // 従来の関数を実行
    generateReservationsList(event);
    
    if (uiExists) {
        // UI操作（通知）を伴わないシンプルな終了メッセージをログに出力
        Logger.log('リスト生成処理が正常に完了しました。');
    }
    
  } catch (e) {
    if (uiExists) {
        // ui.alert() は使わないが、Browser.msgBox()で通知を試みる
        Browser.msgBox('致命的なエラー', 'リスト生成中にエラーが発生しました。ログを確認してください: ' + e.message);
    }
    Logger.log('[FATAL WRAPPER ERROR] ' + e);
  }
}


// ====================================
// 予約可能日リストの生成（残席数を維持する）
// ====================================
/**
 * 予約可能日リストを生成し、reservationsListシートに出力する。
 * (今日から未来 3ヶ月間を生成対象とする)
 */
function generateReservationsList(event) {

  const lock = LockService.getScriptLock();
  const LOCK_TIMEOUT_MS = 30000;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const ui = safelyGetUi();
  const uiExists = ui !== null;
  const listSheet = ss.getSheetByName(SHEET_NAME_RESERVATIONS_LIST) || ss.insertSheet(SHEET_NAME_RESERVATIONS_LIST);
  let generatedCount = 0;

  try {
    const isLockAcquired = lock.tryLock(LOCK_TIMEOUT_MS);

    if (!isLockAcquired) {
      // ロックが取得できなかった場合（他のバッチ処理などが実行中の場合）
      // エラーとして処理を中断する
      throw new Error('他のバッチ処理が実行中のため、ロックを取得できませんでした。処理をスキップします。');
    }

    const templateSheet = ss.getSheetByName(SHEET_NAME_SCHEDULE_TEMPLATE);
    const exceptionsSheet = ss.getSheetByName(SHEET_NAME_EXCEPTIONS);

    if (!templateSheet || !exceptionsSheet) {
      throw new Error(`必要なシートが見つかりません: ${SHEET_NAME_SCHEDULE_TEMPLATE} または ${SHEET_NAME_EXCEPTIONS}`);
    }

    const ssTimezone = ss.getSpreadsheetTimeZone();

    // --- A. 既存の残席数データを読み込む ---
    const existingData = listSheet.getDataRange().getValues();
    const existingReservationsMap = {};
    
    if (existingData.length > 1) {
      existingData.slice(1).forEach(row => {
        const rawRemainingCapacity = row[RES_LIST_COL_REMAINING_CAPACITY];
        const remainingCapacity = parseInt(rawRemainingCapacity, 10); 
        
        if (!isNaN(remainingCapacity) && remainingCapacity >= 0) {
          
          const lessonId = row[RES_LIST_COL_LESSON_ID];
          
          // 日付を yyyy-MM-dd 文字列に変換
          const dateStr = Utilities.formatDate(new Date(row[RES_LIST_COL_DATE]), ssTimezone, 'yyyy-MM-dd');
          
          // ★★★ 修正箇所: 時刻データを文字列化し、trim() して空欄チェック ★★★
          const startTimeRaw = row[RES_LIST_COL_START_TIME];
          // タイムスタンプ（Dateオブジェクト）の場合もあれば、既に文字列の場合もある
          let startTime = '';
          if (startTimeRaw instanceof Date) {
              // Dateオブジェクトの場合、HH:mm文字列にフォーマット
              startTime = Utilities.formatDate(startTimeRaw, ssTimezone, 'HH:mm');
          } else if (typeof startTimeRaw === 'string') {
              // 単なる文字列の場合、trim()のみ
              startTime = startTimeRaw.trim();
          } else {
              // その他のデータ型の場合 (数値など) は文字列化して trim()
              startTime = startTimeRaw ? String(startTimeRaw).trim() : '';
          }
          
          if (startTime === '') {
              // 時刻が空の場合はキーに含めない（キーを成立させない）
              Logger.log(`Skipping row due to empty time: ${lessonId}, ${dateStr}`);
              return; 
          }
          
          const key = `${lessonId}_${dateStr}_${startTime}`;
          existingReservationsMap[key] = remainingCapacity;
        }
      });
    }

    // --- B. テンプレートと例外データの読み込み (省略) ---
    const templates = templateSheet.getDataRange().getValues().slice(1);
    const exceptions = exceptionsSheet.getDataRange().getValues().slice(1);

    // レッスンIDから詳細を取得するためのマップ
    const templateDetailMap = {};
    templates.forEach(template => {
        templateDetailMap[template[0]] = {
            time: template[2],       // 開始時間オブジェクト/数値
            endTime: template[3],    // 終了時間オブジェクト/数値
            capacity: parseInt(template[4], 10), // 定員 (数値化)
            className: template[5],  // クラス名
            dayOfWeek: template[1]   // 曜日番号
        };
    });

    // 例外日を処理しやすいように Map に変換 (省略)
    const exceptionMap = {};
    exceptions.forEach(row => {
      const date = Utilities.formatDate(new Date(row[1]), ssTimezone, 'yyyy-MM-dd');
      const type = row[2];
      const targetId = row[3] || '';
      
      if (!exceptionMap[date]) {
        exceptionMap[date] = [];
      }
      exceptionMap[date].push({ type: type, targetId: targetId });
    });
    
    // --- C. 予約枠の生成期間を設定（3ヶ月） ---
    const today = new Date();
    const endDate = new Date();
    endDate.setMonth(today.getMonth() + GENERATE_RESERVATIONS_LIST_MONTH);

    let reservationsData = [];
    let currentDate = new Date(today.getTime());
    currentDate.setHours(0, 0, 0, 0);

    // 期間内の日付を1日ずつループ
    while (currentDate.getTime() <= endDate.getTime()) {
      const dayOfWeek = currentDate.getDay();
      const dateStr = Utilities.formatDate(currentDate, ssTimezone, 'yyyy-MM-dd');

      // テンプレートを適用して予約枠を生成
      templates.forEach(template => {
        const templateDay = template[1];
        
        if (templateDay === dayOfWeek) {
          const lessonId = template[0];
          const capacity = parseInt(template[4], 10);
          
          // 生成時の時刻は必ず HH:mm 文字列で統一
          const startTime = Utilities.formatDate(template[2], ssTimezone, 'HH:mm'); 
          const endTime = Utilities.formatDate(template[3], ssTimezone, 'HH:mm');
          
          const className = template[5];
          
          // --- D. 例外日の適用 (フィルタリング) (省略) ---
          let isAvailable = true;
          
          if (exceptionMap[dateStr]) {
            for (const ex of exceptionMap[dateStr]) {
              if (ex.type === '全日休講' || (ex.type === '特定ID休講' && ex.targetId === lessonId)) {
                isAvailable = false;
                break;
              }
            }
          }

          if (isAvailable) {
            
            // --- E. 残席数の決定（既存データがあれば引き継ぐ） ---
            const key = `${lessonId}_${dateStr}_${startTime}`; // ★ キー生成を統一
            const initialCapacity = capacity;
            
            const remainingCapacity = existingReservationsMap[key] !== undefined 
                                    ? existingReservationsMap[key] 
                                    : initialCapacity;

            reservationsData.push([
              new Date(dateStr), 
              NUMBER_TO_DAY_NAME[dayOfWeek], 
              startTime,
              endTime,
              className,
              lessonId,
              initialCapacity,
              remainingCapacity
            ]);
          }
        }
      });

      // 次の日へ
      currentDate.setDate(currentDate.getDate() + 1);
    }

    // --- E-EXTRA. 臨時開講（強制追加）の処理 ---
    for (const dateStr in exceptionMap) {
        exceptionMap[dateStr].filter(ex => ex.type === '臨時開講' && ex.targetId).forEach(ex => {
            const lessonId = ex.targetId;
            const templateDetail = templateDetailMap[lessonId];
            
            if (templateDetail) {
                const currentExceptionDate = new Date(dateStr);
                const dayOfWeek = currentExceptionDate.getDay();
                
                const startTime = Utilities.formatDate(templateDetail.time, ssTimezone, 'HH:mm');
                const endTime = Utilities.formatDate(templateDetail.endTime, ssTimezone, 'HH:mm');
                const initialCapacity = templateDetail.capacity;
                const className = templateDetail.className;
                
                // 残席数の引き継ぎチェック
                const key = `${lessonId}_${dateStr}_${startTime}`;
                const remainingCapacity = existingReservationsMap[key] !== undefined 
                                        ? existingReservationsMap[key] 
                                        : initialCapacity;

                // 予約枠を強制的に追加
                reservationsData.push([
                    currentExceptionDate, 
                    NUMBER_TO_DAY_NAME[dayOfWeek], 
                    startTime,
                    endTime,
                    className,
                    lessonId,
                    initialCapacity,
                    remainingCapacity
                ]);
            }
        });
    }

    // --- F. リストシートへの書き込み (省略) ---
    listSheet.clear();
    const header = ['日付', '曜日', '開始時間', '終了時間', 'クラス名', 'レッスンID', '定員', '残席数'];
    listSheet.appendRow(header);

    if (reservationsData.length > 0) {
      listSheet.getRange(2, 1, reservationsData.length, reservationsData[0].length).setValues(reservationsData);
      listSheet.getRange('A:A').setNumberFormat('yyyy/MM/dd');
      generatedCount = reservationsData.length;
    }

    // 成功時の通知: UIが存在する場合のみアラートを表示
    if (uiExists) {
      Utilities.sleep(1000);
      // Browser.msgBox(
      //     '予約可能日リストの生成が完了しました！', 
      //     `${generatedCount} 件の予約枠を書き込みました。`, 
      //     Browser.Buttons.OK
      // );
    } else {
      Logger.log(`[SUCCESS] 予約可能日リスト生成完了: ${generatedCount} 件を書き込みました。`);
    }
    // 生徒の予約回数をスライド（来月の回数を今月の回数へ変更する。来月はデフォルトを設定）
    monthlyMaintenance(event);
    // 過去の予約を削除する。（データ整理）
    deletePastReservations();

    // WorkersKV同期
    // 残席情報
    syncCapacityToWorkers();
    // ユーザの予約情報
    migrateAllUsersToWorkers();

  } catch (e) {
    // エラー時の通知: UIが存在する場合のみアラートを表示
    if (uiExists) {
        Browser.msgBox('エラーが発生しました', 'リスト生成中にエラー: ' + e.message, Browser.Buttons.OK);
    } else {
        Logger.log(`[FATAL ERROR] リスト生成中にエラーが発生しました: ${e.message}`);
    }
    Logger.log(e);
    throw e;
    
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
  return;
}

function deletePastReservations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_RESERVATIONS);
  const data = sheet.getDataRange().getValues();

  // 実行した時点の「今月1日 00:00:00」を取得
  const now = new Date();
  const firstDayOfThisMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  
  console.log(firstDayOfThisMonth + " より前のデータを削除します。");

  // 下から順にループ（行番号がずれないための鉄則）
  // 2行目(i=1)までチェック。1行目はヘッダーなので除外。
  for (let i = data.length - 1; i >= 1; i--) {
    const reservationDate = new Date(data[i][RES_COL_DATE]);
    
    // 予約日が「今月1日」よりも前（＝先月以前）なら削除
    if (reservationDate < firstDayOfThisMonth) {
      sheet.deleteRow(i + 1);
    }
  }
}

// 生徒の来月の授業回数を今月の授業回数に修正.来月の授業回数は初回の選択値を設定.
// generateReservationsList内で呼び出し（月が切り替わるタイミングで呼び出し）
function monthlyMaintenance(event) {
  // トリガー実行の時のみ実行（手動実行時はスルー）
  if (event) {
    // ===========================================
    // 今月の稽古回数のスライド処理実行
    // ===========================================
    const userMonthlySubscriptionsSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USER_MONTHLY_SUBSCRIPTIONS);
    const reservationsSheet = SPREADSHEET.getSheetByName(SHEET_NAME_RESERVATIONS);
    const userTicketSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USER_TICKETS);

    if (!userMonthlySubscriptionsSheet || !reservationsSheet || !userTicketSheet) {
      throw new Error("必要なシートが見つかりません。");
    }

    const userMonthlySubscriptionsData = userMonthlySubscriptionsSheet.getDataRange().getValues();
    
    // 配列上の列番号
    const COL_USER_ID = 0;  // A列
    const COL_DEFAULT = 2;  // C列
    const COL_CURRENT = 3;  // D列
    const COL_NEXT = 4;     // E列
    
    for (let i = 1; i < userMonthlySubscriptionsData.length; i++) {
      const row = userMonthlySubscriptionsData[i];
      const userId = row[COL_USER_ID];

      if (!userId) {
        continue;
      }
      // ===========================================
      // 切り替え前の回数を取得
      // ===========================================
      const newCurrentCount = Number(row[COL_NEXT]) || 0;
      const newNextCount = Number(row[COL_DEFAULT]) || 0;

      // ===========================================
      // 現在のチケット予約を一旦解除
      // ===========================================
      // 今月の稽古回数を調整するため、
      // 最初に既存のチケット予約をすべて解除する。
      const ticketReservationData = getTicketReservations(
        reservationsSheet,
        userTicketSheet,
        userId
      );

      clearTicketReservations(
        reservationsSheet,
        userTicketSheet,
        ticketReservationData
      );
      
      // ===========================================
      // 現在の月謝予約を取得
      // ===========================================
      let monthlyReservations = getMonthlyLessonReservations(reservationsSheet, userId);

      // ===========================================
      // 今月の月謝枠を調整
      // ===========================================
      const currentReservations = monthlyReservations.thisMonth;

      if (currentReservations.length < newCurrentCount) {
        const convertCount =
          newCurrentCount - currentReservations.length;

        const monthlyData =
          getMonthlyTicketReservations(
            reservationsSheet,
            userTicketSheet,
            userId
          );

        convertTicketReservationsToMonthly(
          reservationsSheet,
          userTicketSheet,
          monthlyData.thisMonth,
          convertCount
        );

      } else if (currentReservations.length > newCurrentCount) {
        const convertCount = currentReservations.length - newCurrentCount;
        convertMonthlyReservationsToTicket(
          reservationsSheet,
          userTicketSheet,
          userId,
          currentReservations,
          convertCount
        );
      }

      // ===========================================
      // 今月チケット予約を再配置
      // ===========================================
      const updatedTicketReservationData =
        getTicketReservations(
          reservationsSheet,
          userTicketSheet,
          userId
        );

      reassignTicketReservations(
        reservationsSheet,
        userTicketSheet,
        updatedTicketReservationData
      );

      // ===========================================
      // 月謝回数を更新
      // ===========================================
      row[COL_CURRENT] = newCurrentCount;
      row[COL_NEXT] = newNextCount;

      syncUserFullData(userId);
    }
    
    // ===========================================
    // 月謝回数の更新をまとめて反映
    // ===========================================
    userMonthlySubscriptionsSheet
      .getRange(1, 1, userMonthlySubscriptionsData.length, userMonthlySubscriptionsData[0].length)
      .setValues(userMonthlySubscriptionsData);
    
    // ===========================================
    // チケットステータスを「有効期限切れ」に更新
    // ===========================================

    // 有効期限が切れているチケットでステータスが「有効」のチケットを取得
    const ssTimezone = SPREADSHEET.getSpreadsheetTimeZone();
    const dateStr = Utilities.formatDate(new Date(), ssTimezone, "yyyy-MM-dd");
    const userTicketsRowsWithIndex = userTicketSheet.getDataRange().getValues().slice(1).map((row, index) => ({
      rowIndex: index + 1,
      rowData: row
    }));
    const targetUserTicketRow = userTicketsRowsWithIndex.filter(ticket =>
      // 有効期限が切れている
      Utilities.formatDate(ticket.rowData[USER_TICKETS_COL_EXPIRE_DATE], ssTimezone, 'yyyy-MM-dd') < dateStr &&
      ticket.rowData[USER_TICKETS_COL_STATUS] === "有効"
    )
    
    targetUserTicketRow.forEach(ticket => {
      userTicketSheet
      .getRange(ticket.rowIndex + 1, USER_TICKETS_COL_STATUS + 1)
      .setValue("有効期限切れ");
    });
  }
}

/**
 * 1時間ごとに実行されるカレンダー連携・メンテナンス関数。
 * 予約枠ごとにデータを集計し、単一のイベントを作成・更新する。
 */
function frequentCalendarSyncAndMaintenance() {

  const lock = LockService.getScriptLock();
  const LOCK_TIMEOUT_MS = 30000; // 30秒待機

  try {

    const isLockAcquired = lock.tryLock(LOCK_TIMEOUT_MS);

    if (!isLockAcquired) {
      Logger.log('他のGAS処理が実行中のため、カレンダー同期処理をスキップしました。');
      // ロックを取得できなかったため、処理を中断
      throw new Error('他のGAS処理が実行中のため、カレンダー同期処理をスキップしました。');
    }

    const ssTimezone = SPREADSHEET.getSpreadsheetTimeZone();
    const resSheet = SPREADSHEET.getSheetByName(SHEET_NAME_RESERVATIONS);
    const userTicketSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USER_TICKETS);
    const calendar = CalendarApp.getCalendarById(ADMIN_CALENDAR_ID);
    const now = new Date();

    const dataRange = resSheet.getDataRange();
    const values = dataRange.getValues();

    if (!resSheet || !calendar || !userTicketSheet) {
        if (resSheet) updatePastReservationsToAttended(resSheet, values, now); // 過去の予約更新は継続
        throw new Error('指定のカレンダーが存在しないため、過去の予約の更新のみ実行しました。');
    }
    
    // 予約枠ごとの情報を格納するオブジェクト
    const slotReservations = {}; 
    const HEADER_ROW_COUNT = 1;

    const eventIdToRowIndices = {};

    // ===========================================
    // STEP 1: 予約データを枠ごとにグループ化し、受講済みを更新
    // ===========================================
    for (let i = HEADER_ROW_COUNT; i < values.length; i++) {
        const row = values[i];
        const rowIndex = i + 1;
        
        const status = row[RES_COL_STATUS];
        
        // 日付・時刻情報の取得
        const dateRaw = row[RES_COL_DATE];
        if (!dateRaw || !(dateRaw instanceof Date)) continue;

        const dateStr = Utilities.formatDate(dateRaw, ssTimezone, 'yyyy/MM/dd');
        const timeStr = Utilities.formatDate(row[RES_COL_START_TIME], ssTimezone, 'HH:mm');
        const endTimeStr = Utilities.formatDate(row[RES_COL_END_TIME], ssTimezone, 'HH:mm');
        
        const startDateTime = new Date(`${dateStr} ${timeStr}`);
        const eventId = row[RES_COL_CALENDAR_EVENT_ID];

        if (eventId) {
          if (!eventIdToRowIndices[eventId]) {
              eventIdToRowIndices[eventId] = [];
          }
          eventIdToRowIndices[eventId].push(rowIndex);
        }
        
        // A. 過去の予約の場合: ステータスを「受講済み」に更新
        if (status === '確定' && startDateTime < now) {
            resSheet.getRange(rowIndex, RES_COL_STATUS + 1).setValue('受講済み');
            continue; 
        }

        // B. 未来の確定予約のみをグループ化
        if (status === '確定' && startDateTime >= now) {
            // 💡 修正点: クラスに関わらず、時間帯でキーを作成
            const slotKey = `${dateStr}_${timeStr}`; 
            const className = row[RES_COL_SELECTED_CLASS_NAME]; // クラス名を取得
            
            if (!slotReservations[slotKey]) {
                const endDateTime = new Date(`${dateStr} ${endTimeStr}`);
                
                slotReservations[slotKey] = {
                    startDateTime: startDateTime,
                    endDateTime: endDateTime,
                    // 💡 混在するクラス名を保持するためにオブジェクトに変更
                    classBreakdown: {}, 
                    totalCount: 0,
                    rowIndices: [],
                    eventIds: []
                };
            }
            
            // 💡 クラスごとの集計
            if (!slotReservations[slotKey].classBreakdown[className]) {
                slotReservations[slotKey].classBreakdown[className] = [];
            }
            slotReservations[slotKey].classBreakdown[className].push(row[RES_COL_RESERVER_NAME]);

            slotReservations[slotKey].totalCount++;
            slotReservations[slotKey].rowIndices.push(rowIndex);
            
            if (eventId && !slotReservations[slotKey].eventIds.includes(eventId)) {
                  slotReservations[slotKey].eventIds.push(eventId);
            }
        }
    }
    // ユーザの予約情報のWorkersKV同期
    migrateAllUsersToWorkers();

    // ===========================================
    // STEP 2: 予約枠ごとにカレンダーイベントを作成または更新
    // ===========================================
    const activeEventIds = new Set();
    
    for (const slotKey in slotReservations) {
        const slot = slotReservations[slotKey];
        const totalCount = slot.totalCount;
        
        // 💡 修正点: イベントの説明文にクラス別の内訳を含める
        let titleDetails = [];
        let descriptionDetails = [];
        for (const className in slot.classBreakdown) {
          const count = slot.classBreakdown[className].length;
          const names = slot.classBreakdown[className].join(', ');
          
          // タイトル用の詳細部分を配列に追加
          titleDetails.push(`${className}(${count}名)`); 
          
          // 説明文用の詳細部分を構築
          descriptionDetails.push(
              `--- ${className} (${count}名) ---\n` +
              `予約者: ${names}`
          );
        }
        
        const eventTitle = `【予約${totalCount}名】${titleDetails.join(' ')}`;
        const eventDescription = 
            `**総予約人数: ${totalCount}名**\n` +
            `==========================\n` +
            `${descriptionDetails.join('\n')}\n` +
            `==========================`;

        let currentEventId = slot.eventIds.length > 0 ? slot.eventIds[0] : null;

        try {
            let event;

            // 既存のイベントがある場合は更新
            if (currentEventId) {
                event = calendar.getEventById(currentEventId);
                if (event) {
                    event.setTitle(eventTitle);
                    event.setDescription(eventDescription);
                    Logger.log(`イベントを更新しました (${slotKey})`);
                } else {
                    currentEventId = null; // 見つからない場合は新規作成へ
                }
            }
            
            // イベントがない場合は新規作成
            if (!currentEventId) {
                event = calendar.createEvent(
                    eventTitle,
                    slot.startDateTime,
                    slot.endDateTime,
                    { description: eventDescription, sendInvites: false }
                );
                currentEventId = event.getId();
                Logger.log(`新規イベントを作成しました (${slotKey})`);
            }

            if (currentEventId) {
              activeEventIds.add(currentEventId);
            }
            
            // STEP 3: 該当する全ての予約レコードにイベントIDを記録
            if (currentEventId) {
                for (const rowIndex of slot.rowIndices) {
                    resSheet.getRange(rowIndex, RES_COL_CALENDAR_EVENT_ID + 1).setValue(currentEventId);
                }
            }

        } catch (e) {
            Logger.log(`カレンダー同期エラー (${slotKey}): ${e.toString()}`);
        }
    }

    // ===========================================
    // イベントのクリーンアップと削除 (キャンセル対応)
    // ===========================================
    
    // シートに存在するが、activeEventIdsに含まれない（確定予約が0になった）イベントIDを特定
    const idsToDelete = Object.keys(eventIdToRowIndices).filter(eventId => !activeEventIds.has(eventId));
    
    if (idsToDelete.length > 0) {
        const rangesToClear = []; // シートのIDをクリアするためのRangeオブジェクトを保持
        
        for (const eventId of idsToDelete) {
            try {
                const event = calendar.getEventById(eventId);
                if (event) {
                    event.deleteEvent();
                    Logger.log(`キャンセルにより不要になったカレンダーイベントID: ${eventId} を削除しました。`);
                }

                // 関連するシートの Event ID 列をクリアするための座標を収集
                const rowIndices = eventIdToRowIndices[eventId];
                for (const rowIndex of rowIndices) {
                    // A1表記で Range を収集
                    rangesToClear.push(resSheet.getRange(rowIndex, RES_COL_CALENDAR_EVENT_ID + 1).getA1Notation());
                }
            } catch (e) {
                Logger.log(`カレンダーイベント削除エラー (${eventId}): ${e.toString()}`);
            }
        }
        
        // 💡 バッチ更新でシートのイベントIDをクリア（高速化）
        if (rangesToClear.length > 0) {
            resSheet.getRangeList(rangesToClear).clearContent();
            Logger.log(`${rangesToClear.length} 件のシート上のカレンダーIDをクリアしました。`);
        }
    }

    // ===========================================
    // チケットステータスを「消費済み」に更新
    // ===========================================

    // 有効期限が有効、残数が「0」、ステータスが「有効」のチケットを取得
    const dateStr = Utilities.formatDate(new Date(), ssTimezone, "yyyy-MM-dd");
    const userTicketsRowsWithIndex = userTicketSheet.getDataRange().getValues().slice(1).map((row, index) => ({
      rowIndex: index + 1,
      rowData: row
    }));

    const targetUserTicketRow = userTicketsRowsWithIndex.filter(ticket =>
      Utilities.formatDate(ticket.rowData[USER_TICKETS_COL_EXPIRE_DATE], ssTimezone, 'yyyy-MM-dd') >= dateStr &&
      ticket.rowData[USER_TICKETS_COL_STATUS] === "有効" &&
      Number(ticket.rowData[USER_TICKETS_COL_REMAINING_NUM]) === 0
    )
    
    targetUserTicketRow.forEach(ticket => {
      const existReservation = values.some(reservation =>
        reservation[RES_COL_TICKET_ID] === ticket.rowData[USER_TICKETS_COL_TICKET_ID] &&
        reservation[RES_COL_STATUS] === "確定"
      )
      if (!existReservation) {
        userTicketSheet
        .getRange(ticket.rowIndex + 1, USER_TICKETS_COL_STATUS + 1)
        .setValue("消費済み");
      }
    });

    Logger.log('高頻度カレンダーSyncおよびメンテナンス処理が完了しました。');
  } catch(e) {
    Logger.log(`[FATAL ERROR] カレンダー同期中にエラーが発生しました: ${e.toString()}`);
    throw e;

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

/**
 * 予約シートのステータスを確定⇒受講済みに更新（現在よりも過去の予約）
 */
function updatePastReservationsToAttended(resSheet, values, now) {

  const ssTimezone = SPREADSHEET.getSpreadsheetTimeZone();
  const HEADER_ROW_COUNT = 1;

  for (let i = HEADER_ROW_COUNT; i < values.length; i++) {
    const row = values[i];
    const rowIndex = i + 1;
    const status = row[RES_COL_STATUS];
    
    // 日付・時刻情報の取得
    const dateRaw = row[RES_COL_DATE];
    if (!dateRaw || !(dateRaw instanceof Date)) continue;

    const dateStr = Utilities.formatDate(dateRaw, ssTimezone, 'yyyy/MM/dd');
    const timeStr = Utilities.formatDate(row[RES_COL_START_TIME], ssTimezone, 'HH:mm');
    const startDateTime = new Date(`${dateStr} ${timeStr}`);
    
    // A. 過去の予約の場合: ステータスを「受講済み」に更新
    if (status === '確定' && startDateTime < now) {
        resSheet.getRange(rowIndex, RES_COL_STATUS + 1).setValue('受講済み');
    }
  }
}

/**
 * スプレッドシートが編集された時に自動実行されるトリガー
 */
function handleEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  
  // A: 設定シートの特定セル(B2)が変更された場合
  if (sheetName === SHEET_NAME_CONFIG && range.getA1Notation() === "B2") {
    Logger.log("設定バージョンの更新を検知しました。");
    syncConfigToWorkers();
    return;
  }

  // B: ユーザー情報シートが変更された場合
  if (sheetName === SHEET_NAME_USERS) {
    const row = range.getRow();
    // ヘッダー（1行目）以外の編集なら実行
    if (row > 1) {
      // 編集された行のUserIDを取得（A列にある想定）
      const userId = sheet.getRange(row, 1).getValue();
      if (userId) {
        Logger.log(`ユーザー情報の更新を検知しました (Row: ${row}, UserID: ${userId})`);
        syncUserFullData(userId);
        
        // 成功のサイン（背景色一瞬変更）
        range.setBackground("#e2f0fb");
        Utilities.sleep(500);
        range.setBackground(null);
      }
    }
    return;
  }

  // reservationsList シートが変更されたとき（残席情報の更新）
  if (sheetName === SHEET_NAME_RESERVATIONS_LIST) {
    Logger.log("残席情報の更新を検知しました。");
    syncCapacityToWorkers();
    return;
  }

  // userMonthlySubscriptionsシートが変更された場合
  if (sheetName === SHEET_NAME_USER_MONTHLY_SUBSCRIPTIONS) {
    const row = range.getRow();
    // ヘッダー（1行目）以外の編集なら実行
    if (row > 1) {
      // 編集された行のUserIDを取得（A列にある想定）
      const userId = sheet.getRange(row, 1).getValue();      
      if (userId) {
        Logger.log("稽古回数の更新を検知しました。");

        // デフォルトの稽古回数が変更された場合
        if (range.getColumn() === 3) {
          syncUserFullData(userId);
          return;
        }
        // 今月・来月以外の変更は何もしない
        if (range.getColumn() !== 4 && range.getColumn() !== 5) {
          return;
        }
        const lock = LockService.getScriptLock();
        lock.waitLock(10000);
        try {
          const reservationsSheet = SPREADSHEET.getSheetByName(SHEET_NAME_RESERVATIONS);
          const userTicketSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USER_TICKETS);
          const monthlyData = getMonthlyTicketReservations(reservationsSheet, userTicketSheet, userId);
          const monthlyReservations = getMonthlyLessonReservations(reservationsSheet, userId);
  
          if (range.getColumn() === 4) {
            const currentCount = sheet.getRange(row, 4).getValue();
            const currentMonthlyCount = monthlyReservations.thisMonth.length;
  
            if (currentMonthlyCount < currentCount) {
              convertTicketReservationsToMonthly(reservationsSheet, userTicketSheet, monthlyData.thisMonth, currentCount - currentMonthlyCount);
            } else if (currentMonthlyCount > currentCount) {
              const currentConvertCount = currentMonthlyCount - currentCount;
              convertMonthlyReservationsToTicket(reservationsSheet, userTicketSheet, userId, monthlyReservations.thisMonth, currentConvertCount);
            }
          } else if (range.getColumn() === 5) {
            const nextCount = sheet.getRange(row, 5).getValue();
            const nextMonthlyCount = monthlyReservations.nextMonth.length;
            if (nextMonthlyCount < nextCount) {
              convertTicketReservationsToMonthly(reservationsSheet, userTicketSheet, monthlyData.nextMonth, nextCount - nextMonthlyCount);
            } else if (nextMonthlyCount > nextCount) {
              const nextConvertCount = nextMonthlyCount - nextCount;
              convertMonthlyReservationsToTicket(reservationsSheet, userTicketSheet, userId, monthlyReservations.nextMonth, nextConvertCount);
            }
          }
          syncUserFullData(userId);
        } finally {
          lock.releaseLock();
        }
      }
    }
    return;
  }
  // userTicketsシートが変更された場合
  if (sheetName === SHEET_NAME_USER_TICKETS) {
    const row = range.getRow();
    // ヘッダー（1行目）以外の編集なら実行
    if (row > 1) {
      // 編集された行のUserIDを取得（B列にある想定）
      const userId = sheet.getRange(row, 2).getValue();
      if (userId) {
        Logger.log("チケット情報の更新を検知しました。");
        syncUserFullData(userId);
      }
    }
    return;
  }
}

function getMonthlyTicketReservations(
  reservationsSheet,
  userTicketSheet,
  userId
) {
  const monthlyData = {
    thisMonth: [],
    nextMonth: []
  };
  const ssTimezone = SPREADSHEET.getSpreadsheetTimeZone();
  const now = new Date();
  const thisMonth = Utilities.formatDate(now, ssTimezone, 'yyyy-MM');
  const nextMonthDate = new Date(
    Number(Utilities.formatDate(now, ssTimezone, 'yyyy')),
    Number(Utilities.formatDate(now, ssTimezone, 'MM')),
    1
  );
  const nextMonth = Utilities.formatDate(nextMonthDate, ssTimezone, 'yyyy-MM');

  // スライド処理
  const dateString = Utilities.formatDate(now, ssTimezone, "yyyy-MM-dd");
  // 予約シートから情報取得
  const reservationsRowsWithIndex = reservationsSheet.getDataRange().getValues().map((row, index) => ({
    rowIndex: index,
    rowData: row
  }));
  // チケット管理シートからデータ取得
  const userTicketsRowsWithIndex = userTicketSheet.getDataRange().getValues().map((row, index) => ({
      rowIndex: index,
      rowData: row
  }));

  const validUserTicketsRows = userTicketsRowsWithIndex.filter(ticket =>
    // ユーザー絞り込み
    ticket.rowData[USER_TICKETS_COL_USER_ID] === userId &&
    // 有効期限が残っているチケット
    Utilities.formatDate(ticket.rowData[USER_TICKETS_COL_EXPIRE_DATE], SPREADSHEET.getSpreadsheetTimeZone(), 'yyyy-MM-dd') >= dateString &&
    // 未消費のチケットは対象外
    ticket.rowData[USER_TICKETS_COL_PURCHASE_NUM] !== ticket.rowData[USER_TICKETS_COL_REMAINING_NUM]
  ).sort((a, b) => {
    // 有効期限が近い順
    const expireDiff = a.rowData[USER_TICKETS_COL_EXPIRE_DATE] - b.rowData[USER_TICKETS_COL_EXPIRE_DATE];
    if (expireDiff !== 0) {
      return expireDiff;
    }
    // 同じ有効期限なら登録順（シートの行番号）
    return a.rowIndex - b.rowIndex;
  });

  validUserTicketsRows.forEach(ticket => {

    const ticketId = ticket.rowData[USER_TICKETS_COL_TICKET_ID];

    const reservations = reservationsRowsWithIndex.filter(reservation =>
      reservation.rowData[RES_COL_TICKET_ID] === ticketId &&
      reservation.rowData[RES_COL_STATUS] === "確定"
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

    const thisMonthReservations = reservations.filter(reservation => {
      const reservationDate = Utilities.formatDate(
        reservation.rowData[RES_COL_DATE],
        ssTimezone,
        'yyyy-MM'
      );

      return reservationDate === thisMonth;
    });

    const nextMonthReservations = reservations.filter(reservation => {
      const reservationDate = Utilities.formatDate(
        reservation.rowData[RES_COL_DATE],
        ssTimezone,
        'yyyy-MM'
      );

      return reservationDate === nextMonth;
    });

    if (thisMonthReservations.length > 0) {
      monthlyData.thisMonth.push({
        ticket: ticket,
        reservations: thisMonthReservations
      });
    }

    if (nextMonthReservations.length > 0) {
      monthlyData.nextMonth.push({
        ticket: ticket,
        reservations: nextMonthReservations
      });
    }
  });

  return monthlyData;
}

function getMonthlyLessonReservations(
  reservationsSheet,
  userId
) {
    const monthlyData = {
    thisMonth: [],
    nextMonth: []
  };
  const ssTimezone = SPREADSHEET.getSpreadsheetTimeZone();
  const now = new Date();
  const thisMonth = Utilities.formatDate(now, ssTimezone, 'yyyy-MM');
  const nextMonthDate = new Date(
    Number(Utilities.formatDate(now, ssTimezone, 'yyyy')),
    Number(Utilities.formatDate(now, ssTimezone, 'MM')),
    1
  );
  const nextMonth = Utilities.formatDate(nextMonthDate, ssTimezone, 'yyyy-MM');

  // スライド処理
  // 予約シートから情報取得
  const reservationsRowsWithIndex = reservationsSheet.getDataRange().getValues().map((row, index) => ({
    rowIndex: index,
    rowData: row
  }));

  const reservations = reservationsRowsWithIndex.filter(reservation =>
    reservation.rowData[RES_COL_USER_ID] === userId &&
    reservation.rowData[RES_COL_STATUS] === "確定" &&
    reservation.rowData[RES_COL_USAGE_TYPE] === "月謝"
  ).sort((reservation1, reservation2) => {
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

  const thisMonthReservations = reservations.filter(reservation => {
    const reservationDate = Utilities.formatDate(
      reservation.rowData[RES_COL_DATE],
      ssTimezone,
      'yyyy-MM'
    );

    return reservationDate === thisMonth;
  });

  const nextMonthReservations = reservations.filter(reservation => {
    const reservationDate = Utilities.formatDate(
      reservation.rowData[RES_COL_DATE],
      ssTimezone,
      'yyyy-MM'
    );

    return reservationDate === nextMonth;
  });

  monthlyData.thisMonth = thisMonthReservations;
  monthlyData.nextMonth = nextMonthReservations;

  return monthlyData;
}

// チケット再配置用（チケット追加処理で使用）
function getTicketReservations(
  reservationsSheet,
  userTicketSheet,
  userId
) {

  const userTicketsRowsWithIndex = userTicketSheet.getDataRange().getValues().map((row, index) => ({
    rowIndex: index,
    rowData: row
  }));
  
  const reservationsRowsWithIndex = reservationsSheet.getDataRange().getValues().map((row, index) => ({
    rowIndex: index,
    rowData: row
  }));
  
  const ssTimezone = SPREADSHEET.getSpreadsheetTimeZone();
  const dateStr = Utilities.formatDate(new Date(), ssTimezone, 'yyyy-MM-dd');
  // 再配置では残数「0」のチケットも検索対象
  const validUserTicketsRows = userTicketsRowsWithIndex
    .filter(ticket =>
      // 指定ユーザーで絞り込み
      ticket.rowData[USER_TICKETS_COL_USER_ID] === userId &&
      // 有効期限がキャンセル日以降のチケット
      Utilities.formatDate(ticket.rowData[USER_TICKETS_COL_EXPIRE_DATE], ssTimezone, 'yyyy-MM-dd') >= dateStr
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
    });
  
  return validUserTicketsRows.map(ticket => {

    const ticketId = ticket.rowData[USER_TICKETS_COL_TICKET_ID];

    const reservations = reservationsRowsWithIndex.filter(reservation =>
      reservation.rowData[RES_COL_TICKET_ID] === ticketId &&
      reservation.rowData[RES_COL_STATUS] === "確定"
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

    return {
      ticket: ticket,
      reservations: reservations
    };
  });
}

// 現在のチケット予約の紐付けを一旦解除
function clearTicketReservations(
  reservationsSheet,
  userTicketSheet,
  ticketReservationData
) {

  for (const ticketInfo of ticketReservationData) {

    let remainingNum = Number(ticketInfo.ticket.rowData[USER_TICKETS_COL_REMAINING_NUM]);

    for (const reservation of ticketInfo.reservations) {
      // 利用種別をクリア
      reservationsSheet.getRange(reservation.rowIndex + 1, RES_COL_USAGE_TYPE + 1).setValue("");
      // チケットIDをクリア
      reservationsSheet.getRange(reservation.rowIndex + 1, RES_COL_TICKET_ID + 1).setValue("");
      // チケットの残数を1戻す
      remainingNum++;
    }

    // メモリ上の値も更新
    ticketInfo.ticket.rowData[USER_TICKETS_COL_REMAINING_NUM] = remainingNum;
    // チケット残数をまとめて更新
    userTicketSheet
      .getRange(ticketInfo.ticket.rowIndex + 1, USER_TICKETS_COL_REMAINING_NUM + 1)
      .setValue(remainingNum);
  }
}

// 有効期限の近いチケットから再割り当て
function reassignTicketReservations(
  reservationsSheet,
  userTicketSheet,
  ticketReservationData
) {

  const ssTimezone = SPREADSHEET.getSpreadsheetTimeZone();

  // 全チケットに紐づいていた予約を1つの配列にまとめる
  const reservations = ticketReservationData
    .flatMap(ticketInfo => ticketInfo.reservations)
    .sort((a, b) => {
      const dateA = a.rowData[RES_COL_DATE];
      const dateB = b.rowData[RES_COL_DATE];

      if (dateA.getTime() !== dateB.getTime()) {
        return dateA - dateB;
      }

      const timeA = a.rowData[RES_COL_START_TIME] instanceof Date
        ? Utilities.formatDate(
          a.rowData[RES_COL_START_TIME],
          ssTimezone,
          'HH:mm'
        )
        : String(a.rowData[RES_COL_START_TIME]);      
      
      const timeB = b.rowData[RES_COL_START_TIME] instanceof Date
        ? Utilities.formatDate(
          b.rowData[RES_COL_START_TIME],
          ssTimezone,
          'HH:mm'          
        )
        : String(b.rowData[RES_COL_START_TIME]);

      return timeA < timeB ? -1 : timeA > timeB ? 1 : 0;
    });

  let reservationIndex = 0;

  // 有効期限の近いチケットから割り当て
  for (const ticketInfo of ticketReservationData) {
    let remainingNum = Number(ticketInfo.ticket.rowData[USER_TICKETS_COL_REMAINING_NUM]);

    while (remainingNum > 0 && reservationIndex < reservations.length) {
      const reservation = reservations[reservationIndex];

      // 利用種別を設定
      reservationsSheet
        .getRange(reservation.rowIndex + 1, RES_COL_USAGE_TYPE + 1)
        .setValue("チケット");

      // チケットIDを設定
      reservationsSheet
        .getRange(reservation.rowIndex + 1, RES_COL_TICKET_ID + 1)
        .setValue(ticketInfo.ticket.rowData[USER_TICKETS_COL_TICKET_ID]);
      remainingNum--;
      reservationIndex++;
    }

    // メモリ上の値を更新
    ticketInfo.ticket.rowData[USER_TICKETS_COL_REMAINING_NUM] = remainingNum;

    // シートへ反映
    userTicketSheet
      .getRange(ticketInfo.ticket.rowIndex + 1, USER_TICKETS_COL_REMAINING_NUM + 1)
      .setValue(remainingNum);

    // 全予約に割り当て終わったら終了
    if (reservationIndex >= reservations.length) {
      break;
    }
  }
}