// ================================
// ユーザーのクラス情報を登録
// ================================
function registerUserClassGAS(params) {
  const usersSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USERS);

  const userId = params.userId;
  const displayName = params.displayName;
  const className = params.className;
  const upperLimitNumber = params.upperLimitNumber;
  const timestamp = new Date();

  if (!userId || !displayName || !className || !upperLimitNumber) {
    return { success: false, message: "不足パラメータがあります" };
  }

  const limitNumberInt = parseInt(upperLimitNumber, 10);
  const limitNumberIntThisMonth = limitNumberInt;
  const limitNumberIntNextMonth = limitNumberInt;
  const CONFIG = getConfig();
  if (isNaN(limitNumberInt) || !CONFIG.CLASS_INFO.UPPER_LIMIT_NUMBER.includes(limitNumberInt)) {
    return { success: false, message: "授業回数が不正です。" };
  }

  // 既存ユーザーを確認（userId が存在するか）
  const searchRange = usersSheet.getRange("A:A"); // A列のみを対象
  const foundCell = searchRange.createTextFinder(userId).matchEntireCell(true).findNext();

  if (foundCell) {
    // 既存ユーザーがいる場合は更新
    const actualRow = foundCell.getRow();
    const newData = [displayName, className, limitNumberInt, limitNumberIntThisMonth, limitNumberIntNextMonth, timestamp];
    usersSheet.getRange(actualRow, 2, 1, 7).setValues(newData);
  } else {
    // 新規ユーザーを追加
    usersSheet.appendRow([userId, displayName, className, limitNumberInt, limitNumberIntThisMonth, limitNumberIntNextMonth, timestamp]);
  }

  CacheService.getScriptCache().remove('user_' + params.userId);


  // 1. ユーザーの全予約情報を取得（あなたが作成した関数を呼び出し）
  // これにより、登録直後（通常は空）の予約リストが正しく生成されます
  const reservations = getAllReservationsForUser(userId);

  // 2. 最新の残席データを取得（原本から直接生成）
  const capacityData = getCapacityData();

  // 画面表示用にユーザ情報を返却
  const userInfoFull = {
    data: {
      userId: userId,
      displayName: displayName,
      className: className,
      upperLimitNumber: limitNumberInt,
      upperLimitNumberThisMonth: limitNumberIntThisMonth,
      upperLimitNumberNextMonth: limitNumberIntNextMonth
    },
    myReservedDates: reservations.myReservedDates,
    myAttendedDates: reservations.myAttendedDates
  };

  // 非同期でWorkersも更新（次回の起動用）
  syncUserFullData(userId);

  return { 
    success: true, 
    userInfo: userInfoFull,      // 階層をWorkersの返却形式に合わせる
    capacityData: capacityData,  // これを返すことでWorkersに聞きに行く必要がなくなる
    config: CONFIG,
    message: "クラス情報を登録しました" 
  };
}

/**
 * 補助：現在の残席状況をスプレッドシートから直接取得する（Workersを介さない）
 */
function getCapacityData() {
  const today = new Date();
  const listSheet = SPREADSHEET.getSheetByName(SHEET_NAME_RESERVATIONS_LIST);  
  const ssTimezone = SPREADSHEET.getSpreadsheetTimeZone();

  const result = getCapacityForMonthInternalAll(listSheet, ssTimezone);

  return result.capacityData; 
}

/**
 * LIFFから渡されたアクセストークンを検証し、ユーザーIDを取得後、スプレッドシートから情報を検索する
 * @param {string} accessToken - LIFFから取得したアクセストークン
 * @return {GoogleAppsScript.Content.TextOutput} ユーザー情報を含むJSON
 */
function verifyTokenAndGetUserInfo(accessToken) {
  try {
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

    // 2. プロフィール取得 API を呼び出す (表示名の取得)
    const profileOptions = {
      'headers': { 'Authorization': 'Bearer ' + accessToken },
      'method': 'get'
    };
    const profileResponse = UrlFetchApp.fetch(LINE_PROFILE_URL, profileOptions);
    const profileData = JSON.parse(profileResponse.getContentText());
    
    const lineUserId = profileData.userId;
    const lineDisplayName = profileData.displayName;

    // 2. キャッシュサービスからユーザー情報を取得（既存のロジックを再利用）
    // NOTE: userIdが確定したため、既存のgetUserInfoのロジック(シート検索とキャッシュ登録)を流用する
    const userInfoResult = getAndCacheUserInfo(lineUserId, lineDisplayName);
    
    // LIFF側に返すJSONには、ユーザー情報と検証で得られた displayName などを統合しても良い
    // 現状はスプレッドシートの情報のみを返します。
    return ContentService
      .createTextOutput(userInfoResult)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    Logger.log("Token verification or fetch failed: " + e.message);
    return ContentService.createTextOutput(
      JSON.stringify({ 
        exists: false, 
        error: true, 
        message: "認証または情報取得に失敗しました。",
        detail: e.message 
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ================================
// ユーザ情報取得用関数
// ================================
function getAndCacheUserInfo(lineUserId, lineDisplayName) {
  // configバージョン情報取得
  const configSheet = SPREADSHEET.getSheetByName(SHEET_NAME_CONFIG);
  if (!configSheet) throw new Error(`${SHEET_NAME_CONFIG} シートが見つかりません。`);
  const configVersion = configSheet.getRange('B2').getValue();

  // --- ① CacheService（ユーザーキャッシュ領域） ---
  const cache = CacheService.getUserCache();
  const cacheKey = "user_" + lineUserId;

  // --- ② キャッシュから読み込み ---
  const cached = cache.get(cacheKey);
  if (cached) {
    // キャッシュヒット時も最新のdisplayNameで更新
    const cachedObj = JSON.parse(cached);
    if (cachedObj.data) {
        cachedObj.data.displayName = lineDisplayName;
        cachedObj.configVersion = configVersion;
    }
    return JSON.stringify(cachedObj);
  }

  // --- ③ スプレッドシート検索ロジック ---
  let found = null;
  try {
    const usersSheet = SPREADSHEET.getSheetByName(SHEET_NAME_USERS);
    if (!usersSheet) throw new Error("Sheet 'users' not found.");

    // **速度改善案: createTextFinderで検索**
    const allUserData = usersSheet.getDataRange().getValues();
    const HEADER_ROW_COUNT = 1;

    let targetRowData = null;

    // A列 (インデックス0) の userId でメモリ内を検索
    for (let i = HEADER_ROW_COUNT; i < allUserData.length; i++) {
        if (allUserData[i][0] === lineUserId) {
            targetRowData = allUserData[i];
            break;
        }
    }

    if (targetRowData) {
        // targetRowData は [UserID, DisplayName, ClassName, UpperLimitNumber, ...]
        found = {
            userId: targetRowData[0],
            displayName: targetRowData[1],
            className: targetRowData[2],
            upperLimitNumber: targetRowData[3],
            upperLimitNumberThisMonth: targetRowData[4],
            upperLimitNumberNextMonth: targetRowData[5]
        };
        // LINE APIから取得した最新のdisplayNameで上書き
        found.displayName = lineDisplayName; 
    }
  } catch (e) {
      Logger.log("Spreadsheet error: " + e.message);
      // エラー時もキャッシュには登録しない
  }

  if (!found) {
    found = { userId: lineUserId, displayName: lineDisplayName };
  }
  // 最終的な結果をJSON文字列として準備
  const result = JSON.stringify({ exists: !!found.className, data: found, configVersion: configVersion });  
  // --- ④ キャッシュ保存（5分間） ---
  if (result.exsits) {
    cache.put(cacheKey, result, CACHE_SECONDS);
  }
  return result; // JSON 文字列を返す
}