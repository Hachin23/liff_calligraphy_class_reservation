// ====================================
// 1. グローバル変数と定数
// ====================================

// 既存のグローバル変数
const GAS_BASE_URL = "$$GAS_ENDPOINT_URL_PLACEHOLDER$$";
const LIFF_ID = "$$LIFF_ID_PLACEHOLDER$$";
const WORKERS_BASE_URL = "$$WORKERS_ENDPOINT_URL_PLACEHOLDER$$";
const VERSION_KEY = 'config_version';
const CONFIG_KEY = 'reservation_config_data';

// 予約画面用
let AVAILABLE_CAPACITY_DATA = {}; // { 'YYYY-MM-DD': [{ startTime: 'HH:mm', className: '...', remainingCapacity: N }, ...] }
let MY_RESERVIONS = {};
let MY_ATTENDED_DATES = { data: [], lastFetch: 0 };
let CURRENT_SCREEN_DATE = new Date(); // 予約画面のカレンダー表示月
const MAX_RESERVABLE_MONTHS = 1; // (今月、来月)
const CACHE_EXPIRATION_MS = 2 * 60 * 1000; // 1分(Workersが最新に反映されるまでで問題なし)

// 予約カレンダーの曜日初めの切り替え用
let FIRST_DAY_OF_THE_WEEK = "";

// 予約画面用DOM要素
const reservationArea = document.getElementById("reservationArea");
const calendarContainerRes = document.getElementById('calendar-container-res'); // 予約画面のカレンダーグリッド本体
const currentMonthSpanRes = document.getElementById('current-month-res');       // 予約画面の月表示
const prevMonthBtnRes = document.getElementById('prev-month-btn-res');         // 予約画面の前月ボタン
const nextMonthBtnRes = document.getElementById('next-month-btn-res');         // 予約画面の次月ボタン
const selectionDitailsModel = document.getElementById('selectionDitails-model');         // 予約画面の次月ボタン
const selectionDetails = document.getElementById('selectionDetails'); 
const selectedDateText = document.getElementById('selectedDateText');
const courseName = document.getElementById('courseName');
const closeModalButton = document.getElementById('closeModalButton');
const availableClassesList = document.getElementById('availableClassesList');
const classInfo = document.getElementById('userClassInfo');
const upperLimitMessageArea = document.getElementById('upperLimitMessageArea');

// カスタムモーダル要素
const customModal = document.getElementById('custom-modal');
const modalTitle = document.getElementById('modal-title');
const modalMessage = document.getElementById('modal-message');
const modalConfirmBtn = document.getElementById('modal-confirm-btn');
const modalCancelBtn = document.getElementById('modal-cancel-btn');
let currentConfirmCallback = null;

// ====================================
// 2. メイン処理と初期化
// ====================================

document.addEventListener("DOMContentLoaded", main);
async function main() {
  document.getElementById("main").classList.remove("hidden");
  
  try {
      // 初期化は並列実行不可
      await liff.init({ liffId: LIFF_ID });
      if (!liff.isLoggedIn()) {
          liff.login(); 
          return;
      }
      // モーダル設定
      setupModalListeners();

      console.time("初期表示までの時間 start");

      // GASを叩かず、Workersから全情報を一度に取得する
      await fetchInitialAppData();
      
      console.time("初期表示までの時間 end");

  } catch (err) {
      console.error('LIFF init failed or subsequent process failed:', err);
      document.getElementById("errordisp").textContent = "初期化に失敗しました。LINEアプリの設定をご確認ください。";
  }
}

/**
 * Workers KVから「ユーザー情報」「設定」「残席」「自分の予約」を一括で取得して描画する
 */
async function fetchInitialAppData() {
  const profile = await liff.getProfile();
  const userId = profile.userId;
  const displayName = profile.displayName; // 新規登録時に使用
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth() + 1;
  const monthKey = `${year}-${String(month).padStart(2, '0')}`;

  const cachedJSON = localStorage.getItem("APP_DATA_CACHE");
  let json = null;

  if (cachedJSON) {
    json = getInitDispFullCache(monthKey);
    if (json == null) {
      // 期限切れの場合、Workersから取得しに行く
      json = await getWorkersDataJson(userId);
    }
  } else {
    // ローカルストレージに登録されていない場合、Workersへ問い合わせ
    json = await getWorkersDataJson(userId);
  }
  if (!json.success) throw new Error("データの取得に失敗しました。");

  FIRST_DAY_OF_THE_WEEK = json.config.CALENDAR_INFO.FIRST_DAY_OF_WEEK;

  // --- 分岐点：Workersにユーザー情報があるか ---
  if (json.userInfo && json.userInfo.data) {
    // 【既存ユーザー】
    console.log("登録済みユーザーです。カレンダーを表示します。");

    saveToCache(json.capacityData, json.userInfo, json.config, monthKey);

    switchPage(false, json.userInfo.data);
    renderReservationCalendar(today, 'loaded', AVAILABLE_CAPACITY_DATA[monthKey]?.data, MY_RESERVIONS[monthKey]?.data, MY_ATTENDED_DATES.data);
    document.getElementById('loading').style.display = 'none';

  } else {
    // 【新規ユーザー】
    console.log("未登録ユーザーです。授業選択画面を表示します。");
    
    // Workersから返ってきた config を使ってセットアップ
    document.getElementById("user-select").classList.remove("hidden");
    document.getElementById('loading').style.display = 'none';
    
    // お使いの setupClassSelect を呼び出し
    setupClassSelect(userId, displayName, json.config);
  }
}

// -----------------------------
// 授業選択画面の表示処理 
// -----------------------------
function setupClassSelect(userId, displayName, config) {
  const classSelect = document.getElementById("classSelect");
  const countSelect = document.getElementById("countSelect");
  const submitBtn = document.getElementById("classSubmitBtn");

  const classNames = config.CLASS_INFO.CLASS_NAME;
  classSelect.innerHTML = `<option value="">選択してください</option>`;
  classNames.forEach((name, index) => {
    classSelect.innerHTML += `<option value="${index}">${name}</option>`;
  });

  classSelect.addEventListener("change", () => {
    const idx = classSelect.value;

    if (idx === "") {
      countSelect.disabled = true;
      submitBtn.classList.add("hidden");
      countSelect.innerHTML = `
        <option value="">クラスを先に選んでください</option>`;
      return;
    }

    const upperArray = config.CLASS_INFO.UPPER_LIMIT_NUMBER;
    countSelect.disabled = false;
    countSelect.innerHTML = `<option value="">選択してください</option>`;
    upperArray.forEach(n => {
      countSelect.innerHTML += `<option value="${n}">${n}回</option>`;
    })

    submitBtn.classList.add("hidden");
  });

  countSelect.addEventListener("change", () => {
    if (countSelect.value !== "") {
      submitBtn.classList.remove("hidden");
    } else {
      submitBtn.classList.add("hidden");
    }
  });

  submitBtn.addEventListener("click", () => {
    const selectedClassIndex = classSelect.value;
    const selectedUpperLimitNumber = countSelect.value;
    confirmClassRegister(userId, displayName, selectedClassIndex, selectedUpperLimitNumber, config);
  });
}

// ------------------------------
// クラス登録確認モーダル表示
// ------------------------------
function confirmClassRegister(userId, displayName, classIndex, upperLimit, config) {
    const className = config.CLASS_INFO.CLASS_NAME[classIndex];

    const message = `クラスは「${className} 月${upperLimit}回」でよろしいですか？`;
    showCustomModal(
        'クラス登録',
        message,
        async () => {
            await registerUserClass(userId, displayName, classIndex, upperLimit, config);
        }
    );
}

async function registerUserClass(userId, displayName, classIndex, upperLimitNumber, config) {
  const className = config.CLASS_INFO.CLASS_NAME[classIndex];

  // 送信データをオブジェクトでまとめる
  const payload = {
    mode: "registerUserInfo",
    userId: userId,
    displayName: displayName,
    className: className,
    upperLimitNumber: upperLimitNumber
  };
  console.log("registerUserInfo payload", payload);
  
  const formBody = new URLSearchParams(payload);
  try {
    const res = await fetch(GAS_BASE_URL, { 
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: formBody
    });

    const json = await res.json();

    if (json.success) {      
      // セッションストレージには基本情報(data)を保存
      const today = new Date();
      const monthKey = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
      
      // キャッシュを最新のデータで上書きし、lastFetchを「今」にする
      saveToCache(json.capacityData, json.userInfo, json.config, monthKey);

      alert("クラスの登録が完了しました！");
      switchPage(true, json.userInfo);
    } else {
      alert("クラスの登録に失敗しました！");
    }
  } catch (e) {
    alert("通信エラーが発生しました");
    console.log(e);
  }
}

// 操作するユーザ側でメッセージを送信する
// NOTE: 現状、使用していないが、念のため残しておく
function sendResigterResultMessage(messageText) {
  // 1. LIFFが初期化されているか、かつLINEアプリ内で動作しているかを確認
  if (!liff.isInClient()) {
      alert("メッセージ送信はLINEアプリ内でのみ実行可能です。");
      return;
  }
  // 2. メッセージを送信
  liff.sendMessages([{ type: 'text', text: messageText }])
      .then(() => {
          console.log('Message sent successfully!');
      })
      .catch((err) => {
          console.error('Error sending message:', err);
      });
}

// ------------------------------
// 画面切り替え
// ------------------------------
async function switchPage(registerFlag, userInfoJson = {}) {
  const reservation = reservationArea;
  const userSelect = document.getElementById("user-select");

  if (registerFlag) {
      userSelect.classList.add("hidden");
  }
  reservation.classList.remove("hidden");
  setupReservationScreen();
}

// ====================================
// 予約画面 ロジック
// ====================================

/**
 * 予約画面の初期設定と月移動リスナーのセットアップ
 */
function setupReservationScreen() {
    // 画面切り替え時にカレンダーをリセットして描画開始
    CURRENT_SCREEN_DATE.setDate(1); 
    // 時刻情報もクリアして、0時0分0秒に設定
    CURRENT_SCREEN_DATE.setHours(0, 0, 0, 0);
    fetchAndRenderCapacity(CURRENT_SCREEN_DATE);

    // 予約画面専用のボタンにリスナーを設定
    if (!prevMonthBtnRes.hasAttribute('data-res-listener')) {
        prevMonthBtnRes.addEventListener('click', () => {
            CURRENT_SCREEN_DATE.setMonth(CURRENT_SCREEN_DATE.getMonth() - 1);
            fetchAndRenderCapacity(CURRENT_SCREEN_DATE);
        });
        nextMonthBtnRes.addEventListener('click', () => {
            CURRENT_SCREEN_DATE.setMonth(CURRENT_SCREEN_DATE.getMonth() + 1);
            fetchAndRenderCapacity(CURRENT_SCREEN_DATE);
        });
        prevMonthBtnRes.setAttribute('data-res-listener', 'true');
    }
}

/**
 * 予約画面のカレンダー描画と、残席情報の取得・表示をメインで処理する
 * @param {Date} date - 表示する月
 */
async function fetchAndRenderCapacity(date) {
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
  // セッションストレージからユーザ情報取得
  const currentUser = getSessionUserInfo();

  // キャッシュをチェック
  const fullCache = getValidFullCache(monthKey);

  if (fullCache) {
    console.log("全データ有効なキャッシュがあるため、描画のみ実行");
    updateClassInfoUI(currentUser, monthKey); // UIの文字更新
    renderReservationCalendar(date, 'loaded', fullCache.capacity, fullCache.reserved, fullCache.attended);
    return;
  }

  // キャッシュがない場合、カレンダーのUIを先に描画する (ローディング表示)
  renderReservationCalendar(date, 'loading');
  // ユーザのクラス・回数を画面上部に表示
  updateClassInfoUI(currentUser, monthKey); // UIの文字更新

  // 2. GASから統合されたカレンダー情報を取得する
  try {

    // ★ userId もパラメータに含める！
    const url = `${WORKERS_BASE_URL}?year=${year}&month=${month}&userId=${currentUser.userId}`;    
    // 1回の通信ですべて取得
    const res = await fetch(url);
    const json = await res.json();
    
    if (json.success) {
      saveToCache(json.capacityData, json.userInfo, json.config, monthKey);
      const fullCache = getValidFullCache(monthKey); // キャッシュから最新の形を取得
      renderReservationCalendar(date, 'loaded', fullCache.capacity, fullCache.reserved, fullCache.attended);
    }
  } catch (e) {
      console.error("カレンダー情報取得時の通信エラー", e);
  }
}

// ------------------------------
// 予約画面のカレンダー描画ロジック 
// ------------------------------
function renderReservationCalendar(date, status, capacityData = {}, myReservations = [], myAttendedDates = []) {
  
  console.log("描画データ確認:", { capacityData, myReservations, myAttendedDates});
  // 上限到達エリアの初期化
  upperLimitMessageArea.innerText = "";
  upperLimitMessageArea.classList.add("hidden");

  const year = date.getFullYear();
  const month = date.getMonth(); // 0-11
  
  // 予約画面専用のDOM要素を参照
  currentMonthSpanRes.textContent = `${year}年 ${month + 1}月`; 
  calendarContainerRes.innerHTML = ''; // クリア

  const today = new Date();
  today.setHours(0, 0, 0, 0); 
  const firstDayOfMonth = new Date(year, month, 1);
  const lastDayOfMonth = new Date(year, month + 1, 0);
  
  // 予約可能月制限 (MAX_RESERVABLE_MONTHSに基づく)
  // MAX_RESERVABLE_MONTHS はグローバル変数に定義済みとする
  const maxReservableDateBoundary = new Date(today.getFullYear(), today.getMonth() + MAX_RESERVABLE_MONTHS, 1);

  // 【月移動ボタン制御】
  prevMonthBtnRes.disabled = (year === today.getFullYear() && month === today.getMonth());
  nextMonthBtnRes.disabled = (firstDayOfMonth.getTime() >= maxReservableDateBoundary.getTime());

  // 【曜日のヘッダー作成】
  const daysOfWeek = FIRST_DAY_OF_THE_WEEK === "月曜日" ? ['月', '火', '水', '木', '金', '土', '日'] : ['日', '月', '火', '水', '木', '金', '土'];  
  let calendarHtml = '';
  // daysOfWeek.forEach(day => {
  //   calendarHtml += `<div class="calendar-day-header">${day}</div>`;
  // });
  let day_of_week_red_css = "day_of_week_red";
  let day_of_week_blue_css = "day_of_week_blue";
  for (let i = 0; i < daysOfWeek.length; i++) {
    let css_serector = '';
    if (FIRST_DAY_OF_THE_WEEK === "月曜日") {
      if (i === 5) {
        css_serector = day_of_week_blue_css;
      } else if (i === 6) {
        css_serector = day_of_week_red_css;
      }
    } else {
      if (i === 6) {
        css_serector = day_of_week_blue_css;
      } else if (i === 0) {
        css_serector = day_of_week_red_css;
      }
    }
    calendarHtml += `<div class="calendar-day-header ${css_serector}">${daysOfWeek[i]}</div>`;
  }

  

  // 【1日の開始曜日までの空セルを作成】
  let startDayOfWeek = firstDayOfMonth.getDay();
  if (FIRST_DAY_OF_THE_WEEK === "月曜日") {
    startDayOfWeek = (startDayOfWeek === 0) ? 6 : startDayOfWeek - 1;
  }

  for (let i = 0; i < startDayOfWeek; i++) {
      calendarHtml += '<div class="calendar-cell inactive"></div>';
  }
    
  const currentUser = getSessionUserInfo();
  const upperLimit = today.getMonth() === month ? currentUser.upperLimitNumberThisMonth : currentUser.upperLimitNumberNextMonth;
  const monthString = `${year}-${String(month + 1).padStart(2, '0')}`;
  const AttendedCount = myAttendedDates.filter(dateTimeString => dateTimeString.includes(monthString)).length;
  const reservedCount = myReservations.filter(dateTimeObj => {
          const keys = Object.keys(dateTimeObj);
          return keys.some(key => key.includes(monthString));
        }).length;
  // 受講済みで上限到達か
  const userAttendedLimitReached = AttendedCount === upperLimit;
  // 予約数と受講数の合計で上限到達か
  const userLimitReached = (reservedCount + AttendedCount) == upperLimit;

  // ⭐ 日付セルを作成
  for (let day = 1; day <= lastDayOfMonth.getDate(); day++) {
    const dateString = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    const currentDateOnly = new Date(dateString); 
    
    let dayClass = 'calendar-cell';
    let capacityInfo = ''; 
    let isReservable = false;
    let isMyReserved = false; // 予約済みフラグを追加
    let isMyAttended = false;
    // capacityData は { 'YYYY-MM-DD': [{ ... }] } の形式
    const dayCapacity = capacityData[dateString] || [];
    const myAttendedDateCheck = myAttendedDates.some(dateTimeString => dateTimeString.includes(dateString));
    
    // 過去
    if (currentDateOnly < today || myAttendedDateCheck) {
      dayClass += ' inactive';
      // 受講済みチェック(過去日は授業なし判定と同じになるので、ここでチェック)
      if (myAttendedDateCheck) {
        dayClass += ' my-attended';
        isMyAttended = true;
      }
    } else {
      // --- 授業なしの判定 ---
      if (dayCapacity.length === 0) {
          // 授業なし：提案色（薄い灰色）の inactive を使用
          dayClass += ' no-lesson inactive'; // 授業なしの日
      } else {
        // --- 授業あり（予約可能/満席の判定） ---
        const totalRemaining = dayCapacity.reduce((sum, item) => sum + item.remainingCapacity, 0);

        if (totalRemaining > 0 && !userLimitReached) {
          // 空席あり：緑 (reservable clickable)
          if (currentDateOnly.getFullYear() === today.getFullYear() &&
            currentDateOnly.getMonth() === today.getMonth() &&
          currentDateOnly.getDate() === today.getDate())
          {
            dayClass += ' today-contact available clickable'; 
            capacityInfo = '要連絡';
          } else {
            dayClass += ' available clickable';
            capacityInfo = '予約可';
            isReservable = true;
          }
        } else if (userLimitReached) {
          dayClass += ' limit-reached';
          if (!userAttendedLimitReached) {
            dayClass += ' clickable';
          } else {
            dayClass += ' inactive';
          }
          capacityInfo = '予約不可';
        } else {
          // 満席：赤 (fully-booked full)
          dayClass += ' fully-booked full';
          capacityInfo = '満席';
        }
        
        // --- 予約済みの判定 ---
        // myReservations は 'YYYY-MM-DD' の日付文字列の配列と想定なのでsome + inculudesで判定（実際は、'YYYY-MM-DD HH:mm'　リスト表示で必要）
        const reservedCheck = myReservations.some(dateTimeObj => {
          const keys = Object.keys(dateTimeObj);
          return keys.some(key => key.includes(dateString));
        });
        if (reservedCheck) {
          // 予約済みの日：青 (my-reserved)
          dayClass += ' my-reserved available';
          if (userLimitReached) {
            dayClass += ' clickable';  
          }
          // 予約済みの場合は下線を緑にしたい
          dayClass = dayClass.replace('limit-reached ', '');
          isMyReserved = true;
          capacityInfo = '';
        }
      }
    }
    
    // ローディング中の表示
    // if (status === 'loading') {
    //     capacityInfo = '読込中...';
    //     dayClass = 'calendar-cell loading'; // ロード中は上書き
    // }

    calendarHtml += `
        <div class="${dayClass}" data-date="${dateString}">
            <span class="date-number">${day}</span>
            ${isMyAttended ? '<span class="my-attended-badge">受講済</span>' : ''}
            ${isMyReserved ? '<span class="my-reserved-badge">予約済</span>' : ''} 
            ${isReservable || dayCapacity.length > 0 ? `<div class="capacity-indicator">${capacityInfo}</div>` : ''}
        </div>
    `;
  }
  
  // ⭐️ 予約画面専用のカレンダーコンテナに書き込む
  calendarContainerRes.innerHTML = calendarHtml;

  // ⭐ リスナー再設定 (reservable clickableな要素のみ)
  if (status === 'loaded') {
      calendarContainerRes.querySelectorAll('.calendar-cell.clickable').forEach(cell => {
          cell.addEventListener('click', (event) => selectDate(event.currentTarget.dataset.date));
      });
  }
  // 上限到達時のメッセージ表示
  if (userLimitReached) {
    upperLimitMessageArea.classList.remove("hidden");
    if (!userAttendedLimitReached) {
      // upperLimitMessageArea.innerHTML = `<div class='reservedMsg'>今月の予約上限数（${upperLimit}回）に到達しました。</div>`;
    } else {
      //受講上限到達
      upperLimitMessageArea.innerHTML = `<div class='attendedMsg'>今月の稽古お疲れ様でした🙌</div>`;
    }
  }
}

// ------------------------------
// 日付がクリックされたときの処理
// ------------------------------
function selectDate(dateString) {
  selectedDateText.textContent = `📅 ${dateString} 稽古一覧`;
  courseName.textContent = "クラス：一般・おとな美文字";
  closeModalButton.addEventListener('click', closeReservationModal);
  selectionDitailsModel.classList.remove('hidden');
  
  // 該当日の残席情報を AVAILABLE_CAPACITY_DATA から取得し、リストを描画
  const monthKey = `${CURRENT_SCREEN_DATE.getFullYear()}-${String(CURRENT_SCREEN_DATE.getMonth() + 1).padStart(2, '0')}`;
  const monthCapacity = AVAILABLE_CAPACITY_DATA[monthKey]?.data || {};
  const dayCapacity = monthCapacity[dateString] || [];

  // dateString を渡してボタンのデータ属性に持たせる
  renderAvailableClassesList(dayCapacity, dateString, monthKey);
}

// ------------------------------
// 予約可能クラスのリストを描画
// ------------------------------
function renderAvailableClassesList(classes, dateString, monthKey) {
  // セッションストレージからユーザ情報取得
  const now = new Date();
  const currentUser = getSessionUserInfo();
  const currentMonthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  const upperLimit = currentMonthKey === monthKey ? currentUser.upperLimitNumberThisMonth : currentUser.upperLimitNumberNextMonth;

  let listHtml = '';

  const monthReservation = MY_RESERVIONS[monthKey].data || {};
  const reservedCount = monthReservation.length;
  const AttendedCount = MY_ATTENDED_DATES.data.filter(item => item.includes(monthKey)).length;
  const userLimitReached = (reservedCount + AttendedCount) == upperLimit;

  classes.forEach(item => {
    // MY_RESERVIONSから取得して、予約済み時間を特定
    const reservation = monthReservation.find(dateTimeObj => {
      return Object.keys(dateTimeObj).some(key => key.includes(`${dateString} ${item.startTime}`));
    });
    const isReserved = !reservation ? false : true;
    const isFull = item.remainingCapacity <= 0;
    let buttonHtml = '';
    let noButtonHtml = '';

    // -----------------------------------------------------------------
    // A. 自分が予約済みの場合: キャンセルボタンを表示
    // -----------------------------------------------------------------
    if (isReserved) {
      const reservationId = reservation[`${dateString} ${item.startTime}`].reservationId;
      const cancellableUntil = reservation[`${dateString} ${item.startTime}`].cancellableUntil;
      const cancellableUntilDate = new Date(cancellableUntil);
      if (cancellableUntilDate.getTime() > now.getTime()) {
        buttonHtml = `
              <div class="reservation-area-container">
                <span class="status-text reserved-info">${item.startTime} - ${item.endTime} ✅予約済み</span><br>
                <span class="reserved-class">取消期限:${cancellableUntil}</span>
              </div>
              <button class="class-select-button is-reserved-cancel" 
                      data-action="cancel" 
                      data-date="${dateString}" 
                      data-time="${item.startTime} - ${item.endTime}"
                      data-reservation-id="${reservationId}">
                  キャンセル
              </button>
          `;
      } else {
        noButtonHtml = `
            <span class="status-text is-unavailable">${item.startTime} - ${item.endTime}</span><br>
            <span class="unavailable-reason">※キャンセル期限切れのためキャンセル不可</span>
          `;
      }
    // -----------------------------------------------------------------
    // B. 予約可能で、満席でも上限でもない場合: 予約ボタンを表示
    // -----------------------------------------------------------------
    } else if (!isFull && !userLimitReached) {
      // 残席数はあるけど、当日予約は不可。
      const dateStringDate = new Date(dateString);
      if (dateStringDate.getFullYear() === now.getFullYear() &&
        dateStringDate.getMonth() === now.getMonth() &&
        dateStringDate.getDate() === now.getDate()
      ) {
        const endTimeArray = item.endTime.split(":");
        const isClassIsOver = now.getTime() > new Date(dateStringDate.getFullYear(), dateStringDate.getMonth(), dateStringDate.getDate(), endTimeArray[0], endTimeArray[1]).getTime();
        if (isClassIsOver) {
          noButtonHtml = `
            <span class="status-text is-unavailable">${item.startTime} - ${item.endTime}</span><br>
            <span class="unavailable-reason">※この稽古は終了しているため、予約できません。</span>
          `;
        } else {
          noButtonHtml = `
            <span class="status-text is-unavailable">${item.startTime} - ${item.endTime}</span><br>
            <span class="unavailable-reason">※当日予約はLINEにて直接ご連絡お願いします。</span>
          `;
        }
      } else {
        buttonHtml = `
            <div class="reservation-area-container">
              <span class="status-text available-info">${item.startTime} - ${item.endTime}</span>
              <span class="remaining-class-number"> 👤 残り${item.remainingCapacity}席</span>
            </div>
            <button class="class-select-button is-available-reserve" 
                    data-action="reserve" 
                    data-lesson-id="${item.lessonId}" 
                    data-date="${dateString}" 
                    data-time="${item.startTime}"
                    data-display-time="${item.startTime} - ${item.endTime}">
                予約
            </button>
        `;
      }
    } else {
      let reason = isFull ? '満席' : `稽古予約回数の上限（${upperLimit}回）到達`;
         noButtonHtml = `
            <span class="status-text is-unavailable">${item.startTime} - ${item.endTime}</span>
            <span class="remaining-class-number"> 👤 残${item.remainingCapacity}席</span><br>
            <span class="unavailable-reason">※${reason}のため予約不可</span>
         `;
    }
    if (buttonHtml) {
      listHtml += `<div class="time-slot-container"><div class="time-slot-flex-container">${buttonHtml}</div></div>`;
    } else {
      listHtml += `<div class="time-slot-container">${noButtonHtml}</div>`;
    }
  });
  availableClassesList.innerHTML = listHtml;
  
  // 予約ボタンのリスナー設定
  document.querySelectorAll('.is-available-reserve').forEach(button => {
      button.addEventListener('click', (event) => confirmReservation(event.currentTarget));
  });
  // キャンセルボタンのリスナー設定
  document.querySelectorAll('.is-reserved-cancel').forEach(button => {
    button.addEventListener('click', (event) => confirmReservationCancel(event.currentTarget));
  });
}

// ------------------------------
// 予約確認モーダル表示
// ------------------------------
function confirmReservation(buttonElement) {
  closeReservationModal();
  const lessonId = buttonElement.dataset.lessonId;
  const dateString = buttonElement.dataset.date;
  const time = buttonElement.dataset.time; // 開始時間 HH:mm
  const displayTime = buttonElement.dataset.displayTime // 開始時間 - 終了時間
  // セッションストレージからユーザ情報取得
  const currentUser = getSessionUserInfo();
  const classNameText = currentUser.className; //ユーザのクラス名を送信
  const userId = currentUser.userId; //ユーザIDを送信

  // const message = `${dateString} ${displayTime} の ${classNameText} を予約します。よろしいですか？`;
  const message = `予約日：${dateString}<br>予約時間：${displayTime}<br>クラス名：${classNameText}<br>予約してよろしいですか？`;

  showCustomModal(
      `予約の確定`,
      message,
      async () => {
          await handleReservation(lessonId, dateString, time, classNameText, userId);
      }
  );
}

// ------------------------------
// 予約確定処理（GASと通信）
// ------------------------------
async function handleReservation(lessonId, dateString, time, classNameText, userId) {
  const accessToken = liff.getAccessToken();

  const payload = { 
      mode: "makeReservation", 
      accessToken: accessToken,
      userId: userId, 
      lessonId: lessonId,
      date: dateString, // YYYY-MM-DD
      time: time,       // HH:mm
      className: classNameText
  };
  const formBody = new URLSearchParams(payload);

  try {
      const res = await fetch(GAS_BASE_URL, { 
          method: "POST", 
          headers: { "Content-Type": "application/x-www-form-urlencoded" }, 
          body: formBody 
      });
      const json = await res.json();

      if (json.success) {
        alert("予約が完了しました！");

        // 1. 最新のデータをキャッシュに保存 (複数月対応版の saveToCache を使用)
        // GASのレスポンスに capacityData と userInfo が含まれている必要があります
        const today = new Date();
        const monthKey = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
        saveToCache(json.capacityData, json.userInfo, null, monthKey);

        sendLiffMessage(`稽古予約：${json.reservationDateTime}\n取消期限：${json.cancellableUntil}まで`);
        // 選択エリアは非表示にする
        selectionDitailsModel.classList.add('hidden');
        // 予約成功後、カレンダーを再描画して残席情報を更新
        fetchAndRenderCapacity(CURRENT_SCREEN_DATE);

      } else {
          alert("予約に失敗しました: " + (json.message || "残席がないか、上限を超えています。"));
      }
  } catch (e) {
      alert("通信エラーが発生しました");
      console.error("予約通信エラー:", e);
  }
}

// ------------------------------
// キャンセル処理（カスタムモーダル）
// ------------------------------
function confirmReservationCancel(buttonElement) {
  closeReservationModal();
  const dateString = buttonElement.dataset.date;
  const time = buttonElement.dataset.time;
  const reservationId = buttonElement.dataset.reservationId;
  // セッションストレージからユーザ情報取得
  const currentUser = getSessionUserInfo();
  const classNameText = currentUser.className; //ユーザのクラス名を送信
  const userId = currentUser.userId

  //const message = `${dateString} ${time} の ${classNameText} をキャンセルします。よろしいですか？`;
  const message = `予約日：${dateString}<br>予約時間：${time}<br>クラス名：${classNameText}<br>予約をキャンセルしてよろしいですか？`;

  showCustomModal(
      `予約のキャンセル`,
      message,
      async () => {
          await executeCancellation(userId, reservationId);
      }
  );
}

// ------------------------------
// GASへのキャンセルAPIコール
// ------------------------------
async function executeCancellation(userId, reservationId) {
  const accessToken = liff.getAccessToken();

  const payload = { 
    mode: "cancelReservation",
    accessToken: accessToken,
    userId: userId,
    reservationId: reservationId 
  };
  const formBody = new URLSearchParams(payload);

  try {
      const res = await fetch(GAS_BASE_URL, { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: formBody });
      const json = await res.json();
      
      if (json.success) {
        alert("キャンセルが完了しました。");

        // 1. 最新のデータをキャッシュに保存 (複数月対応版の saveToCache を使用)
        const today = new Date();
        const monthKey = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
        saveToCache(json.capacityData, json.userInfo, null, monthKey);

        sendLiffMessage(`キャンセル：${json.cancelDateTime}`);
        // 選択エリアは非表示にする
        selectionDitailsModel.classList.add('hidden');
        // 予約成功後、カレンダーを再描画して残席情報を更新
        fetchAndRenderCapacity(CURRENT_SCREEN_DATE);
      } else {
          alert("キャンセルに失敗しました: " + json.message);
      }
  } catch (e) {
      alert("通信エラーが発生しました");
      console.error("キャンセル通信エラー:", e);
  }
}

// ------------------------------
// カスタムモーダル表示ロジック
// ------------------------------
const showCustomModal = (title, message, onConfirm) => {
    modalTitle.textContent = title;
    // modalMessage.textContent = message;
    modalMessage.innerHTML = message;

    currentConfirmCallback = onConfirm; 
    customModal.classList.remove('hidden');
};

const hideCustomModal = () => {
    customModal.classList.add('hidden');
    currentConfirmCallback = null;
};

function closeReservationModal() {
    selectionDitailsModel.classList.add('hidden');
}

function showReservationModal() {
    selectionDitailsModel.classList.remove('hidden');
}

// ====================================
// カスタムモーダル イベントリスナー設定
// ====================================
function setupModalListeners() {
    // 承認ボタンの処理
    modalConfirmBtn.addEventListener('click', async () => {
        if (currentConfirmCallback) {
          modalConfirmBtn.disabled = true;
          modalCancelBtn.disabled = true;

            try {
                await currentConfirmCallback();
            } catch (error) {
                console.error("Confirm callback failed:", error);
            } finally {
              modalConfirmBtn.disabled = false;
              modalCancelBtn.disabled = false;
            }
        }
        hideCustomModal();
    });

    // キャンセルボタンの処理
    modalCancelBtn.addEventListener('click', () => {
      hideCustomModal();
      showReservationModal();
    });
}

// ==========================================
// セッションストレージに設定したユーザ情報を取得
// ※セッションストレージからローカルストレージに変更
// ==========================================
function getSessionUserInfo() {
  const localStorageInfo = localStorage.getItem("APP_DATA_CACHE");

  if (!localStorageInfo) {
    alert("ユーザ情報が取得できませんでした。一度、画面を閉じて開き直してください。");
    liff.closeWindow();
    return null;
  }
  // JSON文字列をオブジェクトに戻す
  const localStorageData = JSON.parse(localStorageInfo);
  return localStorageData.userInfo.data;
}

function sendLiffMessage(messageText) {
  // 1. LIFFが初期化されているか、かつLINEアプリ内で動作しているかを確認
  if (!liff.isInClient()) {
      alert("メッセージ送信はLINEアプリ内でのみ実行可能です。");
      return;
  }
  // 2. メッセージを送信
  liff.sendMessages([{ type: 'text', text: messageText }])
      .then(() => {
          console.log('Message sent successfully!');
      })
      .catch((err) => {
          console.error('Error sending message:', err);
      });
}

/**
 * データをキャッシュに保存する関数
 */
function saveToCache(capacityData, userInfoData, configData, monthKey = "") {
  const now = Date.now();

  // 1. 残席情報を保存
  Object.keys(capacityData).forEach(dateStr => {
    const mKey = dateStr.substring(0, 7); // "YYYY-MM"
    if (!AVAILABLE_CAPACITY_DATA[mKey]) {
      AVAILABLE_CAPACITY_DATA[mKey] = { data: {}, lastFetch: now };
    }
    AVAILABLE_CAPACITY_DATA[mKey].data[dateStr] = capacityData[dateStr];
    AVAILABLE_CAPACITY_DATA[mKey].lastFetch = now;
  });

  // 2. 予約情報の保存（画像通りの配列構造に対応）
  const reservedArray = userInfoData.myReservedDates || [];
  
  // 既存のキャッシュをクリア（一括更新のため）
  // 全てのキーに対して data をリセット
  Object.keys(MY_RESERVIONS).forEach(k => MY_RESERVIONS[k].data = []);

  reservedArray.forEach(resObj => {
    // resObj は { "2026-02-01 10:10": {...} } という形
    const dateTimeStr = Object.keys(resObj)[0]; 
    if (!dateTimeStr) return;

    const mKey = dateTimeStr.substring(0, 7); // "YYYY-MM"
    
    if (!MY_RESERVIONS[mKey]) {
      MY_RESERVIONS[mKey] = { data: [], lastFetch: now };
    }
    // 配列の中にオブジェクトをそのままプッシュ
    MY_RESERVIONS[mKey].data.push(resObj);
    MY_RESERVIONS[mKey].lastFetch = now;
  });

  // 予約がない月の対応
  if (!MY_RESERVIONS[monthKey] && monthKey !== "") {
    MY_RESERVIONS[monthKey] = { data: [], lastFetch: now };
  }

  // 3. 出席情報
  MY_ATTENDED_DATES = {
    data: userInfoData.myAttendedDates || [],
    lastFetch: now 
  };
  
  const cachedConfigData = configData == null ? JSON.parse(localStorage.getItem("APP_DATA_CACHE")) : null;

  // ローカルストレージ登録
  const appCache = {
    success: true,
    lastFetch: now,
    capacityData: capacityData,  // 全体の残席情報
    config: configData == null ? cachedConfigData.config : configData,
    userInfo: {
      data: userInfoData.data,
      myAttendedDates: userInfoData.myAttendedDates,
      myReservedDates: userInfoData.myReservedDates
    }
  };
  localStorage.setItem("APP_DATA_CACHE", JSON.stringify(appCache));
}

/**
 * キャッシュが有効か判定し、有効なら一式を返す
 */
function getValidFullCache(monthKey) {
  const now = Date.now();
  const capCache = AVAILABLE_CAPACITY_DATA[monthKey];
  const resCache = MY_RESERVIONS[monthKey];
  const attCache = MY_ATTENDED_DATES;

  // すべてのキャッシュが存在し、かつ期限内かチェック
  if (!capCache?.lastFetch || !resCache?.lastFetch || !attCache?.lastFetch) return null;

  const isCapExpired = (now - capCache.lastFetch) > CACHE_EXPIRATION_MS;
  const isResExpired = (now - resCache.lastFetch) > CACHE_EXPIRATION_MS;
  const isAttExpired = (now - attCache.lastFetch) > CACHE_EXPIRATION_MS;

  if (isCapExpired || isResExpired || isAttExpired) {
    console.log(`キャッシュのいずれかが期限切れです: ${monthKey}`);
    return null;
  }

  return {
    capacity: capCache.data,
    reserved: resCache.data,
    attended: attCache.data
  };
}

/**
 * キャッシュが有効か判定し、有効なら一式を返す
 */
function getInitDispFullCache(monthKey) {
  const now = Date.now();
  const cachedJSON = localStorage.getItem("APP_DATA_CACHE");
  if (!cachedJSON) return null;

  const cacheObject = JSON.parse(cachedJSON);
  const capCache = cacheObject.capacityData;
  const resCache = cacheObject.userInfo.myReservedDates;
  const attCache = cacheObject.userInfo.myAttendedDates;

  // すべてのキャッシュが存在し、かつ期限内かチェック
  if (!cacheObject?.lastFetch) return null;

  const isCacheExpired = (now - cacheObject.lastFetch) > CACHE_EXPIRATION_MS;
  if (isCacheExpired) {
    console.log(`キャッシュのいずれかが期限切れです: ${monthKey}`);
    return null;
  }

  return {
    success: true,
    lastFetch: cacheObject.lastFetch,
    capacityData: capCache,  // 全体の残席情報
    config: cacheObject.config,
    userInfo: {
      data: cacheObject.userInfo.data,
      myAttendedDates: attCache,
      myReservedDates: resCache
    }
  };
}

function updateClassInfoUI(currentUser, monthKey) {
  const currentDate = new Date();
  const currentMonthKey = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}`; 
  const upperLimitLabel = currentMonthKey === monthKey ? currentUser.upperLimitNumberThisMonth : currentUser.upperLimitNumberNextMonth;
  classInfo.innerHTML = `<span id='userName'>   👤 ${currentUser.displayName}</span><span id='userClassName'>  ┊  🖌️ ${currentUser.className} 🗓️ 月${upperLimitLabel}回</span>`;
}

async function getWorkersDataJson(userId) {
  const url = `${WORKERS_BASE_URL}?userId=${userId}`;
  const response = await fetch(url);
  const json = await response.json();
  return json;
}
