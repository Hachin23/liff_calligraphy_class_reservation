function sendLineMessage(userId, message) {
  fetchLineMessage(userId, message);
}

function sendAdminLineMessage(message) {
  const userId = getRequiredProperty("ADMIN_LINE_USER_ID");
  fetchLineMessage(userId, message);
}

function fetchLineMessage(userId, message) {

  const token = getRequiredProperty("CHANNEL_ACCESS_TOKEN");
  const url = 'https://api.line.me/v2/bot/message/push';

  const payload = {
    to: userId,
    messages: [
      {
        type: "text",
        text: message
      }
    ]
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    headers: {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    console.error(response.getContentText());
    throw new Error("LINE通知に失敗しました。");
  }
}