function saveSystemLog(log) {
  const sheet = SPREADSHEET.getSheetByName("systemlog");

  sheet.appendRow([
    new Date(),
    log.level,
    log.type,
    log.targetUserId,
    log.message,
    log.detailInfo ?? ""
  ]);
}

function logWarn(log) {
  saveSystemLog({
    ...log, 
    level: "WARN"
  });
}

function logError(log) {
  saveSystemLog({
    ...log,
    level: "ERROR"
  });
}

function logInfo(log) {
  saveSystemLog({
    ...log,
    level: "INFO"
  });
}