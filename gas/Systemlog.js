function saveSystemLog(log) {
  const sheet = SPREADSHEET.getSheetByName("systemlog");

  sheet.appendRow([
    new Date(),
    log.level,
    log.type,
    log.targetUserId,
    log.message
  ]);
}

function logWarn(log) {
  saveSystemLog(
    ...log, 
    "WARN"
  );
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