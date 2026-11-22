function sendEmailNotification(info) {
  try {
    var myEmail = "dongjin3808@gmail.com";
    var spreadsheetUrl = "https://docs.google.com/spreadsheets/d/1Xcm2Ha8dkrO4FT3IfBYfJlTiXScNQFwFTF8j3J0csQ/edit?gid=0#gid=0";

    var typeVal = (info && info.type === "update") ? "✏️ [예약 변경]" : "🔔 [새 예약 접수]";
    var nickname = (info && info.nickname) ? info.nickname : "고객";
    var pickupTime = (info && info.pickupTime) ? info.pickupTime : "";
    var cookType = (info && info.cookType) ? info.cookType : "";
    var menuSummary = (info && info.menuSummary) ? info.menuSummary : "";
    var sauce = (info && info.sauce) ? info.sauce : "없음";
    var totalPrice = (info && info.totalPrice) ? info.totalPrice : "";

    var subject = typeVal + " " + nickname + "님 (" + pickupTime + " 픽업)";

    var body = "--- [장날동천 예약 내역] ---\n" +
               "• 픽업시간: " + pickupTime + "\n" +
               "• 조리방식: " + cookType + "\n" +
               "• 메뉴: " + menuSummary + "\n" +
               "• 추가소스: " + sauce + "\n" +
               "• 예약 식별명: " + nickname + "\n" +
               "• 주문 총 가격: " + totalPrice + "\n\n" +
               "📊 [예약 스프레드시트 바로가기]\n" + spreadsheetUrl + "\n\n" +
               "장날동천 수제 돈가스 예약 시스템에서 발송된 알림입니다.";

    MailApp.sendEmail(myEmail, subject, body);
  } catch (e) {
    Logger.log("🚨 이메일 발송 실패: " + e.toString());
  }
}

function applyLabelToLatestOrderEmail(myEmail, labelName) {
  try {
    var targetEmail = myEmail || "dongjin3808@gmail.com";
    var targetLabelName = labelName || "돈가스예약";
    var threads = GmailApp.search('to:' + targetEmail, 0, 1);
    if (threads.length > 0) {
      var label = GmailApp.getUserLabelByName(targetLabelName) || GmailApp.createLabel(targetLabelName);
      threads[0].addLabel(label);
    }
  } catch (labelErr) {
    Logger.log("라벨 지정 중 예외 발생: " + labelErr.toString());
  }
}

function runPostOrderBackgroundTasks(emailInfo) {
  if (!emailInfo) return;
  try {
    sendEmailNotification(emailInfo);
  } catch (err) {
    Logger.log("메일 발송 에러: " + err.toString());
  }
  applyLabelToLatestOrderEmail("dongjin3808@gmail.com", "돈가스예약");
}

function parseTimeToMin(timeVal) {
  if (!timeVal) return 9999;

  if (timeVal instanceof Date) {
    var hours = timeVal.getHours();
    var minutes = timeVal.getMinutes();
    return hours * 60 + minutes;
  }

  var str = String(timeVal).trim();
  if (!str) return 9999;

  var isPM = str.indexOf('오후') !== -1;
  var isAM = str.indexOf('오전') !== -1;
  var clean = str.replace('오전', '').replace('오후', '').trim();
  var parts = clean.split(':');
  if (parts.length < 2) return 9999;

  var hours = parseInt(parts[0], 10);
  var minutes = parseInt(parts[1], 10);

  if (isNaN(hours)) hours = 0;
  if (isNaN(minutes)) minutes = 0;

  if (isAM && hours === 12) {
    hours = 0;
  } else if (isPM && hours < 12) {
    hours += 12;
  }

  return hours * 60 + minutes;
}

function getKoreanCurrentTimeVal() {
  var nowStr = Utilities.formatDate(new Date(), "Asia/Seoul", "HH:mm");
  var parts = nowStr.split(":");
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
}

function checkIsClosedNow(settings) {
  var currentTimeVal = getKoreanCurrentTimeVal();
  var openTimeVal = settings.openHour * 60 + settings.openMin;
  var closeTimeVal = settings.closeHour * 60 + settings.closeMin;

  if (closeTimeVal < openTimeVal) {
    return (currentTimeVal > closeTimeVal && currentTimeVal < openTimeVal);
  } else {
    return (currentTimeVal < openTimeVal || currentTimeVal > closeTimeVal);
  }
}

function sortOrders(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    var maxCols = Math.max(sheet.getLastColumn(), 10);
    var range = sheet.getRange(2, 1, lastRow - 1, maxCols);
    var values = range.getValues();

    values.sort(function(a, b) {
      var statusA = String(a[0]).trim();
      var statusB = String(b[0]).trim();

      if (statusA === "취소" && statusB !== "취소") return 1;
      if (statusA !== "취소" && statusB === "취소") return -1;

      var timeA = parseTimeToMin(a[1]);
      var timeB = parseTimeToMin(b[1]);

      if (timeA < 360) timeA += 1440;
      if (timeB < 360) timeB += 1440;

      return timeA - timeB;
    });

    range.setValues(values);
  }
}

function normalizeKakaoId(rawId) {
  if (rawId === null || rawId === undefined) return "";
  var s = String(rawId).trim();
  if (!s) return "";
  if (s.indexOf("kakao_") === 0) s = s.substring(6).trim();
  if (/^\d+\.0$/.test(s)) s = s.replace(/\.0$/, "");
  return s;
}

function formatKakaoDisplayId(rawId) {
  var norm = normalizeKakaoId(rawId);
  if (!norm) return "";
  return /^\d+$/.test(norm) ? ("kakao_" + norm) : norm;
}

function isSameKakaoMember(idA, idB) {
  var a = normalizeKakaoId(idA);
  var b = normalizeKakaoId(idB);
  if (!a || !b) return false;
  return a === b;
}

function exchangeKakaoAuthorizationCode(code, redirectUri) {
  var cleanCode = code ? String(code).trim() : "";
  if (!cleanCode) {
    return { ok: false, message: "인가 코드가 비어 있습니다." };
  }

  var props = PropertiesService.getScriptProperties();
  var clientId = (props.getProperty("KAKAO_REST_API_KEY") || props.getProperty("KAKAO_CLIENT_ID") || "").trim();
  var clientSecret = (props.getProperty("KAKAO_CLIENT_SECRET") || "").trim();

  if (!clientId) {
    return {
      ok: false,
      message: "스크립트 속성에 KAKAO_REST_API_KEY가 설정되어 있지 않습니다."
    };
  }

  try {
    var payload = {
      grant_type: "authorization_code",
      client_id: clientId,
      code: cleanCode
    };
    if (redirectUri) {
      payload.redirect_uri = String(redirectUri).trim();
    }
    if (clientSecret) {
      payload.client_secret = clientSecret;
    }

    var tokenRes = UrlFetchApp.fetch("https://kauth.kakao.com/oauth/token", {
      method: "post",
      contentType: "application/x-www-form-urlencoded;charset=utf-8",
      payload: payload,
      muteHttpExceptions: true
    });

    var tokenCode = tokenRes.getResponseCode();
    var tokenJson = null;
    try { tokenJson = JSON.parse(tokenRes.getContentText()); } catch (e) {}

    if (tokenCode < 200 || tokenCode >= 300 || !tokenJson || !tokenJson.access_token) {
      return {
        ok: false,
        message: "카카오 토큰 발급 실패: " + (tokenJson && tokenJson.error_description ? tokenJson.error_description : tokenRes.getContentText())
      };
    }

    var userRes = UrlFetchApp.fetch("https://kapi.kakao.com/v2/user/me", {
      method: "get",
      headers: {
        Authorization: "Bearer " + tokenJson.access_token
      },
      muteHttpExceptions: true
    });

    var userCode = userRes.getResponseCode();
    var userJson = null;
    try { userJson = JSON.parse(userRes.getContentText()); } catch (e) {}

    if (userCode < 200 || userCode >= 300 || !userJson || !userJson.id) {
      return {
        ok: false,
        message: "카카오 사용자 정보 조회 실패: " + userRes.getContentText()
      };
    }

    var kakaoId = String(userJson.id).trim();
    var kakaoNickname = "";
    if (userJson.kakao_account && userJson.kakao_account.profile && userJson.kakao_account.profile.nickname) {
      kakaoNickname = String(userJson.kakao_account.profile.nickname).trim();
    } else if (userJson.properties && userJson.properties.nickname) {
      kakaoNickname = String(userJson.properties.nickname).trim();
    }
    if (!kakaoNickname) kakaoNickname = "카카오고객";

    return {
      ok: true,
      kakaoId: kakaoId,
      kakaoNickname: kakaoNickname
    };
  } catch (err) {
    return {
      ok: false,
      message: "카카오 인증 처리 중 오류: " + err.toString()
    };
  }
}

function getMemberSheet(ss) {
  var sheet = ss.getSheetByName("회원관리");
  if (!sheet) {
    sheet = ss.insertSheet("회원관리");
    sheet.appendRow(["예약 식별명", "카카오 회원번호", "카톡 닉네임", "전화번호 뒷자리", "최초 가입일시", "최근 예약일시", "고객 고유값"]);
  }
  return sheet;
}

function getAccessLimitSheet(ss) {
  var sheet = ss.getSheetByName("접속제한관리");
  if (!sheet) {
    sheet = ss.insertSheet("접속제한관리");
    sheet.appendRow(["예약 식별명", "카카오 회원번호", "최초 접속일", "최근 구매(예약)일", "현재 이용상태", "마지막 접속 시간", "사장님 예외 승인", "사장님 메모"]);
  }
  return sheet;
}

function checkAllClientsAccess() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var limitSheet = ss.getSheetByName("접속제한관리");
  if (!limitSheet) return;

  var lastRow = limitSheet.getLastRow();
  if (lastRow < 2) return;

  var settings = getStoreSettings(ss);
  var limitDays = settings.accessLimitDays || 30;
  var todayDateObj = new Date(Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd"));

  var maxCols = Math.max(limitSheet.getLastColumn(), 7);
  var range = limitSheet.getRange(2, 1, lastRow - 1, maxCols);
  var values = range.getValues();
  var updated = false;

  for (var i = 0; i < values.length; i++) {
    var manualApproval = String(values[i][6] || "").trim();
    var currentE = String(values[i][4] || "").trim();
    var newStatus = "";

    if (manualApproval === "허용" || manualApproval === "관리자 허용") {
      newStatus = "관리자 허용";
    } else if (manualApproval === "차단" || manualApproval === "관리자 차단") {
      newStatus = "관리자 차단";
    } else {
      var rawC = values[i][2];
      var rawD = values[i][3];
      var rawF = values[i][5];

      var cDate = null;
      var dDate = null;
      if (rawC) {
        var dc = rawC instanceof Date ? rawC : new Date(String(rawC).trim());
        if (!isNaN(dc.getTime())) cDate = new Date(Utilities.formatDate(dc, "Asia/Seoul", "yyyy-MM-dd"));
      }
      if (rawD) {
        var dd = rawD instanceof Date ? rawD : new Date(String(rawD).trim());
        if (!isNaN(dd.getTime())) dDate = new Date(Utilities.formatDate(dd, "Asia/Seoul", "yyyy-MM-dd"));
      }

      var base1Date = null;
      if (cDate && dDate) {
        base1Date = (dDate > cDate) ? dDate : cDate;
      } else if (dDate) {
        base1Date = dDate;
      } else if (cDate) {
        base1Date = cDate;
      }

      var fDate = null;
      if (rawF) {
        var df = rawF instanceof Date ? rawF : new Date(String(rawF).trim());
        if (!isNaN(df.getTime())) fDate = new Date(Utilities.formatDate(df, "Asia/Seoul", "yyyy-MM-dd"));
      }

      if (base1Date && fDate) {
        var diffDays = Math.floor((fDate.getTime() - base1Date.getTime()) / (1000 * 60 * 60 * 24));
        newStatus = (diffDays > limitDays) ? "시스템 차단" : "시스템 허용";
      } else if (base1Date) {
        var diffDays = Math.floor((todayDateObj.getTime() - base1Date.getTime()) / (1000 * 60 * 60 * 24));
        newStatus = (diffDays > limitDays) ? "시스템 차단" : "시스템 허용";
      } else {
        newStatus = "시스템 차단";
      }
    }

    if (currentE !== newStatus) {
      values[i][4] = newStatus;
      updated = true;
    }
  }

  if (updated) {
    range.setValues(values);
  }
}

// 일별 자동 트리거 전용 진입점
function runDailyAccessLimitCheck() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    checkAllClientsAccess();
  } catch (err) {
    Logger.log("일별 접속제한 점검 오류: " + err.toString());
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function onEdit(e) {
  try {
    var range = e.range;
    var sheet = range.getSheet();
    var sheetName = sheet.getName();

    if (sheetName === "시트1" && range.getRow() > 1) {
      var col = range.getColumn();
      if (col === 1 || col === 2) {
        sortOrders(sheet);
      }
    }

    if (sheetName === "메뉴판" && range.getRow() > 1 && range.getColumn() === 6) {
      var val = String(range.getValue() || "").trim().toUpperCase();
      var tsCell = sheet.getRange(range.getRow(), 8);
      if (val === "HOT") {
        tsCell.setValue(new Date().getTime());
      } else {
        tsCell.clearContent();
      }
    }

    if (sheetName === "설정" || sheetName === "메뉴판") {
      try {
        CacheService.getScriptCache().remove("store_settings_cache_v5");
      } catch (cacheErr) {}
    }
  } catch (err) {
    Logger.log("onEdit 오류: " + err.toString());
  }
}

function parseItemsToMap(str) {
  var map = {};
  if (!str || str === "없음") return map;
  var parts = str.split(',');
  for (var i = 0; i < parts.length; i++) {
    var part = parts[i].trim();
    var match = part.match(/(.*?)\s*(\d+)개/);
    if (match) {
      var name = match[1].trim();
      var qty = parseInt(match[2], 10) || 0;
      map[name] = qty;
    }
  }
  return map;
}

function syncKakaoMemberProfile(nickname, kakaoId, kakaoNickname, phone4, isPurchase) {
  var normId = normalizeKakaoId(kakaoId);
  if (!normId) return null;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = getMemberSheet(ss);
  var nowStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  var displayId = formatKakaoDisplayId(normId);
  var cleanIdent = nickname ? String(nickname).trim() : "";
  var cleanKakaoNick = kakaoNickname ? String(kakaoNickname).trim() : "";
  var cleanPhone4 = phone4 ? String(phone4).replace(/[^0-9]/g, "").slice(-4) : "";

  var lastRow = memberSheet.getLastRow();
  var values = lastRow >= 2 ? memberSheet.getRange(2, 1, lastRow - 1, 7).getValues() : [];
  var foundRow = -1;

  for (var i = 0; i < values.length; i++) {
    var rowId = String(values[i][1] || "").trim();
    if (isSameKakaoMember(rowId, normId)) {
      foundRow = i + 2;
      break;
    }
  }

  if (foundRow > 0) {
    var existingRow = values[foundRow - 2];
    var existingIdent = String(existingRow[0] || "").trim();
    var existingKakaoNick = String(existingRow[2] || "").trim();
    var existingPhone4 = String(existingRow[3] || "").replace(/[^0-9]/g, "").slice(-4);
    var hasLockedPhone4 = !!(existingPhone4 && existingPhone4.length === 4);
    var finalKakaoNick = cleanKakaoNick || existingKakaoNick;
    var finalPhone4 = hasLockedPhone4 ? existingPhone4 : cleanPhone4;
    var finalIdent = hasLockedPhone4
      ? (finalKakaoNick ? (finalKakaoNick + finalPhone4) : (existingIdent || cleanIdent))
      : (cleanIdent || existingIdent || (finalKakaoNick && finalPhone4 ? (finalKakaoNick + finalPhone4) : finalKakaoNick));
    var finalFirstJoin = String(existingRow[4] || "").trim() || nowStr;
    var finalLastReserve = isPurchase ? nowStr : String(existingRow[5] || "").trim();

    var rowToWrite = [
      finalIdent || existingIdent,
      displayId,
      finalKakaoNick || existingKakaoNick,
      finalPhone4 ? ("'" + finalPhone4) : (existingPhone4 ? ("'" + existingPhone4) : ""),
      finalFirstJoin,
      finalLastReserve,
      displayId
    ];
    memberSheet.getRange(foundRow, 1, 1, 7).setValues([rowToWrite]);

    return {
      identifier: rowToWrite[0],
      kakaoId: displayId,
      kakaoNickname: rowToWrite[2],
      phone4: finalPhone4,
      firstJoinAt: finalFirstJoin,
      lastReserveAt: finalLastReserve,
      isRegistered: !!(rowToWrite[0] && finalPhone4 && finalPhone4.length === 4)
    };
  } else {
    var initIdent = cleanIdent || (cleanKakaoNick && cleanPhone4 ? (cleanKakaoNick + cleanPhone4) : cleanKakaoNick);
    var initLastReserve = isPurchase ? nowStr : "";
    memberSheet.appendRow([
      initIdent,
      displayId,
      cleanKakaoNick,
      cleanPhone4 ? ("'" + cleanPhone4) : "",
      nowStr,
      initLastReserve,
      displayId
    ]);

    return {
      identifier: initIdent,
      kakaoId: displayId,
      kakaoNickname: cleanKakaoNick,
      phone4: cleanPhone4,
      firstJoinAt: nowStr,
      lastReserveAt: initLastReserve,
      isRegistered: !!(initIdent && cleanPhone4 && cleanPhone4.length === 4)
    };
  }
}

function getKakaoMemberProfile(kakaoId) {
  var normId = normalizeKakaoId(kakaoId);
  if (!normId) return null;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName("회원관리");
  if (!memberSheet) return null;

  var lastRow = memberSheet.getLastRow();
  if (lastRow < 2) return null;

  var values = memberSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    var rowId = String(values[i][1] || "").trim();
    if (isSameKakaoMember(rowId, normId)) {
      var ident = String(values[i][0] || "").trim();
      var kNick = String(values[i][2] || "").trim();
      var p4 = String(values[i][3] || "").replace(/^'/, "").trim();
      return {
        identifier: ident,
        kakaoId: formatKakaoDisplayId(normId),
        kakaoNickname: kNick,
        phone4: p4,
        firstJoinAt: values[i][4] || "",
        lastReserveAt: values[i][5] || "",
        isRegistered: !!(ident && p4 && p4.length === 4)
      };
    }
  }
  return null;
}

function getMemberIdentifiersByKakaoId(kakaoId) {
  var normId = normalizeKakaoId(kakaoId);
  var list = [];
  if (!normId) return list;

  function addUnique(val) {
    var s = val ? String(val).trim() : "";
    if (s && list.indexOf(s) === -1) list.push(s);
  }

  var profile = getKakaoMemberProfile(normId);
  if (profile) {
    addUnique(profile.identifier);
    if (profile.kakaoNickname && profile.phone4) {
      addUnique(profile.kakaoNickname + profile.phone4);
    }
  }

  return list;
}

function isCustomerMatchForKakao(custCol, targetNickname, kakaoId, rowKakaoId, aliasesCache) {
  var normId = normalizeKakaoId(kakaoId);
  var normRowId = normalizeKakaoId(rowKakaoId);
  if (normId && normRowId && normId === normRowId) {
    return true;
  }

  var cleanCust = custCol ? String(custCol).trim() : "";
  if (!cleanCust) return false;

  if (targetNickname && cleanCust === String(targetNickname).trim()) {
    return true;
  }

  if (normId) {
    var aliases = aliasesCache || getMemberIdentifiersByKakaoId(normId);
    if (aliases.indexOf(cleanCust) !== -1) {
      return true;
    }
  }

  return false;
}

function checkClientAccessRestriction(nickname, clientId) {
  var cleanNick = nickname ? String(nickname).trim() : "";
  var cleanClient = clientId ? String(clientId).trim() : "";
  var normClient = normalizeKakaoId(cleanClient);
  if (!cleanNick && !normClient) return { blocked: false };

  var aliases = normClient ? getMemberIdentifiersByKakaoId(normClient) : [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var limitSheet = getAccessLimitSheet(ss);

  var lastRow = limitSheet.getLastRow();
  if (lastRow >= 2) {
    var values = limitSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    for (var i = 0; i < values.length; i++) {
      var rowNick = String(values[i][0] || "").trim();
      var rowClient = String(values[i][1] || "").trim();

      var isMatch = (normClient && rowClient && isSameKakaoMember(rowClient, normClient)) ||
                      (cleanNick && rowNick && rowNick === cleanNick) ||
                      (rowNick && aliases.indexOf(rowNick) !== -1);
      if (isMatch) {
        var currentStatus = String(values[i][4] || "").trim();
        if (currentStatus === "시스템 차단" || currentStatus === "관리자 차단") {
          return { blocked: true, reason: "이용이 제한되었습니다. (" + currentStatus + ")" };
        }
      }
    }
  }
  return { blocked: false };
}

function updateClientRecordOnAction(nickname, clientId, isPurchase) {
  var cleanNick = nickname ? String(nickname).trim() : "";
  var cleanClient = clientId ? String(clientId).trim() : "";
  var normClient = normalizeKakaoId(cleanClient);
  if (!cleanNick && !normClient) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var limitSheet = getAccessLimitSheet(ss);

  var lastRow = limitSheet.getLastRow();
  var values = lastRow >= 2 ? limitSheet.getRange(2, 1, lastRow - 1, 7).getValues() : [];
  var foundRow = -1;
  var nowStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  var todayDateStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
  var formattedId = normClient ? formatKakaoDisplayId(normClient) : cleanClient;

  for (var i = 0; i < values.length; i++) {
    var rowNick = String(values[i][0] || "").trim();
    var rowClient = String(values[i][1] || "").trim();
    if ((normClient && rowClient && isSameKakaoMember(rowClient, normClient)) ||
        (cleanNick && rowNick && rowNick === cleanNick)) {
      foundRow = i + 2;
      break;
    }
  }

  if (foundRow > 0) {
    var existingRow = values[foundRow - 2];
    var nextNick = (cleanNick && cleanNick !== "미입력고객" && !cleanNick.startsWith("손님 ("))
      ? cleanNick
      : existingRow[0];
    var nextId = formattedId || existingRow[1];
    var nextPurchaseDate = isPurchase ? todayDateStr : existingRow[3];

    limitSheet.getRange(foundRow, 1, 1, 6).setValues([[
      nextNick,
      nextId,
      existingRow[2],
      nextPurchaseDate,
      existingRow[4],
      nowStr
    ]]);
  } else {
    limitSheet.appendRow([
      cleanNick || "미입력고객",
      formattedId,
      todayDateStr,
      isPurchase ? todayDateStr : "",
      "시스템 허용",
      nowStr,
      "",
      ""
    ]);
  }
}

function findNicknameByClientId(clientId) {
  var normId = normalizeKakaoId(clientId);
  if (!normId) return "";

  var member = getKakaoMemberProfile(normId);
  if (member && member.identifier) {
    return member.identifier;
  }
  return "";
}

function updateActivePresenceCache(clientId, nickname) {
  var normId = normalizeKakaoId(clientId);
  var cacheKey = normId ? ("kakao_" + normId) : (clientId ? String(clientId).trim() : "");
  if (!cacheKey) return 0;

  try {
    var cache = CacheService.getScriptCache();
    var raw = cache.get("active_users_v1");
    var activeMap = raw ? JSON.parse(raw) : {};
    var now = Date.now();

    if (nickname === 'EXIT') {
      delete activeMap[cacheKey];
    } else {
      var suffix = normId ? normId.slice(-4) : cacheKey.slice(-4);
      var displayNick = (nickname && nickname !== "미입력고객") ? nickname : ("손님 (" + suffix + ")");
      activeMap[cacheKey] = {
        nickname: displayNick,
        lastSeen: now
      };
    }

    var keys = Object.keys(activeMap);
    var activeCount = 0;
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i];
      if (now - activeMap[k].lastSeen > 70000) {
        delete activeMap[k];
      } else {
        activeCount++;
      }
    }

    cache.put("active_users_v1", JSON.stringify(activeMap), 300);
    return activeCount;
  } catch (e) {
    return 0;
  }
}

function recordUserHeartbeat(data) {
  var clientId = data.clientId || data.kakaoId;
  var nickname = data.nickname ? String(data.nickname).trim() : "미입력고객";
  if (!nickname) nickname = "미입력고객";

  var activeCount = updateActivePresenceCache(clientId, nickname);

  if (clientId && nickname !== 'EXIT' && data.updateSheet) {
    updateClientRecordOnAction(nickname, clientId, false);
  }

  var restriction = checkClientAccessRestriction(nickname, clientId);
  return { success: true, count: activeCount, restriction: restriction };
}

function getActiveUserList() {
  try {
    var cache = CacheService.getScriptCache();
    var raw = cache.get("active_users_v1");
    var activeMap = raw ? JSON.parse(raw) : {};
    var now = Date.now();
    var users = [];
    var changed = false;

    var keys = Object.keys(activeMap);
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i];
      if (now - activeMap[k].lastSeen > 70000) {
        delete activeMap[k];
        changed = true;
      } else {
        var userNick = activeMap[k].nickname;
        if (!userNick || userNick === "미입력고객") {
          userNick = "손님 (" + k.slice(-4) + ")";
        }
        users.push(userNick);
      }
    }

    if (changed) {
      cache.put("active_users_v1", JSON.stringify(activeMap), 300);
    }

    return { count: users.length, users: users };
  } catch (e) {
    return { count: 0, users: [] };
  }
}

function buildFullSettingsPayload(ss, sheet, reqNick, reqClient) {
  var restriction = checkClientAccessRestriction(reqNick, reqClient);
  var settings = getStoreSettings(ss);

  var nowStr = Utilities.formatDate(new Date(), "Asia/Seoul", "HH:mm");
  var parts = nowStr.split(":");
  var currentTimeVal = parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
  var openTimeVal = settings.openHour * 60 + settings.openMin;
  var closeTimeVal = settings.closeHour * 60 + settings.closeMin;
  var isClosedNow = (closeTimeVal < openTimeVal)
    ? (currentTimeVal > closeTimeVal && currentTimeVal < openTimeVal)
    : (currentTimeVal < openTimeVal || currentTimeVal > closeTimeVal);

  var rawPickupTimes = [];
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var rangeVals = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < rangeVals.length; i++) {
      var status = String(rangeVals[i][0] || "").trim();
      var pTime = rangeVals[i][1];
      if (status !== "취소" && pTime !== "") {
        var minVal = parseTimeToMin(pTime);
        if (minVal !== 9999) {
          rawPickupTimes.push(minVal);
        }
      }
    }
  }

  var fortuneList = [];
  var fortuneSheet = ss.getSheetByName("포춘쿠키");
  if (fortuneSheet) {
    var fLastRow = fortuneSheet.getLastRow();
    if (fLastRow >= 1) {
      var fValues = fortuneSheet.getRange(1, 1, fLastRow, 1).getValues();
      for (var j = 0; j < fValues.length; j++) {
        var txt = String(fValues[j][0] || "").trim();
        if (txt !== "") {
          fortuneList.push(txt);
        }
      }
    }
  }

  var fortuneEnabledVal = "On";
  var settingSheet = ss.getSheetByName("설정");
  if (settingSheet) {
    fortuneEnabledVal = settingSheet.getRange("B8").getDisplayValue();
  }

  var responseData = Object.assign({}, settings);
  responseData.isClosed = isClosedNow;
  responseData.pickupTimesList = rawPickupTimes;
  responseData.fortuneList = fortuneList;
  responseData.fortuneEnabled = (String(fortuneEnabledVal).trim() !== "Off");
  responseData.isBlocked = restriction.blocked;
  responseData.blockReason = restriction.reason || "";

  return responseData;
}

function buildReservationCheckPayload(ss, sheet, nickname, clientIdCheck, shouldUpdateRecord, kakaoNickname, phone4) {
  var recoveredNickname = "";
  var memberProfile = null;
  var normId = normalizeKakaoId(clientIdCheck);

  if (normId) {
    memberProfile = syncKakaoMemberProfile(nickname, normId, kakaoNickname, phone4, false);
    if (!nickname && memberProfile && memberProfile.identifier) {
      recoveredNickname = memberProfile.identifier;
      nickname = recoveredNickname;
    }
  }

  if (!nickname && clientIdCheck) {
    recoveredNickname = findNicknameByClientId(clientIdCheck);
    if (recoveredNickname) {
      nickname = recoveredNickname;
    }
  }

  if (clientIdCheck && shouldUpdateRecord) {
    updateClientRecordOnAction(nickname || (memberProfile && memberProfile.kakaoNickname) || "미입력고객", clientIdCheck, false);
  }

  var restriction = checkClientAccessRestriction(nickname, clientIdCheck);
  if (restriction.blocked) {
    return {
      "result": "blocked",
      "message": restriction.reason,
      "memberProfile": memberProfile
    };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var maxCols = Math.min(Math.max(sheet.getLastColumn(), 10), 10);
    var rows = sheet.getRange(1, 1, lastRow, maxCols).getValues();
    var aliases = normId ? getMemberIdentifiersByKakaoId(normId) : [];

    for (var i = rows.length - 1; i >= 1; i--) {
      var status = String(rows[i][0] || "").trim();
      if (status === "취소") continue;

      var custCol = String(rows[i][5] || "").trim();
      var rowKakaoId = String(rows[i][9] || "").trim();
      if (isCustomerMatchForKakao(custCol, nickname, normId, rowKakaoId, aliases)) {
        var cellValue = sheet.getRange(i + 1, 2).getDisplayValue();
        var resolvedNick = (memberProfile && memberProfile.identifier) || nickname || custCol;
        return {
          "result": "found",
          "rowIndex": i + 1,
          "status": status,
          "nickname": resolvedNick,
          "recoveredNickname": resolvedNick,
          "memberProfile": memberProfile,
          "pickupTime": cellValue,
          "cookType": String(rows[i][2] || "").trim(),
          "menu": rows[i][3] || "",
          "sauce": rows[i][4] || "없음",
          "totalPrice": rows[i][7] || "",
          "eventPrize": rows[i][8] || ""
        };
      }
    }
  }

  return {
    "result": "not_found",
    "recoveredNickname": recoveredNickname || ((memberProfile && memberProfile.identifier) || ""),
    "memberProfile": memberProfile
  };
}

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({"result": "error"}))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  var action = data.action;

  if (action === "exchangeKakaoCode") {
    var exchanged = exchangeKakaoAuthorizationCode(data.code, data.redirectUri);
    if (!exchanged.ok) {
      return ContentService.createTextOutput(JSON.stringify({
        "result": "error",
        "message": exchanged.message
      })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({
      "result": "ok",
      "authUser": {
        "kakaoId": exchanged.kakaoId,
        "kakaoNickname": exchanged.kakaoNickname
      }
    })).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'heartbeat') {
    var res = recordUserHeartbeat(data);
    if (res.restriction && res.restriction.blocked) {
      return ContentService.createTextOutput(JSON.stringify({ "result": "blocked", "message": res.restriction.reason }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify(res))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'getActiveUsers') {
    var activeRes = getActiveUserList();
    return ContentService.createTextOutput(JSON.stringify(activeRes))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "getSettings") {
    try {
      var ssQuick = SpreadsheetApp.getActiveSpreadsheet();
      var sheetQuick = ssQuick.getSheetByName("시트1");
      if (!sheetQuick) { sheetQuick = ssQuick.getSheets()[0]; }
      var reqNickQuick = String(data.nickname || "").trim();
      var reqClientQuick = String(data.clientId || data.kakaoId || "").trim();
      var responseDataQuick = buildFullSettingsPayload(ssQuick, sheetQuick, reqNickQuick, reqClientQuick);

      return ContentService.createTextOutput(JSON.stringify(responseDataQuick))
                           .setMimeType(ContentService.MimeType.JSON);
    } catch (quickErr) {
      return ContentService.createTextOutput(JSON.stringify({"result": "error", "message": quickErr.toString()}))
                           .setMimeType(ContentService.MimeType.JSON);
    }
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return ContentService.createTextOutput(JSON.stringify({"result": "error", "message": "서버가 혼잡합니다. 잠시 후 다시 시도해 주세요."}))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  var backgroundTaskInfo = null;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("시트1");
    if (!sheet) { sheet = ss.getSheets()[0]; }

    if (action === "syncKakaoMember") {
      var reqNickSync = String(data.nickname || "").trim();
      var reqClientSync = String(data.clientId || data.kakaoId || "").trim();
      var reqKakaoNick = String(data.kakaoNickname || "").trim();
      var reqPhone4 = String(data.phone4 || "").replace(/[^0-9]/g, "").slice(-4);

      var restrictionSync = checkClientAccessRestriction(reqNickSync, reqClientSync);
      if (restrictionSync.blocked) {
        return ContentService.createTextOutput(JSON.stringify({
          "result": "blocked",
          "message": restrictionSync.reason
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var syncedProfile = syncKakaoMemberProfile(reqNickSync, reqClientSync, reqKakaoNick, reqPhone4, false);
      var effectiveIdent = reqNickSync || (syncedProfile && syncedProfile.identifier) || "";
      if (effectiveIdent || reqClientSync) {
        updateClientRecordOnAction(effectiveIdent || reqKakaoNick || "미입력고객", reqClientSync, false);
        if (reqClientSync) {
          updateActivePresenceCache(reqClientSync, effectiveIdent || reqKakaoNick || "미입력고객");
        }
      }

      var resPayloadSync = buildReservationCheckPayload(ss, sheet, effectiveIdent, reqClientSync, false, reqKakaoNick, reqPhone4);
      var settingsPayloadSync = buildFullSettingsPayload(ss, sheet, effectiveIdent, reqClientSync);

      return ContentService.createTextOutput(JSON.stringify({
        "result": "ok",
        "memberProfile": syncedProfile,
        "settings": settingsPayloadSync,
        "reservation": resPayloadSync
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "initAll") {
      var reqNick = String(data.nickname || "").trim();
      var reqClient = String(data.clientId || data.kakaoId || "").trim();
      var reqKakaoNickInit = String(data.kakaoNickname || "").trim();
      var reqPhone4Init = String(data.phone4 || "").trim();

      var resPayload = buildReservationCheckPayload(ss, sheet, reqNick, reqClient, true, reqKakaoNickInit, reqPhone4Init);
      var effectiveNick = (resPayload && resPayload.recoveredNickname) ? resPayload.recoveredNickname : reqNick;

      if (reqClient) {
        updateActivePresenceCache(reqClient, effectiveNick || reqKakaoNickInit || "미입력고객");
      }

      var settingsPayload = buildFullSettingsPayload(ss, sheet, effectiveNick, reqClient);

      return ContentService.createTextOutput(JSON.stringify({
        "settings": settingsPayload,
        "reservation": resPayload,
        "memberProfile": resPayload ? resPayload.memberProfile : null
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'getAdminOrderList') {
      var adminOrders = getAdminOrderList();
      return ContentService.createTextOutput(JSON.stringify(adminOrders))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'updateOrderStatus') {
      var res = updateOrderStatus(data.rowIndex, data.status);
      return ContentService.createTextOutput(JSON.stringify(res))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "create") {
      var customerInfo = String(data.nickname || "").trim();
      var clientIdVal = String(data.clientId || data.kakaoId || "").trim();
      var kakaoNickVal = String(data.kakaoNickname || "").trim();
      var phone4Val = String(data.phone4 || "").replace(/[^0-9]/g, "").slice(-4);
      var formattedKakaoId = formatKakaoDisplayId(clientIdVal);
      var normClientId = normalizeKakaoId(clientIdVal);

      if (!customerInfo || customerInfo === "미입력고객" || customerInfo === "EXIT" || customerInfo.length < 2) {
        return ContentService.createTextOutput(JSON.stringify({
          "result": "blocked",
          "message": "올바른 닉네임과 전화번호 뒷자리를 입력해 주세요!"
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var restrictionCheck = checkClientAccessRestriction(customerInfo, clientIdVal);
      if (restrictionCheck.blocked) {
        return ContentService.createTextOutput(JSON.stringify({
          "result": "blocked",
          "message": restrictionCheck.reason
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var settings = getStoreSettings(ss);
      if (checkIsClosedNow(settings)) {
        return ContentService.createTextOutput(JSON.stringify({"result": "closed"}))
                             .setMimeType(ContentService.MimeType.JSON);
      }

      var menuText = data.menuSummary || data.menu || "";
      var sauceText = data.sauce || "없음";

      for (var m = 0; m < settings.menuList.length; m++) {
        var menuObj = settings.menuList[m];
        if (menuObj.isSoldOut) {
          var sName = menuObj.shortName;
          var regex = new RegExp(sName + "\\s*(\\d+)개");
          if (menuText.match(regex)) {
            return ContentService.createTextOutput(JSON.stringify({
              "result": "out_of_stock",
              "menuName": sName
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }

      for (var s = 0; s < settings.sauceList.length; s++) {
        var sauceObj = settings.sauceList[s];
        if (sauceObj.isSoldOut) {
          var sSauceName = sauceObj.shortName || sauceObj.name;
          var regexSauce = new RegExp(sSauceName + "\\s*(\\d+)개");
          if (sauceText.match(regexSauce)) {
            return ContentService.createTextOutput(JSON.stringify({
              "result": "out_of_stock",
              "menuName": sSauceName
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }

      var lastRowCreate = sheet.getLastRow();
      var aliasesCreate = normClientId ? getMemberIdentifiersByKakaoId(normClientId) : [];
      var todayWinnerCount = 0;
      var userSecretCode = String(data.secretCode || "").trim();
      var shouldCountWinners = !!(settings.secretCode && userSecretCode !== "" && userSecretCode === settings.secretCode);

      if (lastRowCreate >= 2) {
        var orderSlice = sheet.getRange(2, 1, lastRowCreate - 1, 10).getValues();
        for (var idx = 0; idx < orderSlice.length; idx++) {
          var rowStatus = String(orderSlice[idx][0] || "").trim();
          if (rowStatus === "취소") continue;

          var dupNick = String(orderSlice[idx][5] || "").trim();
          var dupKakaoId = String(orderSlice[idx][9] || "").trim();
          if (rowStatus !== "완료" && isCustomerMatchForKakao(dupNick, customerInfo, clientIdVal, dupKakaoId, aliasesCreate)) {
            return ContentService.createTextOutput(JSON.stringify({
              "result": "duplicate_nickname",
              "message": "이미 동일한 계정(" + customerInfo + ")으로 접수된 예약이 존재합니다. 중복 예약은 불가능합니다!"
            })).setMimeType(ContentService.MimeType.JSON);
          }

          if (shouldCountWinners) {
            var eventVal = String(orderSlice[idx][8] || "").trim();
            if (eventVal.indexOf("당첨") !== -1) {
              todayWinnerCount++;
            }
          }
        }
      }

      var rawCook = data.cookType || "";
      var shortCook = "튀김";
      if (rawCook.includes("생고기") || rawCook.includes("안튀김") || rawCook.includes("비튀김")) {
        shortCook = "비튀김";
      }

      var receiptTime = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
      var totalPriceInput = data.totalPrice || "";
      var totalPriceVal = totalPriceInput.toString().replace(/[^0-9]/g, '');
      var finalPriceDisplay = totalPriceVal ? Number(totalPriceVal) : "";

      var isWinner = false;
      if (shouldCountWinners && todayWinnerCount < settings.maxEventWinners) {
        if (Math.random() * 100 < settings.eventWinProbability) {
          isWinner = true;
        }
      }

      sheet.appendRow([
        "미확인",
        data.pickupTime || "",
        shortCook,
        menuText,
        sauceText,
        customerInfo,
        receiptTime,
        finalPriceDisplay,
        isWinner ? "당첨 (선택대기)" : "",
        formattedKakaoId
      ]);

      if (normClientId) {
        syncKakaoMemberProfile(customerInfo, clientIdVal, kakaoNickVal, phone4Val, true);
      }
      updateClientRecordOnAction(customerInfo, clientIdVal, true);
      if (formattedKakaoId) {
        updateActivePresenceCache(formattedKakaoId, customerInfo);
      }

      var lastRowIdx = sheet.getLastRow();
      if (totalPriceVal) {
        sheet.getRange(lastRowIdx, 8).setNumberFormat('#,##0"원"');
      }

      try {
        CacheService.getScriptCache().remove("store_settings_cache_v5");
      } catch(e) {}

      backgroundTaskInfo = {
        type: "create",
        nickname: customerInfo,
        pickupTime: data.pickupTime || "",
        cookType: rawCook,
        menuSummary: menuText,
        sauce: sauceText,
        totalPrice: totalPriceVal ? Number(totalPriceVal).toLocaleString() + "원" : "0원"
      };

      var responseOutput = ContentService.createTextOutput(JSON.stringify({
        "result": "success",
        "isWinner": isWinner,
        "rowIndex": lastRowIdx,
        "eventMenuList": settings.eventMenuList,
        "syncedClientId": formattedKakaoId
      })).setMimeType(ContentService.MimeType.JSON);

      try { lock.releaseLock(); } catch (e) {}
      runPostOrderBackgroundTasks(backgroundTaskInfo);
      return responseOutput;
    }

    if (action === "checkNickname") {
      return ContentService.createTextOutput(JSON.stringify({"result": "ok"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "uploadMenuImage") {
      try {
        var requestData = JSON.parse(e.postData.contents);
        var base64Data = requestData.base64Data;
        var targetName = requestData.targetName;
        var fileName = requestData.fileName || "menu.jpg";

        if (!base64Data) {
          return ContentService.createTextOutput(JSON.stringify({"result": "error", "message": "이미지 데이터가 없습니다."})).setMimeType(ContentService.MimeType.JSON);
        }

        var imageUrl = uploadImageToDriveAndGetUrl(base64Data, fileName);
        var menuSheet = ss.getSheetByName("메뉴판");
        var finalUrlList = imageUrl;

        if (menuSheet) {
          var mLastRow = menuSheet.getLastRow();
          if (mLastRow >= 2) {
            var menuRows = menuSheet.getRange(2, 1, mLastRow - 1, 4).getValues();
            for (var i = 0; i < menuRows.length; i++) {
              var menuName = String(menuRows[i][0] || "").trim();
              if (menuName === targetName) {
                var existingUrl = String(menuRows[i][3] || "").trim();
                if (existingUrl !== "") {
                  finalUrlList = existingUrl + "," + imageUrl;
                }
                menuSheet.getRange(i + 2, 4).setValue(finalUrlList);
                break;
              }
            }
          }
        }

        try {
          CacheService.getScriptCache().remove("store_settings_cache_v5");
        } catch(e) {}

        return ContentService.createTextOutput(JSON.stringify({
          "result": "success",
          "imageUrl": finalUrlList
        })).setMimeType(ContentService.MimeType.JSON);

      } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({
          "result": "error",
          "message": err.toString()
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    if (action === "selectEventPrize") {
      var nickname = String(data.nickname || "").trim();
      var clientIdPrize = String(data.clientId || data.kakaoId || "").trim();
      var prize = String(data.prize || "").trim();
      var pLastRow = sheet.getLastRow();

      if (pLastRow >= 2) {
        var rows = sheet.getRange(2, 1, pLastRow - 1, 10).getValues();
        var aliasesPrize = normalizeKakaoId(clientIdPrize) ? getMemberIdentifiersByKakaoId(clientIdPrize) : [];
        for (var i = rows.length - 1; i >= 0; i--) {
          var status = String(rows[i][0] || "").trim();
          if (status === "취소") continue;

          var custCol = String(rows[i][5] || "").trim();
          var rowKakaoIdPrize = String(rows[i][9] || "").trim();
          if (isCustomerMatchForKakao(custCol, nickname, clientIdPrize, rowKakaoIdPrize, aliasesPrize)) {
            sheet.getRange(i + 2, 9).setValue("당첨 (" + prize + ")");
            return ContentService.createTextOutput(JSON.stringify({"result": "success"}))
                                       .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({"result": "not_found"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "verify") {
      var nickname = String(data.nickname || "").trim();
      var clientIdVerify = String(data.clientId || data.kakaoId || "").trim();
      var vLastRow = sheet.getLastRow();

      if (vLastRow >= 2) {
        var rows = sheet.getRange(2, 1, vLastRow - 1, 10).getValues();
        var aliasesVerify = normalizeKakaoId(clientIdVerify) ? getMemberIdentifiersByKakaoId(clientIdVerify) : [];
        for (var i = rows.length - 1; i >= 0; i--) {
          var status = String(rows[i][0] || "").trim();
          if (status === "취소") continue;

          var custCol = String(rows[i][5] || "").trim();
          var rowKakaoIdVerify = String(rows[i][9] || "").trim();
          if (isCustomerMatchForKakao(custCol, nickname, clientIdVerify, rowKakaoIdVerify, aliasesVerify)) {
            return ContentService.createTextOutput(JSON.stringify({"result": "exists"}))
                                       .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({"result": "not_exists"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "check" || action === "checkSimple") {
      var nickname = String(data.nickname || "").trim();
      var clientIdCheck = String(data.clientId || data.kakaoId || "").trim();
      var checkResult = buildReservationCheckPayload(ss, sheet, nickname, clientIdCheck, true, String(data.kakaoNickname || "").trim(), String(data.phone4 || "").trim());

      return ContentService.createTextOutput(JSON.stringify(checkResult))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "update") {
      var settingsCheck = getStoreSettings(ss);
      if (checkIsClosedNow(settingsCheck)) {
        return ContentService.createTextOutput(JSON.stringify({"result": "closed"}))
                             .setMimeType(ContentService.MimeType.JSON);
      }

      var rowIndex = Number(data.rowIndex);
      if (!rowIndex || rowIndex < 2) {
        return ContentService.createTextOutput(JSON.stringify({"result": "error"}))
                             .setMimeType(ContentService.MimeType.JSON);
      }

      var rowValues = sheet.getRange(rowIndex, 1, 1, 8).getValues()[0];
      var currentStatus = String(rowValues[0] || "").trim();
      if (currentStatus === "확인") {
        return ContentService.createTextOutput(JSON.stringify({"result": "locked"}))
                             .setMimeType(ContentService.MimeType.JSON);
      }

      var newMenuSummary = data.newMenu || "";
      var newSauceSummary = data.newSauce || "없음";

      var oldMenuStr = String(rowValues[3] || "");
      var oldSauceStr = String(rowValues[4] || "");
      var oldMenuMap = parseItemsToMap(oldMenuStr);
      var oldSauceMap = parseItemsToMap(oldSauceStr);

      var newMenuMap = parseItemsToMap(newMenuSummary);
      var newSauceMap = parseItemsToMap(newSauceSummary);

      for (var m = 0; m < settingsCheck.menuList.length; m++) {
        var menuObj = settingsCheck.menuList[m];
        if (menuObj.isSoldOut) {
          var sName = menuObj.shortName;
          var requestedQty = newMenuMap[sName] || 0;
          var previouslyReservedQty = oldMenuMap[sName] || 0;
          if (requestedQty > previouslyReservedQty) {
            return ContentService.createTextOutput(JSON.stringify({
              "result": "out_of_stock",
              "menuName": sName
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }

      for (var s = 0; s < settingsCheck.sauceList.length; s++) {
        var sauceObj = settingsCheck.sauceList[s];
        if (sauceObj.isSoldOut) {
          var sSauceName = sauceObj.shortName || sauceObj.name;
          var requestedQtyS = newSauceMap[sSauceName] || 0;
          var previouslyReservedQtyS = oldSauceMap[sSauceName] || 0;
          if (requestedQtyS > previouslyReservedQtyS) {
            return ContentService.createTextOutput(JSON.stringify({
              "result": "out_of_stock",
              "menuName": sSauceName
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }

      var finalCookTypeStr = "튀겨서 포장";
      if (data.newCookType) {
        var rawCookUpdate = data.newCookType;
        finalCookTypeStr = rawCookUpdate;
        var shortCookUpdate = "튀김";
        if (rawCookUpdate.includes("생고기") || rawCookUpdate.includes("안튀김") || rawCookUpdate.includes("비튀김")) {
          shortCookUpdate = "비튀김";
        }
        sheet.getRange(rowIndex, 3).setValue(shortCookUpdate);
      }

      if (data.newMenu) sheet.getRange(rowIndex, 4).setValue(data.newMenu);
      if (data.newSauce) sheet.getRange(rowIndex, 5).setValue(data.newSauce);
      if (data.newPickupTime) sheet.getRange(rowIndex, 2).setValue(data.newPickupTime);

      var updatePriceNumeric = "";
      if (data.newTotalPrice) {
        var updatePriceVal = data.newTotalPrice.toString().replace(/[^0-9]/g, '');
        if (updatePriceVal) {
          updatePriceNumeric = Number(updatePriceVal);
          sheet.getRange(rowIndex, 8).setValue(updatePriceNumeric);
          sheet.getRange(rowIndex, 8).setNumberFormat('#,##0"원"');
        }
      }

      var custColVal = String(rowValues[5] || "").trim();
      var pickupDisplay = data.newPickupTime || sheet.getRange(rowIndex, 2).getDisplayValue();

      backgroundTaskInfo = {
        type: "update",
        nickname: custColVal,
        pickupTime: pickupDisplay,
        cookType: finalCookTypeStr,
        menuSummary: data.newMenu || oldMenuStr,
        sauce: data.newSauce || oldSauceStr,
        totalPrice: updatePriceNumeric ? updatePriceNumeric.toLocaleString() + "원" : "0원"
      };

      try {
        CacheService.getScriptCache().remove("store_settings_cache_v5");
      } catch(e) {}

      var updateResponseOutput = ContentService.createTextOutput(JSON.stringify({"result": "updated"}))
                                               .setMimeType(ContentService.MimeType.JSON);

      try { lock.releaseLock(); } catch (e) {}
      runPostOrderBackgroundTasks(backgroundTaskInfo);
      return updateResponseOutput;
    }

    return ContentService.createTextOutput(JSON.stringify({"result": "invalid"}))
                         .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({"result": "error", "message": err.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : '';
  var callback = e && e.parameter ? e.parameter.callback : '';

  if (action === 'getAdminOrderList') {
    var data = getAdminOrderList();
    return outputJsonP(data, callback);
  }

  if (action === 'getActiveUsers') {
    var data = getActiveUserList();
    return outputJsonP(data, callback);
  }

  if (action === 'updateOrderStatus') {
    var rowIndex = e.parameter.rowIndex;
    var status = e.parameter.status;
    var res = updateOrderStatus(rowIndex, status);
    return outputJsonP(res, callback);
  }

  var page = e && e.parameter && e.parameter.p ? e.parameter.p : 'index';

  if (page === 'admin_upload') {
    var renderUploadUrl = 'https://donkatsu-booking.onrender.com/admin_upload.html';
    return HtmlService.createHtmlOutput('<script>window.location.replace("' + renderUploadUrl + '");</script>')
      .setTitle('장날동천 메뉴 사진 간편 등록')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var fileName = 'index';
  var titleName = '장날동천 수제 돈가스 간편 예약';

  if (page === 'admin' || page === 'joobang-board') {
    fileName = 'joobang-board';
    titleName = '장날동천 주방 대시보드';
  }

  return HtmlService.createHtmlOutputFromFile(fileName)
      .setTitle(titleName)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function outputJsonP(data, callbackName) {
  var jsonString = JSON.stringify(data);
  var output = callbackName ? (callbackName + '(' + jsonString + ');') : jsonString;
  return ContentService.createTextOutput(output)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function getAdminOrderList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("시트1");
  if (!sheet) { sheet = ss.getSheets()[0]; }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var maxCols = Math.min(Math.max(sheet.getLastColumn(), 10), 10);
  var data = sheet.getRange(1, 1, lastRow, maxCols).getValues();
  if (data.length <= 1) return [];

  var orders = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] && !row[1] && !row[3]) continue;

    var status = String(row[0] || '미확인').trim();
    if (status === '완료' || status === '취소') continue;

    var rawTime = row[1];
    var formattedTime = "";

    if (rawTime instanceof Date) {
      var hours = rawTime.getHours();
      var minutes = rawTime.getMinutes();
      var ampm = hours >= 12 ? '오후' : '오전';
      var displayHours = hours % 12;
      displayHours = displayHours ? displayHours : 12;
      var minStr = minutes < 10 ? '0' + minutes : minutes;
      formattedTime = ampm + ' ' + displayHours + ':' + minStr;
    } else {
      formattedTime = String(rawTime || '');
    }

    var eventVal = String(row[8] || "").trim();
    var totalPriceVal = row[7] || "";

    orders.push({
      rowIndex: i + 1,
      status: status,
      pickupTime: formattedTime,
      cooking: row[2] || '',
      menu: row[3] || '',
      extraSauce: row[4] || '',
      nickname: row[5] || '',
      totalPrice: totalPriceVal,
      isEventWinner: eventVal.indexOf("당첨") !== -1,
      eventPrize: eventVal,
      kakaoId: row[9] || ''
    });
  }

  orders.sort(function(a, b) {
    return parseTimeToMin(a.pickupTime) - parseTimeToMin(b.pickupTime);
  });

  return orders;
}

function updateOrderStatus(rowIndex, newStatus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("시트1");
  if (!sheet) { sheet = ss.getSheets()[0]; }

  var rowNum = Number(rowIndex);
  if (rowNum && rowNum > 1) {
    sheet.getRange(rowNum, 1).setValue(newStatus);
    return { success: true, status: newStatus };
  }
  return { success: false, error: "Invalid row: " + rowIndex };
}

function getStoreSettings(ss) {
  var cache = CacheService.getScriptCache();
  var cachedSettings = cache.get("store_settings_cache_v5");
  var baseSettings = null;

  if (cachedSettings) {
    try {
      baseSettings = JSON.parse(cachedSettings);
    } catch (e) {}
  }

  var settingSheet = ss.getSheetByName("설정");
  var openHour = 13, openMin = 0, closeHour = 18, closeMin = 30;
  var pickupStartHour = 15, pickupStartMin = 0, pickupCloseHour = 20, pickupCloseMin = 0;

  var notificationStatus = "ON";
  var prepMinutes = 0;

  var maxEventWinners = 2;
  var eventWinProbability = 5;
  var eventMenuList = ["돈까스 1개", "새우까스 1개", "안심까스 1개"];
  var isMaintenance = false;
  var secretCode = "";
  var accessLimitDays = 30;

  if (!baseSettings) {
    if (settingSheet) {
      var openObj = parseTimeCell(settingSheet.getRange("B1"));
      openHour = openObj.hour;
      openMin = openObj.min;

      var closeObj = parseTimeCell(settingSheet.getRange("B2"));
      closeHour = closeObj.hour;
      closeMin = closeObj.min;

      var pStartObj = parseTimeCell(settingSheet.getRange("F1"));
      pickupStartHour = pStartObj.hour;
      pickupStartMin = pStartObj.min;

      var pCloseObj = parseTimeCell(settingSheet.getRange("F2"));
      pickupCloseHour = pCloseObj.hour;
      pickupCloseMin = pCloseObj.min;

      var notiVal = String(settingSheet.getRange("B3").getDisplayValue()).trim().toUpperCase();
      if (notiVal === "OFF") { notificationStatus = "OFF"; }

      var delayVal = parseInt(settingSheet.getRange("B4").getValue(), 10);
      if (!isNaN(delayVal) && delayVal >= 0) { prepMinutes = delayVal; }

      var maxW = parseInt(settingSheet.getRange("B5").getValue(), 10);
      if (!isNaN(maxW)) maxEventWinners = maxW;

      var prob = parseFloat(settingSheet.getRange("B6").getValue());
      if (!isNaN(prob)) eventWinProbability = prob;

      var rawEventMenus = String(settingSheet.getRange("B7").getDisplayValue()).trim();
      if (rawEventMenus !== "") {
        eventMenuList = rawEventMenus.split(",").map(function(item) { return item.trim(); });
      }

      var maintVal = String(settingSheet.getRange("B9").getDisplayValue()).trim().toUpperCase();
      if (maintVal === "ON") { isMaintenance = true; }

      var codeCellVal = String(settingSheet.getRange("F5").getDisplayValue()).trim();
      if (codeCellVal !== "") { secretCode = codeCellVal; }

      var limitDaysVal = parseInt(settingSheet.getRange("F10").getValue(), 10);
      if (!isNaN(limitDaysVal) && limitDaysVal > 0) { accessLimitDays = limitDaysVal; }
    }

    var maintCode = "3808";
    if (settingSheet) {
      var codeVal = String(settingSheet.getRange("B10").getDisplayValue()).trim();
      if (codeVal) maintCode = codeVal;
    }

    baseSettings = {
      openHour: openHour, openMin: openMin, closeHour: closeHour, closeMin: closeMin,
      pickupStartHour: pickupStartHour, pickupStartMin: pickupStartMin,
      pickupCloseHour: pickupCloseHour, pickupCloseMin: pickupCloseMin,
      notification: notificationStatus, prepMinutes: prepMinutes,
      maxEventWinners: maxEventWinners, eventWinProbability: eventWinProbability,
      eventMenuList: eventMenuList, isMaintenance: isMaintenance,
      maintenanceCode: maintCode, secretCode: secretCode, accessLimitDays: accessLimitDays
    };
  }

  var menuSheet = ss.getSheetByName("메뉴판");
  var menuList = [];
  var sauceList = [];

  if (menuSheet) {
    var mLastRow = menuSheet.getLastRow();
    var menuRows = mLastRow >= 2 ? menuSheet.getRange(2, 1, mLastRow - 1, 8).getValues() : [];
    var isSauceSection = false;

    for (var i = 0; i < menuRows.length; i++) {
      var colA = String(menuRows[i][0] || "").trim();
      var colB = String(menuRows[i][1] || "").trim();
      var colC = String(menuRows[i][2] || "").trim();
      var colD = String(menuRows[i][3] || "").trim();
      var colE = String(menuRows[i][4] || "").trim();
      var colF = String(menuRows[i][5] || "").trim().toUpperCase();
      var colH = Number(menuRows[i][7]) || 0;

      if (colA.includes("추가소스") || colA.includes("소스분류") || colA.includes("추가소스 분류") || colB.includes("추가소스")) {
        isSauceSection = true;
        continue;
      }

      if (colA === "") continue;
      var isSoldOut = (colE === "품절");
      var isHot = (colF === "HOT");
      var hotOrder = isHot ? (colH > 0 ? colH : (i + 1)) : 0;

      if (isSauceSection) {
        sauceList.push({
          name: colA,
          shortName: colB || colA,
          price: colC || "₩500",
          imageUrl: colD,
          isSoldOut: isSoldOut
        });
      } else {
        menuList.push({
          name: colA,
          shortName: colB || colA,
          price: colC,
          imageUrl: colD,
          isSoldOut: isSoldOut,
          isHot: isHot,
          hotOrder: hotOrder
        });
      }
    }
  }

  if (sauceList.length === 0) {
    sauceList.push({ name: "돈까스 소스", shortName: "돈소스", price: "₩500", imageUrl: "", isSoldOut: false });
  }

  var finalSettings = Object.assign({}, baseSettings);
  finalSettings.menuList = menuList;
  finalSettings.sauceList = sauceList;

  try {
    cache.put("store_settings_cache_v5", JSON.stringify(finalSettings), 30);
  } catch (e) {}

  return finalSettings;
}

function parseTimeCell(cellRange) {
  var val = cellRange.getValue();
  var h = 0, m = 0;
  if (val instanceof Date) {
    h = val.getHours();
    m = val.getMinutes();
  } else {
    var str = String(cellRange.getDisplayValue()).trim();
    var parts = str.split(":");
    if (parts.length >= 2) {
      h = parseInt(parts[0], 10) || 0;
      m = parseInt(parts[1], 10) || 0;
    }
  }
  return { hour: h, min: m };
}

function uploadImageToDriveAndGetUrl(base64Data, fileName) {
  try {
    var splitData = base64Data.split(',');
    var contentType = splitData[0].match(/:(.*?);/)[1];
    var bytes = Utilities.base64Decode(splitData[1]);
    var blob = Utilities.newBlob(bytes, contentType, fileName);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuSheet = ss.getSheetByName("메뉴판");
    var folderName = "돈가스사진";
    if (menuSheet) {
      var g2Value = String(menuSheet.getRange("G2").getValue() || "").trim();
      if (g2Value !== "") {
        folderName = g2Value;
      }
    }

    var folders = DriveApp.getFoldersByName(folderName);
    var folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = folderName ? DriveApp.createFolder(folderName) : DriveApp.createFolder("돈가스사진");
    }

    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return "https://drive.google.com/uc?export=view&id=" + file.getId();
  } catch (e) {
    throw new Error("이미지 업로드 실패: " + e.toString());
  }
}

function backupAndClearDailyOrders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("시트1");
  var dbSheet = ss.getSheetByName("누적DB");

  if (!sheet) {
    SpreadsheetApp.getUi().alert("시트1을 찾을 수 없습니다.");
    return;
  }

  if (!dbSheet) {
    dbSheet = ss.insertSheet("누적DB");
    dbSheet.appendRow(["확인상태", "픽업시간", "조리방식", "메뉴", "추가소스", "닉네임+전화번호", "접수시간", "주문된 총 가격", "이벤트 당첨", "카카오 회원번호"]);
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("백업할 주문 데이터가 없습니다.");
    return;
  }

  var dataRange = sheet.getRange(2, 1, lastRow - 1, 10);
  var values = dataRange.getValues();

  dbSheet.getRange(dbSheet.getLastRow() + 1, 1, values.length, 10).setValues(values);
  dataRange.clearContent();

  SpreadsheetApp.getUi().alert("✅ 성공적으로 영업 마감 및 누적DB 백업이 완료되었습니다!");
}