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

function clearStoreCacheFast() {
  try {
    CacheService.getScriptCache().remove("store_settings_cache_v5");
  } catch (e) { }
}

function flushSheetChangesNow(ss) {
  try {
    SpreadsheetApp.flush();
  } catch (e) { }
  clearStoreCacheFast();
}

function extractHourMinuteFromDate(dateObj) {
  if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) return null;
  var ssTz = "Asia/Seoul";
  try {
    var activeSs = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSs) {
      ssTz = activeSs.getSpreadsheetTimeZone() || "Asia/Seoul";
    }
  } catch (e) { }
  var formatted = Utilities.formatDate(dateObj, ssTz, "HH:mm");
  var parts = formatted.split(":");
  if (parts.length >= 2) {
    return {
      hour: parseInt(parts[0], 10) || 0,
      min: parseInt(parts[1], 10) || 0
    };
  }
  return {
    hour: dateObj.getHours(),
    min: dateObj.getMinutes()
  };
}

function parseTimeToMin(timeVal) {
  if (timeVal === null || timeVal === undefined || timeVal === "") return 9999;

  if (timeVal instanceof Date) {
    var hm = extractHourMinuteFromDate(timeVal);
    if (hm) {
      return hm.hour * 60 + hm.min;
    }
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

function parseKoreanOpenDateCell(cellRange) {
  if (!cellRange) {
    return { raw: "", timestamp: 0, isFuture: false };
  }

  var rawVal = cellRange.getValue();
  var dispVal = String(cellRange.getDisplayValue() || "").trim();

  if (!rawVal && !dispVal) {
    return { raw: "", timestamp: 0, isFuture: false };
  }

  var targetDate = null;

  if (rawVal instanceof Date && !isNaN(rawVal.getTime())) {
    targetDate = new Date(rawVal.getTime());
  } else {
    var sourceStr = dispVal || String(rawVal || "").trim();
    var match = sourceStr.match(
      /(\d{4})\s*[.\-\/년]\s*(\d{1,2})\s*[.\-\/월]\s*(\d{1,2})\s*일?\s*(?:(오전|오후)\s*)?(\d{1,2})?:(\d{1,2})(?::(\d{1,2}))?/
    );

    if (match) {
      var year = parseInt(match[1], 10);
      var month = parseInt(match[2], 10) - 1;
      var day = parseInt(match[3], 10);
      var meridiem = match[4] || "";
      var hour = match[5] ? parseInt(match[5], 10) : 0;
      var minute = match[6] ? parseInt(match[6], 10) : 0;
      var second = match[7] ? parseInt(match[7], 10) : 0;

      if (meridiem === "오전" && hour === 12) {
        hour = 0;
      } else if (meridiem === "오후" && hour < 12) {
        hour += 12;
      }

      var isoLike = year + "-" +
        ("0" + (month + 1)).slice(-2) + "-" +
        ("0" + day).slice(-2) + "T" +
        ("0" + hour).slice(-2) + ":" +
        ("0" + minute).slice(-2) + ":" +
        ("0" + second).slice(-2) + "+09:00";
      var parsedIso = new Date(isoLike);
      if (!isNaN(parsedIso.getTime())) {
        targetDate = parsedIso;
      }
    } else {
      var fallbackDate = new Date(sourceStr);
      if (!isNaN(fallbackDate.getTime())) {
        targetDate = fallbackDate;
      }
    }
  }

  if (!targetDate || isNaN(targetDate.getTime())) {
    return { raw: dispVal || String(rawVal || "").trim(), timestamp: 0, isFuture: false };
  }

  var ts = targetDate.getTime();
  var nowTs = Date.now();
  var formattedLabel = dispVal || Utilities.formatDate(targetDate, "Asia/Seoul", "yyyy. M. d a h:mm:ss")
    .replace("AM", "오전")
    .replace("PM", "오후");

  return {
    raw: formattedLabel,
    timestamp: ts,
    isFuture: ts > nowTs
  };
}

function getKoreanCurrentTimeVal() {
  var nowStr = Utilities.formatDate(new Date(), "Asia/Seoul", "HH:mm");
  var parts = nowStr.split(":");
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
}

function isBeforeScheduledOpenDate(settings) {
  if (!settings) return false;
  var openTs = Number(settings.openDateTimestamp) || 0;
  if (!openTs) return false;
  return Date.now() < openTs;
}

function checkIsClosedNow(settings) {
  if (isBeforeScheduledOpenDate(settings)) {
    return true;
  }

  var currentTimeVal = getKoreanCurrentTimeVal();
  var openTimeVal = settings.openHour * 60 + settings.openMin;
  var closeTimeVal = settings.closeHour * 60 + settings.closeMin;

  if (closeTimeVal < openTimeVal) {
    return (currentTimeVal > closeTimeVal && currentTimeVal < openTimeVal);
  } else {
    return (currentTimeVal < openTimeVal || currentTimeVal > closeTimeVal);
  }
}

function normalizeAndValidatePickupTime(rawPickupTime, settings) {
  var s = String(rawPickupTime || "").trim();
  if (!s) return { valid: false, reason: "픽업시간이 비어 있습니다." };

  var minVal = parseTimeToMin(s);
  if (minVal === 9999 || minVal < 0 || minVal >= 1440) {
    return { valid: false, reason: "올바르지 않은 픽업시간 형식입니다." };
  }

  if (settings) {
    var startVal = (Number(settings.pickupStartHour) || 0) * 60 + (Number(settings.pickupStartMin) || 0);
    var closeVal = (Number(settings.pickupCloseHour) || 23) * 60 + (Number(settings.pickupCloseMin) || 59);
    if (closeVal < startVal) {
      if (minVal < startVal && minVal > closeVal) {
        return { valid: false, reason: "픽업 가능 시간대를 벗어났습니다." };
      }
    } else {
      if (minVal < startVal || minVal > closeVal) {
        return { valid: false, reason: "픽업 가능 시간대를 벗어났습니다." };
      }
    }
  }

  var h24 = Math.floor(minVal / 60);
  var m = minVal % 60;
  var ampm = h24 >= 12 ? "오후" : "오전";
  var h12 = h24 % 12;
  if (h12 === 0) h12 = 12;
  var mStr = m < 10 ? ("0" + m) : String(m);

  return {
    valid: true,
    normalized: ampm + " " + h12 + ":" + mStr,
    minutes: minVal
  };
}

function parsePriceNumber(priceVal) {
  if (typeof priceVal === "number") return isNaN(priceVal) ? 0 : priceVal;
  var digits = String(priceVal || "").replace(/[^0-9]/g, "");
  if (!digits) return 0;
  var n = parseInt(digits, 10);
  return isNaN(n) ? 0 : n;
}

function calculateServerOrderTotal(menuSummaryStr, sauceSummaryStr, settings) {
  if (!settings) return 0;
  var total = 0;

  var menuMap = parseItemsToMap(menuSummaryStr || "");
  var menuNames = Object.keys(menuMap);
  for (var i = 0; i < menuNames.length; i++) {
    var mName = menuNames[i];
    var qty = menuMap[mName] || 0;
    if (qty <= 0) continue;

    var unitPrice = 0;
    for (var m = 0; m < settings.menuList.length; m++) {
      var item = settings.menuList[m];
      var shortN = String(item.shortName || "").trim();
      var fullN = String(item.name || "").trim();
      if ((shortN && shortN === mName) || (fullN && fullN === mName) ||
        (shortN && mName.indexOf(shortN) !== -1) || (fullN && mName.indexOf(fullN) !== -1)) {
        unitPrice = parsePriceNumber(item.price);
        break;
      }
    }
    total += unitPrice * qty;
  }

  var sauceMap = parseItemsToMap(sauceSummaryStr || "");
  var sauceNames = Object.keys(sauceMap);
  for (var j = 0; j < sauceNames.length; j++) {
    var sName = sauceNames[j];
    var sQty = sauceMap[sName] || 0;
    if (sQty <= 0) continue;

    var sUnitPrice = 0;
    for (var s = 0; s < settings.sauceList.length; s++) {
      var sItem = settings.sauceList[s];
      var sShort = String(sItem.shortName || "").trim();
      var sFull = String(sItem.name || "").trim();
      if ((sShort && sShort === sName) || (sFull && sFull === sName) ||
        (sShort && sName.indexOf(sShort) !== -1) || (sFull && sName.indexOf(sFull) !== -1)) {
        sUnitPrice = parsePriceNumber(sItem.price);
        break;
      }
    }
    total += sUnitPrice * sQty;
  }

  return total;
}

function hasPositiveItemQuantity(menuSummaryStr) {
  var map = parseItemsToMap(menuSummaryStr || "");
  var keys = Object.keys(map);
  for (var i = 0; i < keys.length; i++) {
    if ((Number(map[keys[i]]) || 0) > 0) {
      return true;
    }
  }
  return false;
}

function isRapidDuplicateOrder(customerInfo, clientIdVal, menuText, pickupTime) {
  try {
    var cache = CacheService.getScriptCache();
    var normId = normalizeKakaoId(clientIdVal);
    var actorKey = normId ? ("kakao_" + normId) : ("nick_" + String(customerInfo || "").trim());
    if (!actorKey || actorKey === "nick_") return false;

    var sig = [actorKey, String(pickupTime || "").trim(), String(menuText || "").trim()].join("|");
    var cacheKey = "recent_order_v1_" + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, sig, Utilities.Charset.UTF_8)
      .map(function (b) { return ("0" + (b & 0xFF).toString(16)).slice(-2); })
      .join("");

    var prevRaw = cache.get(cacheKey);
    var nowMs = Date.now();
    if (prevRaw) {
      var prevMs = Number(prevRaw);
      if (!isNaN(prevMs) && (nowMs - prevMs) < 12000) {
        return true;
      }
    }
    cache.put(cacheKey, String(nowMs), 20);
    return false;
  } catch (e) {
    return false;
  }
}

function sortOrders(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    var maxCols = Math.max(sheet.getLastColumn(), 10);
    var range = sheet.getRange(2, 1, lastRow - 1, maxCols);
    var values = range.getValues();

    values.sort(function (a, b) {
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
    SpreadsheetApp.flush();
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
    try { tokenJson = JSON.parse(tokenRes.getContentText()); } catch (e) { }

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
    try { userJson = JSON.parse(userRes.getContentText()); } catch (e) { }

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

      var baseDate = dDate || cDate || null;

      if (baseDate) {
        var elapsedDays = Math.floor((todayDateObj.getTime() - baseDate.getTime()) / (1000 * 60 * 60 * 24));
        newStatus = (elapsedDays > limitDays) ? "시스템 차단" : "시스템 허용";
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
    SpreadsheetApp.flush();
  }
}

function runDailyAccessLimitCheck() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    checkAllClientsAccess();
  } catch (err) {
    Logger.log("일별 접속제한 점검 오류: " + err.toString());
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function onEdit(e) {
  try {
    var range = e && e.range;
    var sheet = range ? range.getSheet() : SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    if (sheetName === "시트1" && range && range.getRow() > 1) {
      var col = range.getColumn();
      var rowNum = range.getRow();
      if (col === 1) {
        var editedStatus = String(range.getValue() || "").trim();
        if (editedStatus === "취소") {
          restoreCouponsForCancelledOrder(sheet.getParent(), rowNum);
        }
      }
      if (col === 1 || col === 2) {
        sortOrders(sheet);
      }
    }

    if (sheetName === "메뉴판" && range && range.getRow() > 1 && range.getColumn() === 6) {
      var val = String(range.getValue() || "").trim().toUpperCase();
      var tsCell = sheet.getRange(range.getRow(), 8);
      if (val === "HOT") {
        tsCell.setValue(new Date().getTime());
      } else {
        tsCell.clearContent();
      }
    }

    if (sheetName === "설정" || sheetName === "메뉴판" || sheetName === "시트1") {
      clearStoreCacheFast();
      try { SpreadsheetApp.flush(); } catch (flushErr) { }
    }
  } catch (err) {
    Logger.log("onEdit 오류: " + err.toString());
  }
}

function matchSoldOutItemName(targetText, itemObj) {
  if (!targetText || !itemObj || !itemObj.isSoldOut) return "";
  var candidates = [
    String(itemObj.shortName || "").trim(),
    String(itemObj.name || "").trim()
  ];
  for (var i = 0; i < candidates.length; i++) {
    var label = candidates[i];
    if (!label) continue;
    var escaped = label.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    var regex = new RegExp(escaped + "\\s*(\\d+)개");
    if (regex.test(targetText)) {
      return label;
    }
  }
  return "";
}

function getRequestedQtyForSoldOutItem(qtyMap, itemObj) {
  if (!qtyMap || !itemObj) return 0;
  var candidates = [
    String(itemObj.shortName || "").trim(),
    String(itemObj.name || "").trim()
  ];
  var maxQty = 0;
  for (var i = 0; i < candidates.length; i++) {
    var key = candidates[i];
    if (key && qtyMap[key] !== undefined) {
      var q = Number(qtyMap[key]) || 0;
      if (q > maxQty) maxQty = q;
    }
  }
  return maxQty;
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

function getMenuUnitDifficultyWeight(menuName) {
  var normalized = String(menuName || "").replace(/\s+/g, "").toLowerCase();
  if (!normalized) return 3.0;

  if (normalized.indexOf("세트4") !== -1 || normalized.indexOf("새우5") !== -1) return 2.0;
  if (normalized.indexOf("세트1") !== -1 || normalized.indexOf("등심5") !== -1) return 3.0;
  if (normalized.indexOf("세트2") !== -1 || normalized.indexOf("치즈3") !== -1) return 3.5;
  if (normalized.indexOf("세트3") !== -1 || (normalized.indexOf("등심2") !== -1 && normalized.indexOf("치즈2") !== -1)) return 4.0;
  if (normalized.indexOf("세트6") !== -1 || (normalized.indexOf("등심2") !== -1 && normalized.indexOf("안심8") !== -1)) return 4.5;
  if (normalized.indexOf("모듬") !== -1 || normalized.indexOf("모둠") !== -1) return 5.0;
  if (normalized.indexOf("세트5") !== -1 || normalized.indexOf("닭안심13") !== -1 || normalized.indexOf("안심13") !== -1) return 6.0;

  return 3.0;
}

function calculateOrderFryerLoad(cookType, menuStr) {
  var cType = String(cookType || "").trim();
  if (cType.indexOf("비튀김") !== -1 || cType.indexOf("생고기") !== -1 || cType.indexOf("안튀김") !== -1) {
    return 0;
  }

  var rawMenu = String(menuStr || "").trim();
  if (!rawMenu || rawMenu === "없음") return 0;

  var parts = rawMenu.split(',');
  var totalLoad = 0;
  for (var i = 0; i < parts.length; i++) {
    var part = parts[i].trim();
    if (!part) continue;
    var match = part.match(/(.*?)\s*(\d+)개/);
    if (match) {
      var itemName = match[1].trim();
      var qty = parseInt(match[2], 10) || 0;
      if (qty > 0) {
        totalLoad += getMenuUnitDifficultyWeight(itemName) * qty;
      }
    } else {
      totalLoad += getMenuUnitDifficultyWeight(part);
    }
  }
  return Math.round(totalLoad * 100) / 100;
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

  var isPreOpenWait = isBeforeScheduledOpenDate(settings);
  var isClosedNow = checkIsClosedNow(settings);

  var rawPickupTimes = [];
  var loadByMin = {};
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var rangeVals = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    var dispTimeVals = sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues();
    for (var i = 0; i < rangeVals.length; i++) {
      var status = String(rangeVals[i][0] || "").trim();
      var pTimeDisp = String(dispTimeVals[i][0] || "").trim();
      var pTime = pTimeDisp || rangeVals[i][1];
      var cookType = String(rangeVals[i][2] || "").trim();
      var menuStr = String(rangeVals[i][3] || "").trim();

      // 삭제된 행(빈 값), 완료, 취소 상태는 합산에서 완전히 제외
      if (!pTime || !menuStr || status === "취소" || status === "완료" || status === "삭제") {
        continue;
      }

      var minVal = parseTimeToMin(pTime);
      if (minVal !== 9999 && minVal >= 0 && minVal < 1440) {
        var fryerLoad = calculateOrderFryerLoad(cookType, menuStr);
        if (fryerLoad > 0) {
          // 해당 픽업시간 기준 앞뒤 10분(±10분, 5분 단위 슬롯)에 교집합 누적 합산
          for (var offset = -10; offset <= 10; offset += 5) {
            var targetMin = minVal + offset;
            if (targetMin >= 0 && targetMin < 1440) {
              var prevVal = loadByMin[targetMin] || 0;
              loadByMin[targetMin] = Math.round((prevVal + fryerLoad) * 100) / 100;
            }
          }
        }
      }
    }
  }

  Object.keys(loadByMin)
    .map(function (k) { return Number(k); })
    .sort(function (a, b) { return a - b; })
    .forEach(function (mNum) {
      var totalLoad = loadByMin[mNum];
      var loadLabel = "";
      if (totalLoad >= 30) {
        loadLabel = "몰림";
      } else if (totalLoad >= 15) {
        loadLabel = "혼잡";
      }
      rawPickupTimes.push({
        time: mNum,
        load: loadLabel,
        totalLoad: totalLoad
      });
    });

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
  responseData.isPreOpenWait = isPreOpenWait;
  responseData.isClosed = isClosedNow;
  responseData.serverNowMs = Date.now();
  responseData.pickupTimesList = rawPickupTimes;
  responseData.fortuneList = fortuneList;
  responseData.fortuneEnabled = (String(fortuneEnabledVal).trim() !== "Off");
  responseData.isBlocked = restriction.blocked;
  responseData.blockReason = restriction.reason || "";

  return responseData;
}

function buildReservationCheckPayload(ss, sheet, nickname, clientIdCheck, shouldUpdateRecord, kakaoNickname, phone4) {
  var storeSettings = getStoreSettings(ss);
  var enrichedSettings = Object.assign({}, storeSettings, {
    isClosed: checkIsClosedNow(storeSettings),
    isPreOpenWait: isBeforeScheduledOpenDate(storeSettings)
  });

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
      "memberProfile": memberProfile,
      "availableCoupons": [],
      "settings": enrichedSettings
    };
  }

  var effectiveLookupNick = recoveredNickname || ((memberProfile && memberProfile.identifier) || "") || nickname;
  var availableCoupons = getAvailableCouponsForMember(ss, effectiveLookupNick, normId || clientIdCheck);

  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var maxCols = Math.min(Math.max(sheet.getLastColumn(), 12), 12);
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
        var savedCouponLabel = String(rows[i][10] || "").trim();
        var savedCouponDiscount = parsePriceNumber(rows[i][11] || 0);
        var matchedCouponId = findCouponIdForRow_(ss, i + 1, savedCouponLabel);

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
          "eventPrize": rows[i][8] || "",
          "usedCoupon": savedCouponLabel,
          "couponDiscount": savedCouponDiscount,
          "couponId": matchedCouponId,
          "availableCoupons": availableCoupons,
          "settings": enrichedSettings
        };
      }
    }
  }

  return {
    "result": "not_found",
    "recoveredNickname": recoveredNickname || ((memberProfile && memberProfile.identifier) || ""),
    "memberProfile": memberProfile,
    "availableCoupons": availableCoupons,
    "settings": enrichedSettings
  };
}

function findCouponIdForRow_(ss, rowIndex, savedCouponLabel) {
  var histSheet = ss.getSheetByName("쿠폰발급내역");
  if (!histSheet) return "";
  var lastRow = histSheet.getLastRow();
  if (lastRow < 2) return "";

  var rows = histSheet.getRange(2, 1, lastRow - 1, 10).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (Number(rows[i][9]) === Number(rowIndex)) {
      return String(rows[i][0] || "").trim();
    }
  }
  return "";
}

function renderPreOpenCountdownPage(settings) {
  var targetTs = Number(settings.openDateTimestamp) || 0;
  var openLabel = String(settings.openDateRaw || "오픈 일정 준비 중").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  var openTimeTxt = ("0" + settings.openHour).slice(-2) + ":" + ("0" + settings.openMin).slice(-2);
  var closeTimeTxt = ("0" + settings.closeHour).slice(-2) + ":" + ("0" + settings.closeMin).slice(-2);
  var pickupStartTxt = ("0" + settings.pickupStartHour).slice(-2) + ":" + ("0" + settings.pickupStartMin).slice(-2);
  var pickupCloseTxt = ("0" + settings.pickupCloseHour).slice(-2) + ":" + ("0" + settings.pickupCloseMin).slice(-2);

  var html = '<!DOCTYPE html>' +
    '<html lang="ko">' +
    '<head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
    '<title>장날동천 | 새로운 오픈 카운트다운</title>' +
    '<style>' +
    ':root { --bg-deep: #12100e; --card-bg: #1b1714; --card-border: rgba(212, 175, 110, 0.24); --gold: #d8b26e; --gold-soft: rgba(216, 178, 110, 0.14); --text-main: #f5efe6; --text-sub: #b8aea0; --text-muted: #857b6f; --btn-bg: #25201b; --btn-hover: #2f2822; }' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { min-height: 100vh; background: radial-gradient(circle at top, #241d17 0%, var(--bg-deep) 68%); color: var(--text-main); font-family: -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", "Pretendard", "Noto Sans KR", sans-serif; display: flex; align-items: center; justify-content: center; padding: 24px 16px; }' +
    '.wrap { width: 100%; max-width: 460px; }' +
    '.panel { background: linear-gradient(180deg, rgba(33, 27, 22, 0.96) 0%, rgba(22, 19, 16, 0.98) 100%); border: 1px solid var(--card-border); border-radius: 24px; padding: 34px 24px 28px; box-shadow: 0 24px 60px rgba(0, 0, 0, 0.55); text-align: center; }' +
    '.eyebrow { display: inline-flex; align-items: center; gap: 6px; padding: 6px 14px; border-radius: 999px; background: var(--gold-soft); border: 1px solid rgba(216, 178, 110, 0.35); color: var(--gold); font-size: 12px; font-weight: 600; letter-spacing: 0.08em; margin-bottom: 16px; }' +
    '.brand { font-size: 13px; color: var(--text-sub); letter-spacing: 0.12em; margin-bottom: 8px; }' +
    '.title { font-size: 24px; font-weight: 700; line-height: 1.35; letter-spacing: -0.02em; margin-bottom: 10px; }' +
    '.desc { font-size: 14px; line-height: 1.65; color: var(--text-sub); margin-bottom: 24px; }' +
    '.dday-badge { display: inline-block; font-size: 15px; font-weight: 700; color: var(--bg-deep); background: linear-gradient(135deg, #e5c485 0%, #c69b53 100%); padding: 6px 16px; border-radius: 999px; margin-bottom: 14px; letter-spacing: 0.04em; }' +
    '.timer-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; margin-bottom: 18px; }' +
    '.time-box { background: rgba(15, 12, 10, 0.72); border: 1px solid rgba(216, 178, 110, 0.18); border-radius: 14px; padding: 14px 6px; }' +
    '.time-num { font-size: 23px; font-weight: 700; color: var(--text-main); font-variant-numeric: tabular-nums; line-height: 1.2; }' +
    '.time-unit { font-size: 11px; color: var(--gold); margin-top: 4px; font-weight: 600; }' +
    '.open-schedule { font-size: 12.5px; color: var(--text-sub); background: rgba(255, 255, 255, 0.02); border-radius: 12px; padding: 11px 14px; border: 1px solid rgba(255, 255, 255, 0.05); margin-bottom: 22px; }' +
    '.open-schedule strong { color: var(--gold); font-weight: 600; }' +
    '.action-row { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 14px; }' +
    '.action-btn { appearance: none; border: 1px solid rgba(216, 178, 110, 0.28); background: var(--btn-bg); color: var(--text-main); border-radius: 14px; padding: 14px 12px; font-size: 14px; font-weight: 600; cursor: pointer; transition: all 0.2s ease; display: flex; flex-direction: column; align-items: center; gap: 4px; }' +
    '.action-btn:hover { background: var(--btn-hover); border-color: var(--gold); transform: translateY(-1px); }' +
    '.action-btn span.sub { font-size: 11px; font-weight: 400; color: var(--text-sub); }' +
    '.hours-note { font-size: 12px; color: var(--text-muted); line-height: 1.6; }' +
    '.modal-backdrop { position: fixed; inset: 0; background: rgba(0, 0, 0, 0.72); display: none; align-items: center; justify-content: center; padding: 20px; z-index: 99; }' +
    '.modal-backdrop.open { display: flex; }' +
    '.modal-card { width: 100%; max-width: 400px; background: #1d1814; border: 1px solid var(--card-border); border-radius: 20px; padding: 24px 20px; text-align: left; box-shadow: 0 20px 50px rgba(0, 0, 0, 0.65); }' +
    '.modal-title { font-size: 17px; font-weight: 700; color: var(--gold); margin-bottom: 12px; }' +
    '.modal-list { list-style: none; display: flex; flex-direction: column; gap: 10px; margin-bottom: 18px; font-size: 13.5px; line-height: 1.55; color: var(--text-main); }' +
    '.modal-list li { padding-left: 12px; border-left: 2px solid rgba(216, 178, 110, 0.45); color: var(--text-sub); }' +
    '.modal-close { width: 100%; border: none; border-radius: 12px; padding: 12px; background: linear-gradient(135deg, #e5c485 0%, #c69b53 100%); color: #15120f; font-size: 14px; font-weight: 700; cursor: pointer; }' +
    '</style>' +
    '</head>' +
    '<body>' +
    '<div class="wrap">' +
    '<section class="panel">' +
    '<div class="eyebrow">JANGNAL DONGCHEON · PRE-OPEN</div>' +
    '<div class="brand">장날동천 수제 돈가스</div>' +
    '<h1 class="title">잠시 숨을 고르고<br>더 단단한 맛으로 찾아옵니다</h1>' +
    '<p class="desc">중요한 일정으로 잠시 돈가스 튀김을 쉬어가는 동안입니다.<br>새로고침 없이도 오픈 순간까지 남은 시간이 실시간으로 흐릅니다.</p>' +
    '<div class="dday-badge" id="ddayBadge">D-DAY</div>' +
    '<div class="timer-grid">' +
    '<div class="time-box"><div class="time-num" id="cdDays">00</div><div class="time-unit">일</div></div>' +
    '<div class="time-box"><div class="time-num" id="cdHours">00</div><div class="time-unit">시간</div></div>' +
    '<div class="time-box"><div class="time-num" id="cdMins">00</div><div class="time-unit">분</div></div>' +
    '<div class="time-box"><div class="time-num" id="cdSecs">00</div><div class="time-unit">초</div></div>' +
    '</div>' +
    '<div class="open-schedule">오픈 예정 시각 · <strong>' + openLabel + '</strong></div>' +
    '<div class="action-row">' +
    '<button type="button" class="action-btn" id="btnSamchon">삼촌 1:1 문의<span class="sub">일정 및 예약 안내</span></button>' +
    '<button type="button" class="action-btn" id="btnRegularTip">단골 팁<span class="sub">오픈 직후 이용 요령</span></button>' +
    '</div>' +
    '<p class="hours-note">오픈 이후 예약 운영시간 ' + openTimeTxt + ' ~ ' + closeTimeTxt + ' (픽업 ' + pickupStartTxt + ' ~ ' + pickupCloseTxt + ')가 적용됩니다.</p>' +
    '</section>' +
    '</div>' +
    '<div class="modal-backdrop" id="infoModal">' +
    '<div class="modal-card">' +
    '<div class="modal-title" id="modalTitle">안내</div>' +
    '<ul class="modal-list" id="modalBodyList"></ul>' +
    '<button type="button" class="modal-close" id="modalCloseBtn">확인했습니다</button>' +
    '</div>' +
    '</div>' +
    '<script>' +
    '(function(){' +
    'var targetMs = ' + targetTs + ';' +
    'var elBadge = document.getElementById("ddayBadge");' +
    'var elD = document.getElementById("cdDays");' +
    'var elH = document.getElementById("cdHours");' +
    'var elM = document.getElementById("cdMins");' +
    'var elS = document.getElementById("cdSecs");' +
    'var modal = document.getElementById("infoModal");' +
    'var modalTitle = document.getElementById("modalTitle");' +
    'var modalList = document.getElementById("modalBodyList");' +
    'function pad2(n){ return n < 10 ? "0" + n : String(n); }' +
    'function tick(){' +
    'var diff = targetMs - Date.now();' +
    'if (diff <= 0) {' +
    'if (elBadge) elBadge.textContent = "D-DAY · OPEN";' +
    'if (elD) elD.textContent = "00";' +
    'if (elH) elH.textContent = "00";' +
    'if (elM) elM.textContent = "00";' +
    'if (elS) elS.textContent = "00";' +
    'var baseUrl = window.location.href.split("?")[0];' +
    'window.location.replace(baseUrl);' +
    'return;' +
    '}' +
    'var totalSec = Math.floor(diff / 1000);' +
    'var days = Math.floor(totalSec / 86400);' +
    'var hours = Math.floor((totalSec % 86400) / 3600);' +
    'var mins = Math.floor((totalSec % 3600) / 60);' +
    'var secs = totalSec % 60;' +
    'var ddayCount = Math.max(1, Math.ceil(diff / 86400000));' +
    'if (elBadge) elBadge.textContent = "D-" + ddayCount + " · 오픈 카운트다운";' +
    'if (elD) elD.textContent = pad2(days);' +
    'if (elH) elH.textContent = pad2(hours);' +
    'if (elM) elM.textContent = pad2(mins);' +
    'if (elS) elS.textContent = pad2(secs);' +
    '}' +
    'tick();' +
    'setInterval(tick, 1000);' +
    'function openModal(title, items){' +
    'modalTitle.textContent = title;' +
    'modalList.innerHTML = items.map(function(t){ return "<li>" + t + "</li>"; }).join("");' +
    'modal.classList.add("open");' +
    '}' +
    'document.getElementById("btnSamchon").addEventListener("click", function(){' +
    'openModal("삼촌 1:1 문의 안내", [' +
    '"오픈 전 휴무 및 특별 일정에 대한 개별 문의는 카카오 채널 또는 삼촌 1:1 메시지로 편하게 남겨주세요.",' +
    '"오픈 시각이 도래하면 설정된 예약 운영시간에 맞춰 순차적으로 예약을 도와드립니다."' +
    ']);' +
    '});' +
    'document.getElementById("btnRegularTip").addEventListener("click", function(){' +
    'openModal("장날동천 단골 팁", [' +
    '"오픈 시각이 지나면 시트에 등록된 예약 시작·마감 시간 기준으로 바로 예약이 열립니다.",' +
    '"원하는 픽업 시간대와 조리방식(튀김 / 비튀김), 추가소스를 미리 정해두시면 오픈 직후 빠르고 여유롭게 주문하실 수 있습니다."' +
    ']);' +
    '});' +
    'document.getElementById("modalCloseBtn").addEventListener("click", function(){ modal.classList.remove("open"); });' +
    'modal.addEventListener("click", function(e){ if (e.target === modal) modal.classList.remove("open"); });' +
    '})();' +
    '</script>' +
    '</body>' +
    '</html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('장날동천 오픈 카운트다운')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ "result": "error" }))
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

  if (action === 'submitCustomerFeedback') {
    var feedbackRes = handleSubmitCustomerFeedback(data);
    return ContentService.createTextOutput(JSON.stringify(feedbackRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "getSettings") {
    try {
      clearStoreCacheFast();
      var ssQuick = SpreadsheetApp.getActiveSpreadsheet();
      var sheetQuick = ssQuick.getSheetByName("시트1");
      if (!sheetQuick) { sheetQuick = ssQuick.getSheets()[0]; }
      var reqNickQuick = String(data.nickname || "").trim();
      var reqClientQuick = String(data.clientId || data.kakaoId || "").trim();
      var responseDataQuick = buildFullSettingsPayload(ssQuick, sheetQuick, reqNickQuick, reqClientQuick);

      return ContentService.createTextOutput(JSON.stringify(responseDataQuick))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (quickErr) {
      return ContentService.createTextOutput(JSON.stringify({ "result": "error", "message": quickErr.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return ContentService.createTextOutput(JSON.stringify({ "result": "error", "message": "서버가 혼잡합니다. 잠시 후 다시 시도해 주세요." }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var backgroundTaskInfo = null;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("시트1");
    if (!sheet) { sheet = ss.getSheets()[0]; }

    if (action === "syncKakaoMember") {
      clearStoreCacheFast();
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

      flushSheetChangesNow(ss);
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
      clearStoreCacheFast();
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
        "pickupTimesList": settingsPayload.pickupTimesList || [],
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
      clearStoreCacheFast();
      var customerInfo = String(data.nickname || "").trim();
      var clientIdVal = String(data.clientId || data.kakaoId || "").trim();
      var kakaoNickVal = String(data.kakaoNickname || "").trim();
      var phone4Val = String(data.phone4 || "").replace(/[^0-9]/g, "").slice(-4);
      var formattedKakaoId = formatKakaoDisplayId(clientIdVal);
      var normClientId = normalizeKakaoId(clientIdVal);
      var selectedCouponId = String(
        data.couponId || data.selectedCouponId || data.couponCode || ""
      ).trim();

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
        return ContentService.createTextOutput(JSON.stringify({ "result": "closed" }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      var menuText = data.menuSummary || data.menu || "";
      var sauceText = data.sauce || "없음";

      if (isRapidDuplicateOrder(customerInfo, clientIdVal, menuText, data.pickupTime)) {
        return ContentService.createTextOutput(JSON.stringify({
          "result": "duplicate_nickname",
          "message": "방금 동일한 예약 요청이 접수되었습니다. 잠시만 기다려 주세요!"
        })).setMimeType(ContentService.MimeType.JSON);
      }

      for (var m = 0; m < settings.menuList.length; m++) {
        var matchedMenu = matchSoldOutItemName(menuText, settings.menuList[m]);
        if (matchedMenu) {
          return ContentService.createTextOutput(JSON.stringify({
            "result": "out_of_stock",
            "menuName": matchedMenu
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      for (var s = 0; s < settings.sauceList.length; s++) {
        var matchedSauce = matchSoldOutItemName(sauceText, settings.sauceList[s]);
        if (matchedSauce) {
          return ContentService.createTextOutput(JSON.stringify({
            "result": "out_of_stock",
            "menuName": matchedSauce
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      // 첫 예약 여부 판별 (새 주문 행 추가 전에 먼저 확인)
      var hadPriorOrder = hasAnyPastOrderForMember(sheet, customerInfo, clientIdVal);

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
      var computedTotal = calculateServerOrderTotal(menuText, sauceText, settings);
      var clientPriceNum = parsePriceNumber(data.totalPrice || "");
      var finalPriceNum = computedTotal > 0 ? computedTotal : clientPriceNum;
      var finalPriceDisplay = finalPriceNum > 0 ? finalPriceNum : "";

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
        formattedKakaoId,
        "",
        ""
      ]);

      var lastRowIdx = sheet.getLastRow();
      if (finalPriceDisplay) {
        sheet.getRange(lastRowIdx, 8).setNumberFormat('#,##0"원"');
      }

      // 선택한 쿠폰이 있으면 검증 후 금액 차감 및 사용완료 처리
      var couponApplyRes = applyCouponToOrderRow(
        ss,
        sheet,
        lastRowIdx,
        selectedCouponId,
        finalPriceNum,
        customerInfo,
        clientIdVal
      );
      if (couponApplyRes && couponApplyRes.error) {
        sheet.deleteRow(lastRowIdx);
        flushSheetChangesNow(ss);
        return ContentService.createTextOutput(JSON.stringify({
          "result": "error",
          "message": couponApplyRes.error
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var effectiveFinalPriceNum = (couponApplyRes && couponApplyRes.applied)
        ? couponApplyRes.finalPrice
        : finalPriceNum;

      // 첫예약 쿠폰 / 랜덤 쿠폰 추첨 발급
      var awardedCoupons = processPostOrderCouponRewards(
        ss,
        sheet,
        lastRowIdx,
        customerInfo,
        clientIdVal,
        hadPriorOrder
      );

      if (normClientId) {
        syncKakaoMemberProfile(customerInfo, clientIdVal, kakaoNickVal, phone4Val, true);
      }
      updateClientRecordOnAction(customerInfo, clientIdVal, true);
      if (formattedKakaoId) {
        updateActivePresenceCache(formattedKakaoId, customerInfo);
      }

      flushSheetChangesNow(ss);

      var emailTotalText = effectiveFinalPriceNum
        ? (Number(effectiveFinalPriceNum).toLocaleString() + "원" + ((couponApplyRes && couponApplyRes.applied) ? (" (쿠폰 -" + Number(couponApplyRes.discountAmount).toLocaleString() + "원 적용)") : ""))
        : "0원";

      backgroundTaskInfo = {
        type: "create",
        nickname: customerInfo,
        pickupTime: data.pickupTime || "",
        cookType: rawCook,
        menuSummary: menuText,
        sauce: sauceText,
        totalPrice: emailTotalText
      };

      var firstIssuedCoupon = (Array.isArray(awardedCoupons) && awardedCoupons.length > 0)
        ? awardedCoupons[0]
        : null;

      var responseOutput = ContentService.createTextOutput(JSON.stringify({
        "result": "success",
        "isWinner": isWinner,
        "rowIndex": lastRowIdx,
        "eventMenuList": settings.eventMenuList,
        "syncedClientId": formattedKakaoId,
        "appliedCoupon": couponApplyRes,
        "awardedCoupons": awardedCoupons,
        "issuedCoupon": firstIssuedCoupon
      })).setMimeType(ContentService.MimeType.JSON);

      try { lock.releaseLock(); } catch (e) { }
      runPostOrderBackgroundTasks(backgroundTaskInfo);
      return responseOutput;
    }

    if (action === "checkNickname") {
      return ContentService.createTextOutput(JSON.stringify({ "result": "ok" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "uploadMenuImage") {
      try {
        var requestData = data || JSON.parse(e.postData.contents);
        var base64Data = requestData.base64Data || requestData.imageBase64;
        var targetName = String(requestData.targetName || requestData.menuName || "").trim();
        var fileName = requestData.fileName || (targetName ? (targetName + ".jpg") : "menu.jpg");

        if (!base64Data) {
          return ContentService.createTextOutput(JSON.stringify({ "result": "error", "message": "이미지 데이터가 없습니다." })).setMimeType(ContentService.MimeType.JSON);
        }

        var imageUrl = uploadImageToDriveAndGetUrl(base64Data, fileName);
        var menuSheet = ss.getSheetByName("메뉴판");
        var finalUrlList = imageUrl;

        if (menuSheet && targetName) {
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

        flushSheetChangesNow(ss);

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
            flushSheetChangesNow(ss);
            return ContentService.createTextOutput(JSON.stringify({ "result": "success" }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ "result": "not_found" }))
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
            return ContentService.createTextOutput(JSON.stringify({ "result": "exists" }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ "result": "not_exists" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "check" || action === "checkSimple") {
      clearStoreCacheFast();
      var nickname = String(data.nickname || "").trim();
      var clientIdCheck = String(data.clientId || data.kakaoId || "").trim();
      var shouldUpdateAccessRecord = (action === "check") && (data.isPolling !== true);
      var checkResult = buildReservationCheckPayload(
        ss,
        sheet,
        nickname,
        clientIdCheck,
        shouldUpdateAccessRecord,
        String(data.kakaoNickname || "").trim(),
        String(data.phone4 || "").trim()
      );
      var checkSettingsPayload = buildFullSettingsPayload(ss, sheet, nickname, clientIdCheck);
      checkResult.pickupTimesList = checkSettingsPayload.pickupTimesList || [];
      checkResult.myCoupons = getAvailableCouponsForMember(ss, nickname, clientIdCheck);
      checkResult.isClosed = Boolean(checkSettingsPayload.isClosed);
      checkResult.isPreOpenWait = Boolean(checkSettingsPayload.isPreOpenWait);
      checkResult.settings = Object.assign({}, checkResult.settings || {}, {
        isClosed: Boolean(checkSettingsPayload.isClosed),
        isPreOpenWait: Boolean(checkSettingsPayload.isPreOpenWait),
        openHour: checkSettingsPayload.openHour,
        openMin: checkSettingsPayload.openMin,
        closeHour: checkSettingsPayload.closeHour,
        closeMin: checkSettingsPayload.closeMin
      });

      return ContentService.createTextOutput(JSON.stringify(checkResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "update") {
      clearStoreCacheFast();
      var settingsCheck = getStoreSettings(ss);
      if (checkIsClosedNow(settingsCheck)) {
        return ContentService.createTextOutput(JSON.stringify({ "result": "closed" }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      var rowIndex = Number(data.rowIndex);
      if (!rowIndex || rowIndex < 2) {
        return ContentService.createTextOutput(JSON.stringify({ "result": "error" }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      var maxReadCols = Math.min(Math.max(sheet.getLastColumn(), 12), 12);
      var rowValues = sheet.getRange(rowIndex, 1, 1, maxReadCols).getValues()[0];
      var currentStatus = String(rowValues[0] || "").trim();
      if (currentStatus === "확인") {
        return ContentService.createTextOutput(JSON.stringify({ "result": "locked" }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      var oldMenuStr = String(rowValues[3] || "");
      var oldSauceStr = String(rowValues[4] || "");
      var savedUsedCoupon = String(rowValues[10] || "").trim();
      var savedCouponDiscount = Math.max(0, parsePriceNumber(rowValues[11] || 0));

      var rawNewMenu = (data.newMenu !== undefined) ? data.newMenu : data.menu;
      var rawNewSauce = (data.newSauce !== undefined) ? data.newSauce : data.sauce;
      var rawNewPickupTime = (data.newPickupTime !== undefined) ? data.newPickupTime : data.pickupTime;
      var rawNewCookType = (data.newCookType !== undefined) ? data.newCookType : data.cookType;
      var rawNewTotalPrice = (data.newTotalPrice !== undefined) ? data.newTotalPrice : data.totalPrice;

      var newMenuSummary = rawNewMenu !== undefined ? String(rawNewMenu).trim() : oldMenuStr;
      var newSauceSummary = rawNewSauce !== undefined ? (String(rawNewSauce).trim() || "없음") : (oldSauceStr || "없음");

      if (!newMenuSummary || newMenuSummary === "없음" || !hasPositiveItemQuantity(newMenuSummary)) {
        return ContentService.createTextOutput(JSON.stringify({
          "result": "invalid",
          "message": "최소 1개 이상의 메뉴를 선택해 주세요."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var validatedPickup = null;
      if (rawNewPickupTime !== undefined && String(rawNewPickupTime).trim() !== "") {
        validatedPickup = normalizeAndValidatePickupTime(rawNewPickupTime, settingsCheck);
        if (!validatedPickup.valid) {
          return ContentService.createTextOutput(JSON.stringify({
            "result": "invalid_time",
            "message": validatedPickup.reason
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }

      var oldMenuMap = parseItemsToMap(oldMenuStr);
      var oldSauceMap = parseItemsToMap(oldSauceStr);

      var newMenuMap = parseItemsToMap(newMenuSummary);
      var newSauceMap = parseItemsToMap(newSauceSummary);

      for (var m = 0; m < settingsCheck.menuList.length; m++) {
        var menuObj = settingsCheck.menuList[m];
        if (menuObj.isSoldOut) {
          var requestedQty = getRequestedQtyForSoldOutItem(newMenuMap, menuObj);
          var previouslyReservedQty = getRequestedQtyForSoldOutItem(oldMenuMap, menuObj);
          if (requestedQty > previouslyReservedQty) {
            return ContentService.createTextOutput(JSON.stringify({
              "result": "out_of_stock",
              "menuName": menuObj.shortName || menuObj.name
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }

      for (var s = 0; s < settingsCheck.sauceList.length; s++) {
        var sauceObj = settingsCheck.sauceList[s];
        if (sauceObj.isSoldOut) {
          var requestedQtyS = getRequestedQtyForSoldOutItem(newSauceMap, sauceObj);
          var previouslyReservedQtyS = getRequestedQtyForSoldOutItem(oldSauceMap, sauceObj);
          if (requestedQtyS > previouslyReservedQtyS) {
            return ContentService.createTextOutput(JSON.stringify({
              "result": "out_of_stock",
              "menuName": sauceObj.shortName || sauceObj.name
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }

      var finalCookTypeStr = "튀겨서 포장";
      if (rawNewCookType) {
        var rawCookUpdate = rawNewCookType;
        finalCookTypeStr = rawCookUpdate;
        var shortCookUpdate = "튀김";
        if (rawCookUpdate.includes("생고기") || rawCookUpdate.includes("안튀김") || rawCookUpdate.includes("비튀김")) {
          shortCookUpdate = "비튀김";
        }
        sheet.getRange(rowIndex, 3).setValue(shortCookUpdate);
      }

      sheet.getRange(rowIndex, 4).setValue(newMenuSummary);
      sheet.getRange(rowIndex, 5).setValue(newSauceSummary);
      if (validatedPickup) {
        sheet.getRange(rowIndex, 2).setValue(validatedPickup.normalized);
      }

      var recomputedBasePrice = calculateServerOrderTotal(newMenuSummary, newSauceSummary, settingsCheck);
      var clientUpdatePrice = parsePriceNumber(rawNewTotalPrice || "");
      var basePriceBeforeCoupon = recomputedBasePrice > 0 ? recomputedBasePrice : clientUpdatePrice;

      var custColVal = String(rowValues[5] || "").trim();
      var reqNickname = String(data.nickname || custColVal || "").trim();
      var reqClientId = String(data.clientId || data.kakaoId || rowValues[9] || "").trim();

      var hasCouponSelectionArg = (
        data.couponCode !== undefined ||
        data.couponId !== undefined ||
        data.appliedCouponCode !== undefined ||
        data.selectedCouponId !== undefined
      );
      var reqCouponId = String(
        (data.couponCode !== undefined ? data.couponCode :
          (data.couponId !== undefined ? data.couponId :
            (data.appliedCouponCode !== undefined ? data.appliedCouponCode :
              (data.selectedCouponId !== undefined ? data.selectedCouponId : "")))) || ""
      ).trim();

      var updatePriceNumeric = basePriceBeforeCoupon;
      var appliedDiscountNow = 0;
      var finalUsedCouponText = savedUsedCoupon;
      var finalCouponCodeOut = reqCouponId;
      var finalCouponNameOut = String(data.couponName || "").trim();

      if (hasCouponSelectionArg) {
        if (typeof restoreCouponsForOrderRow === "function") {
          restoreCouponsForOrderRow(ss, rowIndex);
        }
        if (reqCouponId) {
          var couponApplyRes = applyCouponToOrderRow(ss, sheet, rowIndex, reqCouponId, basePriceBeforeCoupon, reqNickname, reqClientId);
          if (couponApplyRes && couponApplyRes.error) {
            sheet.getRange(rowIndex, 8).setValue(basePriceBeforeCoupon > 0 ? basePriceBeforeCoupon : "");
            if (basePriceBeforeCoupon > 0) sheet.getRange(rowIndex, 8).setNumberFormat('#,##0"원"');
            sheet.getRange(rowIndex, 11).setValue("");
            sheet.getRange(rowIndex, 12).setValue("");
            flushSheetChangesNow(ss);
            return ContentService.createTextOutput(JSON.stringify({
              "result": "error",
              "message": couponApplyRes.error
            })).setMimeType(ContentService.MimeType.JSON);
          }
          if (couponApplyRes && couponApplyRes.applied) {
            appliedDiscountNow = Math.max(0, Number(couponApplyRes.discountAmount) || 0);
            updatePriceNumeric = Math.max(0, Number(couponApplyRes.finalPrice) || 0);
            finalCouponCodeOut = couponApplyRes.couponId || reqCouponId;
            finalCouponNameOut = couponApplyRes.couponName || finalCouponNameOut;
            finalUsedCouponText = finalCouponNameOut + " (-" + appliedDiscountNow.toLocaleString() + "원)";
          }
        } else {
          sheet.getRange(rowIndex, 8).setValue(basePriceBeforeCoupon > 0 ? basePriceBeforeCoupon : "");
          if (basePriceBeforeCoupon > 0) sheet.getRange(rowIndex, 8).setNumberFormat('#,##0"원"');
          sheet.getRange(rowIndex, 11).setValue("");
          sheet.getRange(rowIndex, 12).setValue("");
          appliedDiscountNow = 0;
          updatePriceNumeric = basePriceBeforeCoupon;
          finalUsedCouponText = "";
          finalCouponCodeOut = "";
          finalCouponNameOut = "";
        }
      } else {
        var effectiveSavedDiscount = savedCouponDiscount > 0 ? savedCouponDiscount : parsePriceNumber(savedUsedCoupon);
        appliedDiscountNow = effectiveSavedDiscount > 0 ? Math.min(basePriceBeforeCoupon, effectiveSavedDiscount) : 0;
        updatePriceNumeric = Math.max(0, basePriceBeforeCoupon - appliedDiscountNow);

        if (basePriceBeforeCoupon > 0) {
          sheet.getRange(rowIndex, 8).setValue(updatePriceNumeric);
          sheet.getRange(rowIndex, 8).setNumberFormat('#,##0"원"');
        }
        if (savedUsedCoupon) {
          sheet.getRange(rowIndex, 11).setValue(savedUsedCoupon);
          sheet.getRange(rowIndex, 12).setValue(appliedDiscountNow);
        }
      }

      var pickupDisplay = validatedPickup ? validatedPickup.normalized : sheet.getRange(rowIndex, 2).getDisplayValue();

      flushSheetChangesNow(ss);

      var emailUpdateTotalText = updatePriceNumeric
        ? (updatePriceNumeric.toLocaleString() + "원" + (appliedDiscountNow > 0 ? (" (쿠폰 -" + appliedDiscountNow.toLocaleString() + "원 적용)") : ""))
        : "0원";

      backgroundTaskInfo = {
        type: "update",
        nickname: custColVal,
        pickupTime: pickupDisplay,
        cookType: finalCookTypeStr,
        menuSummary: newMenuSummary,
        sauce: newSauceSummary,
        totalPrice: emailUpdateTotalText
      };

      var updateResponseOutput = ContentService.createTextOutput(JSON.stringify({
        "result": "success",
        "finalPrice": updatePriceNumeric,
        "couponCode": finalCouponCodeOut,
        "couponName": finalCouponNameOut,
        "couponDiscount": appliedDiscountNow,
        "usedCoupon": finalUsedCouponText
      })).setMimeType(ContentService.MimeType.JSON);

      try { lock.releaseLock(); } catch (e) { }
      runPostOrderBackgroundTasks(backgroundTaskInfo);
      return updateResponseOutput;
    }

    return ContentService.createTextOutput(JSON.stringify({ "result": "invalid" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ "result": "error", "message": err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function handleSubmitCustomerFeedback(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "고객의견_로그";
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["접수일시", "닉네임", "카카오ID", "의견내용", "반영상태", "삼촌메모"]);
  }

  var nowStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  var nickname = String((data && data.nickname) || "고객").trim();
  var kakaoId = String((data && data.kakaoId) || "").trim();
  var feedback = String((data && data.feedback) || "").trim();

  if (!feedback) {
    return { "result": "error", "message": "empty_feedback" };
  }

  sheet.appendRow([nowStr, nickname, kakaoId, feedback, "대기", ""]);
  return { "result": "success" };
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

  if (page === 'admin' || page === 'joobang-board') {
    return HtmlService.createHtmlOutputFromFile('joobang-board')
      .setTitle('장날동천 주방 대시보드')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var ssGet = SpreadsheetApp.getActiveSpreadsheet();
  var settingsGet = getStoreSettings(ssGet);
  if (page === 'dday' || isBeforeScheduledOpenDate(settingsGet)) {
    return renderPreOpenCountdownPage(settingsGet);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('장날동천 수제 돈가스 간편 예약')
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

  var maxCols = Math.min(Math.max(sheet.getLastColumn(), 12), 12);
  var data = sheet.getRange(1, 1, lastRow, maxCols).getValues();
  var dispTimes = sheet.getRange(1, 2, lastRow, 1).getDisplayValues();
  if (data.length <= 1) return [];

  var orders = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] && !row[1] && !row[3]) continue;

    var status = String(row[0] || '미확인').trim();
    if (status === '완료' || status === '취소' || status === '삭제') continue;

    var dispTimeStr = String(dispTimes[i][0] || '').trim();
    var formattedTime = dispTimeStr;

    if (!formattedTime) {
      var rawTime = row[1];
      if (rawTime instanceof Date) {
        var hm = extractHourMinuteFromDate(rawTime);
        var hours = hm ? hm.hour : rawTime.getHours();
        var minutes = hm ? hm.min : rawTime.getMinutes();
        var ampm = hours >= 12 ? '오후' : '오전';
        var displayHours = hours % 12;
        displayHours = displayHours ? displayHours : 12;
        var minStr = minutes < 10 ? '0' + minutes : minutes;
        formattedTime = ampm + ' ' + displayHours + ':' + minStr;
      } else {
        formattedTime = String(rawTime || '');
      }
    }

    var eventVal = String(row[8] || "").trim();
    var totalPriceVal = row[7] || "";
    var usedCouponVal = String(row[10] || "").trim();
    var couponDiscountVal = Number(row[11]) || 0;

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
      kakaoId: row[9] || '',
      usedCoupon: usedCouponVal,
      couponDiscount: couponDiscountVal
    });
  }

  orders.sort(function (a, b) {
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
    var cleanStatus = String(newStatus || "").trim();
    sheet.getRange(rowNum, 1).setValue(cleanStatus);
    if (cleanStatus === "취소") {
      restoreCouponsForCancelledOrder(ss, rowNum);
    }
    flushSheetChangesNow(ss);
    return { success: true, status: cleanStatus };
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
    } catch (e) { }
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
  var openDateInfo = { raw: "", timestamp: 0, isFuture: false };

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
        eventMenuList = rawEventMenus.split(",").map(function (item) { return item.trim(); });
      }

      var maintVal = String(settingSheet.getRange("B9").getDisplayValue()).trim().toUpperCase();
      if (maintVal === "ON") { isMaintenance = true; }

      var codeCellVal = String(settingSheet.getRange("F5").getDisplayValue()).trim();
      if (codeCellVal !== "") { secretCode = codeCellVal; }

      var limitDaysVal = parseInt(settingSheet.getRange("F10").getValue(), 10);
      if (!isNaN(limitDaysVal) && limitDaysVal > 0) { accessLimitDays = limitDaysVal; }

      openDateInfo = parseKoreanOpenDateCell(settingSheet.getRange("J1"));
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
      maintenanceCode: maintCode, secretCode: secretCode, accessLimitDays: accessLimitDays,
      openDateRaw: openDateInfo.raw, openDateTimestamp: openDateInfo.timestamp
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
          image: colD,
          imageUrl: colD,
          isSoldOut: isSoldOut
        });
      } else {
        menuList.push({
          name: colA,
          shortName: colB || colA,
          price: colC,
          image: colD,
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
    cache.put("store_settings_cache_v5", JSON.stringify(finalSettings), 10);
  } catch (e) { }

  return finalSettings;
}

function parseTimeCell(cellRange) {
  var dispStr = String(cellRange.getDisplayValue() || "").trim();
  if (dispStr) {
    var parsedMin = parseTimeToMin(dispStr);
    if (parsedMin !== 9999 && parsedMin >= 0 && parsedMin < 1440) {
      return {
        hour: Math.floor(parsedMin / 60),
        min: parsedMin % 60
      };
    }
  }

  var val = cellRange.getValue();
  var h = 0, m = 0;
  if (val instanceof Date) {
    var hm = extractHourMinuteFromDate(val);
    if (hm) {
      h = hm.hour;
      m = hm.min;
    } else {
      h = val.getHours();
      m = val.getMinutes();
    }
  } else {
    var str = dispStr || String(val || "").trim();
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
  flushSheetChangesNow(ss);

  SpreadsheetApp.getUi().alert("✅ 성공적으로 영업 마감 및 누적DB 백업이 완료되었습니다!");
}

// ==========================================
// [장날동천 할인쿠폰 엔진] 시트 기반 발급·차감·복구
// ==========================================

var COUPON_STATUS_AVAILABLE = "사용가능";
var COUPON_STATUS_USED = "사용완료";
var COUPON_STATUS_EXPIRED = "만료";
var COUPON_STATUS_CANCELLED = "취소무효";

function getCouponSettingsSheet(ss) {
  return ss.getSheetByName("쿠폰설정");
}

function getCouponHistorySheet(ss) {
  var sheet = ss.getSheetByName("쿠폰발급내역");
  if (!sheet) {
    sheet = ss.insertSheet("쿠폰발급내역");
    sheet.appendRow([
      "쿠폰ID", "카카오 회원번호", "예약 식별명", "쿠폰명",
      "할인금액", "발급유형", "발급일시", "만료일시",
      "상태", "사용주문행"
    ]);
  }
  return sheet;
}

function parseBoolOnOff(val, defaultTrue) {
  if (val === null || val === undefined || String(val).trim() === "") {
    return !!defaultTrue;
  }
  var s = String(val).trim().toUpperCase();
  if (s === "OFF" || s === "FALSE" || s === "0" || s === "N" || s === "NO") return false;
  if (s === "ON" || s === "TRUE" || s === "1" || s === "Y" || s === "YES") return true;
  return !!defaultTrue;
}

function getCouponConfig(ss) {
  var config = {
    firstOrderEnabled: true,
    firstOrderAmount: 1000,
    randomEnabled: true,
    randomWinProbability: 15,
    randomMaxWinnersPerDay: 3,
    validDays: 14,
    tiers: []
  };

  var sheet = getCouponSettingsSheet(ss);
  if (!sheet) {
    config.tiers = [
      { name: "천원 할인쿠폰", amount: 1000, weight: 10, enabled: true },
      { name: "이천원 할인쿠폰", amount: 2000, weight: 5, enabled: true },
      { name: "삼천원 할인쿠폰", amount: 3000, weight: 1, enabled: true }
    ];
    return config;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return config;

  // 상단 설정표 읽기 (A열 항목명, B열 설정값)
  var topRows = sheet.getRange(2, 1, Math.min(lastRow - 1, 20), 2).getValues();
  for (var i = 0; i < topRows.length; i++) {
    var key = String(topRows[i][0] || "").trim();
    var rawVal = topRows[i][1];
    if (!key) continue;

    if (key === "첫예약쿠폰_사용여부") {
      config.firstOrderEnabled = parseBoolOnOff(rawVal, true);
    } else if (key === "첫예약쿠폰_금액") {
      var fAmt = parsePriceNumber(rawVal);
      if (fAmt >= 0) config.firstOrderAmount = fAmt;
    } else if (key === "랜덤쿠폰_사용여부") {
      config.randomEnabled = parseBoolOnOff(rawVal, true);
    } else if (key === "랜덤쿠폰_당첨확률") {
      var prob = parseFloat(rawVal);
      if (!isNaN(prob) && prob >= 0) config.randomWinProbability = prob;
    } else if (key === "랜덤쿠폰_일최대당첨수") {
      var maxW = parseInt(rawVal, 10);
      if (!isNaN(maxW) && maxW >= 0) config.randomMaxWinnersPerDay = maxW;
    } else if (key === "쿠폰_유효기간일수") {
      var vDays = parseInt(rawVal, 10);
      if (!isNaN(vDays) && vDays > 0) config.validDays = vDays;
    }
  }

  // 하단 가중치 권종 목록 탐색 ("쿠폰명" 헤더 행 아래부터 행 추가 무제한 지원)
  var allRows = sheet.getRange(1, 1, lastRow, 4).getValues();
  var tierStartRowIdx = -1;
  for (var r = 0; r < allRows.length; r++) {
    if (String(allRows[r][0] || "").trim() === "쿠폰명") {
      tierStartRowIdx = r + 1;
      break;
    }
  }

  if (tierStartRowIdx >= 0 && tierStartRowIdx < allRows.length) {
    for (var t = tierStartRowIdx; t < allRows.length; t++) {
      var tName = String(allRows[t][0] || "").trim();
      var tAmt = parsePriceNumber(allRows[t][1]);
      var tWeight = parseFloat(allRows[t][2]);
      var tEnabled = parseBoolOnOff(allRows[t][3], true);

      if (!tName || tAmt <= 0 || isNaN(tWeight) || tWeight <= 0 || !tEnabled) {
        continue;
      }

      config.tiers.push({
        name: tName,
        amount: tAmt,
        weight: tWeight,
        enabled: true
      });
    }
  }

  return config;
}

function pickWeightedCouponTier(tiers) {
  if (!tiers || tiers.length === 0) return null;

  var totalWeight = 0;
  for (var i = 0; i < tiers.length; i++) {
    var w = Number(tiers[i].weight) || 0;
    if (w > 0 && (Number(tiers[i].amount) || 0) > 0) {
      totalWeight += w;
    }
  }
  if (totalWeight <= 0) return null;

  var roll = Math.random() * totalWeight;
  var acc = 0;
  for (var j = 0; j < tiers.length; j++) {
    var itemW = Number(tiers[j].weight) || 0;
    if (itemW <= 0) continue;
    acc += itemW;
    if (roll < acc) {
      return tiers[j];
    }
  }
  return tiers[tiers.length - 1];
}

function formatDateTimeSeoul(dateObj) {
  return Utilities.formatDate(dateObj, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
}

function parseDateSafeSeoul(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  var str = String(val).trim();
  if (!str) return null;
  // "YYYY-MM-DD HH:mm:ss" 형태를 KST(+09:00)로 해석
  if (/^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}(:\d{2})?$/.test(str)) {
    var isoStr = str.replace(/\s+/, "T") + (str.length <= 16 ? ":00" : "") + "+09:00";
    var d1 = new Date(isoStr);
    if (!isNaN(d1.getTime())) return d1;
  }
  var d2 = new Date(str);
  return isNaN(d2.getTime()) ? null : d2;
}

function refreshCouponExpirations(histSheet) {
  if (!histSheet) return;
  var lastRow = histSheet.getLastRow();
  if (lastRow < 2) return;

  var nowMs = Date.now();
  var range = histSheet.getRange(2, 1, lastRow - 1, 10);
  var rows = range.getValues();
  var dirty = false;

  for (var i = 0; i < rows.length; i++) {
    var status = String(rows[i][8] || "").trim();
    if (status === COUPON_STATUS_AVAILABLE || status === "AVAILABLE") {
      var expDate = parseDateSafeSeoul(rows[i][7]);
      if (expDate && nowMs > expDate.getTime()) {
        rows[i][8] = COUPON_STATUS_EXPIRED;
        dirty = true;
      } else if (status === "AVAILABLE") {
        rows[i][8] = COUPON_STATUS_AVAILABLE;
        dirty = true;
      }
    }
  }

  if (dirty) {
    range.setValues(rows);
  }
}

function getAvailableCouponsForMember(ss, nickname, clientId) {
  var normId = normalizeKakaoId(clientId);
  var cleanNick = String(nickname || "").trim();
  if (!normId && !cleanNick) return [];

  var histSheet = getCouponHistorySheet(ss);
  refreshCouponExpirations(histSheet);

  var lastRow = histSheet.getLastRow();
  if (lastRow < 2) return [];

  var aliases = normId ? getMemberIdentifiersByKakaoId(normId) : [];
  var rows = histSheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var nowMs = Date.now();
  var list = [];

  for (var i = 0; i < rows.length; i++) {
    var status = String(rows[i][8] || "").trim();
    if (status !== COUPON_STATUS_AVAILABLE && status !== "AVAILABLE") continue;

    var rowKakaoId = String(rows[i][1] || "").trim();
    var rowIdent = String(rows[i][2] || "").trim();
    if (!isCustomerMatchForKakao(rowIdent, cleanNick, normId, rowKakaoId, aliases)) continue;

    var expDate = parseDateSafeSeoul(rows[i][7]);
    if (expDate && nowMs > expDate.getTime()) continue;

    var amountNum = parsePriceNumber(rows[i][4]);
    if (amountNum <= 0) continue;

    list.push({
      couponId: String(rows[i][0] || "").trim(),
      couponName: String(rows[i][3] || "").trim(),
      discountAmount: amountNum,
      issueType: String(rows[i][5] || "").trim(),
      issuedAt: String(rows[i][6] || "").trim(),
      expiresAt: String(rows[i][7] || "").trim(),
      status: COUPON_STATUS_AVAILABLE
    });
  }

  return list;
}

function hasAnyPastOrderForMember(orderSheet, nickname, clientId) {
  var normId = normalizeKakaoId(clientId);
  var cleanNick = String(nickname || "").trim();
  if (!normId && !cleanNick) return false;

  var aliases = normId ? getMemberIdentifiersByKakaoId(normId) : [];

  // 현재 예약 시트(취소 제외) 이력 확인
  var lastRow = orderSheet.getLastRow();
  if (lastRow >= 2) {
    var rows = orderSheet.getRange(2, 1, lastRow - 1, 10).getValues();
    for (var i = 0; i < rows.length; i++) {
      var st = String(rows[i][0] || "").trim();
      if (st === "취소" || st === "삭제") continue;
      if (isCustomerMatchForKakao(rows[i][5], cleanNick, normId, rows[i][9], aliases)) {
        return true;
      }
    }
  }

  // 누적DB 이력도 함께 확인 (이미 기존에 주문했던 단골을 첫예약으로 오인하지 않도록)
  var ss = orderSheet.getParent();
  var dbSheet = ss.getSheetByName("누적DB");
  if (dbSheet && dbSheet.getLastRow() >= 2) {
    var dbRows = dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, Math.min(Math.max(dbSheet.getLastColumn(), 10), 10)).getValues();
    for (var j = 0; j < dbRows.length; j++) {
      var dbSt = String(dbRows[j][0] || "").trim();
      if (dbSt === "취소" || dbSt === "삭제") continue;
      if (isCustomerMatchForKakao(dbRows[j][5], cleanNick, normId, dbRows[j][9], aliases)) {
        return true;
      }
    }
  }

  // 과거 첫예약쿠폰 발급 이력이 있는지도 확인
  var histSheet = getCouponHistorySheet(ss);
  var hLast = histSheet.getLastRow();
  if (hLast >= 2) {
    var hRows = histSheet.getRange(2, 1, hLast - 1, 6).getValues();
    for (var k = 0; k < hRows.length; k++) {
      if (String(hRows[k][5] || "").trim() === "FIRST_ORDER") {
        if (isCustomerMatchForKakao(hRows[k][2], cleanNick, normId, hRows[k][1], aliases)) {
          return true;
        }
      }
    }
  }

  return false;
}

function getTodayRandomWinnerCount(histSheet) {
  var lastRow = histSheet.getLastRow();
  if (lastRow < 2) return 0;

  var todayPrefix = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
  var rows = histSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  var count = 0;

  for (var i = 0; i < rows.length; i++) {
    var issueType = String(rows[i][5] || "").trim();
    var issuedAt = String(rows[i][6] || "").trim();
    var status = String(rows[i][8] || "").trim();
    if (issueType === "RANDOM_WIN" && status !== COUPON_STATUS_CANCELLED && issuedAt.indexOf(todayPrefix) === 0) {
      count++;
    }
  }
  return count;
}

function issueCouponRecord(histSheet, kakaoId, nickname, couponName, amount, issueType, validDays) {
  var nowDate = new Date();
  var expDate = new Date(nowDate.getTime() + (Math.max(1, Number(validDays) || 14) * 24 * 60 * 60 * 1000));
  var couponId = "CPN-" + Utilities.formatDate(nowDate, "Asia/Seoul", "yyyyMMddHHmmss") + "-" + Math.floor(100 + Math.random() * 900);
  var issuedStr = formatDateTimeSeoul(nowDate);
  var expiresStr = formatDateTimeSeoul(expDate);
  var displayKakaoId = formatKakaoDisplayId(kakaoId);

  histSheet.appendRow([
    couponId,
    displayKakaoId,
    String(nickname || "").trim(),
    String(couponName || "").trim(),
    Number(amount) || 0,
    issueType,
    issuedStr,
    expiresStr,
    COUPON_STATUS_AVAILABLE,
    ""
  ]);

  return {
    couponId: couponId,
    kakaoId: displayKakaoId,
    nickname: String(nickname || "").trim(),
    couponName: String(couponName || "").trim(),
    discountAmount: Number(amount) || 0,
    issueType: issueType,
    issuedAt: issuedStr,
    expiresAt: expiresStr,
    status: COUPON_STATUS_AVAILABLE
  };
}

function applyCouponToOrderRow(ss, orderSheet, rowIndex, couponId, basePrice, nickname, clientId) {
  var cleanCouponId = String(couponId || "").trim();
  var baseNum = Math.max(0, Number(basePrice) || 0);

  if (!cleanCouponId) {
    orderSheet.getRange(rowIndex, 8).setValue(baseNum > 0 ? baseNum : "");
    if (baseNum > 0) orderSheet.getRange(rowIndex, 8).setNumberFormat('#,##0"원"');
    orderSheet.getRange(rowIndex, 11).setValue("");
    orderSheet.getRange(rowIndex, 12).setValue("");
    return {
      applied: false,
      couponId: "",
      couponName: "",
      discountAmount: 0,
      finalPrice: baseNum
    };
  }

  var histSheet = getCouponHistorySheet(ss);
  refreshCouponExpirations(histSheet);

  var lastRow = histSheet.getLastRow();
  var normId = normalizeKakaoId(clientId);
  var cleanNick = String(nickname || "").trim();
  var aliases = normId ? getMemberIdentifiersByKakaoId(normId) : [];

  if (lastRow >= 2) {
    var rows = histSheet.getRange(2, 1, lastRow - 1, 10).getValues();
    for (var i = 0; i < rows.length; i++) {
      var rowCouponId = String(rows[i][0] || "").trim();
      if (rowCouponId !== cleanCouponId) continue;

      var status = String(rows[i][8] || "").trim();
      var rowKakaoId = String(rows[i][1] || "").trim();
      var rowIdent = String(rows[i][2] || "").trim();

      var ownerMatch = (!normId && !cleanNick) || isCustomerMatchForKakao(rowIdent, cleanNick, normId, rowKakaoId, aliases);
      if (!ownerMatch) {
        return { error: "본인 계정의 쿠폰만 사용할 수 있습니다." };
      }
      if (status !== COUPON_STATUS_AVAILABLE) {
        return { error: "이미 사용되었거나 만료된 쿠폰입니다." };
      }

      var expDate = parseDateSafeSeoul(rows[i][7]);
      if (expDate && Date.now() > expDate.getTime()) {
        histSheet.getRange(i + 2, 9).setValue(COUPON_STATUS_EXPIRED);
        return { error: "유효기간이 지난 쿠폰입니다." };
      }

      var cName = String(rows[i][3] || "할인쿠폰").trim();
      var rawAmt = parsePriceNumber(rows[i][4]);
      var discountApplied = Math.min(baseNum, Math.max(0, rawAmt));
      var finalPrice = Math.max(0, baseNum - discountApplied);

      // 쿠폰 상태 -> 사용완료 및 사용주문행 기록
      histSheet.getRange(i + 2, 9).setValue(COUPON_STATUS_USED);
      histSheet.getRange(i + 2, 10).setValue(rowIndex);

      // 시트1 H열(차감된 최종 결제금액), K열(사용쿠폰), L열(쿠폰할인액) 기록
      orderSheet.getRange(rowIndex, 8).setValue(finalPrice);
      orderSheet.getRange(rowIndex, 8).setNumberFormat('#,##0"원"');
      orderSheet.getRange(rowIndex, 11).setValue(cName + " (-" + discountApplied.toLocaleString() + "원)");
      orderSheet.getRange(rowIndex, 12).setValue(discountApplied);

      return {
        applied: true,
        couponId: cleanCouponId,
        couponName: cName,
        discountAmount: discountApplied,
        finalPrice: finalPrice
      };
    }
  }

  return { error: "선택한 쿠폰을 찾을 수 없습니다." };
}

function processPostOrderCouponRewards(ss, orderSheet, rowIndex, nickname, clientId, hadPriorOrder) {
  var cfg = getCouponConfig(ss);
  var histSheet = getCouponHistorySheet(ss);
  var awardedCoupons = [];

  // 1) 최초 예약 완료 쿠폰 (기존 주문/발급 이력이 없는 경우만 확정 지급)
  if (cfg.firstOrderEnabled && !hadPriorOrder && cfg.firstOrderAmount > 0) {
    var firstCoupon = issueCouponRecord(
      histSheet,
      clientId,
      nickname,
      "첫예약 감사 " + Number(cfg.firstOrderAmount).toLocaleString() + "원 쿠폰",
      cfg.firstOrderAmount,
      "FIRST_ORDER",
      cfg.validDays
    );
    awardedCoupons.push(firstCoupon);
  }

// 2) 예약 완료 시 확률 랜덤 쿠폰
  if (cfg.randomEnabled && cfg.randomWinProbability > 0 && cfg.randomMaxWinnersPerDay > 0) {
    var lock = LockService.getScriptLock();
    var hasLock = false;
    try {
      hasLock = lock.tryLock(5000);
    } catch (e) {
      hasLock = false;
    }

    try {
      SpreadsheetApp.flush();
      var todayKey = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
      var cache = CacheService.getScriptCache();
      var cacheKey = "RANDOM_WIN_COUNT_" + todayKey;
      var sheetCount = getTodayRandomWinnerCount(histSheet);
      var cachedCount = parseInt(cache.get(cacheKey) || "0", 10) || 0;
      var todayCount = sheetCount > 0 ? Math.max(cachedCount, sheetCount) : sheetCount;
      if (sheetCount === 0 && cachedCount > 0) {
        cache.remove(cacheKey);
      }
      var alreadyWonThisOrder = hasAlreadyWonRandomCouponForOrder_(histSheet, clientId, nickname);

      if (!alreadyWonThisOrder && todayCount < cfg.randomMaxWinnersPerDay && cfg.tiers.length > 0) {
        if (Math.random() * 100 < cfg.randomWinProbability) {
          var pickedTier = pickWeightedCouponTier(cfg.tiers);
          if (pickedTier) {
            var randCoupon = issueCouponRecord(
              histSheet,
              clientId,
              nickname,
              pickedTier.name,
              pickedTier.amount,
              "RANDOM_WIN",
              cfg.validDays
            );
            SpreadsheetApp.flush();
            var nextCount = Math.max(getTodayRandomWinnerCount(histSheet), todayCount + 1);
            cache.put(cacheKey, String(nextCount), 21600);
            if (randCoupon) {
              awardedCoupons.push(randCoupon);
            }
          }
        }
      }
    } catch (err) {
      Logger.log("RANDOM_WIN 발급 중 오류: " + err);
    } finally {
      if (hasLock) {
        try {
          lock.releaseLock();
        } catch (releaseErr) {}
      }
    }
  }

  return awardedCoupons;
}

function hasAlreadyWonRandomCouponForOrder_(histSheet, clientId, nickname) {
  if (!histSheet) return false;
  var lastRow = histSheet.getLastRow();
  var lastCol = histSheet.getLastColumn();
  if (lastRow < 2 || lastCol < 7) return false;

  var cleanClientId = String(clientId || "").trim();
  var cleanNick = String(nickname || "").trim();
  var todayPrefix = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
  var rows = histSheet.getRange(2, 1, lastRow - 1, Math.min(lastCol, 9)).getValues();

  for (var i = 0; i < rows.length; i++) {
    var rowClientId = String(rows[i][1] || "").trim();
    var rowNick = String(rows[i][2] || "").trim();
    var issueType = String(rows[i][5] || "").trim();
    var issuedAt = String(rows[i][6] || "").trim();
    var status = rows[i].length > 8 ? String(rows[i][8] || "").trim() : "";

    if (issueType === "RANDOM_WIN" && status !== COUPON_STATUS_CANCELLED && issuedAt.indexOf(todayPrefix) === 0) {
      if (cleanClientId && rowClientId === cleanClientId) return true;
      if (!cleanClientId && cleanNick && rowNick === cleanNick) return true;
    }
  }
  return false;
}

// 주문이 '취소'로 바뀔 때: 해당 주문에서 사용한 쿠폰은 '사용가능'으로 복구(만료 지났으면 '만료'),
// 쿠폰 없는 주문은 조용히 예외 처리하여 통과
function restoreCouponsForCancelledOrder(ss, rowIndex) {
  var targetRow = Number(rowIndex);
  if (!targetRow || targetRow < 2) return { restoredCount: 0 };

  var histSheet = ss.getSheetByName("쿠폰발급내역");
  if (!histSheet) return { restoredCount: 0 };

  var lastRow = histSheet.getLastRow();
  if (lastRow < 2) return { restoredCount: 0 };

  var range = histSheet.getRange(2, 1, lastRow - 1, 10);
  var rows = range.getValues();
  var nowMs = Date.now();
  var restoredCount = 0;
  var dirty = false;

  for (var i = 0; i < rows.length; i++) {
    var usedRow = Number(rows[i][9]);
    if (usedRow === targetRow) {
      var curStatus = String(rows[i][8] || "").trim();
      if (curStatus === COUPON_STATUS_USED || curStatus === "USED") {
        var expDate = parseDateSafeSeoul(rows[i][7]);
        rows[i][8] = (expDate && nowMs > expDate.getTime()) ? COUPON_STATUS_EXPIRED : COUPON_STATUS_AVAILABLE;
        rows[i][9] = "";
        restoredCount++;
        dirty = true;
      }
    }
  }

  if (dirty) {
    range.setValues(rows);
  }
  return { restoredCount: restoredCount };
}

function restoreCouponsForOrderRow(ss, rowIndex) {
  var targetRow = Number(rowIndex);
  if (!targetRow || targetRow < 2) return { restoredCount: 0 };

  var histSheet = ss.getSheetByName("쿠폰발급내역");
  if (!histSheet) return { restoredCount: 0 };

  var lastRow = histSheet.getLastRow();
  if (lastRow < 2) return { restoredCount: 0 };

  var range = histSheet.getRange(2, 1, lastRow - 1, 10);
  var rows = range.getValues();
  var nowMs = Date.now();
  var restoredCount = 0;
  var dirty = false;

  for (var i = 0; i < rows.length; i++) {
    var usedRow = Number(rows[i][9]);
    if (usedRow === targetRow) {
      var curStatus = String(rows[i][8] || "").trim();
      if (curStatus === COUPON_STATUS_USED || curStatus === "USED") {
        var expDate = parseDateSafeSeoul(rows[i][7]);
        rows[i][8] = (expDate && nowMs > expDate.getTime()) ? COUPON_STATUS_EXPIRED : COUPON_STATUS_AVAILABLE;
        rows[i][9] = "";
        restoredCount++;
        dirty = true;
      }
    }
  }

  if (dirty) {
    range.setValues(rows);
  }
  return { restoredCount: restoredCount };
}

function syncCouponForUpdatedOrder(ss, orderSheet, rowIndex, couponId, basePrice, nickname, clientId) {
  restoreCouponsForOrderRow(ss, rowIndex);
  return applyCouponToOrderRow(ss, orderSheet, rowIndex, couponId, basePrice, nickname, clientId);
}