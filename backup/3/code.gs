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
               "• 닉네임/전화번호: " + nickname + "\n" +
               "• 주문 총 가격: " + totalPrice + "\n\n" +
               "📊 [예약 스프레드시트 바로가기]\n" + spreadsheetUrl + "\n\n" +
               "장날동천 수제 돈가스 예약 시스템에서 발송된 알림입니다.";
    
    MailApp.sendEmail(myEmail, subject, body);
    
    try {
      var threads = GmailApp.search('to:' + myEmail, 0, 1);
      if (threads.length > 0) {
        var label = GmailApp.getUserLabelByName("돈가스예약");
        if (label) {
          threads[0].addLabel(label);
        } else {
          var newLabel = GmailApp.createLabel("돈가스예약");
          threads[0].addLabel(newLabel);
        }
      }
    } catch (labelErr) {
      Logger.log("라벨 지정 중 예외 발생: " + labelErr.toString());
    }
  } catch (e) {
    Logger.log("🚨 이메일 발송 실패: " + e.toString());
  }
}

function sortOrders(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    var range = sheet.getRange(2, 1, lastRow - 1, 9);
    var values = range.getValues();
    
    values.sort(function(a, b) {
      var statusA = String(a[0]).trim();
      var statusB = String(b[0]).trim();
      
      if (statusA === "취소" && statusB !== "취소") return 1;
      if (statusA !== "취소" && statusB === "취소") return -1;
      
      var timeA = String(a[1]);
      var timeB = String(b[1]);
      return timeA.localeCompare(timeB);
    });
    
    range.setValues(values);
  }
}

function onEdit(e) {
  try {
    var range = e.range;
    var sheet = range.getSheet();
    if (sheet.getName() === "시트1" && range.getColumn() === 1 && range.getRow() > 1) {
      sortOrders(sheet);
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
      var qty = parseInt(match[2]) || 0;
      map[name] = qty;
    }
  }
  return map;
}

// --------------------------------------------------------------------------
// 📌 중복ID관리 시트 기반 엄격한 닉네임 중복 체크 함수
// --------------------------------------------------------------------------
function checkNicknameExistsInRegistry(nickname, clientId) {
  if (!nickname) return false;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var idSheet = ss.getSheetByName("중복ID관리");
  
  // 1. 중복ID관리 시트 검사
  if (idSheet) {
    var lastRow = idSheet.getLastRow();
    if (lastRow >= 1) {
      var values = idSheet.getRange(1, 1, lastRow, 2).getValues();
      for (var i = 0; i < values.length; i++) {
        var regNick = String(values[i][0] || "").trim();
        var regClient = String(values[i][1] || "").trim();
        
        if (regNick === nickname) {
          // 동일한 닉네임이 존재할 때, 기기 고유 ID(clientId)가 정확히 일치하면 본인임 (중복 아님)
          if (clientId && regClient && regClient === clientId) {
            return false;
          }
          // 기기가 다르다면 다른 사람이 이미 사용하는 닉네임이므로 중복 차단!
          return true;
        }
      }
    }
  }
  
  // 2. 만약 중복ID관리 시트에 없더라도 현재 시트1 또는 누적DB에 이미 사용 중인 활성 닉네임인지 2중 체크
  return checkNicknameExistsFallback(nickname);
}

/**
 * 닉네임 등록 또는 기존 본인 기기 갱신 처리
 */
function registerOrUpdateIdRegistry(nickname, clientId) {
  if (!nickname) return clientId || "";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var idSheet = ss.getSheetByName("중복ID관리");
  if (!idSheet) {
    idSheet = ss.insertSheet("중복ID관리");
    idSheet.appendRow(["고객 ID", "고유 ID", "확정시간"]);
  }
  
  var lastRow = idSheet.getLastRow();
  var values = lastRow > 0 ? idSheet.getRange(1, 1, lastRow, 3).getValues() : [];
  var foundRowByNick = -1;
  var foundRowByClient = -1;
  
  for (var i = 0; i < values.length; i++) {
    var regNick = String(values[i][0] || "").trim();
    var regClient = String(values[i][1] || "").trim();
    
    if (regNick === nickname) {
      foundRowByNick = i + 1;
    }
    if (clientId && regClient && regClient === clientId) {
      foundRowByClient = i + 1;
    }
  }
  
  var now = new Date();

  // 이미 내 기기(clientId)로 등록되어 있거나 닉네임이 일치하는 경우 갱신
  if (foundRowByNick > 0) {
    if (clientId) {
      idSheet.getRange(foundRowByNick, 2).setValue(clientId);
    }
    idSheet.getRange(foundRowByNick, 3).setValue(now);
    return clientId || String(values[foundRowByNick - 1][1] || "").trim();
  }
  
  if (foundRowByClient > 0) {
    idSheet.getRange(foundRowByClient, 1).setValue(nickname);
    idSheet.getRange(foundRowByClient, 3).setValue(now);
    return clientId;
  }
  
  // 완전 신규 등록
  idSheet.appendRow([nickname, clientId || "", now]);
  return clientId || "";
}

/**
 * clientId만으로 중복ID관리 시트에서 기존 고객 ID(닉네임+전화번호) 역추적 복구
 */
function findNicknameByClientId(clientId) {
  if (!clientId) return "";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var idSheet = ss.getSheetByName("중복ID관리");
  if (!idSheet) return "";
  
  var lastRow = idSheet.getLastRow();
  if (lastRow < 1) return "";
  
  var values = idSheet.getRange(1, 1, lastRow, 2).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    var regNick = String(values[i][0] || "").trim();
    var regClient = String(values[i][1] || "").trim();
    if (regClient === clientId && regNick !== "") {
      return regNick;
    }
  }
  return "";
}

function checkNicknameExistsFallback(nickname) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsToCheck = ["시트1", "누적DB"];
  for (var s = 0; s < sheetsToCheck.length; s++) {
    var sheet = ss.getSheetByName(sheetsToCheck[s]);
    if (!sheet) continue;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    var values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    for (var i = 0; i < values.length; i++) {
      var status = String(values[i][0] || "").trim();
      if (status === "취소") continue;
      var custCol = String(values[i][5] || "").trim();
      if (custCol === nickname) return true;
    }
  }
  return false;
}

// --------------------------------------------------------------------------
// 📌 실시간 접속자 세션 관리 함수
// --------------------------------------------------------------------------
function recordUserHeartbeat(data) {
  var cache = CacheService.getScriptCache();
  var clientId = data.clientId;
  var nickname = data.nickname ? String(data.nickname).trim() : "미입력고객";
  if (!nickname) nickname = "미입력고객";
  
  var sessionsJson = cache.get("active_user_sessions");
  var sessions = sessionsJson ? JSON.parse(sessionsJson) : {};
  
  var now = new Date().getTime();
  if (clientId && nickname !== 'EXIT') {
    sessions[clientId] = {
      nickname: nickname,
      lastActive: now
    };
  } else if (clientId && nickname === 'EXIT') {
    delete sessions[clientId];
  }
  
  var activeSessions = {};
  for (var id in sessions) {
    if (now - sessions[id].lastActive < 20000) {
      activeSessions[id] = sessions[id];
    }
  }
  
  cache.put("active_user_sessions", JSON.stringify(activeSessions), 60);
  return { success: true, count: Object.keys(activeSessions).length };
}

function getActiveUserList() {
  var cache = CacheService.getScriptCache();
  var sessionsJson = cache.get("active_user_sessions");
  var sessions = sessionsJson ? JSON.parse(sessionsJson) : {};
  
  var now = new Date().getTime();
  var userList = [];
  
  for (var id in sessions) {
    if (now - sessions[id].lastActive < 20000) {
      userList.push(sessions[id].nickname);
    }
  }
  
  return { count: userList.length, users: userList };
}

// --------------------------------------------------------------------------
// 📌 POST 요청 처리
// --------------------------------------------------------------------------
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("시트1");
    if (!sheet) { sheet = ss.getSheets()[0]; }
    
    var data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({"result": "error"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    var action = data.action;

    if (action === 'heartbeat') {
      var res = recordUserHeartbeat(data);
      return ContentService.createTextOutput(JSON.stringify(res))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'getActiveUsers') {
      var res = getActiveUserList();
      return ContentService.createTextOutput(JSON.stringify(res))
                           .setMimeType(ContentService.MimeType.JSON);
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

    function getKoreanCurrentTimeVal() {
      var nowStr = Utilities.formatDate(new Date(), "Asia/Seoul", "HH:mm");
      var parts = nowStr.split(":");
      return parseInt(parts[0]) * 60 + parseInt(parts[1]);
    }

    if (action === "getSettings") {
      var settings = getStoreSettings(ss);
      
      var currentTimeVal = getKoreanCurrentTimeVal();
      var openTimeVal = settings.openHour * 60 + settings.openMin;
      var closeTimeVal = settings.closeHour * 60 + settings.closeMin;

      var isClosedNow = false;
      if (closeTimeVal > 1440 && currentTimeVal < openTimeVal) {
          if ((currentTimeVal + 1440) > closeTimeVal) {
              isClosedNow = true;
          }
      } else {
          isClosedNow = (currentTimeVal < openTimeVal || currentTimeVal > closeTimeVal);
      }

      var rawPickupTimes = [];
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        var timeValues = sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues();
        var statusValues = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();

        for (var i = 0; i < timeValues.length; i++) {
          var status = String(statusValues[i][0] || "").trim();
          var pTime = String(timeValues[i][0] || "").trim();
          
          if (status !== "취소" && pTime !== "") {
            var minVal = convertTimeToMinutes(pTime);
            if (minVal !== null) {
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

      var responseData = settings;
      responseData.isClosed = isClosedNow;
      responseData.pickupTimesList = rawPickupTimes;
      responseData.fortuneList = fortuneList;
      responseData.fortuneEnabled = (String(fortuneEnabledVal).trim() !== "Off");

      return ContentService.createTextOutput(JSON.stringify(responseData))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "create") {
      var settings = getStoreSettings(ss);

      var currentTimeVal = getKoreanCurrentTimeVal();
      var openTimeVal = settings.openHour * 60 + settings.openMin;
      var closeTimeVal = settings.closeHour * 60 + settings.closeMin;
      
      var isClosedNow = false;
      if (closeTimeVal > 1440 && currentTimeVal < openTimeVal) {
          if ((currentTimeVal + 1440) > closeTimeVal) {
              isClosedNow = true;
          }
      } else {
          isClosedNow = (currentTimeVal < openTimeVal || currentTimeVal > closeTimeVal);
      }

      if (isClosedNow) {
        return ContentService.createTextOutput(JSON.stringify({"result": "closed"}))
                             .setMimeType(ContentService.MimeType.JSON);
      }

      var customerInfo = String(data.nickname || "").trim();
      var clientIdVal = String(data.clientId || "").trim();

      // [핵심 방어] 예약 생성 시점에 다른 사람이 이미 사용 중인 닉네임인지 엄격 재확인
      if (checkNicknameExistsInRegistry(customerInfo, clientIdVal)) {
        return ContentService.createTextOutput(JSON.stringify({
          "result": "duplicate_nickname"
        })).setMimeType(ContentService.MimeType.JSON);
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
      var userSecretCode = String(data.secretCode || "").trim();
      
      if (settings.secretCode && userSecretCode !== "" && userSecretCode === settings.secretCode) {
        var todayWinnerCount = 0;
        var rows = sheet.getDataRange().getValues();
        
        for (var r = 1; r < rows.length; r++) {
          var eventVal = String(rows[r][8] || "").trim();
          var rowStatus = String(rows[r][0] || "").trim();
          if (rowStatus !== "취소" && eventVal.indexOf("당첨") !== -1) {
            todayWinnerCount++;
          }
        }

        if (todayWinnerCount < settings.maxEventWinners) {
          var rand = Math.random() * 100;
          if (rand < settings.eventWinProbability) {
            isWinner = true;
          }
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
        isWinner ? "당첨 (선택대기)" : ""
      ]);

      var syncedClientId = registerOrUpdateIdRegistry(customerInfo, clientIdVal);

      var lastRowIdx = sheet.getLastRow();
      if (totalPriceVal) {
        sheet.getRange(lastRowIdx, 8).setNumberFormat('#,##0"원"');
      }

      sortOrders(sheet);

      try {
        sendEmailNotification({
          type: "create",
          nickname: customerInfo,
          pickupTime: data.pickupTime || "",
          cookType: rawCook,
          menuSummary: menuText,
          sauce: sauceText,
          totalPrice: totalPriceVal ? Number(totalPriceVal).toLocaleString() + "원" : "0원"
        });
      } catch (err) {
        Logger.log("메일 함수 호출 에러: " + err.toString());
      }

      return ContentService.createTextOutput(JSON.stringify({
        "result": "success",
        "isWinner": isWinner,
        "rowIndex": lastRowIdx,
        "eventMenuList": settings.eventMenuList,
        "syncedClientId": syncedClientId
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "checkNickname") {
      var checkNick = String(data.nickname || "").trim();
      var checkClient = String(data.clientId || "").trim();
      var isDup = checkNicknameExistsInRegistry(checkNick, checkClient);
      return ContentService.createTextOutput(JSON.stringify({"result": isDup ? "duplicate" : "ok"}))
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
          var menuRows = menuSheet.getDataRange().getValues();
          for (var i = 1; i < menuRows.length; i++) {
            var menuName = String(menuRows[i][0] || "").trim();
            if (menuName === targetName) {
              var existingUrl = String(menuRows[i][3] || "").trim();
              if (existingUrl !== "") {
                finalUrlList = existingUrl + "," + imageUrl;
              }
              menuSheet.getRange(i + 1, 4).setValue(finalUrlList);
              break;
            }
          }
        }

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
      var prize = String(data.prize || "").trim();
      var rows = sheet.getDataRange().getValues();

      for (var i = rows.length - 1; i >= 1; i--) {
        var status = String(rows[i][0] || "").trim();
        if (status === "취소") continue;

        var custCol = String(rows[i][5] || "").trim();
        if (custCol === nickname) {
          sheet.getRange(i + 1, 9).setValue("당첨 (" + prize + ")");
          return ContentService.createTextOutput(JSON.stringify({"result": "success"}))
                                     .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({"result": "not_found"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "verify") {
      var nickname = String(data.nickname || "").trim();
      var rows = sheet.getDataRange().getValues();
      
      for (var i = rows.length - 1; i >= 1; i--) {
        var status = String(rows[i][0] || "").trim();
        if (status === "취소") continue;

        var custCol = String(rows[i][5] || "").trim();
        if (custCol === nickname) {
          return ContentService.createTextOutput(JSON.stringify({"result": "exists"}))
                                     .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({"result": "not_exists"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "check" || action === "checkSimple") {
      var nickname = String(data.nickname || "").trim();
      var clientIdCheck = String(data.clientId || "").trim();
      var recoveredNickname = "";

      if (!nickname && clientIdCheck) {
        recoveredNickname = findNicknameByClientId(clientIdCheck);
        if (recoveredNickname) {
          nickname = recoveredNickname;
        }
      }

      var rows = sheet.getDataRange().getValues();
      
      for (var i = rows.length - 1; i >= 1; i--) {
        var status = String(rows[i][0] || "").trim();
        if (status === "취소") continue;

        var custCol = String(rows[i][5] || "").trim();
        if (nickname && custCol === nickname) {
          var cellValue = sheet.getRange(i + 1, 2).getDisplayValue();
          return ContentService.createTextOutput(JSON.stringify({
            "result": "found",
            "rowIndex": i + 1,
            "status": status,
            "nickname": custCol,
            "recoveredNickname": recoveredNickname || custCol,
            "pickupTime": cellValue, 
            "cookType": String(rows[i][2] || "").trim(),
            "menu": rows[i][3] || "",
            "sauce": rows[i][4] || "없음",
            "totalPrice": rows[i][7] || "",
            "eventPrize": rows[i][8] || ""
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({
        "result": "not_found",
        "recoveredNickname": recoveredNickname
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "update") {
      var settingsCheck = getStoreSettings(ss);
      var currentTimeVal = getKoreanCurrentTimeVal();
      var openTimeVal = settingsCheck.openHour * 60 + settingsCheck.openMin;
      var closeTimeVal = settingsCheck.closeHour * 60 + settingsCheck.closeMin;

      var isClosedNow = false;
      if (closeTimeVal > 1440 && currentTimeVal < openTimeVal) {
          if ((currentTimeVal + 1440) > closeTimeVal) {
              isClosedNow = true;
          }
      } else {
          isClosedNow = (currentTimeVal < openTimeVal || currentTimeVal > closeTimeVal);
      }

      if (isClosedNow) {
        return ContentService.createTextOutput(JSON.stringify({"result": "closed"}))
                             .setMimeType(ContentService.MimeType.JSON);
      }

      var rowIndex = Number(data.rowIndex);
      if (!rowIndex || rowIndex < 2) {
        return ContentService.createTextOutput(JSON.stringify({"result": "error"}))
                             .setMimeType(ContentService.MimeType.JSON);
      }

      var currentStatus = String(sheet.getRange(rowIndex, 1).getValue()).trim();
      if (currentStatus === "확인") {
        return ContentService.createTextOutput(JSON.stringify({"result": "locked"}))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      
      var newMenuSummary = data.newMenu || "";
      var newSauceSummary = data.newSauce || "없음";

      var oldMenuStr = String(sheet.getRange(rowIndex, 4).getValue() || "");
      var oldSauceStr = String(sheet.getRange(rowIndex, 5).getValue() || "");
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
      
      sortOrders(sheet);

      try {
        var custColVal = String(sheet.getRange(rowIndex, 6).getValue()).trim();

        sendEmailNotification({
          type: "update",
          nickname: custColVal,
          pickupTime: data.newPickupTime || sheet.getRange(rowIndex, 2).getDisplayValue(),
          cookType: finalCookTypeStr,
          menuSummary: data.newMenu || sheet.getRange(rowIndex, 4).getValue(),
          sauce: data.newSauce || sheet.getRange(rowIndex, 5).getValue(),
          totalPrice: updatePriceNumeric ? updatePriceNumeric.toLocaleString() + "원" : "0원"
        });
      } catch (err) {
        Logger.log("변경 메일 함수 호출 에러: " + err.toString());
      }
      
      return ContentService.createTextOutput(JSON.stringify({"result": "updated"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({"result": "invalid"}))
                         .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({"result": "error", "message": err.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
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
  
  var data = sheet.getDataRange().getValues();
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
      eventPrize: eventVal
    });
  }

  orders.sort(function(a, b) {
    return parseTimeToMin(a.pickupTime) - parseTimeToMin(b.pickupTime);
  });

  return orders;
}

function parseTimeToMin(timeStr) {
  if (!timeStr) return 9999;
  var isPM = timeStr.indexOf('오후') !== -1;
  var isAM = timeStr.indexOf('오전') !== -1;
  var clean = timeStr.replace('오전', '').replace('오후', '').trim();
  var parts = clean.split(':');
  if (parts.length < 2) return 9999;

  var hours = parseInt(parts[0], 10);
  var minutes = parseInt(parts[1], 10);

  if (isPM && hours < 12) hours += 12;
  if (isAM && hours === 12) hours = 0;

  return hours * 60 + minutes;
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
    if (notiVal === "OFF") {
      notificationStatus = "OFF";
    }

    var delayVal = parseInt(settingSheet.getRange("B4").getValue());
    if (!isNaN(delayVal) && delayVal >= 0) {
      prepMinutes = delayVal;
    }

    var maxW = parseInt(settingSheet.getRange("B5").getValue());
    if (!isNaN(maxW)) maxEventWinners = maxW;

    var prob = parseFloat(settingSheet.getRange("B6").getValue());
    if (!isNaN(prob)) eventWinProbability = prob;

    var rawEventMenus = String(settingSheet.getRange("B7").getDisplayValue()).trim();
    if (rawEventMenus !== "") {
      eventMenuList = rawEventMenus.split(",").map(function(item) { return item.trim(); });
    }

    var maintVal = String(settingSheet.getRange("B9").getDisplayValue()).trim().toUpperCase();
    if (maintVal === "ON") {
      isMaintenance = true;
    }

    var codeCellVal = String(settingSheet.getRange("F5").getDisplayValue()).trim();
    if (codeCellVal !== "") {
      secretCode = codeCellVal;
    }
  }

  var menuSheet = ss.getSheetByName("메뉴판");
  var menuList = [];
  var sauceList = [];

  if (menuSheet) {
    var menuRows = menuSheet.getDataRange().getValues();
    var isSauceSection = false;

    for (let i = 1; i < menuRows.length; i++) {
      let colA = String(menuRows[i][0] || "").trim();
      let colB = String(menuRows[i][1] || "").trim();
      let colC = String(menuRows[i][2] || "").trim();
      let colD = String(menuRows[i][3] || "").trim();
      let colE = String(menuRows[i][4] || "").trim();

      if (colA.includes("추가소스") || colA.includes("소스분류") || colA.includes("추가소스 분류") || colB.includes("추가소스")) {
        isSauceSection = true;
        continue;
      }

      if (colA === "") continue;

      var isSoldOut = (colE === "품절");

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
          isSoldOut: isSoldOut
        });
      }
    }
  }

  if (sauceList.length === 0) {
    sauceList.push({ name: "돈까스 소스", shortName: "돈소스", price: "₩500", imageUrl: "", isSoldOut: false });
  }

  var maintCode = "3808";
  if (settingSheet) {
    var codeVal = String(settingSheet.getRange("B10").getDisplayValue()).trim();
    if (codeVal) maintCode = codeVal;
  }

  return {
    openHour: openHour,
    openMin: openMin,
    closeHour: closeHour,
    closeMin: closeMin,
    pickupStartHour: pickupStartHour,
    pickupStartMin: pickupStartMin,
    pickupCloseHour: pickupCloseHour,
    pickupCloseMin: pickupCloseMin,
    notification: notificationStatus,
    prepMinutes: prepMinutes, 
    menuList: menuList,
    sauceList: sauceList,
    maxEventWinners: maxEventWinners,
    eventWinProbability: eventWinProbability,
    eventMenuList: eventMenuList,
    isMaintenance: isMaintenance,
    maintenanceCode: maintCode,
    secretCode: secretCode
  };
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
      h = parseInt(parts[0]) || 0;
      m = parseInt(parts[1]) || 0;
    }
  }
  return { hour: h, min: m };
}

function convertTimeToMinutes(timeStr) {
  try {
    var isPM = timeStr.includes("오후");
    var cleanStr = timeStr.replace("오전", "").replace("오후", "").trim();
    var parts = cleanStr.split(":");
    var h = parseInt(parts[0]) || 0;
    var m = parseInt(parts[1]) || 0;

    if (isPM && h < 12) h += 12;
    if (!isPM && h === 12) h = 0;

    return h * 60 + m;
  } catch (e) {
    return null;
  }
}

function backupAndClearDailyOrders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getSheetByName("시트1");
  var dbSheet = ss.getSheetByName("누적DB");
  
  if (!activeSheet || !dbSheet) {
    SpreadsheetApp.getUi().alert("🚨 '시트1' 또는 '누적DB' 시트를 찾을 수 없습니다!");
    return;
  }
  
  var lastRow = activeSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("💡 백업할 예약 내역이 없습니다!");
    return;
  }
  
  var rangeToCopy = activeSheet.getRange(2, 1, lastRow - 1, 9);
  var values = rangeToCopy.getValues();
  
  var nextDbRow = dbSheet.getLastRow() + 1;
  dbSheet.getRange(nextDbRow, 1, values.length, values[0].length).setValues(values);
  
  activeSheet.deleteRows(2, lastRow - 1);
  
  SpreadsheetApp.getUi().alert("🎉 픽업 시간과 주문 내역이 누적DB에 완벽하게 백업되었습니다!");
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