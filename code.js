function triggerAuth() {
  MailApp.sendEmail("test@example.com", "test", "test");
}
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('新北市卡介苗工作人員查詢系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function normalizeText(value) {
  return String(value || '').replace(/\s+/g, '').trim();
}

function normalizeDateValue(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'GMT+8', 'yyyy-MM-dd');
  }
  const date = new Date(value);
  if (!isNaN(date.getTime())) {
    return Utilities.formatDate(date, 'GMT+8', 'yyyy-MM-dd');
  }
  return String(value).trim();
}

function maskIdentityNumber(value) {
  const identityNumber = String(value || '').trim();
  if (!identityNumber) return '';
  if (identityNumber.length < 5) return '*'.repeat(identityNumber.length);
  return `${identityNumber.slice(0, 3)}*****${identityNumber.slice(-2)}`;
}

function normalizeUnit(value) {
  return String(value || '').replace(/\s+/g, '').trim();
}

function isAdminRole(roleValue, unitValue) {
  const role = String(roleValue || '').trim().toLowerCase();
  const unit = normalizeUnit(unitValue);
  const isHealthDepartmentAdmin = unit === '新北市政府衛生局';
  if (isHealthDepartmentAdmin) return true;
  if (!role) return false;
  return ['admin', '管理者', '主管', '系統管理者', '系統管理', 'administrator'].includes(role) || role.includes('admin');
}

function isValidTesterSession(testToken) {
  if (!testToken) return false;
  return CacheService.getScriptCache().get(`tester:${testToken}`) === '1';
}

function revokeTesterSession(testToken) {
  if (testToken) CacheService.getScriptCache().remove(`tester:${testToken}`);
}

function getUserAccessProfile(searchName, searchBirthday, testToken) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profileSheet = ss.getSheetByName('權限管理') || ss.getSheetByName('UserAccess') || ss.getSheetByName('使用者權限');
  const staffSheet = ss.getSheetByName('新北市在職清冊');

  if (!staffSheet) {
    return { role: 'user', unit: '', canViewAll: false, error: 'SHEET_NOT_FOUND' };
  }
  if (isValidTesterSession(testToken)) {
    return { role: 'admin', unit: '全部單位', canViewAll: true, isTester: true };
  }

  const nameKey = normalizeText(searchName);
  const birthdayKey = normalizeDateValue(searchBirthday);

  if (profileSheet) {
    const lastRow = profileSheet.getLastRow();
    if (lastRow >= 2) {
      const headers = profileSheet.getRange(1, 1, 1, profileSheet.getLastColumn()).getValues()[0];
      const data = profileSheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

      const nameIdx = headers.indexOf('姓名');
      const bdayIdx = headers.indexOf('生日');
      const roleIdx = headers.indexOf('角色');
      const unitIdx = headers.indexOf('服務單位');
      const activeIdx = headers.indexOf('啟用狀態');

      for (let i = 0; i < data.length; i++) {
        const rowName = normalizeText(data[i][nameIdx]);
        const rowBirthday = normalizeDateValue(data[i][bdayIdx]);
        const rowActive = String(data[i][activeIdx] || '').trim();

        if (rowName === nameKey && rowBirthday === birthdayKey) {
          const role = String(data[i][roleIdx] || '').trim();
          const unit = String(data[i][unitIdx] || '').trim();
          const isActive = activeIdx === -1 || rowActive === '' || rowActive === '啟用' || rowActive === 'Y' || rowActive === '是';

          if (!isActive) {
            return { role: 'user', unit: unit || '', canViewAll: false, error: 'ACCOUNT_DISABLED' };
          }

          const isAdmin = isAdminRole(role, unit);
          return {
            role: isAdmin ? 'admin' : 'user',
            unit: unit || '',
            canViewAll: isAdmin,
            name: searchName,
            birthday: searchBirthday
          };
        }
      }
    }
  }

  const headers = staffSheet.getRange(2, 1, 1, staffSheet.getLastColumn()).getValues()[0];
  const data = staffSheet.getRange(3, 1, staffSheet.getLastRow() - 2, headers.length).getValues();
  const nameIdx = headers.indexOf('姓名');
  const bdayIdx = headers.indexOf('生日');
  const unitIdx = headers.indexOf('服務單位');
  const roleIdx = headers.indexOf('角色');

  for (let i = 0; i < data.length; i++) {
    const rowName = normalizeText(data[i][nameIdx]);
    const rowBirthday = normalizeDateValue(data[i][bdayIdx]);
    if (rowName === nameKey && rowBirthday === birthdayKey) {
      const role = String(data[i][roleIdx] || '').trim();
      const unit = String(data[i][unitIdx] || '').trim();
      const isAdmin = isAdminRole(role, unit);
      return {
        role: isAdmin ? 'admin' : 'user',
        unit: unit || '',
        canViewAll: isAdmin,
        name: searchName,
        birthday: searchBirthday
      };
    }
  }

  return { role: 'user', unit: '', canViewAll: false, error: 'NOT_FOUND' };
}

function getTesterLoginProfile(password) {
  if (String(password || '') !== '1971') {
    return { error: 'INVALID_TESTER_PASSWORD' };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('新北市在職清冊');
  if (!sheet || sheet.getLastRow() < 3) return { error: 'SHEET_NOT_FOUND' };

  const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getRange(3, 1, sheet.getLastRow() - 2, headers.length).getValues();
  const nameIdx = headers.indexOf('姓名');
  const birthdayIdx = headers.indexOf('生日');
  const historyIdx = headers.indexOf('歷年指導紀錄');
  const activeIdx = headers.indexOf('目前是否執行業務');
  if ([nameIdx, birthdayIdx, historyIdx].includes(-1)) return { error: 'SHEET_NOT_FOUND' };

  const candidates = data.filter(row =>
    normalizeText(row[nameIdx]) && normalizeDateValue(row[birthdayIdx]) &&
    String(row[historyIdx] || '').trim() &&
    (activeIdx === -1 || String(row[activeIdx] || '').trim() !== '否')
  );
  if (!candidates.length) return { error: 'NO_TEST_STAFF' };

  const row = candidates[Math.floor(Math.random() * candidates.length)];
  const testToken = Utilities.getUuid();
  CacheService.getScriptCache().put(`tester:${testToken}`, '1', 21600);
  return {
    role: 'admin',
    unit: '全部單位',
    canViewAll: true,
    isTester: true,
    testToken,
    name: String(row[nameIdx]).trim(),
    birthday: normalizeDateValue(row[birthdayIdx])
  };
}

//查詢當年度參訓成績
function searchAnnualResults(searchName, searchBirthday) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("新北市在職清冊"); 
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const nameIdx = headers.indexOf("姓名");
  const bdayIdx = headers.indexOf("生日");
  
  if (nameIdx === -1 || bdayIdx === -1) return { error: "找不到姓名或生日欄位" };
  const rowCount = lastRow - 2;
  if (rowCount <= 0) return null;

  const names = sheet.getRange(3, nameIdx + 1, rowCount, 1).getValues();
  const birthdays = sheet.getRange(3, bdayIdx + 1, rowCount, 1).getValues();

  // 2. 尋找符合的資料行索引
  let matchedRow = -1;
  for (let i = 0; i < rowCount; i++) {
    if (normalizeText(names[i][0]) === normalizeText(searchName) &&
        normalizeDateValue(birthdays[i][0]) === normalizeDateValue(searchBirthday)) {
      matchedRow = i + 3;
      break;
    }
  }

  // 找不到人員回傳 null
  if (matchedRow === -1) return null;

  const rowRange = sheet.getRange(matchedRow, 1, 1, lastCol);
  const rowData = rowRange.getValues()[0];
  const rowColors = rowRange.getFontColors()[0];
  
  // 3. 修改輔助函式：現在同時回傳「值」與「顏色」
  const getValueAndColor = (colName) => {
    const idx = headers.indexOf(colName);
    if (idx !== -1 && rowData[idx] !== "") {
      return { 
        val: rowData[idx], 
        color: rowColors[idx] // 將對應的顏色一併存入
      };
    }
    return { val: null, color: "" };
  };

  // 將資料分好類別回傳給前端
  return {
    master: {
      read: getValueAndColor("師資統合-判讀成績"),
      inject: getValueAndColor("師資統合-皮內注射技術成績"),
      pass: getValueAndColor("師資統合-是否合格")
    },
    seed: {
      written: getValueAndColor("種子師資-筆試"),
      teaching: getValueAndColor("種子師資-教學評量"),
      total: getValueAndColor("種子師資-總成績"),
      pass: getValueAndColor("種子師資-是否合格")
    },
    skill: {
      read: getValueAndColor("技術評價-判讀成績"),
      readMakeup: getValueAndColor("技術評價-判讀成績（補考）"),
      inject: getValueAndColor("技術評價-皮內注射技術成績"),
      injectMakeup: getValueAndColor("技術評價-皮內注射技術成績（補考）"),
      pass: getValueAndColor("技術評價-是否合格")
    }
  };
}

function searchData(searchName, searchBirthday, testToken) {
  const access = getUserAccessProfile(searchName, searchBirthday, testToken);
  if (access.error) return { error: access.error };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("新北市在職清冊");

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const dataRange = sheet.getRange(3, 1, lastRow - 2, lastCol);
  const dataValues = dataRange.getValues();
  const dataColors = dataRange.getFontColors();

  const displayOrder = [
    "師資統合獲證年度",
    "種子師資獲證年度",
    "最近一次帶初訓日",
    "師資年限",
    "技術評價獲證年度",
    "初訓獲證年度",
    "施打年限",
    "檢核",
    "建議參訓項目"
  ];

  const nameIdx = headers.indexOf("姓名");
  const bdayIdx = headers.indexOf("生日");

  if (nameIdx === -1 || bdayIdx === -1) return { error: "找不到姓名或生日欄位" };

  let rowIndex = -1;
  for (let i = 0; i < dataValues.length; i++) {
    const rowName = normalizeText(dataValues[i][nameIdx]);
    const rowBday = normalizeDateValue(dataValues[i][bdayIdx]);

    if (rowName === normalizeText(searchName) && rowBday === normalizeDateValue(searchBirthday)) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex === -1) return null;

  let output = [];
  displayOrder.forEach(field => {
    const colIdx = headers.indexOf(field);
    if (colIdx !== -1) {
      let value = dataValues[rowIndex][colIdx];
      if (value instanceof Date) {
        value = Utilities.formatDate(value, "GMT+8", "yyyy-MM-dd");
      }
      output.push({
        label: field,
        value: value || "",
        color: dataColors[rowIndex][colIdx]
      });
    }
  });

  return output;
}

  function getUnitStaffList(searchName, searchBirthday, testToken) {
    const access = getUserAccessProfile(searchName, searchBirthday, testToken);
    if (access.error) return { error: access.error };

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName("新北市在職清冊");
    const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(3, 1, sheet.getLastRow() - 2, headers.length).getValues();

    const nameIdx = headers.indexOf("姓名");
    const bdayIdx = headers.indexOf("生日");
    const unitIdx = headers.indexOf("服務單位");
    const activeIdx = headers.indexOf("目前是否執行業務");

    const ownRowIndex = data.findIndex(row => {
      const rowName = normalizeText(row[nameIdx]);
      const rowBday = normalizeDateValue(row[bdayIdx]);
      return rowName === normalizeText(searchName) && rowBday === normalizeDateValue(searchBirthday);
    });

    if (ownRowIndex === -1) return { error: "NOT_FOUND" };

    const row = data[ownRowIndex];
    const unit = String(row[unitIdx] || '').trim();
    const isActive = String(row[activeIdx] || '').trim();

    const identityIdx = headers.findIndex(header => ["身分證號", "身分證字號"].includes(header));
    const visibleRows = access.canViewAll
      ? data
      : isActive === "否"
        ? [row]
        : data.filter((staffRow, index) =>
            index === ownRowIndex ||
            (unit && normalizeUnit(staffRow[unitIdx]) === normalizeUnit(unit))
          );
    const result = visibleRows.map(staffRow => {
      const obj = {};
      headers.forEach((header, index) => {
        let value = staffRow[index];
        if (value instanceof Date) {
          value = Utilities.formatDate(value, "GMT+8", "yyyy-MM-dd");
        }
        if (!access.canViewAll && index === identityIdx) {
          value = staffRow === row ? maskIdentityNumber(value) : "";
        }
        obj[header] = value;
      });
      return obj;
    });

    return { unit: access.canViewAll ? '全部單位' : unit, list: result, role: access.role, canViewAll: access.canViewAll };
  }

  function updateStaffFieldAsAdmin(editorName, editorBirthday, targetUnit, targetName, targetBirthday, fieldName, newValue, testToken) {
    return updateStaffFieldInternal(editorName, editorBirthday, targetUnit, targetName, targetBirthday, fieldName, newValue, false, testToken);
  }

  function updateStaffField(editorName, editorBirthday, targetUnit, targetName, targetBirthday, fieldName, newValue, testToken) {
    return updateStaffFieldInternal(editorName, editorBirthday, targetUnit, targetName, targetBirthday, fieldName, newValue, true, testToken);
  }

  function updateStaffFieldInternal(editorName, editorBirthday, targetUnit, targetName, targetBirthday, fieldName, newValue, allowSelfEdit, testToken) {
    const access = getUserAccessProfile(editorName, editorBirthday, testToken);
    const selfEditableFields = new Set(["姓名", "生日", "聯絡電話", "E-Mail"]);
    const isOwnRecord = normalizeText(editorName) === normalizeText(targetName) &&
      normalizeDateValue(editorBirthday) === normalizeDateValue(targetBirthday) &&
      normalizeUnit(access.unit) === normalizeUnit(targetUnit);
    if (access.error || (!access.canViewAll &&
        (!allowSelfEdit || !isOwnRecord || !selfEditableFields.has(fieldName)))) {
      throw new Error("沒有權限直接修改此筆資料。");
    }

    const fieldTypes = {
      "姓名": "text",
      "生日": "date",
      "聯絡電話": "text",
      "E-Mail": "text",
      "師資統合獲證年度": "number",
      "種子師資獲證年度": "number",
      "最近一次帶初訓日": "date",
      "技術評價獲證年度": "number",
      "初訓獲證年度": "number",
      "歷年指導紀錄": "json",
      "目前是否執行業務": "select",
      "未執行業務原因": "text"
    };
    const fieldType = fieldTypes[fieldName];
    if (!fieldType) throw new Error("此欄位不開放直接修改。");

    let value = String(newValue == null ? "" : newValue).trim();
    if (!value && fieldType !== "json") throw new Error("修改內容不可空白。");
    if (fieldType === "json") {
      let history;
      try {
        history = JSON.parse(value);
      } catch (error) {
        throw new Error("歷年指導紀錄不是有效的 JSON 格式。");
      }
      if (!history || typeof history !== "object" || Array.isArray(history) ||
          !Array.isArray(history.assistantRecords) || !Array.isArray(history.traineeRecords)) {
        throw new Error("歷年指導紀錄格式不正確，請確認 assistantRecords 與 traineeRecords 均為陣列。");
      }
      if ([...history.assistantRecords, ...history.traineeRecords].some(record =>
        !record || typeof record !== "object" || Array.isArray(record)
      )) {
        throw new Error("歷年指導紀錄包含格式不正確的項目。");
      }
      value = JSON.stringify(history);
    } else if (fieldType === "number") {
      value = Number(value);
      if (!Number.isFinite(value)) throw new Error("請輸入有效的數字。");
    } else if (fieldType === "date") {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) throw new Error("日期格式不正確。");
      const date = new Date(`${value}T00:00:00+08:00`);
      if (Utilities.formatDate(date, "GMT+8", "yyyy-MM-dd") !== value) {
        throw new Error("日期格式不正確。");
      }
      value = date;
    } else if (fieldType === "select" && !["是", "否"].includes(value)) {
      throw new Error("業務狀態僅可設定為「是」或「否」。");
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("新北市在職清冊");
    if (!sheet) throw new Error("找不到工作表。");
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 3) throw new Error("找不到工作人員資料。");

    const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    const nameIdx = headers.indexOf("姓名");
    const birthdayIdx = headers.indexOf("生日");
    const unitIdx = headers.indexOf("服務單位");
    const fieldIdx = headers.indexOf(fieldName);
    if ([nameIdx, birthdayIdx, unitIdx, fieldIdx].includes(-1)) {
      throw new Error("找不到必要欄位。");
    }

    const data = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
    const matchedRows = [];
    data.forEach((row, index) => {
      if (normalizeText(row[nameIdx]) === normalizeText(targetName) &&
          normalizeDateValue(row[birthdayIdx]) === normalizeDateValue(targetBirthday) &&
          normalizeUnit(row[unitIdx]) === normalizeUnit(targetUnit)) {
        matchedRows.push(index + 3);
      }
    });
    if (matchedRows.length !== 1) {
      throw new Error(matchedRows.length ? "符合條件的資料不唯一，未進行修改。" : "找不到對應人員資料。");
    }

    const profileSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("權限管理") ||
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UserAccess") ||
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("使用者權限");
    let profileRow = -1;
    let profileNameIdx = -1;
    let profileBirthdayIdx = -1;
    if (profileSheet && ["姓名", "生日"].includes(fieldName)) {
      const profileLastRow = profileSheet.getLastRow();
      const profileLastCol = profileSheet.getLastColumn();
      if (profileLastRow >= 2 && profileLastCol > 0) {
        const profileHeaders = profileSheet.getRange(1, 1, 1, profileLastCol).getValues()[0];
        profileNameIdx = profileHeaders.indexOf("姓名");
        profileBirthdayIdx = profileHeaders.indexOf("生日");
        if (profileNameIdx !== -1 && profileBirthdayIdx !== -1) {
          const profileData = profileSheet.getRange(2, 1, profileLastRow - 1, profileLastCol).getValues();
          const matchedProfiles = [];
          profileData.forEach((row, index) => {
            if (normalizeText(row[profileNameIdx]) === normalizeText(targetName) &&
                normalizeDateValue(row[profileBirthdayIdx]) === normalizeDateValue(targetBirthday)) {
              matchedProfiles.push(index + 2);
            }
          });
          if (matchedProfiles.length > 1) throw new Error("權限資料符合條件的帳號不唯一，未進行修改。");
          if (matchedProfiles.length === 1) profileRow = matchedProfiles[0];
        }
      }
    }

    sheet.getRange(matchedRows[0], fieldIdx + 1).setValue(value);
    if (profileRow !== -1) {
      const profileFieldIdx = fieldName === "姓名" ? profileNameIdx : profileBirthdayIdx;
      profileSheet.getRange(profileRow, profileFieldIdx + 1).setValue(value);
    }
    const updatedName = fieldName === "姓名" ? String(value) : String(targetName);
    const updatedBirthday = fieldName === "生日"
      ? Utilities.formatDate(value, "GMT+8", "yyyy-MM-dd")
      : normalizeDateValue(targetBirthday);
    return {
      updated: true,
      currentUserName: isOwnRecord ? updatedName : "",
      currentUserBirthday: isOwnRecord ? updatedBirthday : ""
    };
  }

  function addEditCommentByName(unit, targetName, fieldName, newValue, requesterName) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName("新北市在職清冊");
      if (!sheet) throw new Error("找不到工作表");

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow < 3) return false;

      const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
      const nameIdx = headers.indexOf("姓名");
      const unitIdx = headers.indexOf("服務單位");
      const fieldIdx = headers.indexOf(fieldName);

      if (nameIdx === -1 || unitIdx === -1 || fieldIdx === -1) {
        throw new Error("欄位名稱不存在");
      }

      const data = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

      for (let i = 0; i < data.length; i++) {
        const rowUnit = String(data[i][unitIdx]).trim();
        const rowName = String(data[i][nameIdx]).trim();

        if (rowUnit === String(unit).trim() &&
            rowName === String(targetName).trim()) {

          const rowNum = i + 3;
          const colNum = fieldIdx + 1;
          const fileId = ss.getId();
          const sheetId = sheet.getSheetId();

          // 1. 準備註解內容
          const content =
            `${unit}-${requesterName} 申請修改\n` +
            `對象：${targetName}\n` +
            `欄位：「${fieldName}」\n` +
            `建議值：${newValue}`;

          const anchorObj = {
            type: "spreadsheets#region",
            data: {
              range: {
                sheetId: sheetId,
                startRowIndex: rowNum - 1,
                endRowIndex: rowNum,
                startColumnIndex: colNum - 1,
                endColumnIndex: colNum
              }
            }
          };

          // 2. 建立 Drive 註解
          const resp = Drive.Comments.create(
            {
              content: content,
              anchor: JSON.stringify(anchorObj)
            },
            fileId,
            { fields: "id,anchor,content" }
          );

          // 3. 新增：發送 Email 通知管理者
          // 建立一個簡單的 HTML 表格讓郵件好閱讀
          const emailHtml = `
            <div style="font-family: sans-serif; border: 1px solid #ccc; padding: 15px; border-radius: 10px;">
              <h3 style="color: #d9534f;">資料修改申請通知</h3>
              <p>來自 <b>${unit}</b> 的 <b>${requesterName}</b> 提出了一項修改申請：</p>
              <table style="border-collapse: collapse; width: 100%;">
                <tr><td style="padding: 8px; border-bottom: 1px solid #eee;"><b>修改對象</b></td><td style="padding: 8px; border-bottom: 1px solid #eee;">${targetName}</td></tr>
                <tr><td style="padding: 8px; border-bottom: 1px solid #eee;"><b>修改欄位</b></td><td style="padding: 8px; border-bottom: 1px solid #eee;">${fieldName}</td></tr>
                <tr><td style="padding: 8px; border-bottom: 1px solid #eee;"><b>建議新值</b></td><td style="padding: 8px; border-bottom: 1px solid #eee; color: blue;">${newValue}</td></tr>
              </table>
            </div>
          `;

          MailApp.sendEmail({
            to: "AF7422@ntpc.gov.tw",
            subject: `【BCG在職清冊內容修改申請】`,
            htmlBody: emailHtml
          });

          console.log(`註解建立並郵件發送成功：ID=${resp.id}`);
          return true;
        }
      }
      return false;
    } catch (e) {
      console.error("addEditCommentByName 錯誤：", e);
      throw new Error("申請失敗：" + e.message);
    }
  }

  function getStaffDetail(searchName, searchBirthday) {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName("新北市在職清冊");
    const lastCol = sheet.getLastColumn();

    const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    const data = sheet.getRange(3, 1, sheet.getLastRow() - 2, lastCol).getValues();

    const nameIdx = headers.indexOf("姓名");
    const bdayIdx = headers.indexOf("生日");

    for (let i = 0; i < data.length; i++) {
      const name = data[i][nameIdx];
      const bday = Utilities.formatDate(
        new Date(data[i][bdayIdx]),
        "GMT+8",
        "yyyy-MM-dd"
      );

      if (name === searchName && bday === searchBirthday) {
        let result = {};
        headers.forEach((h, idx) => {
          let v = data[i][idx];
          if (v instanceof Date) {
            v = Utilities.formatDate(v, "GMT+8", "yyyy-MM-dd");
          }
          result[h] = v;
        });
        return result;
      }
    }
    return null;
  }


  function sendNewStaffRequest(formData) {
    try {
      // 1. 定義固定的欄位順序
      const fieldOrder = [
        "服務單位",
        "姓名",
        "生日",
        "師資統合獲證年度",
        "種子師資獲證年度",
        "最近一次帶初訓日",
        "技術評價獲證年度",
        "初訓獲證年度",
        "目前是否執行業務",
        "未執行業務原因"
      ];

      // 2. 依照排序模板建立 HTML 表格內容
      // 如果 formData 裡面缺少某個欄位，會顯示 "—"
      let rowsHtml = fieldOrder.map(key => {
        const value = formData[key] || "—";
        return `
          <tr>
            <td style="padding: 12px; border: 1px solid #dee2e6; background-color: #f8f9fa; width: 40%; color: #495057;">
              <b>${key}</b>
            </td>
            <td style="padding: 12px; border: 1px solid #dee2e6; color: #212529;">
              ${value}
            </td>
          </tr>`;
      }).join("");

      // 3. 組合美化過的 HTML 郵件內容
      const htmlBody = `
        <div style="font-family: 'Microsoft JhengHei', sans-serif; max-width: 600px; margin: auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
          <div style="background-color: #007bff; color: #ffffff; padding: 20px; text-align: center;">
            <h2 style="margin: 0; font-size: 24px;">📌 新增工作人員申請</h2>
          </div>
          <div style="padding: 20px;">
            <p style="color: #666; font-size: 14px;">此郵件由系統自動發送，申請明細如下：</p>
            <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">
              ${rowsHtml}
            </table>
            <div style="margin-top: 25px; padding-top: 15px; border-top: 1px dashed #ccc; font-size: 12px; color: #888;">
              發送時間：${new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })}
            </div>
          </div>
        </div>
      `;

      // 4. 使用 GmailApp 發送郵件 (比 MailApp 更穩定)
      MailApp.sendEmail({
        to: "AF7422@ntpc.gov.tw",
        subject: "【BCG在職清冊內容新增申請】新增工作人員申請 - " + (formData["姓名"] || ""),
        htmlBody: htmlBody // 依然可以使用美化後的表格
      });
      console.log("郵件發送成功！欄位已按指定順序排列。");
      return "SUCCESS";

    } catch (e) {
      console.error("發送失敗詳細資訊: " + e.toString());
      // 如果還是權限錯誤，這裡會拋出明確訊息
      throw new Error("發送失敗，請確保已在腳本編輯器中點擊「執行」並通過授權!");
    }
  }
