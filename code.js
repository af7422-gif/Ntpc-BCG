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

function getUserAccessProfile(searchName, searchBirthday) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profileSheet = ss.getSheetByName('權限管理') || ss.getSheetByName('UserAccess') || ss.getSheetByName('使用者權限');
  const staffSheet = ss.getSheetByName('新北市在職清冊');

  if (!staffSheet) {
    return { role: 'user', unit: '', canViewAll: false, error: 'SHEET_NOT_FOUND' };
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

//查詢當年度參訓成績
function searchAnnualResults(searchName, searchBirthday) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("新北市在職清冊"); 
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  // 1. 取得標題、內容與「字體顏色」
  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const dataRange = sheet.getRange(3, 1, lastRow - 2, lastCol);
  const dataValues = dataRange.getValues();
  const dataColors = dataRange.getFontColors(); // 新增這行來抓取顏色

  const nameIdx = headers.indexOf("姓名");
  const bdayIdx = headers.indexOf("生日");
  
  if (nameIdx === -1 || bdayIdx === -1) return { error: "找不到姓名或生日欄位" };

  // 2. 尋找符合的資料行索引
  let rowIndex = -1;
  for (let i = 0; i < dataValues.length; i++) {
    const rowName = dataValues[i][nameIdx].toString().trim();
    const rowBday = Utilities.formatDate(new Date(dataValues[i][bdayIdx]), "GMT+8", "yyyy-MM-dd");
    
    if (rowName === searchName && rowBday === searchBirthday) {
      rowIndex = i;
      break;
    }
  }

  // 找不到人員回傳 null
  if (rowIndex === -1) return null;

  const rowData = dataValues[rowIndex];
  const rowColors = dataColors[rowIndex]; // 取得該行的所有顏色
  
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
      score: getValueAndColor("種子師資-成績"),
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

function searchData(searchName, searchBirthday) {
  const access = getUserAccessProfile(searchName, searchBirthday);
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

  function getUnitStaffList(searchName, searchBirthday) {
    const access = getUserAccessProfile(searchName, searchBirthday);
    if (access.error) return { error: access.error };

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName("新北市在職清冊");
    const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(3, 1, sheet.getLastRow() - 2, headers.length).getValues();

    const nameIdx = headers.indexOf("姓名");
    const bdayIdx = headers.indexOf("生日");
    const unitIdx = headers.indexOf("服務單位");
    const activeIdx = headers.indexOf("目前是否執行業務");

    const matchedRows = data.filter((row, i) => {
      const rowName = normalizeText(row[nameIdx]);
      const rowBday = normalizeDateValue(row[bdayIdx]);
      const match = rowName === normalizeText(searchName) && rowBday === normalizeDateValue(searchBirthday);
      return match;
    });

    if (!matchedRows.length) return { error: "NOT_FOUND" };

    const row = matchedRows[0];
    const unit = String(row[unitIdx] || '').trim();
    const isActive = String(row[activeIdx] || '').trim();

    if (isActive === "否") {
      return { error: "NOT_ACTIVE" };
    }

    let result = [];
    if (access.canViewAll) {
      result = data.map(r => {
        let obj = {};
        headers.forEach((h, idx) => {
          let v = r[idx];
          if (v instanceof Date) {
            v = Utilities.formatDate(v, "GMT+8", "yyyy-MM-dd");
          }
          obj[h] = v;
        });
        return obj;
      });
    } else {
      result = [matchedRows[0]].map(r => {
        let obj = {};
        headers.forEach((h, idx) => {
          let v = r[idx];
          if (v instanceof Date) {
            v = Utilities.formatDate(v, "GMT+8", "yyyy-MM-dd");
          }
          obj[h] = v;
        });
        return obj;
      });
    }

    return { unit: access.canViewAll ? '全部單位' : unit, list: result, role: access.role, canViewAll: access.canViewAll };
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
