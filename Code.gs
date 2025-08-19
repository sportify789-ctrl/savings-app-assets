/**
 * @fileoverview Backend logic for the "ออมเก่งเหล่าประชา" web application.
 * Handles user authentication, data retrieval, and data manipulation with Google Sheets.
 */

// 🔴 สำคัญ: ให้แทนที่ 'YOUR_SHEET_ID' ด้วย ID ของ Google Sheet ของคุณ
const SHEET_ID = '1fQ1iKyN_55ysGcjgbdYk8rwk4H-3swUW7zw6S5gWtbg';

/**
 * Handles CORS preflight requests (OPTIONS method).
 * This is crucial for allowing the frontend on GitHub Pages to call the API.
 */
function doOptions(e) {
  return ContentService.createTextOutput()
    .setMimeType(ContentService.MimeType.JSON)
    .addHttpHeader('Access-Control-Allow-Origin', '*')
    .addHttpHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
    .addHttpHeader('Access-Control-Allow-Headers', 'Content-Type');
}

/**
 * Handles POST requests from the web application, acting as the main API endpoint.
 */
function doPost(e) {
  let response;
  try {
    const body = JSON.parse(e.postData.contents);
    const { action, payload } = body;

    switch (action) {
      case 'getUserData':
        response = getUserData(payload.id, payload.password);
        break;
      // ... เคสอื่นๆ ของคุณเหมือนเดิม ...
      case 'addMultipleTransactions':
        response = addMultipleTransactions(payload.adminId, payload.dateString, payload.type, payload.transactions);
        break;
      default:
        throw new Error(`Unknown action: ${action}`);
    }
    // [แก้ไข] เพิ่ม .addHttpHeader ที่นี่
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: response }))
      .setMimeType(ContentService.MimeType.JSON)
      .addHttpHeader('Access-Control-Allow-Origin', '*'); 

  } catch (error) {
    Logger.log(error);
    // [แก้ไข] เพิ่ม .addHttpHeader ที่นี่ด้วย
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON)
      .addHttpHeader('Access-control-Allow-Origin', '*');
  }
}

/**
 * Serves the main HTML file ONLY when running directly from Apps Script URL (for testing).
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  template.url = ScriptApp.getService().getUrl();
  return template.evaluate()
    .setTitle('โครงการออมเก่งเหล่าประชา')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}


// 🔽🔽🔽 ส่วนที่เหลือของโค้ดเหมือนเดิม ไม่มีการเปลี่ยนแปลง 🔽🔽🔽

/**
 * Authenticates a user based on their ID and password.
 */
function getUserData(id, password) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) throw new Error('ไม่พบชีตข้อมูล "Users"');
    
    const data = usersSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const userId = row[0].toString();
      const userPassword = row[1].toString();
      const userStatus = row[5].toString();

      if (userId.trim() === id.trim()) {
        // Authenticate user first
        const isNewUser = (userStatus === 'new_user' || userStatus === '') && userPassword === '' && password === '';
        const isActiveUser = userPassword !== '' && userPassword === password;

        if (isNewUser || isActiveUser) {
          // If authenticated, update their level before fetching final data
          // This ensures the level is always current on login.
          if (row[3] !== 'admin') { // Don't run for admins
             updateUserLevel(userId.trim(), ss);
          }
          
          // Re-fetch the specific user's row to get the updated level
          const updatedRow = usersSheet.getRange(i + 1, 1, 1, usersSheet.getLastColumn()).getValues()[0];

          const userData = {
            id: updatedRow[0],
            name: updatedRow[2],
            role: updatedRow[3],
            balance: updatedRow[4],
            status: updatedRow[5],
            level: updatedRow[6] || 'ยังไม่มีระดับ', // Column G
            rewards: updatedRow[7] || ''      // Column H
          };

          if (isNewUser) {
            userData.status = 'new_user';
          }
          
          return userData;
        }
      }
    }
    throw new Error("รหัสนักเรียนหรือรหัสผ่านไม่ถูกต้อง");
  } catch (e) {
    Logger.log(`Login Error for ID [${id}]: ${e.message}`);
    throw new Error("เกิดข้อผิดพลาดในการตรวจสอบข้อมูลผู้ใช้");
  }
}

/**
 * Sets a new password for a user and activates their account.
 */
function setNewPassword(userId, newPassword) {
  if (!newPassword || newPassword.length < 4) {
    throw new Error("รหัสผ่านต้องมีความยาวอย่างน้อย 4 ตัวอักษร");
  }
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) throw new Error('ไม่พบชีตข้อมูล "Users"');

    const data = usersSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() == userId.toString().trim()) {
        const userRowIndex = i + 1;
        usersSheet.getRange(userRowIndex, 2).setValue(newPassword);
        usersSheet.getRange(userRowIndex, 6).setValue('active');
        SpreadsheetApp.flush();
        return { success: true, message: "ตั้งรหัสผ่านใหม่สำเร็จ!" };
      }
    }
    throw new Error("ไม่พบข้อมูลผู้ใช้งาน");
  } catch (e) {
    Logger.log(`[CRITICAL ERROR] in setNewPassword: ${e.toString()}`);
    throw new Error(`เกิดข้อผิดพลาดฝั่งเซิร์ฟเวอร์: ${e.message}`);
  }
}

/**
 * Resets a student's password to a default value.
 */
function adminResetPassword(adminId, studentId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  _ensureAdmin(adminId, ss);
  const usersSheet = ss.getSheetByName('Users');
  const data = usersSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() == studentId.toString().trim()) {
      if (data[i][3] === 'admin') throw new Error("ไม่สามารถรีเซ็ตรหัสผ่านของแอดมินได้");
      const userRowIndex = i + 1;
      const defaultPassword = studentId;
      usersSheet.getRange(userRowIndex, 2).setValue(defaultPassword);
      usersSheet.getRange(userRowIndex, 6).setValue('pending_reset');
      SpreadsheetApp.flush();
      return { success: true, message: `รีเซ็ตรหัสผ่านของ ${studentId} สำเร็จ รหัสผ่านใหม่คือ ${defaultPassword}` };
    }
  }
  throw new Error("ไม่พบรหัสนักเรียนที่ระบุ");
}

/**
 * Ensures the provided user ID belongs to an admin.
 */
function _ensureAdmin(userId, ss) {
  const usersSheet = ss.getSheetByName('Users');
  const usersData = usersSheet.getDataRange().getValues();
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0].toString().trim() === userId.toString().trim()) {
      if (usersData[i][3] === 'admin') {
        return; // User is an admin, continue execution.
      } else {
        throw new Error('Permission denied. User is not an admin.');
      }
    }
  }
  throw new Error('Permission denied. User not found.');
}

/**
 * Resets all student levels and rewards, and sets the semester start date to today.
 */
function startNewSemester(adminId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    _ensureAdmin(adminId, ss); // Security check
    const usersSheet = ss.getSheetByName('Users');
    const settingsSheet = ss.getSheetByName('Settings');

    if (!usersSheet || !settingsSheet) {
      throw new Error("ไม่พบชีต Users หรือ Settings");
    }

    // Clear levels (column G) and rewards (column H) for all users except admins
    const usersData = usersSheet.getDataRange().getValues();
    for (let i = 1; i < usersData.length; i++) {
      const role = usersData[i][3]; // Role is in column D
      if (role !== 'admin') {
        const userRowIndex = i + 1;
        usersSheet.getRange(userRowIndex, 7).setValue(''); // Clear level
        usersSheet.getRange(userRowIndex, 8).setValue(''); // Clear rewards
      }
    }

    // Update SemesterStartDate to today's date
    // Assumes SemesterStartDate is in B2
    settingsSheet.getRange('B2').setValue(new Date());
    
    SpreadsheetApp.flush();
    return { success: true, message: "เริ่มต้นภาคเรียนใหม่สำเร็จ! ระบบได้รีเซ็ตระดับและของรางวัลของนักเรียนทั้งหมดแล้ว" };

  } catch (e) {
    Logger.log(`[CRITICAL ERROR] in startNewSemester: ${e.message}`);
    throw new Error(`เกิดข้อผิดพลาดในการเริ่มต้นภาคเรียนใหม่: ${e.message}`);
  }
}

/**
 * Retrieves all transactions for a specific student.
 */
function getTransactions(studentId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const transactionsSheet = ss.getSheetByName('Transactions');
  if (!transactionsSheet) return [];
  const data = transactionsSheet.getDataRange().getValues();
  const transactions = [];
  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][2].toString().trim() == studentId.toString().trim()) {
      transactions.push({
        date: new Date(data[i][1]).toLocaleDateString('th-TH', { year: 'numeric', month: 'short', day: 'numeric' }),
        type: data[i][3],
        amount: data[i][4],
        note: data[i][5] || ''
      });
    }
  }
  return transactions;
}

/**
 * Counts deposit transactions for a user, determines their level, and updates the Users sheet.
 */
function updateUserLevel(studentId, ss) {
  const usersSheet = ss.getSheetByName('Users');
  const transactionsSheet = ss.getSheetByName('Transactions');
  const settingsSheet = ss.getSheetByName('Settings');
  if (!usersSheet || !transactionsSheet || !settingsSheet) {
    Logger.log('Update Level failed: Missing required sheets.');
    return;
  }

  let semesterStartDate;
  try {
    const dateValue = settingsSheet.getRange('B2').getValue();
    semesterStartDate = new Date(dateValue);
    if (isNaN(semesterStartDate.getTime())) throw new Error('Invalid date format');
  } catch (e) {
    Logger.log(`Could not read SemesterStartDate: ${e.message}. Defaulting to beginning of time.`);
    semesterStartDate = new Date(0); // Default to a very old date if not set
  }

  const txData = transactionsSheet.getDataRange().getValues();
  let depositCount = 0;
  let depositAmount = 0;

  for (let i = 1; i < txData.length; i++) {
    const transactionDate = new Date(txData[i][1]);
    // Only count transactions within the current semester
    if (transactionDate < semesterStartDate) continue;

    // Check studentId in column C (index 2) and type in column D (index 3)
    if (txData[i][2].toString().trim() === studentId.toString().trim() && txData[i][3] === 'ฝาก') {
      depositCount++;
      depositAmount += parseFloat(txData[i][4]) || 0; // Sum amount from column E (index 4)
    }
  }

  // Determine level by deposit count
  let levelByCount = 'ยังไม่มีระดับ';
  if (depositCount >= 100) levelByCount = 'Gold';
  else if (depositCount >= 30) levelByCount = 'Silver';
  else if (depositCount >= 10) levelByCount = 'Bronze';

  // Determine level by deposit amount
  let levelByAmount = 'ยังไม่มีระดับ';
  if (depositAmount >= 1000) levelByAmount = 'Gold';
  else if (depositAmount >= 500) levelByAmount = 'Silver';
  else if (depositAmount >= 200) levelByAmount = 'Bronze';

  // Compare levels and pick the higher one
  const levelValues = { 'ยังไม่มีระดับ': 0, 'Bronze': 1, 'Silver': 2, 'Gold': 3 };
  const finalLevel = levelValues[levelByCount] > levelValues[levelByAmount] ? levelByCount : levelByAmount;

  const usersData = usersSheet.getDataRange().getValues();
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0].toString().trim() === studentId.toString().trim()) {
      const userRowIndex = i + 1;
      // Only update if the level has changed
      if (usersData[i][6] !== finalLevel) {
        usersSheet.getRange(userRowIndex, 7).setValue(finalLevel); // Column G for level
      }
      break;
    }
  }
}

/**
 * Allows a student to claim a reward based on their current level.
 */
function claimReward(studentId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) throw new Error('ไม่พบชีตข้อมูล "Users"');

    const data = usersSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === studentId.toString().trim()) {
        const userRowIndex = i + 1;
        const currentLevel = data[i][6]; // level is in Column G
        const rewardStatus = data[i][7]; // rewards is in Column H

        if (currentLevel === 'ยังไม่มีระดับ') {
          throw new Error('คุณยังไม่มีระดับที่จะรับรางวัลได้');
        }
        if (rewardStatus.includes(currentLevel)) {
          throw new Error(`คุณได้รับรางวัลสำหรับระดับ ${currentLevel} ไปแล้ว`);
        }
        
        // Append the new reward to existing ones
        const newRewardStatus = rewardStatus ? `${rewardStatus}, ${currentLevel}` : currentLevel;
        usersSheet.getRange(userRowIndex, 8).setValue(newRewardStatus); // Column H for rewards
        SpreadsheetApp.flush();
        return { success: true, message: `ยินดีด้วย! คุณได้รับรางวัลระดับ ${currentLevel} สำเร็จ!` };
      }
    }
    throw new Error('ไม่พบข้อมูลผู้ใช้งาน');
  } catch (e) {
    Logger.log(`Claim Reward Error for [${studentId}]: ${e.message}`);
    throw new Error(e.message);
  }
}

/**
 * Adds a new transaction with a specified date and updates the student's balance.
 */
function addTransaction(adminId, studentId, type, amount, note, transactionDateString) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    _ensureAdmin(adminId, ss);
    const usersSheet = ss.getSheetByName('Users');
    const transactionsSheet = ss.getSheetByName('Transactions');
    const userData = usersSheet.getDataRange().getValues();
    
    let userRowIndex = -1;
    for (let i = 1; i < userData.length; i++) {
      if (userData[i][0].toString().trim() == studentId.toString().trim()) {
        userRowIndex = i + 1;
        break;
      }
    }
    if (userRowIndex === -1) throw new Error("ไม่พบรหัสนักเรียนนี้ในระบบ");

    const transactionDate = transactionDateString ? new Date(transactionDateString) : new Date();

    transactionsSheet.appendRow(['', transactionDate, studentId, type, amount, note, 'admin']);

    const balanceCell = usersSheet.getRange(userRowIndex, 5);
    const currentBalance = parseFloat(balanceCell.getValue()) || 0;
    const newBalance = (type === 'ฝาก') ? currentBalance + amount : currentBalance - amount;
    balanceCell.setValue(newBalance);

    // If a deposit was made, check and update the user's level.
    if (type === 'ฝาก') {
      updateUserLevel(studentId, ss);
    }
    
    SpreadsheetApp.flush();
    return { success: true, message: `บันทึกรายการสำเร็จ` };
  } catch (e) {
      Logger.log(`Add Transaction Error: ${e.message}`);
      throw new Error(e.message);
  }
}

/**
 * Calculates and retrieves dashboard summary data for all users.
 */
function getAdminDashboardData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const transactionsSheet = ss.getSheetByName('Transactions');
  const data = transactionsSheet.getDataRange().getValues();
  
  const now = new Date();
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const weekStart = new Date(todayStart);
  weekStart.setDate(todayStart.getDate() - todayStart.getDay());
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);

  let daily = { deposits: 0, withdrawals: 0, count: 0 };
  let weekly = { deposits: 0, withdrawals: 0, count: 0 };
  let monthly = { deposits: 0, withdrawals: 0, count: 0 };

  for (let i = 1; i < data.length; i++) {
    const txDate = new Date(data[i][1]);
    if (isNaN(txDate.getTime())) continue;
    
    const type = data[i][3];
    const amount = parseFloat(data[i][4]);
    if (txDate.getTime() >= todayStart.getTime()) {
      daily.count++;
      if (type === 'ฝาก') daily.deposits += amount; else daily.withdrawals += amount;
    }
    if (txDate.getTime() >= weekStart.getTime()) {
      weekly.count++;
      if (type === 'ฝาก') weekly.deposits += amount; else weekly.withdrawals += amount;
    }
    if (txDate.getTime() >= monthStart.getTime()) {
      monthly.count++;
      if (type === 'ฝาก') monthly.deposits += amount; else monthly.withdrawals += amount;
    }
  }
  return { daily, weekly, monthly };
}

/**
 * Retrieves the 20 most recent deposit transactions.
 */
function getRecentDeposits() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const transactionsSheet = ss.getSheetByName('Transactions');
    const usersSheet = ss.getSheetByName('Users');

    if (!transactionsSheet || !usersSheet) {
      throw new Error("ไม่พบชีต Transactions หรือ Users");
    }

    // Create a map of user IDs to names for quick lookup
    const usersData = usersSheet.getDataRange().getValues();
    const userNames = new Map();
    for (let i = 1; i < usersData.length; i++) {
      userNames.set(usersData[i][0].toString().trim(), usersData[i][2]); // Map ID to Name
    }

    const txData = transactionsSheet.getDataRange().getValues();
    const recentDeposits = [];
    
    // Iterate backwards from the last row
    for (let i = txData.length - 1; i > 0 && recentDeposits.length < 20; i--) {
      const row = txData[i];
      const type = row[3];

      if (type === 'ฝาก') {
        const studentId = row[2].toString().trim();
        recentDeposits.push({
          date: new Date(row[1]).toLocaleDateString('th-TH', { year: 'numeric', month: 'short', day: 'numeric' }),
          studentId: studentId,
          studentName: userNames.get(studentId) || 'ไม่พบชื่อ',
          amount: row[4],
          note: row[5] || ''
        });
      }
    }
    return recentDeposits;
  } catch (e) {
    Logger.log(`Error in getRecentDeposits: ${e.message}`);
    throw new Error('เกิดข้อผิดพลาดในการดึงข้อมูลการฝากล่าสุด');
  }
}

/**
 * Calculates and retrieves dashboard summary data for a single student.
 */
function getStudentDashboardData(studentId) {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const transactionsSheet = ss.getSheetByName('Transactions');
    const data = transactionsSheet.getDataRange().getValues();
    
    const now = new Date();
    const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const weekStart = new Date(todayStart);
    weekStart.setDate(todayStart.getDate() - todayStart.getDay());
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);

    let daily = { deposits: 0, withdrawals: 0 };
    let weekly = { deposits: 0, withdrawals: 0 };
    let monthly = { deposits: 0, withdrawals: 0 };

    for (let i = 1; i < data.length; i++) {
        if(data[i][2].toString().trim() !== studentId.toString().trim()) continue;
        const txDate = new Date(data[i][1]);
        if (isNaN(txDate.getTime())) continue;
        const type = data[i][3];
        const amount = parseFloat(data[i][4]);
        if (txDate.getTime() >= todayStart.getTime()) {
            if (type === 'ฝาก') daily.deposits += amount; else daily.withdrawals += amount;
        }
        if (txDate.getTime() >= weekStart.getTime()) {
            if (type === 'ฝาก') weekly.deposits += amount; else weekly.withdrawals += amount;
        }
        if (txDate.getTime() >= monthStart.getTime()) {
            if (type === 'ฝาก') monthly.deposits += amount; else monthly.withdrawals += amount;
        }
    }
    return { daily, weekly, monthly };
}

/**
 * Creates a native Excel (.xlsx) file for direct download.
 */
function exportTransactionsAsExcel(adminId) {
  let tempSpreadsheet;
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    _ensureAdmin(adminId, ss);
    const transactionsSheet = ss.getSheetByName('Transactions');
    const data = transactionsSheet.getDataRange().getValues();
    
    const fileName = `ประวัติธุรกรรมทั้งหมด ณ ${new Date().toLocaleString('th-TH')}`;
    tempSpreadsheet = SpreadsheetApp.create(fileName);
    const exportSheet = tempSpreadsheet.getSheets()[0];
    
    exportSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    exportSheet.getRange(1, 1, 1, data[0].length).setFontWeight('bold');
    for (let i = 1; i <= data[0].length; i++) { exportSheet.autoResizeColumn(i); }
    SpreadsheetApp.flush();

    const url = `https://docs.google.com/spreadsheets/d/${tempSpreadsheet.getId()}/export?format=xlsx`;
    const options = { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } };
    const excelBlob = UrlFetchApp.fetch(url, options).getBlob().setName(fileName + ".xlsx");
    
    const fileData = Utilities.base64Encode(excelBlob.getBytes());
    
    DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
    
    return {
      fileData: fileData,
      fileName: excelBlob.getName()
    };
  } catch (e) {
    if(tempSpreadsheet) {
      DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
    }
    Logger.log(`Excel Export Error: ${e.message}`);
    throw new Error('ไม่สามารถสร้างไฟล์ Excel ได้');
  }
}

/**
 * Creates a PDF report of the admin summary dashboard.
 */
function exportSummaryAsPdf(adminId) {
  let tempSpreadsheet;
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    _ensureAdmin(adminId, ss);
    const summaryData = getAdminDashboardData();
    const now = new Date();
    const thaiMonthName = now.toLocaleDateString('th-TH', { month: 'long', year: 'numeric' });
    
    tempSpreadsheet = SpreadsheetApp.create('Temporary Summary Report');
    const sheet = tempSpreadsheet.getSheets()[0];
    
    sheet.getRange('A1:D1').merge().setValue('รายงานสรุปผล โครงการออมเก่งเหล่าประชา').setHorizontalAlignment('center').setFontWeight('bold').setFontSize(14);
    sheet.getRange('A2:D2').merge().setValue(`ข้อมูล ณ วันที่: ${now.toLocaleDateString('th-TH', {dateStyle: 'full'})}`).setHorizontalAlignment('center');
    
    sheet.getRange('A4').setValue('ช่วงเวลา').setFontWeight('bold');
    sheet.getRange('B4:D4').setValues([['ยอดฝาก', 'ยอดถอน', 'จำนวนรายการ']]).setFontWeight('bold').setHorizontalAlignment('center');
    
    const data = [
      ['วันนี้', summaryData.daily.deposits, summaryData.daily.withdrawals, summaryData.daily.count],
      ['สัปดาห์นี้', summaryData.weekly.deposits, summaryData.weekly.withdrawals, summaryData.weekly.count],
      [`เดือน${thaiMonthName}`, summaryData.monthly.deposits, summaryData.monthly.withdrawals, summaryData.monthly.count]
    ];
    sheet.getRange('A5:D7').setValues(data).setHorizontalAlignment('center');
    sheet.getRange('A5:A7').setHorizontalAlignment('left');
    
    sheet.getRange('B5:C7').setNumberFormat('#,##0.00');
    for (let i = 1; i <= 4; i++) { sheet.autoResizeColumn(i); }
    SpreadsheetApp.flush();

    const pdfBlob = tempSpreadsheet.getAs('application/pdf').setName(`สรุปผล-${thaiMonthName}.pdf`);
    const pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
    
    DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
    
    return {
      pdfData: pdfBase64,
      fileName: pdfBlob.getName()
    };
    
  } catch (e) {
    if (tempSpreadsheet) {
      DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
    }
    Logger.log(`PDF Export Error: ${e.message}`);
    throw new Error('ไม่สามารถสร้างไฟล์ PDF ได้');
  }
}

/**
 * Adds multiple transactions at once and updates balances efficiently.
 */
function addMultipleTransactions(adminId, dateString, type, transactions) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    _ensureAdmin(adminId, ss);
    const usersSheet = ss.getSheetByName('Users');
    const transactionsSheet = ss.getSheetByName('Transactions');
    
    const usersData = usersSheet.getDataRange().getValues();
    const usersMap = new Map();
    for (let i = 1; i < usersData.length; i++) {
      const userId = usersData[i][0].toString().trim();
      usersMap.set(userId, {
        rowIndex: i + 1,
        balance: parseFloat(usersData[i][4]) || 0
      });
    }

    const transactionDate = dateString ? new Date(dateString) : new Date();
    const transactionsToAppend = [];
    let successfulCount = 0;
    
    for (const tx of transactions) {
      const studentId = tx.studentId.toString().trim();
      if (usersMap.has(studentId)) {
        const userData = usersMap.get(studentId);
        const amount = tx.amount;
        
        const newBalance = (type === 'ฝาก') ? userData.balance + amount : userData.balance - amount;
        userData.balance = newBalance;
        usersMap.set(studentId, userData);
        
        transactionsToAppend.push(['', transactionDate, studentId, type, amount, tx.note || '', 'admin']);
        successfulCount++;
      }
    }
    
    if (transactionsToAppend.length > 0) {
      transactionsSheet.getRange(transactionsSheet.getLastRow() + 1, 1, transactionsToAppend.length, 7).setValues(transactionsToAppend);
      for (const [userId, userData] of usersMap.entries()) {
        if (usersData[userData.rowIndex - 1][4] != userData.balance) {
            usersSheet.getRange(userData.rowIndex, 5).setValue(userData.balance);
        }
      }
      
      // After balances are updated, update levels for all affected users if it was a deposit
      if (type === 'ฝาก') {
        // Create a unique list of student IDs to avoid redundant updates
        const uniqueStudentIds = [...new Set(transactions.map(tx => tx.studentId.toString().trim()))];
        for (const studentId of uniqueStudentIds) {
           if (usersMap.has(studentId)) {
             updateUserLevel(studentId, ss);
           }
        }
      }
    }
    
    SpreadsheetApp.flush();
    return { message: `บันทึกข้อมูลสำเร็จ ${successfulCount} รายการ จากทั้งหมด ${transactions.length} รายการ` };

  } catch (e) {
    Logger.log(`Bulk Add Error: ${e.message}`);
    throw new Error('เกิดข้อผิดพลาดในการบันทึกข้อมูลหลายรายการ');
  }
}