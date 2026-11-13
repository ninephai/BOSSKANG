// อ้างอิง : https://www.youtube.com/@Pormodtanoy/videos / tanong.edu

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// สร้างหรือค้นหาโฟลเดอร์ใน Google Drive
function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

// ฟังก์ชันอัปโหลดไฟล์ - แก้ไขให้ return URL ที่ shareable จริง
function uploadFile(fileData, fileName, folderName) {
  try {
    Logger.log("Starting upload for file: " + fileName);
    const folder = getOrCreateFolder(folderName);
    const contentType = fileData.substring(5, fileData.indexOf(';'));
    const base64Data = fileData.substring(fileData.indexOf(',') + 1);
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, fileName);
    const file = folder.createFile(blob);
    
    // แก้ไข: Set sharing ถูกต้องและ return URL ที่เปิดได้จริง (viewer mode)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = file.getUrl();  // URL มาตรฐานที่ shareable
    
    Logger.log("Upload success. File URL: " + fileUrl);
    return fileUrl;
  } catch (e) {
    Logger.log("Error in uploadFile: " + e.toString());
    throw new Error("Upload failed: " + e.toString());
  }
}

// ฟังก์ชันสร้าง ID แบบ UUID (ปรับจาก timestamp แล้ว)
function generateId() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random() * 16 | 0;
    var v = c == 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

// แก้ไข saveToSpreadsheet - ใช้ priority ย่อและ log URL
function saveToSpreadsheet(item) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Sheet1');
    if (!sheet) {
      Logger.log("Sheet1 not found in saveToSpreadsheet");
      throw new Error("Sheet1 not found");
    }

    // สร้าง ID ใหม่หากไม่มี
    const id = item.id || generateId();
    
    // Map priority เต็ม → ย่อ (สำหรับบันทึกใน F)
    let priorityAbbrev = item.priority;
    if (priorityAbbrev === 'สูง') priorityAbbrev = 'ส';
    else if (priorityAbbrev === 'กลาง') priorityAbbrev = 'ป';
    else if (priorityAbbrev === 'ต่ำ') priorityAbbrev = 'น';
    else priorityAbbrev = 'ป';  // default

    const status = item.checked ? 'YES' : 'NO';

    let fileUrl = item.note || '';
    if (item.fileData && item.fileName) {
      fileUrl = uploadFile(item.fileData, item.fileName, "TodoFiles");
      Logger.log("File URL saved: " + fileUrl);
    }

    sheet.appendRow([
      id,
      item.date,
      item.time,
      item.name,
      status,
      priorityAbbrev,  // บันทึกย่อใน F
      fileUrl
    ]);

    Logger.log("Item saved: ID=" + id + ", PriorityAbbrev=" + priorityAbbrev + ", FileURL=" + fileUrl);
    
    return {
      id: id,
      date: item.date,
      time: item.time,
      name: item.name,
      checked: item.checked,
      priority: item.priority,  // return เต็มสำหรับ frontend
      note: fileUrl
    };
  } catch (e) {
    Logger.log("Error in saveToSpreadsheet: " + e.toString());
    throw e;
  }
}

// แก้ไข getAllItems - Map priority ย่อ → เต็มสำหรับ frontend
function getAllItems() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Sheet1');
    if (!sheet) {
      Logger.log("Sheet1 not found in getAllItems");
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (!data || data.length === 0) {
      Logger.log("No data in Sheet1");
      return [];
    }

    const items = data.slice(1).map(row => {
      let dateValue = row[1];
      if (dateValue instanceof Date) {
        dateValue = Utilities.formatDate(dateValue, "Asia/Bangkok", "yyyy-MM-dd");
      } else if (dateValue && typeof dateValue === 'string') {
        const date = new Date(dateValue);
        if (!isNaN(date.getTime())) {
          dateValue = Utilities.formatDate(date, "Asia/Bangkok", "yyyy-MM-dd");
        } else {
          dateValue = dateValue.trim();
        }
      } else {
        dateValue = '';
      }

      let timeValue = row[2];
      if (timeValue instanceof Date) {
        timeValue = Utilities.formatDate(timeValue, "Asia/Bangkok", "HH:mm");
      } else {
        timeValue = timeValue ? timeValue.toString().trim() : '';
      }

      // Map priority ย่อ → เต็ม (จาก F)
      let priorityFull = row[5] || 'กลาง';
      if (priorityFull === 'ส') priorityFull = 'สูง';
      else if (priorityFull === 'ป') priorityFull = 'กลาง';
      else if (priorityFull === 'น') priorityFull = 'ต่ำ';

      return {
        id: String(row[0]) || '',
        date: dateValue,
        time: timeValue,
        name: row[3] || '',
        checked: row[4] === 'YES',
        priority: priorityFull,  // return เต็มสำหรับ frontend
        note: row[6] || ''  // URL ไฟล์
      };
    });

    Logger.log("Items retrieved: " + items.length + " items, sample note: " + (items[0] ? items[0].note : 'none'));
    return items;
  } catch (e) {
    Logger.log("Error in getAllItems: " + e.toString());
    return [];
  }
}

// ฟังก์ชันดึงข้อมูลตาม ID
function getItemById(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Sheet1');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(id)) {  // เปลี่ยนเป็น String() ทั้งคู่
        let dateValue = data[i][1];
        if (dateValue instanceof Date) {
          dateValue = Utilities.formatDate(dateValue, "Asia/Bangkok", "yyyy-MM-dd");
        }

        let timeValue = data[i][2];
        if (timeValue instanceof Date) {
          timeValue = Utilities.formatDate(timeValue, "Asia/Bangkok", "HH:mm");
        }

        return {
          id: data[i][0],
          date: dateValue,
          time: timeValue,
          name: data[i][3],
          checked: data[i][4] === 'YES',
          priority: data[i][5] || 'กลาง',
          note: data[i][6] || ''
        };
      }
    }
    throw new Error("Item not found");
  } catch (e) {
    Logger.log("Error in getItemById: " + e.toString());
    throw e;
  }
}

// ฟังก์ชันกรองข้อมูลตามวันที่
function getItemsByDate(date) {
  try {
    const allItems = getAllItems();
    return allItems.filter(item => item.date === date);
  } catch (e) {
    Logger.log("Error in getItemsByDate: " + e.toString());
    return [];
  }
}

// ฟังก์ชันค้นหาข้อมูล
function searchItems(searchText) {
  try {
    const allItems = getAllItems();
    const searchLower = searchText.toLowerCase();
    return allItems.filter(item => 
      item.name.toLowerCase().includes(searchLower) ||
      item.date.includes(searchLower) ||
      item.time.includes(searchLower) ||
      item.priority.toLowerCase().includes(searchLower)
    );
  } catch (e) {
    Logger.log("Error in searchItems: " + e.toString());
    return [];
  }
}

// ปรับฟังก์ชัน updateItem เพื่อจัดการไฟล์ใหม่
// Map priority เต็ม → ย่อ (ตาม Sheets)
function mapPriorityToAbbrev(priority) {
  switch(priority) {
    case 'สูง': return 'ส';
    case 'กลาง': return 'ป';
    case 'ต่ำ': return 'น';
    default: return 'ป';
  }
}

// แก้ไข updateItem - Map priority และ log URL
function updateItem(item) {
  try {
    if (!item || !item.id) {
      Logger.log("Error: Invalid item or missing id: " + JSON.stringify(item));
      throw new Error("Invalid item or missing id");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Sheet1');
    const data = sheet.getDataRange().getValues();

    // Map priority เต็ม → ย่อ
    let priorityAbbrev = item.priority;
    if (priorityAbbrev === 'สูง') priorityAbbrev = 'ส';
    else if (priorityAbbrev === 'กลาง') priorityAbbrev = 'ป';
    else if (priorityAbbrev === 'ต่ำ') priorityAbbrev = 'น';
    else priorityAbbrev = 'ป';

    let fileUrl = item.note || '';
    if (item.fileData && item.fileName) {
      fileUrl = uploadFile(item.fileData, item.fileName, "TodoFiles");
      Logger.log("Updated file URL: " + fileUrl);
    }

    let found = false;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(item.id)) {
        sheet.getRange(i + 1, 1, 1, 7).setValues([[
          item.id,
          item.date,
          item.time,
          item.name,
          data[i][4],  // คงสถานะเดิม (E)
          priorityAbbrev,  // ย่อ (F)
          fileUrl
        ]]);
        found = true;
        Logger.log("Item updated: ID=" + item.id + ", PriorityAbbrev=" + priorityAbbrev + ", FileURL=" + fileUrl);
        break;
      }
    }

    if (!found) {
      Logger.log("Error: Item not found for id: " + item.id);
      throw new Error("Item not found");
    }

    return {
      id: item.id,
      date: item.date,
      time: item.time,
      name: item.name,
      checked: item.checked,
      priority: item.priority,  // return เต็ม
      note: fileUrl
    };
  } catch (e) {
    Logger.log("Error in updateItem: " + e.toString());
    throw e;
  }
}


//========================= ปรับ updateItemStatus (เพิ่ม log และ return) =====================================
function updateItemStatus(id, checked) {
  try {
    Logger.log("Updating status for ID: " + id + " to " + (checked ? 'YES' : 'NO'));
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Sheet1');
    if (!sheet) throw new Error("Sheet1 not found");
    
    const data = sheet.getDataRange().getValues();
    let found = false;
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, 5).setValue(checked ? 'YES' : 'NO'); // E: Status
        found = true;
        Logger.log("Status updated successfully for ID: " + id);
        return true;
      }
    }
    
    if (!found) {
      Logger.log("Error: Item not found for status update ID: " + id);
      throw new Error("Item not found for status: " + id);
    }
  } catch (e) {
    Logger.log("Error in updateItemStatus: " + e.toString());
    throw e;
  }
}

// ฟังก์ชันส่งข้อความไปยัง Telegram
function sendTelegramMessage(message) {
  const botToken = '9056094469:AAEi9lP9Y7ffPKDgywmnGRHqocBzt3VhyDk'; //เปลี่ยนเป็น ToKen telegram ท่าน
  const chatId = '-1003743409575'; // chat ID Telegram
  const url = `https://api.telegram.org/bot${botToken}/sendMessage`;

  const payload = {
    chat_id: chatId,
    text: message,
    parse_mode: 'HTML'
  };

  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    Logger.log('Telegram Response: ' + response.getContentText());
    return true;
  } catch (e) {
    Logger.log('Error sending Telegram message: ' + e.toString());
    return false;
  }
}

function sendDailyNotification() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Sheet2');
    if (!sheet) {
      Logger.log('Sheet2 not found');
      sendTelegramMessage('Error: Sheet2 not found');
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('No data in Sheet2');
      sendTelegramMessage('No data found in Sheet2');
      return;
    }

    const headers = sheet.getRange(1, 1, 1, 4).getValues()[0];
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    let message = '📋 <b>แจ้งเตือนภาระงานคงค้าง</b>\n\n';
    let hasNotification = false;

    data.forEach((row, index) => {
      if (row[4] && row[4].toString().trim() === 'แจ้งเตือน') {
        hasNotification = true;
        let dateCol1 = row[0];
        let timeCol2 = row[1];

        if (dateCol1 instanceof Date && !isNaN(dateCol1)) {
          dateCol1 = Utilities.formatDate(dateCol1, "Asia/Bangkok", "yyyy-MM-dd");
        } else {
          dateCol1 = row[0] || 'Invalid Date';
        }

        if (timeCol2 && typeof timeCol2 === 'string') {
          timeCol2 = timeCol2.trim();
        } else if (timeCol2 instanceof Date && !isNaN(timeCol2)) {
          timeCol2 = Utilities.formatDate(timeCol2, "Asia/Bangkok", "HH:mm:ss");
        } else {
          timeCol2 = 'Invalid Time';
        }

        message += `<b>รายการ ${index + 1}</b>\n`;
        message += `${headers[0] || 'Column 1'}: ${dateCol1}\n`;
        message += `${headers[1] || 'Column 2'}: ${timeCol2}\n`;
        message += `${headers[2] || 'Column 3'}: ${row[2] || '-'}\n`;
        message += `${headers[3] || 'Column 4'}: ${row[3] || '-'}\n\n`;
      }
    });

    if (!hasNotification) {
      Logger.log('No rows with "แจ้งเตือน" found');
      sendTelegramMessage('No tasks requiring notification found');
      return;
    }

    Logger.log('Message length: ' + message.length);
    if (message.length > 4000) {
      message = message.substring(0, 4000) + '...';
      Logger.log('Message truncated to 4000 characters');
    }
    sendTelegramMessage(message);

    Logger.log('Notification sent: ' + message);
    return true;
  } catch (e) {
    Logger.log('Error in sendDailyNotification: ' + e.toString());
    sendTelegramMessage('Error in notification: ' + e.toString());
    return false;
  }
}



function deleteItemFromSheet(id) {
  try {
    if (!id) {
      Logger.log("Error: Invalid id for deletion: " + id);
      throw new Error("Invalid id for deletion");
    }

    Logger.log("Attempting to delete ID: " + id);  // เพิ่ม log

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Sheet1');  // ตรวจสอบชื่อ sheet ตรงไหม (จากภาพคือ Sheet1)
    if (!sheet) {
      Logger.log("Error: Sheet1 not found in deleteItemFromSheet");
      throw new Error("Sheet1 not found");
    }

    const data = sheet.getDataRange().getValues();
    let found = false;

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {  // ใช้ String() เปรียบเทียบ
        sheet.deleteRow(i + 1);
        found = true;
        Logger.log("Item deleted successfully: row " + (i + 1) + ", ID: " + id);
        break;
      }
    }

    if (!found) {
      Logger.log("Error: Item not found for deletion ID: " + id + ". Available IDs: " + data.slice(1).map(row => String(row[0])).join(', '));
      throw new Error("Item not found for deletion: " + id);
    }

    return true;
  } catch (e) {
    Logger.log("Error in deleteItemFromSheet: " + e.toString());
    throw e;
  }
}

// ==========================================
// 1. API ROUTER (รับคำสั่งจาก GitHub Pages)
// ==========================================
function doPost(e) {
  // ต้องตั้งค่า headers เป็น text/plain จากฝั่งหน้าเว็บ เพื่อเลี่ยงปัญหา CORS Preflight
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    let result;

    switch (action) {
      case 'getAllItems':
        result = getAllItems();
        break;
      case 'getItemsByDate':
        result = getItemsByDate(payload.date);
        break;
      case 'saveItem':
        result = saveToSpreadsheet(payload.item);
        break;
      case 'updateItem':
        result = updateItem(payload.item);
        break;
      case 'updateItemStatus':
        result = updateItemStatus(payload.id, payload.checked);
        break;
      case 'deleteItemFromSheet':
        result = deleteItemFromSheet(payload.id);
        break;
      default:
        throw new Error("Action ไม่ถูกต้อง: " + action);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      data: result
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log("API Error: " + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("API is running...").setMimeType(ContentService.MimeType.TEXT);
}



// function verifyPassword(password) {
//   var validPassword = '1970'; 
//   return password === validPassword;
// }
