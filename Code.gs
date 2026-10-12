/**
 * Cinema Room Booking System - Backend Google Apps Script (Code.gs)
 * Full CRUD for Bookings, Rooms, and Users + LINE Push Notification
 */

const TAB_USERS = "Users";
const TAB_BOOKINGS = "Bookings";
const TAB_ROOMS = "Rooms";

// Config ค่า LINE Messaging API (สามารถใส่ค่าตรงนี้ หรือตั้งค่าใน Script Properties ก็ได้)
const LINE_CHANNEL_ACCESS_TOKEN = "YOUR_LINE_CHANNEL_ACCESS_TOKEN";
const LINE_TARGET_ID = "YOUR_LINE_TARGET_ID"; // ใส่ User ID (Uxxx) หรือ Group ID (Cxxx)

function doGet(e) {
  initDatabase();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ระบบจองห้องดูหนังออนไลน์ - CINEBOOK VIP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function initDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Users Sheet
  let userSheet = ss.getSheetByName(TAB_USERS);
  if (!userSheet) {
    userSheet = ss.insertSheet(TAB_USERS);
    userSheet.appendRow(['id', 'fullname', 'username', 'password', 'telephone', 'role', 'status']);
    userSheet.appendRow(['USR-001', 'ผู้ดูแลระบบสูงสุด', 'admin', 'admin123', '0812345678', 'Admin', 'Active']);
    userSheet.appendRow(['USR-002', 'สมชาย ใจดี', 'user1', '123456', '0898765432', 'User', 'Active']);
    userSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1f1938').setFontColor('#ffffff');
  }

  // 2. Bookings Sheet
  let bookingSheet = ss.getSheetByName(TAB_BOOKINGS);
  if (!bookingSheet) {
    bookingSheet = ss.insertSheet(TAB_BOOKINGS);
    bookingSheet.appendRow(['booking_id', 'username', 'fullname', 'movie_title', 'room_number', 'date', 'time', 'status', 'created_at']);
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    bookingSheet.appendRow(['BK-1001', 'user1', 'สมชาย ใจดี', 'Godzilla x Kong: The New Empire', '101', todayStr, '16:00 - 18:30', 'Confirmed', new Date().toLocaleString('th-TH')]);
    bookingSheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#1f1938').setFontColor('#ffffff');
  }

  // 3. Rooms Sheet
  let roomSheet = ss.getSheetByName(TAB_ROOMS);
  if (!roomSheet || roomSheet.getLastRow() <= 1) {
    if (!roomSheet) roomSheet = ss.insertSheet(TAB_ROOMS);
    if (roomSheet.getLastRow() === 0) {
      roomSheet.appendRow(['room_number', 'room_name', 'total_capacity', 'available_status']);
    }
    if (roomSheet.getLastRow() === 1) {
      roomSheet.appendRow(['101', 'VIP Laser Cinema Suite', 4, 'Available']);
      roomSheet.appendRow(['102', 'Premium Couple Recliner', 2, 'Available']);
      roomSheet.appendRow(['103', 'IMAX Mini Private Lounge', 8, 'Available']);
      roomSheet.appendRow(['104', 'Family Ultimate 4K Room', 10, 'Available']);
      roomSheet.appendRow(['105', 'Standard VIP Screen', 6, 'Available']);
    }
    roomSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1f1938').setFontColor('#ffffff');
  }
}

/* ==================== LINE NOTIFICATION API ==================== */

/**
 * ส่งข้อความไปยัง LINE Target ID ผ่าน Messaging API Push
 * @param {string} textMessage - ข้อความที่ต้องการส่ง
 */
function sendLineMessage(textMessage) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN') || LINE_CHANNEL_ACCESS_TOKEN;
  const targetId = props.getProperty('LINE_TARGET_ID') || LINE_TARGET_ID;

  if (!token || token === "YOUR_LINE_CHANNEL_ACCESS_TOKEN" || !targetId || targetId === "YOUR_LINE_TARGET_ID") {
    Logger.log("กรุณาตั้งค่า LINE_CHANNEL_ACCESS_TOKEN และ LINE_TARGET_ID");
    return false;
  }

  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = {
    to: targetId,
    messages: [
      {
        type: 'text',
        text: textMessage
      }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const resCode = response.getResponseCode();
    Logger.log('LINE Push Response Code: ' + resCode);
    return resCode === 200;
  } catch (error) {
    Logger.log('Error sending LINE message: ' + error.toString());
    return false;
  }
}

/* ==================== AUTHENTICATION & USERS ==================== */

function checkLogin(username, password) {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TAB_USERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[2]).trim() === String(username).trim() && String(row[3]).trim() === String(password).trim()) {
        if (String(row[6]).trim() !== 'Active') {
          return { success: false, message: 'บัญชีของคุณถูกระงับการใช้งาน (Inactive)' };
        }
        return {
          success: true,
          user: { id: row[0], fullname: row[1], username: row[2], telephone: row[4], role: String(row[5]).trim(), status: String(row[6]).trim() }
        };
      }
    }
    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง!' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function registerUser(userData) {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TAB_USERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim().toLowerCase() === String(userData.username).trim().toLowerCase()) {
        return { success: false, message: 'Username นี้ถูกใช้งานไปแล้ว' };
      }
    }
    
    const newId = 'USR-' + Math.floor(1000 + Math.random() * 9000);
    sheet.appendRow([newId, userData.fullname, userData.username, userData.password, userData.telephone, 'User', 'Active']);
    return { success: true, message: 'ลงทะเบียนสำเร็จ!' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function getUsers() {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TAB_USERS);
    const data = sheet.getDataRange().getValues();
    const users = [];
    
    for (let i = 1; i < data.length; i++) {
      users.push({
        id: String(data[i][0]),
        fullname: String(data[i][1]),
        username: String(data[i][2]),
        telephone: String(data[i][4]),
        role: String(data[i][5]),
        status: String(data[i][6])
      });
    }
    return users;
  } catch (error) {
    return [];
  }
}

function updateUser(userId, userData) {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TAB_USERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userId)) {
        const rowNum = i + 1;
        sheet.getRange(rowNum, 2).setValue(userData.fullname);
        sheet.getRange(rowNum, 5).setValue(userData.telephone);
        sheet.getRange(rowNum, 6).setValue(userData.role);
        sheet.getRange(rowNum, 7).setValue(userData.status);
        if (userData.password && userData.password.length > 0) {
          sheet.getRange(rowNum, 4).setValue(userData.password);
        }
        return { success: true, message: 'อัปเดตข้อมูลผู้ใช้งานเรียบร้อยแล้ว' };
      }
    }
    return { success: false, message: 'ไม่พบผู้ใช้งาน' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function deleteUser(userId) {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TAB_USERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userId)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'ลบผู้ใช้งานเรียบร้อยแล้ว' };
      }
    }
    return { success: false, message: 'ไม่พบรหัสสมาชิกนี้' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

/* ==================== ROOM MANAGEMENT ==================== */

function getAvailableRooms() {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TAB_ROOMS);
    if (!sheet) return getFallbackRooms();

    const data = sheet.getDataRange().getValues();
    const rooms = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        rooms.push({
          room_number: String(data[i][0]),
          room_name: String(data[i][1]),
          total_capacity: Number(data[i][2]),
          available_status: String(data[i][3])
        });
      }
    }
    return rooms.length > 0 ? rooms : getFallbackRooms();
  } catch (error) {
    return getFallbackRooms();
  }
}

function getFallbackRooms() {
  return [
    { room_number: '101', room_name: 'VIP Laser Cinema Suite', total_capacity: 4, available_status: 'Available' },
    { room_number: '102', room_name: 'Premium Couple Recliner', total_capacity: 2, available_status: 'Available' },
    { room_number: '103', room_name: 'IMAX Mini Private Lounge', total_capacity: 8, available_status: 'Available' },
    { room_number: '104', room_name: 'Family Ultimate 4K Room', total_capacity: 10, available_status: 'Available' },
    { room_number: '105', room_name: 'Standard VIP Screen', total_capacity: 6, available_status: 'Available' }
  ];
}

function saveRoom(roomData) {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TAB_ROOMS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(roomData.room_number)) {
        const rowNum = i + 1;
        sheet.getRange(rowNum, 2).setValue(roomData.room_name);
        sheet.getRange(rowNum, 3).setValue(roomData.total_capacity);
        sheet.getRange(rowNum, 4).setValue(roomData.available_status);
        return { success: true, message: 'อัปเดตข้อมูลห้องเรียบร้อย' };
      }
    }
    
    sheet.appendRow([roomData.room_number, roomData.room_name, roomData.total_capacity, roomData.available_status]);
    return { success: true, message: 'เพิ่มห้องภาพยนตร์ใหม่เรียบร้อย' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function deleteRoom(roomNumber) {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TAB_ROOMS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(roomNumber)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'ลบห้องพักเรียบร้อยแล้ว' };
      }
    }
    return { success: false, message: 'ไม่พบห้องพักนี้' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

/* ==================== BOOKING MANAGEMENT ==================== */

function saveBooking(bookingData) {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bookingSheet = ss.getSheetByName(TAB_BOOKINGS);
    const data = bookingSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][4]) === String(bookingData.room_number) && 
          String(data[i][5]) === String(bookingData.date) && 
          String(data[i][6]) === String(bookingData.time) && 
          String(data[i][7]) !== 'Canceled') {
        return { success: false, message: 'ห้องพักถูกจองไว้แล้วในช่วงเวลาดังกล่าว!' };
      }
    }
    
    const newBookingId = 'BK-' + Math.floor(10000 + Math.random() * 90000);
    const createdAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const movieTitle = bookingData.movie_title || 'Godzilla x Kong: The New Empire';

    bookingSheet.appendRow([
      newBookingId,
      bookingData.username,
      bookingData.fullname,
      movieTitle,
      bookingData.room_number,
      bookingData.date,
      bookingData.time,
      'Confirmed',
      createdAt
    ]);

    // แจ้งเตือนไปยัง LINE
    const lineNotifyText = `🎬 [แจ้งเตือนการจองใหม่]\n` +
                           `🎟 รหัสจอง: ${newBookingId}\n` +
                           `👤 ผู้จอง: ${bookingData.fullname} (@${bookingData.username})\n` +
                           `🍿 ภาพยนตร์: ${movieTitle}\n` +
                           `🚪 ห้อง: ${bookingData.room_number}\n` +
                           `📅 วันที่: ${bookingData.date}\n` +
                           `⏰ เวลา: ${bookingData.time}\n` +
                           `สถานะ: Confirmed ✅`;
    sendLineMessage(lineNotifyText);
    
    return { success: true, message: 'ทำการจองห้องสำเร็จ!', booking_id: newBookingId };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function updateBookingDetails(bookingId, updatedData) {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TAB_BOOKINGS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(bookingId)) {
        const rowNum = i + 1;
        sheet.getRange(rowNum, 4).setValue(updatedData.movie_title);
        sheet.getRange(rowNum, 5).setValue(updatedData.room_number);
        sheet.getRange(rowNum, 6).setValue(updatedData.date);
        sheet.getRange(rowNum, 7).setValue(updatedData.time);
        sheet.getRange(rowNum, 8).setValue(updatedData.status);

        // แจ้งเตือนไปยัง LINE
        const lineNotifyText = `✏️ [แจ้งเตือนอัปเดตการจอง]\n` +
                               `🎟 รหัสจอง: ${bookingId}\n` +
                               `🍿 ภาพยนตร์: ${updatedData.movie_title}\n` +
                               `🚪 ห้อง: ${updatedData.room_number}\n` +
                               `📅 วันที่: ${updatedData.date}\n` +
                               `⏰ เวลา: ${updatedData.time}\n` +
                               `สถานะ: ${updatedData.status}`;
        sendLineMessage(lineNotifyText);

        return { success: true, message: 'อัปเดตรายละเอียดการจองสำเร็จ' };
      }
    }
    return { success: false, message: 'ไม่พบรายการจองนี้' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function deleteBooking(bookingId) {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TAB_BOOKINGS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(bookingId)) {
        sheet.deleteRow(i + 1);

        // แจ้งเตือนไปยัง LINE
        const lineNotifyText = `🗑️ [แจ้งเตือนลบรายการจอง]\n` +
                               `🎟 รหัสจอง: ${bookingId}\n` +
                               `สถานะ: ถูกลบออกจากระบบแล้ว`;
        sendLineMessage(lineNotifyText);

        return { success: true, message: 'ลบข้อมูลการจองถาวรเรียบร้อยแล้ว' };
      }
    }
    return { success: false, message: 'ไม่พบรายการจอง' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function getUserBookings(username) {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TAB_BOOKINGS);
    const data = sheet.getDataRange().getValues();
    const userBookings = [];
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(username).trim()) {
        userBookings.push({
          booking_id: String(data[i][0]),
          username: String(data[i][1]),
          fullname: String(data[i][2]),
          movie_title: String(data[i][3]),
          room_number: String(data[i][4]),
          date: String(data[i][5]),
          time: String(data[i][6]),
          status: String(data[i][7]),
          created_at: String(data[i][8])
        });
      }
    }
    return userBookings.reverse();
  } catch (error) {
    return [];
  }
}

function getBookings() {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TAB_BOOKINGS);
    const data = sheet.getDataRange().getValues();
    const bookings = [];
    
    for (let i = 1; i < data.length; i++) {
      bookings.push({
        booking_id: String(data[i][0]),
        username: String(data[i][1]),
        fullname: String(data[i][2]),
        movie_title: String(data[i][3]),
        room_number: String(data[i][4]),
        date: String(data[i][5]),
        time: String(data[i][6]),
        status: String(data[i][7]),
        created_at: String(data[i][8])
      });
    }
    return bookings.reverse();
  } catch (error) {
    return [];
  }
}

function cancelBookingByUser(bookingId, username) {
  try {
    initDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TAB_BOOKINGS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(bookingId) && String(data[i][1]) === String(username)) {
        sheet.getRange(i + 1, 8).setValue('Canceled');

        // แจ้งเตือนไปยัง LINE
        const lineNotifyText = `❌ [แจ้งเตือนการยกเลิกจอง]\n` +
                               `🎟 รหัสจอง: ${bookingId}\n` +
                               `👤 ผู้ยกเลิก: @${username}\n` +
                               `สถานะ: Canceled 🚫`;
        sendLineMessage(lineNotifyText);

        return { success: true, message: 'ยกเลิกการจองเรียบร้อยแล้ว' };
      }
    }
    return { success: false, message: 'ไม่พบรายการจองนี้ หรือคุณไม่มีสิทธิ์ยกเลิก' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}
