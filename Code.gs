/**
 * 研究室輔導預約系統 - Google Apps Script 後端
 * 
 * 功能：
 * - Google OAuth 登入
 * - 預約申請（多人可申請同時段）
 * - 教師審核（確認/拒絕）
 * - Calendar 整合
 * - Email 通知
 */

// ========== 設定 ==========
const CONFIG = {
  TEACHER_EMAIL: 'scwu@gms.npu.edu.tw',
  CALENDAR_ID: 'scwu@gms.npu.edu.tw',  // 預約寫入的日曆（主日曆）
  
  // 要讀取忙碌時段的所有行事曆
  CALENDAR_IDS: [
    'scwu@gms.npu.edu.tw',  // 主日曆
    'pccu.edu.tw_fqd7jue2r0kt74pdlbot2gm4ns@group.calendar.google.com',  // 仕傑的工作行事曆
    '84fqqmqu8hcsimq8n6au7od8mg@group.calendar.google.com',  // 我的家庭真可愛
    'gms.npu.edu.tw_dktg99t2p2sqlvcoorpd67h7ok@group.calendar.google.com',  // 澎科大校園行事曆
    'dtlnpu@gms.npu.edu.tw',  // 觀休系共同(新)
    'win9363s@gms.npu.edu.tw',  // 觀休系共同行事曆
    'coralwu1215@gmail.com',  // 家人
    'wyattwu0409@gmail.com',  // 家人
    'evolymwu0810@gmail.com',  // 家人
  ],
  
  ALLOWED_DOMAIN: 'gms.npu.edu.tw',
  BOOKING_HOURS: { start: 8, end: 20 },  // 08:00-20:00
  DEFAULT_DURATION: 60,  // 預設 1 小時（分鐘）
  MIN_ADVANCE_MINUTES: 30,  // 至少提前 30 分鐘
  SPREADSHEET_ID: '19fe3PYmruja_kwqshUwplTMimvqJtfJ5CkeY5ukdyX4',
};

// ========== Web App 入口 ==========
function doGet(e) {
  const action = e.parameter.action;
  
  switch(action) {
    case 'getSlots':
      return jsonResponse(getAvailableSlots(e.parameter.date));
    case 'getMyBookings':
      return jsonResponse(getMyBookings(e.parameter.email));
    case 'getPendingRequests':
      return jsonResponse(getPendingRequests());
    default:
      return HtmlService.createHtmlOutputFromFile('index');
  }
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const action = data.action;
  
  switch(action) {
    case 'submitBooking':
      return jsonResponse(submitBooking(data));
    case 'confirmBooking':
      return jsonResponse(confirmBooking(data.bookingId));
    case 'rejectBooking':
      return jsonResponse(rejectBooking(data.bookingId, data.reason));
    case 'cancelBooking':
      return jsonResponse(cancelBooking(data.bookingId, data.email));
    default:
      return jsonResponse({ error: 'Unknown action' });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========== 預約相關 ==========

/**
 * 取得指定週的忙碌時段
 * 
 * ⚠️ 重要：此函數讀取的是 CONFIG.CALENDAR_ID (scwu@gms.npu.edu.tw) 的行事曆
 *    Web App 必須以「我」(Me) 身份執行，而不是「存取使用者」
 *    部署設定：Execute as → Me, Who has access → Anyone
 */
function getAvailableSlots(dateStr) {
  try {
    console.log('收到日期參數:', dateStr);
    
    // 解析日期（處理各種格式）
    let inputDate;
    if (!dateStr) {
      inputDate = new Date();
    } else {
      // 移除時區後綴，只取日期部分
      const dateOnly = dateStr.split('T')[0];
      inputDate = new Date(dateOnly + 'T00:00:00');
    }
    
    console.log('解析後日期:', inputDate);
    
    if (isNaN(inputDate.getTime())) {
      console.error('日期解析失敗:', dateStr);
      return { busySlots: [], pendingSlots: [], error: '日期格式錯誤: ' + dateStr };
    }
    
    // 取得該週的範圍
    const startOfWeek = new Date(inputDate);
    startOfWeek.setDate(inputDate.getDate() - inputDate.getDay());
    startOfWeek.setHours(0, 0, 0, 0);
    
    const endOfWeek = new Date(startOfWeek);
    endOfWeek.setDate(startOfWeek.getDate() + 7);
    endOfWeek.setHours(23, 59, 59, 999);
    
    console.log('查詢範圍:', startOfWeek, '~', endOfWeek);
    
    // 從所有行事曆取得忙碌時段
    var busySlots = [];
    var totalEvents = 0;
    var loadedCalendars = [];
    
    for (var i = 0; i < CONFIG.CALENDAR_IDS.length; i++) {
      var calendarId = CONFIG.CALENDAR_IDS[i];
      try {
        var calendar = CalendarApp.getCalendarById(calendarId);
        if (!calendar) {
          console.log('跳過無法存取的日曆: ' + calendarId);
          continue;
        }
        
        var calendarName = calendar.getName();
        var events = calendar.getEvents(startOfWeek, endOfWeek);
        console.log('從 ' + calendarName + ' 讀取 ' + events.length + ' 個行程');
        
        for (var j = 0; j < events.length; j++) {
          busySlots.push({
            start: events[j].getStartTime().toISOString(),
            end: events[j].getEndTime().toISOString(),
            title: '已佔用'
          });
        }
        
        totalEvents += events.length;
        loadedCalendars.push(calendarName);
      } catch(err) {
        console.error('讀取日曆失敗 ' + calendarId + ':', err);
      }
    }
    
    console.log('總共從 ' + loadedCalendars.length + ' 個日曆讀取 ' + totalEvents + ' 個行程');
    
    // 取得待審核的申請
    const pendingBookings = getPendingBookingsForWeek(startOfWeek, endOfWeek);
    
    return {
      date: dateStr,
      calendarId: CONFIG.CALENDAR_ID,  // 回傳日曆 ID 供前端確認
      busySlots: busySlots,
      pendingSlots: pendingBookings,
      bookingHours: CONFIG.BOOKING_HOURS
    };
  } catch(e) {
    console.error('getAvailableSlots 錯誤:', e);
    return { busySlots: [], pendingSlots: [], error: e.toString() };
  }
}

/**
 * 提交預約申請
 */
function submitBooking(data) {
  const { email, name, startTime, endTime, purpose } = data;
  
  // 驗證 Email（不限網域）
  if (!email || !email.includes('@')) {
    return { success: false, error: '請使用有效的 Email 登入' };
  }
  
  // 驗證預約事由（必填）
  if (!purpose || purpose.trim() === '') {
    return { success: false, error: '請填寫預約事由' };
  }
  
  // 驗證時間（至少提前 30 分鐘）
  const start = new Date(startTime);
  const now = new Date();
  const minBookingTime = new Date(now.getTime() + CONFIG.MIN_ADVANCE_MINUTES * 60000);
  
  if (start < minBookingTime) {
    return { success: false, error: '請至少提前 30 分鐘預約' };
  }
  
  // 建立預約記錄
  const sheet = getBookingSheet();
  const bookingId = Utilities.getUuid();
  
  sheet.appendRow([
    bookingId,
    email,
    name,
    startTime,
    endTime,
    purpose || '',
    'pending',  // status
    '',  // calendarEventId
    '',  // rejectReason
    false,  // reminderSent
    new Date().toISOString(),  // createdAt
    ''  // confirmedAt
  ]);
  
  // 發送申請確認 Email 給學生
  sendEmail(email, 
    '【輔導預約】申請已送出',
    `${name} 您好，\n\n您已申請以下輔導時段：\n\n日期時間：${formatDateTime(start)} ~ ${formatTime(new Date(endTime))}\n預約事由：${purpose || '未填寫'}\n\n請等待教師確認，確認結果將另行通知。\n\n研究室輔導預約系統`
  );
  
  // 發送通知給教師
  sendEmail(CONFIG.TEACHER_EMAIL,
    '【輔導預約】新的預約申請',
    `有新的輔導預約申請：\n\n學生：${name} (${email})\n日期時間：${formatDateTime(start)} ~ ${formatTime(new Date(endTime))}\n預約事由：${purpose || '未填寫'}\n\n請至系統確認或拒絕此申請。`
  );
  
  return { success: true, bookingId: bookingId };
}

/**
 * 教師確認預約
 */
function confirmBooking(bookingId) {
  const sheet = getBookingSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === bookingId) {
      const email = data[i][1];
      const name = data[i][2];
      const startTime = new Date(data[i][3]);
      const endTime = new Date(data[i][4]);
      const purpose = data[i][5];
      
      // 建立日曆事件
      const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
      const event = calendar.createEvent(
        `輔導預約：${name}`,
        startTime,
        endTime,
        {
          description: `學生：${name}\nEmail：${email}\n事由：${purpose}`,
          guests: email,
          sendInvites: true
        }
      );
      
      // 設定 10 分鐘前提醒
      event.addPopupReminder(10);
      event.addEmailReminder(10);
      
      // 更新狀態
      sheet.getRange(i + 1, 7).setValue('confirmed');
      sheet.getRange(i + 1, 8).setValue(event.getId());
      sheet.getRange(i + 1, 12).setValue(new Date().toISOString());
      
      // 發送確認 Email
      sendEmail(email,
        '【輔導預約】預約成功！',
        `${name} 您好，\n\n預約成功！請記得在 ${formatDateSimple(startTime)} ${formatTime(startTime)} 至 E719 研究室找吳仕傑老師聊天。\n\n系統將於開始前 10 分鐘發送提醒。\n\n研究室輔導預約系統`
      );
      
      // 拒絕同時段的其他申請
      rejectOtherPendingBookings(bookingId, startTime, endTime);
      
      return { success: true };
    }
  }
  
  return { success: false, error: '找不到此預約' };
}

/**
 * 教師拒絕預約
 */
function rejectBooking(bookingId, reason) {
  const sheet = getBookingSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === bookingId) {
      const email = data[i][1];
      const name = data[i][2];
      const startTime = new Date(data[i][3]);
      
      sheet.getRange(i + 1, 7).setValue('rejected');
      sheet.getRange(i + 1, 9).setValue(reason || '');
      
      sendEmail(email,
        '【輔導預約】預約未成功',
        `${name} 您好，\n\n預約未成功，請重新預約其他時段。\n\n研究室輔導預約系統`
      );
      
      return { success: true };
    }
  }
  
  return { success: false, error: '找不到此預約' };
}

/**
 * 學生取消預約
 */
function cancelBooking(bookingId, email) {
  const sheet = getBookingSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === bookingId && data[i][1] === email) {
      const status = data[i][6];
      const eventId = data[i][7];
      
      // 如果已確認，刪除日曆事件
      if (status === 'confirmed' && eventId) {
        try {
          const calendar = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
          const event = calendar.getEventById(eventId);
          if (event) event.deleteEvent();
        } catch(e) {}
      }
      
      sheet.getRange(i + 1, 7).setValue('cancelled');
      return { success: true };
    }
  }
  
  return { success: false, error: '無法取消此預約' };
}

/**
 * 取得我的預約
 */
function getMyBookings(email) {
  const sheet = getBookingSheet();
  const data = sheet.getDataRange().getValues();
  const bookings = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === email && data[i][6] !== 'cancelled') {
      bookings.push({
        id: data[i][0],
        startTime: data[i][3],
        endTime: data[i][4],
        purpose: data[i][5],
        status: data[i][6]
      });
    }
  }
  
  return bookings;
}

/**
 * 取得待審核申請（教師用）
 */
function getPendingRequests() {
  const sheet = getBookingSheet();
  const data = sheet.getDataRange().getValues();
  const pending = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] === 'pending') {
      pending.push({
        id: data[i][0],
        email: data[i][1],
        name: data[i][2],
        startTime: data[i][3],
        endTime: data[i][4],
        purpose: data[i][5],
        createdAt: data[i][10]
      });
    }
  }
  
  return pending;
}

// ========== 輔助函數 ==========

function getBookingSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Bookings');
  
  if (!sheet) {
    sheet = ss.insertSheet('Bookings');
    sheet.appendRow([
      'id', 'email', 'name', 'startTime', 'endTime', 
      'purpose', 'status', 'calendarEventId', 'rejectReason',
      'reminderSent', 'createdAt', 'confirmedAt'
    ]);
  }
  
  return sheet;
}

function getPendingBookingsForWeek(startOfWeek, endOfWeek) {
  try {
    const sheet = getBookingSheet();
    const data = sheet.getDataRange().getValues();
    const pending = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][6] === 'pending') {
        const start = new Date(data[i][3]);
        if (start >= startOfWeek && start <= endOfWeek) {
          pending.push({
            start: data[i][3],
            end: data[i][4],
            count: 1
          });
        }
      }
    }
    
    return pending;
  } catch(e) {
    console.error('getPendingBookingsForWeek 錯誤:', e);
    return [];
  }
}

function rejectOtherPendingBookings(confirmedId, startTime, endTime) {
  const sheet = getBookingSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== confirmedId && data[i][6] === 'pending') {
      const bookingStart = new Date(data[i][3]);
      const bookingEnd = new Date(data[i][4]);
      
      // 檢查時段是否重疊
      if (bookingStart < endTime && bookingEnd > startTime) {
        rejectBooking(data[i][0], '該時段已被其他同學預約');
      }
    }
  }
}

function sendEmail(to, subject, body) {
  try {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: body
    });
  } catch(e) {
    console.error('Email 發送失敗:', e);
  }
}

function formatDateTime(date) {
  return Utilities.formatDate(date, 'Asia/Taipei', 'yyyy/MM/dd (E) HH:mm');
}

function formatDateSimple(date) {
  return Utilities.formatDate(date, 'Asia/Taipei', 'M月d日');
}

function formatTime(date) {
  return Utilities.formatDate(date, 'Asia/Taipei', 'HH:mm');
}

// ========== 定時提醒 ==========

/**
 * 設定觸發器：每分鐘檢查即將開始的預約
 * 在 Apps Script 編輯器中執行一次 setupTrigger() 來建立
 */
function setupTrigger() {
  ScriptApp.newTrigger('checkUpcomingBookings')
    .timeBased()
    .everyMinutes(1)
    .create();
}

function checkUpcomingBookings() {
  const sheet = getBookingSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const tenMinutesLater = new Date(now.getTime() + 10 * 60000);
  const elevenMinutesLater = new Date(now.getTime() + 11 * 60000);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] === 'confirmed' && !data[i][9]) {  // reminderSent = false
      const startTime = new Date(data[i][3]);
      
      // 在開始前 10-11 分鐘之間發送提醒
      if (startTime >= tenMinutesLater && startTime < elevenMinutesLater) {
        const email = data[i][1];
        const name = data[i][2];
        
        // 提醒學生
        sendEmail(email,
          '【輔導預約】10 分鐘後開始！',
          `${name} 您好，\n\n提醒您：輔導預約將於 10 分鐘後開始。\n\n時間：${formatDateTime(startTime)}\n地點：研究室 E719\n\n請準時前往！`
        );
        
        // 提醒教師
        sendEmail(CONFIG.TEACHER_EMAIL,
          '【輔導預約】10 分鐘後有學生來訪',
          `提醒您：\n\n學生 ${name} 預約的輔導將於 10 分鐘後開始。\n\n時間：${formatDateTime(startTime)}`
        );
        
        // 標記已發送
        sheet.getRange(i + 1, 10).setValue(true);
      }
    }
  }
}
