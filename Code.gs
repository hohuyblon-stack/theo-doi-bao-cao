/**
 * ============================================================================
 * HỆ THỐNG THEO DÕI BÁO CÁO - PHIÊN BẢN 2.2
 * ============================================================================
 *
 * TÍNH NĂNG:
 * 1. Quét Gmail tự động, nhận diện file đính kèm từ BẤT KỲ AI
 * 2. Tự động thêm người gửi mới vào Danh Sách Nhân Viên
 * 3. Phát hiện file trùng lặp (chống đối phó nộp bản cũ)
 * 4. Deadline BC TUẦN: NV (23:59 Thứ 2), TP (12:00 trưa Thứ 2)
 * 5. Deadline BC THÁNG: NV KD (23:59 mùng 3), TP KH+KPI (23:59 ngày 28)
 * 6. Nhắc nhở tự động: T2 nhắc NV+TP, T3 nhắc trễ hạn
 * 7. Nhắc BC tháng: Mùng 1 nhắc NV KD, Ngày 26 nhắc TP
 * 8. Báo cáo tổng hợp gửi sếp hàng tuần (Thứ 4 sáng)
 *
 * CẤU TRÚC SHEET "Nhật Ký Báo Cáo":
 *   STT | Email | Họ Tên | Phòng Ban | Loại BC | Thời Gian Nộp | Hạn Nộp | Trạng Thái | Link File | GmailMsgId(ẩn) | FileHash(ẩn)
 *
 * HƯỚNG DẪN:
 * 1. Cập nhật CONFIG bên dưới
 * 2. Chạy migrateToV2() MỘT LẦN
 * 3. Chạy setupTriggers() một lần
 * 4. Hệ thống tự chạy hoàn toàn
 */

// ============================================================================
// CẤU HÌNH HỆ THỐNG
// ============================================================================
const CONFIG = {
  SHEET_ID: '1H6HL8Gvx27Jo5x9K_jRDhmC_l_QTQKE5bLA4IPuXMNI',
  DRIVE_FOLDER_ID: '',
  GMAIL_LABEL: 'BAOCAO',
  SEARCH_DAYS: 30,
  TIMEZONE: 'Asia/Ho_Chi_Minh',

  // ── DEADLINE TUẦN ──
  // NV (tất cả): 12h đêm Thứ 2 (= 23:59 Thứ 2)
  // Trưởng phòng: 12h trưa Thứ 2 (= 12:00 Thứ 2 tuần kế tiếp)
  DEADLINE_NV: { day: 1, hour: 23, minute: 59 },   // day 1 = Thứ 2, 23:59
  DEADLINE_TP: { day: 1, hour: 12, minute: 0 },     // day 1 = Thứ 2, 12:00 trưa

  // ── DEADLINE THÁNG ──
  // NV KD: Báo cáo tháng - hạn mùng 3 lúc 23:59
  // TP: Kế hoạch tháng + KPI - hạn ngày 28 lúc 23:59
  DEADLINE_NV_MONTHLY: { day: 3, hour: 23, minute: 59 },
  DEADLINE_TP_MONTHLY: { day: 28, hour: 23, minute: 59 },

  // Tên phòng ban Kinh Doanh (dùng để xác định NV KD từ cột PhongBan)
  PHONG_KD_NAMES: ['Kinh doanh', 'KD', 'Sales', 'Kinh Doanh', 'KINH DOANH'],

  // ── TRƯỞNG PHÒNG (deadline 12:00 trưa Thứ 2) ──
  // MAST PRO (Hà Nội)
  // MAST SG: GĐ, KTT, TBP KD, TBP Điều vận, HCNS, TP Kho đều deadline 12h TRƯA T2
  TRUONG_PHONG_EMAILS: [
    // ── MAST PRO ──
    'linhdt@mast.com.vn',         // Đinh Tùng Lĩnh - TP Kho
    'tunv@mast.com.vn',           // Nguyễn Văn Tú  - TP Điều vận
    'chienmt@mast.com.vn',        // Mai Thị Chiên  - TP Kế toán
    'hoangvd@mast.com.vn',        // Vũ Đức Hoàng   - TP Nhập mua
    'luongpt@mast.com.vn',        // Phạm Thế Lượng - TP Kinh doanh
    'khanhxuan@mast.com.vn',      // Nguyễn Khánh Xuân - TP HCNS

    // ── MAST SG ──
    'nhanntt@mastsaigon.com.vn',  // Ngô Thị Thanh Nhàn - GĐ
    'anhntt@mastsaigon.com.vn',   // Nguyễn Thị Tuyết Anh - KTT
    'sonpt@mast.com.vn',          // Phạm Thanh Sơn - TBP KD
    'toanlv@mastsaigon.com.vn',   // Lê Văn Toàn - TP Điều vận
    'nhansu@mastsaigon.com.vn',   // Nguyễn Diễm My - NV HCNS - deadline 12h trưa T2
    'baond@mastsaigon.com.vn'     // Nguyễn Đình Bảo - TP Kho
  ],

  // Sheet names
  SHEET_EMPLOYEES: 'Danh Sách Nhân Viên',
  SHEET_LOG: 'Nhật Ký Báo Cáo',

  // Cột trong "Danh Sách Nhân Viên" (0-indexed)
  COL_EMAIL: 0,
  COL_HOTEN: 1,
  COL_PHONGBAN: 2,

  // File extensions hợp lệ
  VALID_EXTENSIONS: ['.xlsx', '.xls', '.pdf', '.doc', '.docx', '.ppt', '.pptx', '.csv'],

  // Trạng thái
  STATUS: {
    ON_TIME: 'Đúng hạn',
    LATE: 'Trễ hạn',
    NO_ATTACHMENT: 'Thiếu file đính kèm',
    DUPLICATE_FILE: 'File trùng lặp'
  },

  // Email nhắc nhở & báo cáo
  HR_EMAIL: 'nhansu@mast.com.vn',
  MANAGER_EMAILS: ['nhansu@mast.com.vn'],
  COMPANY_NAME: 'MAST PRO',

  // ── DEADLINE BC THÁNG ĐẶC BIỆT (override theo email cụ thể) ──
  // Key: email, Value: {day, hour, minute}
  // Dùng khi deadline tháng khác với DEADLINE_NV_MONTHLY hoặc DEADLINE_TP_MONTHLY
  CUSTOM_MONTHLY_DEADLINES: {
    'sonpt@mast.com.vn':             { day: 5,  hour: 23, minute: 59 }, // TBP KD SG: ngày 5
    'nhanntt@mastsaigon.com.vn':     { day: 28, hour: 23, minute: 59 }, // GĐ SG: BC KH tháng - ngày 28
  },

  // ── DEADLINE BC TỔNG KẾT (riêng cho GĐ SG) ──
  // nhanntt@mastsaigon.com.vn: BC Tổng kết hoạt động tháng trước - ngày 9 lúc 23:59
  SUMMARY_REPORT_DEADLINE: { day: 9, hour: 23, minute: 59 },
  SUMMARY_REPORT_EMAILS: ['nhanntt@mastsaigon.com.vn'],

  // Danh sách email LOẠI TRỪ (không theo dõi)
  EXCLUDED_EMAILS: [
    // MAST PRO
    'nhansu@mast.com.vn',
    // MAST SG - không có BC tuần/tháng
    'tiendmt@mastsaigon.com.vn',  // Đào Minh Thủy Tiên - NVKT (không có BC tuần/tháng)
    'ketoan@mastsaigon.com.vn',   // Phạm Ngọc Kim Ngân - NVKT (mail kế toán SG, không có BC)
    'trungtt@mastsaigon.com.vn',  // Trần Trịnh Tuấn Trung - VC (không có BC)
    'sint@mast.com.vn',           // Nguyễn Thành Sĩ - NV KD (không có BC)
    'conghienn.lian@gmail.com',   // Võ Công Hiển (mail cũ gmail - đã đổi sang hienvc@mastsaigon.com.vn)
    // Pattern chung
    'noreply@',
    'notification@',
    'hoadondientu',
    'steveshim@gmckorea.net',
    'thongbao@baohiemxahoi.gov.vn'
  ]
};

// ============================================================================
// MIGRATION V2 - CHẠY MỘT LẦN
// ============================================================================

/**
 * ★★★ CHẠY HÀM NÀY MỘT LẦN ĐỂ NÂNG CẤP LÊN V2 ★★★
 * - Backup sheet cũ
 * - Tạo sheet mới với cấu trúc 8 cột (thêm FileHash)
 * - Giữ nguyên dữ liệu Danh Sách Nhân Viên
 */
function migrateToV2() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const oldSheet = ss.getSheetByName(CONFIG.SHEET_LOG);

    if (oldSheet) {
      const backupName = CONFIG.SHEET_LOG + ' (Backup V1 ' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd-MM-yyyy HH-mm') + ')';
      oldSheet.setName(backupName);
      Logger.log('✓ Backup → "' + backupName + '"');
    }

    createFreshSheetV2_(ss);

    // Đảm bảo sheet nhân viên tồn tại
    let empSheet = ss.getSheetByName(CONFIG.SHEET_EMPLOYEES);
    if (!empSheet) {
      empSheet = ss.insertSheet(CONFIG.SHEET_EMPLOYEES);
      empSheet.getRange(1, 1, 1, 3).setValues([['Email', 'HoTen', 'PhongBan']]);
      empSheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
    }

    Logger.log('★★★ NÂNG CẤP V2 HOÀN THÀNH ★★★');
    Logger.log('Tiếp theo: chạy setupTriggers() rồi rescanAllEmails()');
  } catch (error) {
    Logger.log('✗ LỖI: ' + error.message);
    Logger.log(error.stack);
  }
}

function createFreshSheetV2_(ss) {
  const existing = ss.getSheetByName(CONFIG.SHEET_LOG);
  if (existing) ss.deleteSheet(existing);

  const sheet = ss.insertSheet(CONFIG.SHEET_LOG);
  const headers = [
    'STT', 'Email', 'Họ Tên', 'Phòng Ban', 'Loại BC',
    'Thời Gian Nộp', 'Hạn Nộp', 'Trạng Thái', 'Link File',
    'GmailMsgId', 'FileHash'
  ];

  const currentCols = sheet.getMaxColumns();
  if (currentCols < headers.length) {
    sheet.insertColumnsAfter(currentCols, headers.length - currentCols);
  } else if (currentCols > headers.length) {
    sheet.deleteColumns(headers.length + 1, currentCols - headers.length);
  }

  // Header row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('white')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(10)
    .setFontFamily('Arial')
    .setBorder(true, true, true, true, true, true, '#1a73e8', SpreadsheetApp.BorderStyle.SOLID);

  // Đặt chiều cao header
  sheet.setRowHeight(1, 36);

  sheet.setFrozenRows(1);
  sheet.hideColumns(10, 2); // Ẩn GmailMsgId + FileHash

  // Đặt chiều rộng cột cho dễ nhìn
  sheet.setColumnWidth(1, 50);   // STT
  sheet.setColumnWidth(2, 200);  // Email
  sheet.setColumnWidth(3, 160);  // Họ Tên
  sheet.setColumnWidth(4, 130);  // Phòng Ban
  sheet.setColumnWidth(5, 110);  // Loại BC
  sheet.setColumnWidth(6, 160);  // Thời Gian Nộp
  sheet.setColumnWidth(7, 160);  // Hạn Nộp
  sheet.setColumnWidth(8, 130);  // Trạng Thái
  sheet.setColumnWidth(9, 300);  // Link File

  // Căn giữa cột STT
  sheet.getRange('A:A').setHorizontalAlignment('center');
  // Căn giữa cột Trạng Thái
  sheet.getRange('H:H').setHorizontalAlignment('center');
  // Căn giữa cột Loại BC
  sheet.getRange('E:E').setHorizontalAlignment('center');

  Logger.log('✓ Tạo sheet V2.2: ' + CONFIG.SHEET_LOG);
  return sheet;
}

// ============================================================================
// TRIGGERS - CHẠY MỘT LẦN
// ============================================================================

function setupTriggers() {
  // Xóa trigger cũ
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // 1. Quét email mỗi 10 phút
  ScriptApp.newTrigger('processEmails')
    .timeBased().everyMinutes(10).create();

  // 2. Thứ 2 sáng 9h: Nhắc NV (deadline 23:59 tối nay) + TP (deadline 12:00 trưa nay)
  ScriptApp.newTrigger('sendMondayReminder')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9).create();

  // 3. Thứ 3 sáng 9h: Nhắc mạnh ai trễ hạn BC tuần (NV + TP)
  ScriptApp.newTrigger('sendTuesdayLateReminder')
    .timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .atHour(9).create();

  // 4. Báo cáo tổng hợp tuần: Thứ 4 lúc 8h sáng (buffer 1 ngày sau deadline T2)
  ScriptApp.newTrigger('sendWeeklyReport')
    .timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
    .atHour(8).create();

  // 5. Mùng 1 hàng tháng: Nhắc NV KD nộp BC tháng (deadline mùng 3)
  ScriptApp.newTrigger('sendMonthlyReminder_NV')
    .timeBased().onMonthDay(1)
    .atHour(9).create();

  // 6. Ngày 26 hàng tháng: Nhắc TP nộp KH + KPI (deadline 28)
  ScriptApp.newTrigger('sendMonthlyReminder_TP')
    .timeBased().onMonthDay(26)
    .atHour(9).create();

  // 7. Ngày 7 hàng tháng: Nhắc GĐ SG nộp BC Tổng kết (deadline ngày 9)
  ScriptApp.newTrigger('sendMonthlySummaryReminder_GD')
    .timeBased().onMonthDay(7)
    .atHour(9).create();

  Logger.log('✓ 7 triggers đã thiết lập');
  ScriptApp.getProjectTriggers().forEach(t =>
    Logger.log('  → ' + t.getHandlerFunction() + ' | ' + t.getTriggerSource())
  );
}

// ============================================================================
// HÀM TIỆN ÍCH
// ============================================================================

function formatDateTime(date) {
  if (!date) return '';
  if (date instanceof Date && !isNaN(date.getTime())) {
    return Utilities.formatDate(date, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
  }
  return date.toString();
}

function formatDateShort(date) {
  if (!date || !(date instanceof Date)) return '';
  return Utilities.formatDate(date, CONFIG.TIMEZONE, 'dd/MM');
}

function isValidFile(filename) {
  if (!filename) return false;
  const lower = filename.toLowerCase();
  // Bỏ qua file ảnh nhỏ (signature, icon)
  if (lower.match(/^image\d*\.(png|jpg|jpeg|gif)$/)) return false;
  return CONFIG.VALID_EXTENSIONS.some(ext => lower.endsWith(ext));
}

function extractEmail(fromField) {
  if (!fromField) return '';
  const match = fromField.match(/<(.+?)>/);
  return match ? match[1].toLowerCase().trim() : fromField.toLowerCase().trim();
}

function extractSenderName(fromField) {
  if (!fromField) return '';
  const match = fromField.match(/^"?(.+?)"?\s*<.+?>$/);
  return match ? match[1].trim() : fromField.split('@')[0];
}

function isExcludedEmail(email) {
  if (!email) return true;
  return CONFIG.EXCLUDED_EMAILS.some(exc => email.includes(exc));
}

/**
 * Kiểm tra email có phải Trưởng phòng không
 */
function isTruongPhong(email) {
  if (!email) return false;
  return CONFIG.TRUONG_PHONG_EMAILS.includes(email.toLowerCase().trim());
}

/**
 * Kiểm tra NV có thuộc phòng Kinh Doanh không (qua cột PhongBan trong sheet)
 */
function isNhanVienKD(email) {
  if (!email) return false;
  const employees = getActiveEmployees();
  const emp = employees.get(email.toLowerCase().trim());
  if (!emp) return false;
  const pb = (emp.phongBan || '').trim();
  return CONFIG.PHONG_KD_NAMES.some(name => pb.toLowerCase() === name.toLowerCase());
}

/**
 * Tính deadline THÁNG cho NV KD (mùng 3, 23:59)
 * Nếu email có trong CUSTOM_MONTHLY_DEADLINES → dùng deadline riêng
 * @param {Date} date - Ngày gửi email
 * @param {string} [email] - Email người gửi (để kiểm tra custom deadline)
 * @returns {Date} deadline tháng hiện tại hoặc tháng sau
 */
function getMonthlyDeadlineNV(date, email) {
  const d = (date instanceof Date && !isNaN(date.getTime())) ? new Date(date) : new Date();

  // Kiểm tra xem email có deadline tháng riêng không
  let deadlineCfg = CONFIG.DEADLINE_NV_MONTHLY;
  if (email) {
    const emailLower = email.toLowerCase().trim();
    if (CONFIG.CUSTOM_MONTHLY_DEADLINES[emailLower]) {
      deadlineCfg = CONFIG.CUSTOM_MONTHLY_DEADLINES[emailLower];
    }
  }

  const deadlineDay = deadlineCfg.day;
  const deadlineHour = deadlineCfg.hour;
  const deadlineMinute = deadlineCfg.minute;

  let target = new Date(d.getFullYear(), d.getMonth(), deadlineDay, deadlineHour, deadlineMinute, 0, 0);
  if (d.getTime() > target.getTime()) {
    target = new Date(d.getFullYear(), d.getMonth() + 1, deadlineDay, deadlineHour, deadlineMinute, 0, 0);
  }
  return target;
}

/**
 * Tính deadline THÁNG cho TP (ngày 28, 23:59)
 * @param {Date} date - Ngày gửi email
 * @returns {Date} deadline ngày 28 tháng hiện tại hoặc tháng sau
 */
function getMonthlyDeadlineTP(date) {
  const d = (date instanceof Date && !isNaN(date.getTime())) ? new Date(date) : new Date();
  const deadlineDay = CONFIG.DEADLINE_TP_MONTHLY.day;
  const deadlineHour = CONFIG.DEADLINE_TP_MONTHLY.hour;
  const deadlineMinute = CONFIG.DEADLINE_TP_MONTHLY.minute;

  let target = new Date(d.getFullYear(), d.getMonth(), deadlineDay, deadlineHour, deadlineMinute, 0, 0);
  if (d.getTime() > target.getTime()) {
    target = new Date(d.getFullYear(), d.getMonth() + 1, deadlineDay, deadlineHour, deadlineMinute, 0, 0);
  }
  return target;
}

/**
 * Tính deadline TUẦN theo loại nhân viên:
 *  - NV (tất cả):  23:59 Thứ 2 (12h đêm Thứ 2)
 *  - Trưởng phòng: 12:00 trưa Thứ 2 (12h trưa Thứ 2 tuần kế tiếp)
 * @param {Date} date - Ngày gửi email
 * @param {string} email - Email người gửi (để xác định NV hay TP)
 */
function getDeadlineForEmail(date, email) {
  const deadlineConfig = isTruongPhong(email) ? CONFIG.DEADLINE_TP : CONFIG.DEADLINE_NV;
  return _calcDeadline(date, deadlineConfig);
}

/**
 * Backward-compatible: deadline Thứ 6 cho NV thường
 */
function getFridayDeadline(date) {
  return _calcDeadline(date, CONFIG.DEADLINE_NV);
}

/**
 * Tính deadline TUẦN chính xác theo config {day, hour, minute}
 * Logic: deadline là Thứ 2 (day=1)
 *   - Nếu email gửi vào T3 đến CN → deadline = T2 tuần kế tiếp
 *   - Nếu email gửi vào T2 → deadline = T2 cùng ngày (nếu trước giờ deadline) hoặc T2 tuần sau (nếu sau giờ deadline)
 */
function _calcDeadline(date, deadlineConfig) {
  let d;
  if (date instanceof Date && !isNaN(date.getTime())) {
    d = new Date(date);
  } else {
    d = new Date();
  }

  const day = d.getDay(); // 0=CN, 1=T2, 2=T3, ..., 6=T7
  const targetDay = deadlineConfig.day; // 1 = Thứ 2
  let diff;

  if (day === targetDay) {
    // Cùng ngày với deadline (T2) → check giờ: nếu trước deadline thì hôm nay, sau deadline thì T2 tuần sau
    const deadlineToday = new Date(d);
    deadlineToday.setHours(deadlineConfig.hour, deadlineConfig.minute, 0, 0);
    diff = (d.getTime() <= deadlineToday.getTime()) ? 0 : 7;
  } else if (day === 0) {
    // Chủ nhật → deadline T2 ngày mai
    diff = 1;
  } else if (day > targetDay) {
    // T3(2), T4(3), T5(4), T6(5), T7(6) → deadline T2 tuần kế tiếp
    diff = 7 - day + targetDay; // VD: T3(2) → 7-2+1=6 ngày nữa, T7(6) → 7-6+1=2 ngày nữa
  } else {
    // day < targetDay (không xảy ra vì targetDay=1, chỉ day=0 mà đã xử lý ở trên)
    diff = targetDay - day;
  }

  const target = new Date(d);
  target.setDate(d.getDate() + diff);
  target.setHours(deadlineConfig.hour, deadlineConfig.minute, 0, 0);
  return target;
}

/**
 * Tính hash MD5 của file đính kèm để phát hiện trùng lặp
 */
function computeFileHash(attachment) {
  try {
    const bytes = attachment.copyBlob().getBytes();
    const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, bytes);
    return rawHash.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
  } catch (e) {
    return '';
  }
}

// ============================================================================
// TRUY CẬP DỮ LIỆU
// ============================================================================

function getActiveEmployees() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_EMPLOYEES);
    if (!sheet) return new Map();

    const data = sheet.getDataRange().getValues();
    const employees = new Map();
    for (let i = 1; i < data.length; i++) {
      const email = (data[i][CONFIG.COL_EMAIL] || '').toString().toLowerCase().trim();
      if (email) {
        employees.set(email, {
          email: email,
          hoTen: (data[i][CONFIG.COL_HOTEN] || '').toString().trim(),
          phongBan: (data[i][CONFIG.COL_PHONGBAN] || '').toString().trim()
        });
      }
    }
    return employees;
  } catch (error) {
    Logger.log('✗ Lỗi getActiveEmployees: ' + error.message);
    return new Map();
  }
}

/**
 * TỰ ĐỘNG thêm nhân viên mới vào danh sách
 */
function autoAddEmployee(email, senderName) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_EMPLOYEES);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_EMPLOYEES);
      sheet.getRange(1, 1, 1, 3).setValues([['Email', 'HoTen', 'PhongBan']]);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
    }
    sheet.appendRow([email, senderName, 'MỚI_TỰ_ĐỘNG']);
    Logger.log('✓ Tự động thêm nhân viên: ' + email + ' (' + senderName + ')');
  } catch (error) {
    Logger.log('✗ Lỗi autoAddEmployee: ' + error.message);
  }
}

function getLoggedMessageIds() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_LOG);
    if (!sheet || sheet.getLastRow() < 2) return new Set();

    const data = sheet.getDataRange().getValues();
    const msgIds = new Set();
    for (let i = 1; i < data.length; i++) {
      // Cột 10 (index 9) = GmailMsgId
      const val = (data[i][9] || '').toString().trim();
      if (val) msgIds.add(val);
    }
    return msgIds;
  } catch (error) {
    return new Set();
  }
}

/**
 * Lấy tất cả FileHash đã lưu (để phát hiện trùng lặp)
 * Trả về Map: hash → {email, submitTime}
 */
function getExistingFileHashes() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_LOG);
    if (!sheet || sheet.getLastRow() < 2) return new Map();

    const data = sheet.getDataRange().getValues();
    const hashes = new Map();
    for (let i = 1; i < data.length; i++) {
      // Cấu trúc mới: STT(0), Email(1), HoTen(2), PhongBan(3), LoaiBC(4),
      //               ThoiGianNop(5), HanNop(6), TrangThai(7), LinkFile(8), MsgId(9), FileHash(10)
      const hashField = (data[i][10] || '').toString().trim();
      const email = (data[i][1] || '').toString().trim();
      const time = (data[i][5] || '').toString().trim();
      const status = (data[i][7] || '').toString().trim();
      if (hashField && status !== CONFIG.STATUS.DUPLICATE_FILE) {
        // Hỗ trợ multi-hash (phân tách bằng dấu phẩy)
        hashField.split(',').forEach(function(h) {
          h = h.trim();
          if (h) hashes.set(h, { email: email, submitTime: time });
        });
      }
    }
    return hashes;
  } catch (error) {
    return new Map();
  }
}

// ============================================================================
// GOOGLE DRIVE
// ============================================================================

function getOrCreateRootFolder() {
  if (CONFIG.DRIVE_FOLDER_ID) {
    return DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  }
  const rootFolders = DriveApp.getRootFolder().getFoldersByName('BAOCAO_TUAN');
  if (rootFolders.hasNext()) return rootFolders.next();
  return DriveApp.getRootFolder().createFolder('BAOCAO_TUAN');
}

function saveAttachmentToDrive(attachment, senderName) {
  try {
    const rootFolder = getOrCreateRootFolder();
    const monthKey = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM');

    let targetFolder;
    const subFolders = rootFolder.getFoldersByName(monthKey);
    if (subFolders.hasNext()) {
      targetFolder = subFolders.next();
    } else {
      targetFolder = rootFolder.createFolder(monthKey);
    }

    const originalName = attachment.getName();
    let newName = originalName;
    if (senderName) {
      const safeName = senderName.replace(/[^\w\sÀ-ỹ]/gi, '').trim();
      if (safeName) {
        const dotIdx = originalName.lastIndexOf('.');
        if (dotIdx > 0) {
          newName = safeName + '_' + originalName.substring(0, dotIdx) + originalName.substring(dotIdx);
        }
      }
    }

    const blob = attachment.copyBlob();
    blob.setName(newName);
    return targetFolder.createFile(blob).getUrl();
  } catch (error) {
    Logger.log('✗ Lỗi lưu file: ' + error.message);
    return '';
  }
}

// ============================================================================
// XỬ LÝ EMAIL CHÍNH
// ============================================================================

/**
 * ★ HÀM CHÍNH - Trigger gọi mỗi 10 phút ★
 * Quét Gmail, xử lý BẤT KỲ email nào có đính kèm file hợp lệ.
 * Tự động thêm người gửi mới vào danh sách nhân viên.
 * Phát hiện file trùng lặp.
 */
function processEmails() {
  Logger.log('=== QUÉT EMAIL: ' + formatDateTime(new Date()) + ' ===');

  try {
    const query = 'has:attachment newer_than:' + CONFIG.SEARCH_DAYS + 'd to:' + CONFIG.HR_EMAIL;
    const threads = GmailApp.search(query);

    if (threads.length === 0) {
      Logger.log('Không có email mới');
      return;
    }

    const employees = getActiveEmployees();
    const loggedMsgIds = getLoggedMessageIds();
    const existingHashes = getExistingFileHashes();
    let processed = 0, skipped = 0, errors = 0, newEmployees = 0;

    for (const thread of threads) {
      for (const message of thread.getMessages()) {
        try {
          const msgId = message.getId();
          if (loggedMsgIds.has(msgId)) continue;

          const senderEmail = extractEmail(message.getFrom());
          if (isExcludedEmail(senderEmail)) { skipped++; continue; }

          // Kiểm tra có file đính kèm hợp lệ không
          const attachments = message.getAttachments();
          const validFiles = attachments.filter(att => isValidFile(att.getName()));
          if (validFiles.length === 0) { skipped++; continue; }

          // Tự động thêm nhân viên mới nếu chưa có
          let employee = employees.get(senderEmail);
          if (!employee) {
            const senderName = extractSenderName(message.getFrom());
            autoAddEmployee(senderEmail, senderName);
            employee = { email: senderEmail, hoTen: senderName, phongBan: 'MỚI_TỰ_ĐỘNG' };
            employees.set(senderEmail, employee);
            newEmployees++;
          }

          const result = processSingleEmail(message, employee, existingHashes);
          if (result) {
            writeLogEntry(result);
            loggedMsgIds.add(msgId);
            if (result.fileHash) {
              existingHashes.set(result.fileHash, { email: result.email, submitTime: formatDateTime(result.submitTime) });
            }
            processed++;
            message.markRead();

            try {
              let label = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL);
              if (!label) label = GmailApp.createLabel(CONFIG.GMAIL_LABEL);
              thread.addLabel(label);
            } catch (e) { /* bỏ qua */ }
          }
        } catch (emailError) {
          errors++;
          Logger.log('✗ Lỗi email: ' + emailError.message);
        }
      }
    }

    Logger.log('=== KẾT QUẢ: ' + processed + ' xử lý | ' + skipped + ' bỏ qua | ' + errors + ' lỗi | ' + newEmployees + ' NV mới ===');
  } catch (error) {
    Logger.log('✗ LỖI HỆ THỐNG: ' + error.message);
  }
}

/**
 * Xử lý một email - có kiểm tra file trùng lặp
 */
function processSingleEmail(message, employee, existingHashes) {
  const senderEmail = employee.email;
  const senderName = employee.hoTen || extractSenderName(message.getFrom());
  const submitTime = message.getDate();
  const msgId = message.getId();
  const deadline = getDeadlineForEmail(submitTime, senderEmail);

  const attachments = message.getAttachments();
  const validFiles = attachments.filter(att => isValidFile(att.getName()));

  if (validFiles.length === 0) {
    return {
      email: senderEmail,
      senderName: senderName,
      submitTime: submitTime,
      deadline: deadline,
      status: CONFIG.STATUS.NO_ATTACHMENT,
      fileLink: '',
      gmailMsgId: msgId,
      fileHash: '',
      reportType: 'BC Tuần'
    };
  }

  // Kiểm tra trùng lặp + lưu file
  const fileLinks = [];
  let isDuplicate = false;
  let duplicateInfo = '';
  const allHashes = [];

  for (const file of validFiles) {
    const hash = computeFileHash(file);
    if (hash) allHashes.push(hash);

    // Kiểm tra trùng lặp
    if (hash && existingHashes.has(hash)) {
      const prev = existingHashes.get(hash);
      isDuplicate = true;
      duplicateInfo = 'Trùng với ' + prev.email + ' lúc ' + prev.submitTime;
      Logger.log('⚠ FILE TRÙNG: ' + senderEmail + ' | ' + file.getName() + ' | ' + duplicateInfo);
    }

    const link = saveAttachmentToDrive(file, senderName);
    if (link) fileLinks.push(link);
  }

  let status;
  if (isDuplicate) {
    status = CONFIG.STATUS.DUPLICATE_FILE;
  } else {
    const isLate = submitTime.getTime() > deadline.getTime();
    status = isLate ? CONFIG.STATUS.LATE : CONFIG.STATUS.ON_TIME;
  }

  return {
    email: senderEmail,
    senderName: senderName,
    submitTime: submitTime,
    deadline: deadline,
    status: status,
    fileLink: fileLinks.length > 0 ? fileLinks[0] : '',
    gmailMsgId: msgId,
    fileHash: allHashes.join(','),
    reportType: 'BC Tuần'
  };
}

// ============================================================================
// GHI DỮ LIỆU VÀO SHEET
// ============================================================================

function writeLogEntry(entry) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_LOG);
    if (!sheet) {
      Logger.log('✗ Sheet không tồn tại!');
      return;
    }

    // Lấy thông tin phòng ban từ danh sách nhân viên
    const employees = getActiveEmployees();
    const emp = employees.get(entry.email.toLowerCase().trim());
    const phongBan = emp ? emp.phongBan : '';

    // Xác định loại báo cáo
    const loaiBC = entry.reportType || 'BC Tuần';

    // STT sẽ được đánh lại sau khi chèn
    const row = [
      '',
      entry.email,
      entry.senderName,
      phongBan,
      loaiBC,
      formatDateTime(entry.submitTime),
      formatDateTime(entry.deadline),
      entry.status,
      entry.fileLink,
      entry.gmailMsgId,
      entry.fileHash || ''
    ];

    // Chèn dòng mới ngay sau header (dòng 2) → mới nhất luôn ở trên
    sheet.insertRowAfter(1);
    const targetRow = 2;
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);

    const rowRange = sheet.getRange(targetRow, 1, 1, 11);

    // Font chung cho dòng
    rowRange.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');

    // Đặt chiều cao dòng
    sheet.setRowHeight(targetRow, 30);

    // Tô màu Trạng Thái (cột 8)
    const statusCell = sheet.getRange(targetRow, 8);
    statusCell.setFontWeight('bold');

    switch (entry.status) {
      case CONFIG.STATUS.ON_TIME:
        statusCell.setBackground('#ceead6').setFontColor('#137333');
        break;
      case CONFIG.STATUS.LATE:
        statusCell.setBackground('#fad2cf').setFontColor('#c5221f');
        break;
      case CONFIG.STATUS.NO_ATTACHMENT:
        statusCell.setBackground('#fef7e0').setFontColor('#b05a00');
        break;
      case CONFIG.STATUS.DUPLICATE_FILE:
        statusCell.setBackground('#f4c7c3').setFontColor('#85200c');
        break;
    }

    // Viền nhẹ cho dòng
    rowRange.setBorder(null, null, true, null, null, null, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

    // Đánh STT cho dòng mới = 1 (vì nó luôn ở trên cùng)
    sheet.getRange(targetRow, 1).setValue(1);
    // Cập nhật STT cho các dòng phía dưới (+1)
    const totalRows = sheet.getLastRow() - 1;
    if (totalRows > 1) {
      const sttValues = [];
      for (let i = 1; i <= totalRows; i++) {
        sttValues.push([i]);
      }
      sheet.getRange(2, 1, totalRows, 1).setValues(sttValues);
    }

    Logger.log('✓ ' + entry.email + ' | ' + entry.status + ' | ' + loaiBC);
  } catch (error) {
    Logger.log('✗ Lỗi writeLogEntry: ' + error.message);
  }
}

// ============================================================================
// TÍNH NĂNG 2: NHẮC NHỞ TỰ ĐỘNG
// ============================================================================

/**
 * Lấy danh sách ai đã nộp tuần này
 * Kiểm tra cả deadline NV (Thứ 6) và TP (Thứ 7)
 */
function getSubmittedThisWeek() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_LOG);
    if (!sheet || sheet.getLastRow() < 2) return new Set();

    const data = sheet.getDataRange().getValues();
    const now = new Date();

    // Tìm T2 gần nhất ĐÃ QUA (hoặc hôm nay nếu là T2)
    // Nếu hôm nay là T3 → T2 = hôm qua
    // Nếu hôm nay là T4 → T2 = 2 ngày trước
    // Nếu hôm nay là T2 → T2 = hôm nay
    const dayOfWeek = now.getDay(); // 0=CN, 1=T2, 2=T3, ...
    const lastMonday = new Date(now);
    if (dayOfWeek === 0) {
      lastMonday.setDate(now.getDate() - 6); // CN → T2 tuần trước
    } else if (dayOfWeek === 1) {
      // T2 → chính là hôm nay
    } else {
      lastMonday.setDate(now.getDate() - (dayOfWeek - 1)); // T3-T7 → lùi về T2
    }

    // Deadline NV = T2 gần nhất lúc 23:59, TP = T2 gần nhất lúc 12:00
    const dlNV = new Date(lastMonday);
    dlNV.setHours(CONFIG.DEADLINE_NV.hour, CONFIG.DEADLINE_NV.minute, 0, 0);
    const dlTP = new Date(lastMonday);
    dlTP.setHours(CONFIG.DEADLINE_TP.hour, CONFIG.DEADLINE_TP.minute, 0, 0);

    // Dùng timestamp range: từ T2 tuần trước 00:00 đến T2 này 23:59
    // Bao phủ cả deadline T2 tuần trước (cho người nộp sớm) và T2 tuần này
    const prevMonday = new Date(lastMonday);
    prevMonday.setDate(prevMonday.getDate() - 7);
    const rangeStart = new Date(prevMonday.getFullYear(), prevMonday.getMonth(), prevMonday.getDate(), 0, 0, 0);
    const rangeEnd = new Date(lastMonday.getFullYear(), lastMonday.getMonth(), lastMonday.getDate(), 23, 59, 59);

    const submitted = new Set();

    for (let i = 1; i < data.length; i++) {
      const rawDeadline = data[i][6];
      let deadlineDate;
      if (rawDeadline instanceof Date) {
        deadlineDate = rawDeadline;
      } else {
        // Parse "dd/MM/yyyy HH:mm"
        const parts = (rawDeadline || '').toString().trim().match(/(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2})/);
        if (parts) {
          deadlineDate = new Date(parseInt(parts[3]), parseInt(parts[2]) - 1, parseInt(parts[1]), parseInt(parts[4]), parseInt(parts[5]));
        }
      }
      if (!deadlineDate) continue;

      const email = (data[i][1] || '').toString().toLowerCase().trim();
      const status = (data[i][7] || '').toString().trim();
      if (deadlineDate >= rangeStart && deadlineDate <= rangeEnd && status !== CONFIG.STATUS.DUPLICATE_FILE) {
        submitted.add(email);
      }
    }
    return submitted;
  } catch (error) {
    return new Set();
  }
}

/**
 * Thứ 2 sáng 9h: Nhắc nhở BC TUẦN
 *  - NV: deadline tối nay 23:59
 *  - TP: deadline trưa nay 12:00
 */
function sendMondayReminder() {
  Logger.log('=== THỨ 2: NHẮC NHỞ BC TUẦN ===');
  try {
    const employees = getActiveEmployees();
    const submitted = getSubmittedThisWeek();
    let sentNV = 0, sentTP = 0;

    for (const [email, emp] of employees) {
      if (isExcludedEmail(email)) continue;
      if (submitted.has(email)) continue;

      const name = emp.hoTen || email.split('@')[0];

      if (isTruongPhong(email)) {
        // ─── TP: deadline 12:00 trưa hôm nay ───
        try {
          const subject = '[' + CONFIG.COMPANY_NAME + '] Nhắc nhở TP: Báo cáo tuần - Hạn nộp TRƯA NAY (Thứ 2) 12:00';
          const body = 'Xin chào ' + name + ',\n\n'
            + 'Nhắc nhở: Deadline nộp báo cáo tuần (Trưởng phòng) là HÔM NAY (Thứ 2) lúc 12:00 trưa.\n\n'
            + 'Vui lòng gửi báo cáo tuần qua email về: ' + CONFIG.HR_EMAIL + '\n\n'
            + 'Trân trọng,\n'
            + 'Hệ thống theo dõi báo cáo ' + CONFIG.COMPANY_NAME;
          GmailApp.sendEmail(email, subject, body);
          sentTP++;
        } catch (e) {
          Logger.log('✗ Không gửi được cho TP ' + email + ': ' + e.message);
        }
      } else {
        // ─── NV: deadline 23:59 tối nay ───
        try {
          const subject = '[' + CONFIG.COMPANY_NAME + '] Nhắc nhở: Báo cáo tuần - Hạn nộp tối nay (Thứ 2) 23:59';
          const body = 'Xin chào ' + name + ',\n\n'
            + 'Nhắc nhở: Deadline nộp báo cáo tuần là HÔM NAY (Thứ 2) lúc 23:59.\n\n'
            + 'Vui lòng gửi báo cáo tuần qua email về: ' + CONFIG.HR_EMAIL + '\n'
            + '(Email có file đính kèm .xlsx, .pdf, .doc...)\n\n'
            + 'Trân trọng,\n'
            + 'Hệ thống theo dõi báo cáo ' + CONFIG.COMPANY_NAME;
          GmailApp.sendEmail(email, subject, body);
          sentNV++;
        } catch (e) {
          Logger.log('✗ Không gửi được cho ' + email + ': ' + e.message);
        }
      }
    }
    Logger.log('✓ Thứ 2: nhắc ' + sentNV + ' NV + ' + sentTP + ' TP');
  } catch (error) {
    Logger.log('✗ Lỗi sendMondayReminder: ' + error.message);
  }
}

/**
 * Thứ 3 sáng 9h: Nhắc MẠNH ai trễ hạn BC tuần
 *  - NV đã trễ (deadline T2 23:59 đã qua)
 *  - TP đã trễ (deadline T2 12:00 đã qua)
 */
function sendTuesdayLateReminder() {
  Logger.log('=== THỨ 3: NHẮC MẠNH NGƯỜI TRỄ HẠN BC TUẦN ===');
  try {
    const employees = getActiveEmployees();
    const submitted = getSubmittedThisWeek();
    let sent = 0;
    const treHan = [];

    for (const [email, emp] of employees) {
      if (isExcludedEmail(email)) continue;
      if (submitted.has(email)) continue;

      const name = emp.hoTen || email.split('@')[0];
      const role = isTruongPhong(email) ? 'TP' : 'NV';
      const deadlineStr = isTruongPhong(email) ? 'Thứ 2 lúc 12:00 trưa' : 'Thứ 2 lúc 23:59';
      treHan.push({ email: email, name: name, role: role });

      try {
        const roleLabel = role === 'TP' ? 'TRƯỞNG PHÒNG' : 'NHÂN VIÊN';
        const subject = '⚠ [' + CONFIG.COMPANY_NAME + '] ' + roleLabel + ': BẠN ĐÃ TRỄ HẠN nộp báo cáo tuần!';
        const body = 'Xin chào ' + name + ',\n\n'
          + '⚠ Bạn đã TRỄ HẠN nộp báo cáo tuần!\n'
          + 'Deadline đã qua: ' + deadlineStr + '.\n\n'
          + 'Vui lòng gửi báo cáo NGAY qua email: ' + CONFIG.HR_EMAIL + '\n\n'
          + 'Lưu ý: Việc nộp trễ sẽ được ghi nhận trong hệ thống.\n\n'
          + 'Trân trọng,\n'
          + 'Hệ thống theo dõi báo cáo ' + CONFIG.COMPANY_NAME;
        GmailApp.sendEmail(email, subject, body);
        sent++;
      } catch (e) {
        Logger.log('✗ Không gửi được cho ' + email + ': ' + e.message);
      }
    }

    // CC quản lý nếu có người trễ
    if (treHan.length > 0 && CONFIG.MANAGER_EMAILS.length > 0) {
      const listStr = treHan.map((p, i) => (i + 1) + '. [' + p.role + '] ' + p.name + ' (' + p.email + ')').join('\n');
      const subject = '⚠ [' + CONFIG.COMPANY_NAME + '] ' + treHan.length + ' người trễ hạn báo cáo tuần';
      const body = 'Kính gửi Ban Quản Lý,\n\n'
        + 'Tính đến Thứ 3 sáng, có ' + treHan.length + ' người TRỄ HẠN nộp báo cáo tuần:\n\n'
        + listStr + '\n\n'
        + 'Trân trọng,\n'
        + 'Hệ thống theo dõi báo cáo ' + CONFIG.COMPANY_NAME;

      CONFIG.MANAGER_EMAILS.forEach(mgr => {
        try { GmailApp.sendEmail(mgr, subject, body); } catch (e) { }
      });
    }

    Logger.log('✓ Thứ 3: nhắc mạnh ' + sent + ' người trễ hạn');
  } catch (error) {
    Logger.log('✗ Lỗi sendTuesdayLateReminder: ' + error.message);
  }
}

// ============================================================================
// TÍNH NĂNG 2B: NHẮC NHỞ BÁO CÁO THÁNG
// ============================================================================

/**
 * Mùng 1 hàng tháng: Nhắc NV KD nộp báo cáo tháng
 * - NV KD MAST PRO: deadline mùng 3
 * - NV KD MAST SG (hotd, vypt, vuvh, nhannt, truongdq, tuyetns): deadline mùng 3
 * - TBP KD SG (sonpt@mast.com.vn): deadline ngày 5 (CUSTOM_MONTHLY_DEADLINES)
 * - GĐ SG (nhanntt): deadline ngày 28 cho BC KH (xử lý riêng ở sendMonthlyReminder_TP)
 */
function sendMonthlyReminder_NV() {
  Logger.log('=== MÙNG 1: NHẮC NV KD NỘP BÁO CÁO THÁNG ===');
  try {
    const employees = getActiveEmployees();
    let sent = 0;
    const now = new Date();
    const thangTruoc = now.getMonth() === 0 ? 12 : now.getMonth(); // tháng trước (1-12)

    for (const [email, emp] of employees) {
      if (isExcludedEmail(email)) continue;

      // TP có custom monthly deadline (VD: sonpt ngày 5) → vẫn nhắc
      const hasCustomDeadline = !!CONFIG.CUSTOM_MONTHLY_DEADLINES[email.toLowerCase().trim()];
      if (isTruongPhong(email) && !hasCustomDeadline) continue;

      // Chỉ nhắc NV phòng Kinh Doanh (hoặc TP có custom deadline)
      if (!hasCustomDeadline) {
        const pb = (emp.phongBan || '').trim();
        const isKD = CONFIG.PHONG_KD_NAMES.some(name => pb.toLowerCase() === name.toLowerCase());
        if (!isKD) continue;
      }

      // Xác định deadline tháng (có thể custom theo từng email)
      const deadlineCfg = CONFIG.CUSTOM_MONTHLY_DEADLINES[email] || CONFIG.DEADLINE_NV_MONTHLY;
      const deadlineStr = 'ngày ' + deadlineCfg.day + ' lúc ' + deadlineCfg.hour + ':' + String(deadlineCfg.minute).padStart(2, '0');

      try {
        const name = emp.hoTen || email.split('@')[0];
        const subject = '[' + CONFIG.COMPANY_NAME + '] Nhắc nhở: Báo cáo tháng ' + thangTruoc + ' - Hạn nộp ' + deadlineStr;
        const body = 'Xin chào ' + name + ',\n\n'
          + 'Nhắc nhở: Bạn cần nộp BÁO CÁO THÁNG ' + thangTruoc + '.\n'
          + 'Deadline: ' + deadlineStr + ' tháng này.\n\n'
          + 'Vui lòng gửi báo cáo qua email về: ' + CONFIG.HR_EMAIL + '\n\n'
          + 'Trân trọng,\n'
          + 'Hệ thống theo dõi báo cáo ' + CONFIG.COMPANY_NAME;
        GmailApp.sendEmail(email, subject, body);
        sent++;
      } catch (e) {
        Logger.log('✗ Không gửi được cho ' + email + ': ' + e.message);
      }
    }
    Logger.log('✓ Mùng 1: nhắc ' + sent + ' NV KD nộp BC tháng');
  } catch (error) {
    Logger.log('✗ Lỗi sendMonthlyReminder_NV: ' + error.message);
  }
}

/**
 * Ngày 26 hàng tháng: Nhắc TP nộp Kế hoạch tháng + KPI (deadline 28, 23:59)
 * Bao gồm GĐ SG (nhanntt): BC Kế hoạch tháng tới - deadline ngày 28
 */
function sendMonthlyReminder_TP() {
  Logger.log('=== NGÀY 26: NHẮC TP NỘP KẾ HOẠCH + KPI ===');
  try {
    const employees = getActiveEmployees();
    let sent = 0;
    const now = new Date();
    const thangSau = now.getMonth() === 11 ? 1 : now.getMonth() + 2; // tháng tiếp theo (1-12)
    const thangTruoc = now.getMonth() === 0 ? 12 : now.getMonth(); // tháng trước (1-12)

    for (const [email, emp] of employees) {
      if (isExcludedEmail(email)) continue;
      if (!isTruongPhong(email)) continue;

      try {
        const name = emp.hoTen || email.split('@')[0];

        // GĐ SG (nhanntt): nhắc BC Kế hoạch tháng tới + BC Tổng kết tháng trước
        if (email === 'nhanntt@mastsaigon.com.vn') {
          const subject = '[' + CONFIG.COMPANY_NAME + '] Nhắc nhở GĐ: BC Kế hoạch tháng ' + thangSau + ' - Hạn nộp ngày 28 lúc 23:59';
          const body = 'Xin chào ' + name + ',\n\n'
            + 'Nhắc nhở nộp báo cáo:\n'
            + '1. BC KẾ HOẠCH THÁNG ' + thangSau + ' - Deadline: Ngày 28 tháng này lúc 23:59.\n\n'
            + 'Lưu ý: BC TỔNG KẾT hoạt động tháng ' + thangTruoc + ' có deadline ngày 9 tháng sau.\n\n'
            + 'Vui lòng gửi qua email về: ' + CONFIG.HR_EMAIL + '\n\n'
            + 'Trân trọng,\n'
            + 'Hệ thống theo dõi báo cáo ' + CONFIG.COMPANY_NAME;
          GmailApp.sendEmail(email, subject, body);
        } else {
          // TP MAST PRO + TP SG khác: nhắc KH + KPI thông thường
          const subject = '[' + CONFIG.COMPANY_NAME + '] Nhắc nhở TP: Kế hoạch + KPI tháng ' + thangSau + ' - Hạn nộp ngày 28 lúc 23:59';
          const body = 'Xin chào ' + name + ',\n\n'
            + 'Nhắc nhở: Bạn cần nộp KẾ HOẠCH THÁNG ' + thangSau + ' + KPI THÁNG ' + thangSau + '.\n'
            + 'Deadline: Ngày 28 tháng này lúc 23:59.\n\n'
            + 'Vui lòng gửi qua email về: ' + CONFIG.HR_EMAIL + '\n\n'
            + 'Trân trọng,\n'
            + 'Hệ thống theo dõi báo cáo ' + CONFIG.COMPANY_NAME;
          GmailApp.sendEmail(email, subject, body);
        }
        sent++;
      } catch (e) {
        Logger.log('✗ Không gửi được cho TP ' + email + ': ' + e.message);
      }
    }
    Logger.log('✓ Ngày 26: nhắc ' + sent + ' TP nộp KH + KPI');
  } catch (error) {
    Logger.log('✗ Lỗi sendMonthlyReminder_TP: ' + error.message);
  }
}

/**
 * Ngày 7 hàng tháng: Nhắc GĐ SG nộp BC Tổng kết hoạt động tháng trước
 * nhanntt@mastsaigon.com.vn - deadline ngày 9, 23:59
 */
function sendMonthlySummaryReminder_GD() {
  Logger.log('=== NGÀY 7: NHẮC GĐ SG NỘP BC TỔNG KẾT ===');
  try {
    const now = new Date();
    const thangTruoc = now.getMonth() === 0 ? 12 : now.getMonth(); // tháng trước (1-12)

    for (const email of CONFIG.SUMMARY_REPORT_EMAILS) {
      try {
        const employees = getActiveEmployees();
        const emp = employees.get(email.toLowerCase().trim());
        const name = emp ? (emp.hoTen || email.split('@')[0]) : email.split('@')[0];

        const subject = '[' + CONFIG.COMPANY_NAME + '] Nhắc nhở GĐ: BC Tổng kết tháng ' + thangTruoc + ' - Hạn nộp ngày 9 lúc 23:59';
        const body = 'Xin chào ' + name + ',\n\n'
          + 'Nhắc nhở: Bạn cần nộp BÁO CÁO TỔNG KẾT HOẠT ĐỘNG THÁNG ' + thangTruoc + '.\n'
          + 'Deadline: Ngày ' + CONFIG.SUMMARY_REPORT_DEADLINE.day + ' tháng này lúc 23:59.\n\n'
          + 'Vui lòng gửi qua email về: ' + CONFIG.HR_EMAIL + '\n\n'
          + 'Trân trọng,\n'
          + 'Hệ thống theo dõi báo cáo ' + CONFIG.COMPANY_NAME;
        GmailApp.sendEmail(email, subject, body);
        Logger.log('✓ Đã nhắc BC Tổng kết cho: ' + email);
      } catch (e) {
        Logger.log('✗ Không gửi được cho ' + email + ': ' + e.message);
      }
    }
  } catch (error) {
    Logger.log('✗ Lỗi sendMonthlySummaryReminder_GD: ' + error.message);
  }
}

// ============================================================================
// TÍNH NĂNG 3: BÁO CÁO TỔNG HỢP HÀNG TUẦN
// ============================================================================

/**
 * Thứ 4 sáng 8h: Gửi báo cáo tổng hợp tuần cho quản lý
 * (Dời từ T2 sang T4 vì deadline NV/TP đều là T2, cần 1 ngày buffer)
 */
function sendWeeklyReport() {
  Logger.log('=== GỬI BÁO CÁO TỔNG HỢP TUẦN ===');
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_LOG);
    if (!sheet || sheet.getLastRow() < 2) {
      Logger.log('Không có dữ liệu');
      return;
    }

    const employees = getActiveEmployees();
    const data = sheet.getDataRange().getValues();

    // Tìm T2 gần nhất đã qua (giống getSubmittedThisWeek)
    const now = new Date();
    const dayOfWeek = now.getDay();
    const lastMonday = new Date(now);
    if (dayOfWeek === 0) {
      lastMonday.setDate(now.getDate() - 6);
    } else if (dayOfWeek === 1) {
      // T2 → chính hôm nay
    } else {
      lastMonday.setDate(now.getDate() - (dayOfWeek - 1));
    }
    const dlNV = new Date(lastMonday);
    dlNV.setHours(CONFIG.DEADLINE_NV.hour, CONFIG.DEADLINE_NV.minute, 0, 0);
    const dlTP = new Date(lastMonday);
    dlTP.setHours(CONFIG.DEADLINE_TP.hour, CONFIG.DEADLINE_TP.minute, 0, 0);
    const lastDeadlineNV = formatDateTime(dlNV);
    const lastDeadlineTP = formatDateTime(dlTP);

    // Phân loại theo tuần vừa qua (cả NV và TP)
    const onTime = [], late = [], duplicate = [], noFile = [];
    const reportedEmails = new Set();

    for (let i = 1; i < data.length; i++) {
      // Cấu trúc mới: HanNop(6), Email(1), HoTen(2), TrangThai(7)
      const rawDeadline = data[i][6];
      const deadline = (rawDeadline instanceof Date) ? formatDateTime(rawDeadline) : (rawDeadline || '').toString().trim();
      if (deadline !== lastDeadlineNV && deadline !== lastDeadlineTP) continue;

      const email = (data[i][1] || '').toString().toLowerCase().trim();
      const name = (data[i][2] || '').toString().trim();
      const status = (data[i][7] || '').toString().trim();

      reportedEmails.add(email);

      const entry = name + ' (' + email + ')';
      switch (status) {
        case CONFIG.STATUS.ON_TIME: onTime.push(entry); break;
        case CONFIG.STATUS.LATE: late.push(entry); break;
        case CONFIG.STATUS.DUPLICATE_FILE: duplicate.push(entry); break;
        case CONFIG.STATUS.NO_ATTACHMENT: noFile.push(entry); break;
      }
    }

    // Ai chưa nộp?
    const notSubmitted = [];
    for (const [email, emp] of employees) {
      if (isExcludedEmail(email)) continue;
      if (!reportedEmails.has(email)) {
        notSubmitted.push((emp.hoTen || email) + ' (' + email + ')');
      }
    }

    // Tạo nội dung email
    let total = 0;
    for (const [email] of employees) {
      if (!isExcludedEmail(email)) total++;
    }
    const deadlineDisplay = 'NV: ' + lastDeadlineNV + ' | TP: ' + lastDeadlineTP;
    const subject = '[' + CONFIG.COMPANY_NAME + '] Bao cao tong hop tuan - Deadline T2 ' + lastDeadlineNV.split(' ')[0];

    let body = 'BAO CAO TONG HOP NOP BAO CAO TUAN\n';
    body += 'Deadline NV: ' + lastDeadlineNV + ' (23:59)\n';
    body += 'Deadline TP: ' + lastDeadlineTP + ' (12:00 trua)\n';
    body += '=======================================\n\n';

    body += 'THONG KE:\n';
    body += '- Tong nhan vien: ' + total + '\n';
    body += '- Dung han: ' + onTime.length + '\n';
    body += '- Tre han: ' + late.length + '\n';
    body += '- File trung lap: ' + duplicate.length + '\n';
    body += '- Chua nop: ' + notSubmitted.length + '\n\n';

    if (onTime.length > 0) {
      body += 'DUNG HAN (' + onTime.length + '):\n';
      onTime.forEach((e, i) => body += '  ' + (i + 1) + '. ' + e + '\n');
      body += '\n';
    }

    if (late.length > 0) {
      body += 'TRE HAN (' + late.length + '):\n';
      late.forEach((e, i) => body += '  ' + (i + 1) + '. ' + e + '\n');
      body += '\n';
    }

    if (duplicate.length > 0) {
      body += 'FILE TRUNG LAP (' + duplicate.length + '):\n';
      duplicate.forEach((e, i) => body += '  ' + (i + 1) + '. ' + e + '\n');
      body += '\n';
    }

    if (notSubmitted.length > 0) {
      body += 'CHUA NOP (' + notSubmitted.length + '):\n';
      notSubmitted.forEach((e, i) => body += '  ' + (i + 1) + '. ' + e + '\n');
      body += '\n';
    }

    body += '=======================================\n';
    body += 'Tu dong gui boi he thong theo doi bao cao ' + CONFIG.COMPANY_NAME + '\n';
    body += 'Link sheet: https://docs.google.com/spreadsheets/d/' + CONFIG.SHEET_ID + '\n';

    // Gửi cho quản lý
    CONFIG.MANAGER_EMAILS.forEach(mgr => {
      try {
        GmailApp.sendEmail(mgr, subject, body);
        Logger.log('✓ Đã gửi báo cáo cho: ' + mgr);
      } catch (e) {
        Logger.log('✗ Không gửi được cho ' + mgr + ': ' + e.message);
      }
    });

    Logger.log('✓ Báo cáo tuần: ' + onTime.length + ' đúng hạn, ' + late.length + ' trễ, ' + notSubmitted.length + ' chưa nộp');
  } catch (error) {
    Logger.log('✗ Lỗi sendWeeklyReport: ' + error.message);
  }
}

// ============================================================================
// QUÉT LẠI & TIỆN ÍCH
// ============================================================================

/**
 * ★ QUÉT LẠI TOÀN BỘ EMAIL (30 ngày, kể cả đã đọc) ★
 * Chạy SAU migrateToV2() để lấy lại dữ liệu.
 * Tự động thêm nhân viên mới.
 */
function rescanAllEmails() {
  Logger.log('=== QUÉT LẠI TOÀN BỘ EMAIL (30 ngày) ===');

  try {
    const query = 'has:attachment newer_than:30d to:' + CONFIG.HR_EMAIL;
    const threads = GmailApp.search(query);

    if (threads.length === 0) {
      Logger.log('Không tìm thấy email nào');
      return;
    }

    Logger.log('Tìm thấy ' + threads.length + ' threads');

    const employees = getActiveEmployees();
    const loggedMsgIds = getLoggedMessageIds();
    const existingHashes = getExistingFileHashes();
    let processed = 0, skipped = 0, errors = 0, newEmployees = 0;

    for (const thread of threads) {
      for (const message of thread.getMessages()) {
        try {
          const msgId = message.getId();
          if (loggedMsgIds.has(msgId)) continue;

          const senderEmail = extractEmail(message.getFrom());
          if (isExcludedEmail(senderEmail)) { skipped++; continue; }

          const attachments = message.getAttachments();
          const validFiles = attachments.filter(att => isValidFile(att.getName()));
          if (validFiles.length === 0) { skipped++; continue; }

          // Tự động thêm nhân viên mới
          let employee = employees.get(senderEmail);
          if (!employee) {
            const senderName = extractSenderName(message.getFrom());
            autoAddEmployee(senderEmail, senderName);
            employee = { email: senderEmail, hoTen: senderName, phongBan: 'MỚI_TỰ_ĐỘNG' };
            employees.set(senderEmail, employee);
            newEmployees++;
          }

          const result = processSingleEmail(message, employee, existingHashes);
          if (result) {
            writeLogEntry(result);
            loggedMsgIds.add(msgId);
            if (result.fileHash) {
              existingHashes.set(result.fileHash, { email: result.email, submitTime: formatDateTime(result.submitTime) });
            }
            processed++;
          }
        } catch (emailError) {
          errors++;
          Logger.log('✗ Lỗi: ' + emailError.message);
        }
      }
    }

    Logger.log('=== RESCAN XONG: ' + processed + ' xử lý | ' + skipped + ' bỏ qua | ' + errors + ' lỗi | ' + newEmployees + ' NV mới ===');
  } catch (error) {
    Logger.log('✗ LỖI: ' + error.message);
  }
}

function checkSystemStatus() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const logSheet = ss.getSheetByName(CONFIG.SHEET_LOG);
    const employees = getActiveEmployees();

    Logger.log('=== TRẠNG THÁI HỆ THỐNG V2.2 ===');
    Logger.log('Giờ VN: ' + formatDateTime(new Date()));
    Logger.log('Deadline NV tuần này (T2 23:59): ' + formatDateTime(_calcDeadline(new Date(), CONFIG.DEADLINE_NV)));
    Logger.log('Deadline TP tuần này (T2 12:00): ' + formatDateTime(_calcDeadline(new Date(), CONFIG.DEADLINE_TP)));
    Logger.log('Nhân viên: ' + employees.size);

    if (logSheet) {
      Logger.log('Nhật ký: ' + Math.max(0, logSheet.getLastRow() - 1) + ' dòng');
    }

    const submitted = getSubmittedThisWeek();
    Logger.log('Đã nộp tuần này: ' + submitted.size + '/' + employees.size);

    const triggers = ScriptApp.getProjectTriggers();
    Logger.log('Triggers: ' + triggers.length);
    triggers.forEach(t => Logger.log('  → ' + t.getHandlerFunction()));

    // Liệt kê ai chưa nộp
    let count = 0;
    for (const [email, emp] of employees) {
      if (isExcludedEmail(email)) continue;
      if (!submitted.has(email)) {
        count++;
        Logger.log('  ✗ Chưa nộp: ' + (emp.hoTen || email));
      }
    }
    Logger.log('Tổng chưa nộp: ' + count);
  } catch (error) {
    Logger.log('✗ Lỗi: ' + error.message);
  }
}

function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  Logger.log('Đã xóa ' + triggers.length + ' triggers');
}

/**
 * Test nhắc nhở (không cần đợi đúng ngày)
 */
function testReminders() {
  Logger.log('=== TEST NHẮC NHỞ ===');
  const employees = getActiveEmployees();
  const submitted = getSubmittedThisWeek();

  Logger.log('Nhân viên: ' + employees.size);
  Logger.log('Đã nộp: ' + submitted.size);

  for (const [email, emp] of employees) {
    if (isExcludedEmail(email)) continue;
    const status = submitted.has(email) ? '✓ Đã nộp' : '✗ Chưa nộp → SẼ NHẮC';
    Logger.log('  ' + (emp.hoTen || email) + ': ' + status);
  }
}

/**
 * ★ TEST CẤU HÌNH MAST SG ★
 * Kiểm tra toàn bộ deadline được áp dụng đúng cho từng email SG.
 * Chạy hàm này để verify trước khi deploy.
 */
function testMastSGConfig() {
  Logger.log('=== TEST CẤU HÌNH MAST SG ===');
  const now = new Date();
  const testDate = now; // Dùng ngày hiện tại để test

  // Danh sách email SG cần kiểm tra
  const sgEmails = [
    { email: 'nhanntt@mastsaigon.com.vn', name: 'Ngô Thị Thanh Nhàn',      chucvu: 'GĐ' },
    { email: 'anhntt@mastsaigon.com.vn',  name: 'Nguyễn Thị Tuyết Anh',    chucvu: 'KTT' },
    { email: 'sonpt@mast.com.vn',         name: 'Phạm Thanh Sơn',          chucvu: 'TBP KD' },
    { email: 'toanlv@mastsaigon.com.vn',  name: 'Lê Văn Toàn',             chucvu: 'TP Điều vận' },
    { email: 'nhansu@mastsaigon.com.vn',  name: 'Nguyễn Diễm My',          chucvu: 'HCNS' },
    { email: 'baond@mastsaigon.com.vn',   name: 'Nguyễn Đình Bảo',         chucvu: 'TP Kho' },
    { email: 'nhannt@mast.com.vn',        name: 'Nguyễn Thành Nhân',       chucvu: 'NVKD' },
    { email: 'truongdq@mast.com.vn',      name: 'Đào Quang Trường',        chucvu: 'NVKD' },
    { email: 'tuyetns@mast.com.vn',       name: 'Nguyễn Sơn Tuyết',        chucvu: 'NVKD' },
    { email: 'hotd@mastsaigon.com.vn',    name: 'Trương Đoan Hồ',          chucvu: 'NVKD' },
    { email: 'vypt@mastsaigon.com.vn',    name: 'Phạm Thanh Vy',           chucvu: 'NVKD' },
    { email: 'vuvh@mastsaigon.com.vn',    name: 'Vi Hoàng Vũ',             chucvu: 'NVKD' },
    { email: 'hienvc@mastsaigon.com.vn',  name: 'Võ Công Hiển',            chucvu: 'NVKD' },
  ];

  // Danh sách email SG bị loại trừ (không có BC)
  const excludedSG = [
    'tiendmt@mastsaigon.com.vn',
    'ketoan@mastsaigon.com.vn',
    'trungtt@mastsaigon.com.vn',
    'sint@mast.com.vn',
    'conghienn.lian@gmail.com',
  ];

  let pass = 0, fail = 0;

  Logger.log('\n── DEADLINE BC TUẦN ──');
  for (const item of sgEmails) {
    const isTP = isTruongPhong(item.email);
    const deadlineWeek = getDeadlineForEmail(testDate, item.email);
    const expectedHour = isTP ? 12 : 23;
    const expectedMin  = isTP ? 0  : 59;
    const ok = deadlineWeek.getHours() === expectedHour && deadlineWeek.getMinutes() === expectedMin;
    const label = ok ? '✓ PASS' : '✗ FAIL';
    if (ok) pass++; else fail++;
    Logger.log(label + ' | ' + item.name + ' [' + item.chucvu + '] → deadline tuần: '
      + formatDateTime(deadlineWeek) + ' (kỳ vọng ' + expectedHour + ':' + String(expectedMin).padStart(2,'0') + ')');
  }

  Logger.log('\n── DEADLINE BC THÁNG ──');
  // NV KD thường: ngày 3
  const nvkdList = ['nhannt@mast.com.vn','truongdq@mast.com.vn','tuyetns@mast.com.vn',
                    'hotd@mastsaigon.com.vn','vypt@mastsaigon.com.vn','vuvh@mastsaigon.com.vn',
                    'hienvc@mastsaigon.com.vn'];
  for (const email of nvkdList) {
    const dl = getMonthlyDeadlineNV(testDate, email);
    const ok = dl.getDate() === 3;
    if (ok) pass++; else fail++;
    Logger.log((ok ? '✓ PASS' : '✗ FAIL') + ' | ' + email + ' → deadline tháng ngày: ' + dl.getDate() + ' (kỳ vọng: 3)');
  }
  // TBP KD SG: ngày 5
  const sonDl = getMonthlyDeadlineNV(testDate, 'sonpt@mast.com.vn');
  const sonOk = sonDl.getDate() === 5;
  if (sonOk) pass++; else fail++;
  Logger.log((sonOk ? '✓ PASS' : '✗ FAIL') + ' | sonpt@mast.com.vn (TBP KD SG) → deadline tháng ngày: ' + sonDl.getDate() + ' (kỳ vọng: 5)');

  // GĐ SG: ngày 28
  const gdDl = getMonthlyDeadlineNV(testDate, 'nhanntt@mastsaigon.com.vn');
  const gdOk = gdDl.getDate() === 28;
  if (gdOk) pass++; else fail++;
  Logger.log((gdOk ? '✓ PASS' : '✗ FAIL') + ' | nhanntt@mastsaigon.com.vn (GĐ SG) → deadline tháng ngày: ' + gdDl.getDate() + ' (kỳ vọng: 28)');

  Logger.log('\n── EXCLUDED EMAILS SG ──');
  for (const email of excludedSG) {
    const excluded = isExcludedEmail(email);
    const ok = excluded === true;
    if (ok) pass++; else fail++;
    Logger.log((ok ? '✓ PASS' : '✗ FAIL') + ' | ' + email + ' bị loại trừ: ' + excluded + ' (kỳ vọng: true)');
  }

  Logger.log('\n── TRUONG_PHONG CHECK ──');
  const tpExpected = {
    'nhanntt@mastsaigon.com.vn': true,
    'anhntt@mastsaigon.com.vn':  true,
    'sonpt@mast.com.vn':         true,
    'toanlv@mastsaigon.com.vn':  true,
    'nhansu@mastsaigon.com.vn':  true,
    'baond@mastsaigon.com.vn':   true,
    'nhannt@mast.com.vn':        false,
    'hotd@mastsaigon.com.vn':    false,
    'hienvc@mastsaigon.com.vn':  false,
  };
  for (const [email, expected] of Object.entries(tpExpected)) {
    const actual = isTruongPhong(email);
    const ok = actual === expected;
    if (ok) pass++; else fail++;
    Logger.log((ok ? '✓ PASS' : '✗ FAIL') + ' | isTruongPhong(' + email + ') = ' + actual + ' (kỳ vọng: ' + expected + ')');
  }

  Logger.log('\n=== KẾT QUẢ: ' + pass + ' PASS | ' + fail + ' FAIL ===');
  if (fail === 0) {
    Logger.log('★★★ TẤT CẢ TESTS PASSED - CẤU HÌNH MAST SG CHÍNH XÁC ★★★');
  } else {
    Logger.log('⚠ CÓ ' + fail + ' TEST THẤT BẠI - Kiểm tra lại cấu hình!');
  }
}

// ============================================================================
// ĐÍNH CHÍNH - CHẠY MỘT LẦN
// ============================================================================

/**
 * ★ CHẠY MỘT LẦN ★
 * Gửi email đính chính cho 6 người MAST SG bị hệ thống nhắc SAI deadline.
 * Lý do: trước đây họ chưa được cấu hình là Trưởng phòng nên hệ thống
 * nhắc deadline 23:59 thay vì 12:00 TRƯA (Thứ 2).
 * NVKD (Nhân, Trường, Tuyết, Hồ, Vy, Vũ) deadline 23:59 không đổi → không gửi.
 */
function sendCorrectionEmail() {
  // Chỉ những người bị nhắc SAI: deadline thực là 12:00 TRƯA nhưng hệ thống cũ nhắc 23:59
  const affected = [
    { email: 'nhansu@mastsaigon.com.vn',  name: 'My'       },
    { email: 'nhanntt@mastsaigon.com.vn', name: 'chị Nhàn' },
    { email: 'anhntt@mastsaigon.com.vn',  name: 'chị Anh'  },
    { email: 'sonpt@mast.com.vn',         name: 'anh Sơn'  },
    { email: 'toanlv@mastsaigon.com.vn',  name: 'anh Toàn' },
    { email: 'baond@mastsaigon.com.vn',   name: 'anh Bảo'  },
  ];

  let sent = 0;
  for (const r of affected) {
    try {
      const subject = '[MAST SG] Đính chính: Deadline báo cáo tuần là 12:00 TRƯA Thứ 2';
      const body = 'Xin chào ' + r.name + ',\n\n'
        + 'Trước đó hệ thống đã gửi nhắc nhở với deadline 23:59 Thứ 2 cho bạn — đó là thông tin sai.\n\n'
        + 'Deadline nộp báo cáo tuần đúng của bạn là: 12:00 TRƯA (Thứ 2) hàng tuần.\n\n'
        + 'Hệ thống đã được cập nhật và sẽ nhắc đúng từ tuần tới trở đi.\n\n'
        + 'Xin lỗi vì sự bất tiện này.\n\n'
        + 'Trân trọng,\n'
        + 'Hệ thống theo dõi báo cáo MAST';
      GmailApp.sendEmail(r.email, subject, body);
      Logger.log('✓ Đã gửi đính chính cho: ' + r.name + ' (' + r.email + ')');
      sent++;
    } catch (e) {
      Logger.log('✗ Không gửi được cho ' + r.email + ': ' + e.message);
    }
  }
  Logger.log('=== ĐÍNH CHÍNH XONG: đã gửi ' + sent + '/' + affected.length + ' email ===');
}

