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
  DEADLINE_NV: { day: 1, hour: 23, minute: 59 },
  DEADLINE_TP: { day: 1, hour: 12, minute: 0 },

  // ── DEADLINE THÁNG ──
  DEADLINE_NV_MONTHLY: { day: 3, hour: 23, minute: 59 },
  DEADLINE_TP_MONTHLY: { day: 28, hour: 23, minute: 59 },

  PHONG_KD_NAMES: ['Kinh doanh', 'KD', 'Sales', 'Kinh Doanh', 'KINH DOANH'],

  // ── TRƯỞNG PHÒNG ──
  TRUONG_PHONG_EMAILS: [
    'linhdt@mast.com.vn',
    'tunv@mast.com.vn',
    'chienmt@mast.com.vn',
    'hoangvd@mast.com.vn',
    'luongpt@mast.com.vn',
    'khanhxuan@mast.com.vn',
    'nhanntt@mastsaigon.com',
    'anhntt@mastsaigon.com',
    'sonpt@mast.com.vn',
    'toanlv@mastsaigon.com',
    'nhansu@mastsaigon.com',
    'baond@mastsaigon.com'
  ],

  SHEET_EMPLOYEES: 'Danh Sách Nhân Viên',
  SHEET_LOG: 'Nhật Ký Báo Cáo',

  COL_EMAIL: 0,
  COL_HOTEN: 1,
  COL_PHONGBAN: 2,

  VALID_EXTENSIONS: ['.xlsx', '.xls', '.pdf', '.doc', '.docx', '.ppt', '.pptx', '.csv'],

  STATUS: {
    ON_TIME: 'Đúng hạn',
    LATE: 'Trễ hạn',
    NO_ATTACHMENT: 'Thiếu file đính kèm',
    DUPLICATE_FILE: 'File trùng lặp'
  },

  HR_EMAIL: 'nhansu@mast.com.vn',
  MANAGER_EMAILS: ['nhansu@mast.com.vn'],
  COMPANY_NAME: 'MAST PRO',

  CUSTOM_MONTHLY_DEADLINES: {
    'sonpt@mast.com.vn':          { day: 5,  hour: 23, minute: 59 },
    'nhanntt@mastsaigon.com':     { day: 28, hour: 23, minute: 59 },
  },

  SUMMARY_REPORT_DEADLINE: { day: 9, hour: 23, minute: 59 },
  SUMMARY_REPORT_EMAILS: ['nhanntt@mastsaigon.com'],

  EXCLUDED_EMAILS: [
    'nhansu@mast.com.vn',
    'tiendm@mastsaigon.com',
    'trungtt@mastsaigon.com',
    'ketoan@mastsaigon.com',
    'sint@mast.com.vn',
    'conghienn.lian@gmail.com',
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

  sheet.setRowHeight(1, 36);

  sheet.setFrozenRows(1);
  sheet.hideColumns(10, 2);

  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 160);
  sheet.setColumnWidth(7, 160);
  sheet.setColumnWidth(8, 130);
  sheet.setColumnWidth(9, 300);

  sheet.getRange('A:A').setHorizontalAlignment('center');
  sheet.getRange('H:H').setHorizontalAlignment('center');
  sheet.getRange('E:E').setHorizontalAlignment('center');

  Logger.log('✓ Tạo sheet V2.2: ' + CONFIG.SHEET_LOG);
  return sheet;
}

// ============================================================================
// TRIGGERS - CHẠY MỘT LẦN
// ============================================================================

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('processEmails')
    .timeBased().everyMinutes(10).create();

  ScriptApp.newTrigger('sendMondayReminder')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9).create();

  ScriptApp.newTrigger('sendTuesdayLateReminder')
    .timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .atHour(9).create();

  ScriptApp.newTrigger('sendWeeklyReport')
    .timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
    .atHour(8).create();

  ScriptApp.newTrigger('sendMonthlyReminder_NV')
    .timeBased().onMonthDay(1)
    .atHour(9).create();

  ScriptApp.newTrigger('sendMonthlyReminder_TP')
    .timeBased().onMonthDay(26)
    .atHour(9).create();

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
  if (lower.match(/^image\d*\.(png|jpg|jpeg|gif)$/)) return false;
  return CONFIG.VALID_EXTENSIONS.some(ext => lower.endsWith(ext));
}

// Bảng chuẩn hoá các email bị gõ sai / alias — giúp khớp đúng TP/NV
// khi người gửi dùng biến thể domain nhỏ (tránh HR phải sửa tay cột Email).
const EMAIL_ALIASES_ = {
  'toanlv@mastsaigon.com.vn': 'toanlv@mastsaigon.com'
};

function extractEmail(fromField) {
  if (!fromField) return '';
  const match = fromField.match(/<(.+?)>/);
  const raw = match ? match[1].toLowerCase().trim() : fromField.toLowerCase().trim();
  return EMAIL_ALIASES_[raw] || raw;
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

function isTruongPhong(email) {
  if (!email) return false;
  return CONFIG.TRUONG_PHONG_EMAILS.includes(email.toLowerCase().trim());
}

function isNhanVienKD(email) {
  if (!email) return false;
  const employees = getActiveEmployees();
  const emp = employees.get(email.toLowerCase().trim());
  if (!emp) return false;
  const pb = (emp.phongBan || '').trim();
  return CONFIG.PHONG_KD_NAMES.some(name => pb.toLowerCase() === name.toLowerCase());
}

function getMonthlyDeadlineNV(date, email) {
  const d = (date instanceof Date && !isNaN(date.getTime())) ? new Date(date) : new Date();

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

function getDeadlineForEmail(date, email) {
  const deadlineConfig = isTruongPhong(email) ? CONFIG.DEADLINE_TP : CONFIG.DEADLINE_NV;
  return _calcDeadline(date, deadlineConfig);
}

function getFridayDeadline(date) {
  return _calcDeadline(date, CONFIG.DEADLINE_NV);
}

function _calcDeadline(date, deadlineConfig) {
  let d;
  if (date instanceof Date && !isNaN(date.getTime())) {
    d = new Date(date);
  } else {
    d = new Date();
  }

  const day = d.getDay();
  const targetDay = deadlineConfig.day;
  let diff;

  if (day === targetDay) {
    const deadlineToday = new Date(d);
    deadlineToday.setHours(deadlineConfig.hour, deadlineConfig.minute, 0, 0);
    diff = (d.getTime() <= deadlineToday.getTime()) ? 0 : 7;
  } else if (day === 0) {
    diff = 1;
  } else if (day > targetDay) {
    diff = 7 - day + targetDay;
  } else {
    diff = targetDay - day;
  }

  const target = new Date(d);
  target.setDate(d.getDate() + diff);
  target.setHours(deadlineConfig.hour, deadlineConfig.minute, 0, 0);
  return target;
}

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
// PHÂN LOẠI BÁO CÁO & ĐỒNG BỘ PHÒNG BAN (bổ sung v2.2.1)
// ============================================================================

// Map cố định: email TP → tên phòng ban chính thức.
// Dùng để tự điền cột PhongBan cho các dòng đang là "MỚI_TỰ_ĐỘNG" hoặc rỗng
// → HR không cần điền thủ công cho những người đã biết vị trí trong công ty.
const TP_PHONGBAN_MAP_ = {
  'linhdt@mast.com.vn':      'Kho',
  'tunv@mast.com.vn':        'Điều vận',
  'chienmt@mast.com.vn':     'Kế toán',
  'hoangvd@mast.com.vn':     'Nhập mua',
  'luongpt@mast.com.vn':     'Kinh doanh',
  'khanhxuan@mast.com.vn':   'HCNS',
  'nhanntt@mastsaigon.com':  'Ban Giám đốc SG',
  'anhntt@mastsaigon.com':   'Kế toán SG',
  'sonpt@mast.com.vn':       'Kinh doanh SG',
  'toanlv@mastsaigon.com':   'Điều vận SG',
  'nhansu@mastsaigon.com':   'HCNS SG',
  'baond@mastsaigon.com':    'Kho SG'
};

/**
 * Phân loại báo cáo dựa trên tiêu đề email.
 * Chỉ bắt các tín hiệu CHẮC CHẮN — nếu không rõ thì mặc định 'BC Tuần'
 * để giữ nguyên hành vi cũ, tránh phân loại sai.
 * @param {string} subject
 * @returns {'BC Tuần'|'BC Tháng'|'BC KH Tháng'|'BC KPI'|'BC Tổng kết'}
 */
function classifyReportType(subject) {
  const s = (subject || '').toLowerCase().trim();
  if (!s) return 'BC Tuần';
  const n = s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd');

  // BC Tổng kết (ưu tiên cao nhất)
  if (/\btong\s*ket\b/.test(n)) return 'BC Tổng kết';
  // BC Kế hoạch tháng
  if (/\bke\s*hoach\s*thang\b|\bkh\s*thang\b/.test(n)) return 'BC KH Tháng';
  // BC KPI — chỉ khớp khi có kèm "tháng" để tránh false-positive với BC tuần có nhắc KPI
  if (/\bkpi\b/.test(n) && /\bthang\b/.test(n)) return 'BC KPI';
  // BC tháng tổng quát
  if (/\bbao\s*cao\s*thang\b|\bbc\s*thang\b|\bbct\s*\d/.test(n)) return 'BC Tháng';

  return 'BC Tuần';
}

/**
 * Tính deadline theo loại báo cáo đã phân loại.
 * @param {'BC Tuần'|'BC Tháng'|'BC KH Tháng'|'BC KPI'|'BC Tổng kết'} reportType
 * @param {Date} submitTime
 * @param {string} email
 */
function getDeadlineForReport_(reportType, submitTime, email) {
  const emailLower = (email || '').toLowerCase().trim();

  if (reportType === 'BC Tuần') {
    return getDeadlineForEmail(submitTime, emailLower);
  }

  if (reportType === 'BC Tổng kết') {
    const d = (submitTime instanceof Date && !isNaN(submitTime.getTime()))
      ? new Date(submitTime) : new Date();
    const cfg = CONFIG.SUMMARY_REPORT_DEADLINE;
    let target = new Date(d.getFullYear(), d.getMonth(), cfg.day, cfg.hour, cfg.minute, 0, 0);
    if (d.getTime() > target.getTime()) {
      target = new Date(d.getFullYear(), d.getMonth() + 1, cfg.day, cfg.hour, cfg.minute, 0, 0);
    }
    return target;
  }

  // BC Tháng / BC KH Tháng / BC KPI
  // Thứ tự ưu tiên: custom deadline → TP dùng ngày 28 → NV dùng ngày 3
  if (CONFIG.CUSTOM_MONTHLY_DEADLINES[emailLower]) {
    return getMonthlyDeadlineNV(submitTime, emailLower);
  }
  if (isTruongPhong(emailLower)) {
    return getMonthlyDeadlineTP(submitTime);
  }
  return getMonthlyDeadlineNV(submitTime, emailLower);
}

/**
 * Nếu email là TP trong TP_PHONGBAN_MAP_ và dòng trong sheet đang để
 * "MỚI_TỰ_ĐỘNG" hoặc rỗng → tự điền phòng ban đúng. Không ghi đè nếu
 * HR đã nhập giá trị hợp lệ.
 * @param {string} email
 */
function syncPhongBanForTP_(email) {
  try {
    const key = (email || '').toLowerCase().trim();
    const expected = TP_PHONGBAN_MAP_[key];
    if (!expected) return;

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_EMPLOYEES);
    if (!sheet || sheet.getLastRow() < 2) return;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowEmail = (data[i][CONFIG.COL_EMAIL] || '').toString().toLowerCase().trim();
      if (rowEmail !== key) continue;

      const current = (data[i][CONFIG.COL_PHONGBAN] || '').toString().trim();
      if (!current || current === 'MỚI_TỰ_ĐỘNG') {
        sheet.getRange(i + 1, CONFIG.COL_PHONGBAN + 1).setValue(expected);
        Logger.log('✓ Tự cập nhật phòng ban: ' + key + ' → ' + expected);
      }
      return;
    }
  } catch (e) {
    Logger.log('✗ Lỗi syncPhongBanForTP_: ' + e.message);
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

function autoAddEmployee(email, senderName) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_EMPLOYEES);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_EMPLOYEES);
      sheet.getRange(1, 1, 1, 3).setValues([['Email', 'HoTen', 'PhongBan']]);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
    }
    // Nếu là TP đã biết trước → điền luôn phòng ban đúng, HR không phải sửa tay
    const phongBan = TP_PHONGBAN_MAP_[(email || '').toLowerCase().trim()] || 'MỚI_TỰ_ĐỘNG';
    sheet.appendRow([email, senderName, phongBan]);
    Logger.log('✓ Tự động thêm nhân viên: ' + email + ' (' + senderName + ') [' + phongBan + ']');
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
      const val = (data[i][9] || '').toString().trim();
      if (val) msgIds.add(val);
    }
    return msgIds;
  } catch (error) {
    return new Set();
  }
}

function getExistingFileHashes() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_LOG);
    if (!sheet || sheet.getLastRow() < 2) return new Map();

    const data = sheet.getDataRange().getValues();
    const hashes = new Map();
    for (let i = 1; i < data.length; i++) {
      const hashField = (data[i][10] || '').toString().trim();
      const email = (data[i][1] || '').toString().trim();
      const time = (data[i][5] || '').toString().trim();
      const status = (data[i][7] || '').toString().trim();
      if (hashField && status !== CONFIG.STATUS.DUPLICATE_FILE) {
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

          const attachments = message.getAttachments();
          const validFiles = attachments.filter(att => isValidFile(att.getName()));
          if (validFiles.length === 0) { skipped++; continue; }

          let employee = employees.get(senderEmail);
          if (!employee) {
            const senderName = extractSenderName(message.getFrom());
            autoAddEmployee(senderEmail, senderName);
            const inferredPhongBan = TP_PHONGBAN_MAP_[senderEmail] || 'MỚI_TỰ_ĐỘNG';
            employee = { email: senderEmail, hoTen: senderName, phongBan: inferredPhongBan };
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

function processSingleEmail(message, employee, existingHashes) {
  const senderEmail = employee.email;
  const senderName = employee.hoTen || extractSenderName(message.getFrom());
  const submitTime = message.getDate();
  const msgId = message.getId();

  // Phân loại báo cáo theo tiêu đề (trả về 'BC Tuần' nếu không có signal rõ)
  let subject = '';
  try { subject = message.getSubject() || ''; } catch (e) { /* ignore */ }
  const reportType = classifyReportType(subject);
  const deadline = getDeadlineForReport_(reportType, submitTime, senderEmail);

  // Nếu người gửi là TP đã biết trước (trong TP_PHONGBAN_MAP_) mà dòng sheet
  // đang để "MỚI_TỰ_ĐỘNG" hoặc rỗng → tự điền phòng ban giúp HR.
  if (TP_PHONGBAN_MAP_[senderEmail] &&
      (!employee.phongBan || employee.phongBan === 'MỚI_TỰ_ĐỘNG')) {
    syncPhongBanForTP_(senderEmail);
    employee.phongBan = TP_PHONGBAN_MAP_[senderEmail];
  }

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
      reportType: reportType
    };
  }

  const fileLinks = [];
  let isDuplicate = false;
  let duplicateInfo = '';
  const allHashes = [];

  for (const file of validFiles) {
    const hash = computeFileHash(file);
    if (hash) allHashes.push(hash);

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
    reportType: reportType
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

    const employees = getActiveEmployees();
    const emp = employees.get(entry.email.toLowerCase().trim());
    const phongBan = emp ? emp.phongBan : '';

    const loaiBC = entry.reportType || 'BC Tuần';

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

    sheet.insertRowAfter(1);
    const targetRow = 2;
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);

    const rowRange = sheet.getRange(targetRow, 1, 1, 11);

    rowRange.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');

    sheet.setRowHeight(targetRow, 30);

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

    rowRange.setBorder(null, null, true, null, null, null, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(targetRow, 1).setValue(1);
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

function getSubmittedThisWeek() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_LOG);
    if (!sheet || sheet.getLastRow() < 2) return new Set();

    const data = sheet.getDataRange().getValues();
    const now = new Date();

    const dayOfWeek = now.getDay();
    const lastMonday = new Date(now);
    if (dayOfWeek === 0) {
      lastMonday.setDate(now.getDate() - 6);
    } else if (dayOfWeek === 1) {
      // T2 → chính là hôm nay
    } else {
      lastMonday.setDate(now.getDate() - (dayOfWeek - 1));
    }

    const dlNV = new Date(lastMonday);
    dlNV.setHours(CONFIG.DEADLINE_NV.hour, CONFIG.DEADLINE_NV.minute, 0, 0);
    const dlTP = new Date(lastMonday);
    dlTP.setHours(CONFIG.DEADLINE_TP.hour, CONFIG.DEADLINE_TP.minute, 0, 0);

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

function sendMonthlyReminder_NV() {
  Logger.log('=== MÙNG 1: NHẮC NV KD NỘP BÁO CÁO THÁNG ===');
  try {
    const employees = getActiveEmployees();
    let sent = 0;
    const now = new Date();
    const thangTruoc = now.getMonth() === 0 ? 12 : now.getMonth();

    for (const [email, emp] of employees) {
      if (isExcludedEmail(email)) continue;

      const hasCustomDeadline = !!CONFIG.CUSTOM_MONTHLY_DEADLINES[email.toLowerCase().trim()];
      if (isTruongPhong(email) && !hasCustomDeadline) continue;

      if (!hasCustomDeadline) {
        const pb = (emp.phongBan || '').trim();
        const isKD = CONFIG.PHONG_KD_NAMES.some(name => pb.toLowerCase() === name.toLowerCase());
        if (!isKD) continue;
      }

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

function sendMonthlyReminder_TP() {
  Logger.log('=== NGÀY 26: NHẮC TP NỘP KẾ HOẠCH + KPI ===');
  try {
    const employees = getActiveEmployees();
    let sent = 0;
    const now = new Date();
    const thangSau = now.getMonth() === 11 ? 1 : now.getMonth() + 2;
    const thangTruoc = now.getMonth() === 0 ? 12 : now.getMonth();

    for (const [email, emp] of employees) {
      if (isExcludedEmail(email)) continue;
      if (!isTruongPhong(email)) continue;

      try {
        const name = emp.hoTen || email.split('@')[0];

        if (email === 'nhanntt@mastsaigon.com') {
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

function sendMonthlySummaryReminder_GD() {
  Logger.log('=== NGÀY 7: NHẮC GĐ SG NỘP BC TỔNG KẾT ===');
  try {
    const now = new Date();
    const thangTruoc = now.getMonth() === 0 ? 12 : now.getMonth();

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

    const onTime = [], late = [], duplicate = [], noFile = [];
    const reportedEmails = new Set();

    for (let i = 1; i < data.length; i++) {
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

    const notSubmitted = [];
    for (const [email, emp] of employees) {
      if (isExcludedEmail(email)) continue;
      if (!reportedEmails.has(email)) {
        notSubmitted.push((emp.hoTen || email) + ' (' + email + ')');
      }
    }

    let total = 0;
    for (const [email] of employees) {
      if (!isExcludedEmail(email)) total++;
    }
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

          let employee = employees.get(senderEmail);
          if (!employee) {
            const senderName = extractSenderName(message.getFrom());
            autoAddEmployee(senderEmail, senderName);
            const inferredPhongBan = TP_PHONGBAN_MAP_[senderEmail] || 'MỚI_TỰ_ĐỘNG';
            employee = { email: senderEmail, hoTen: senderName, phongBan: inferredPhongBan };
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

function testMastSGConfig() {
  Logger.log('=== TEST CẤU HÌNH MAST SG ===');
  const now = new Date();
  const testDate = now;

  const sgEmails = [
    { email: 'nhanntt@mastsaigon.com',  name: 'Ngô Thị Thanh Nhàn',      chucvu: 'GĐ' },
    { email: 'anhntt@mastsaigon.com',   name: 'Nguyễn Thị Tuyết Anh',    chucvu: 'KTT' },
    { email: 'sonpt@mast.com.vn',       name: 'Phạm Thanh Sơn',          chucvu: 'TBP KD' },
    { email: 'toanlv@mastsaigon.com',   name: 'Lê Văn Toàn',             chucvu: 'TBP Điều vận' },
    { email: 'nhansu@mastsaigon.com',   name: 'Nhân sự SG',              chucvu: 'HCNS' },
    { email: 'baond@mastsaigon.com',    name: 'Nguyễn Đình Bảo',         chucvu: 'TP Kho' },
    { email: 'nhannt@mast.com.vn',      name: 'Nguyễn Thành Nhân',       chucvu: 'NVKD' },
    { email: 'truongdq@mast.com.vn',    name: 'Đào Quang Trường',        chucvu: 'NVKD' },
    { email: 'tuyetns@mast.com.vn',     name: 'Nguyễn Sơn Tuyết',        chucvu: 'NVKD' },
    { email: 'hotd@mastsaigon.com',     name: 'Trương Đoan Hồ',          chucvu: 'NVKD' },
    { email: 'vypt@mastsaigon.com',     name: 'Phạm Thanh Vy',           chucvu: 'NVKD' },
    { email: 'vuvh@mastsaigon.com',     name: 'Vi Hoàng Vũ',             chucvu: 'NVKD' },
  ];

  const excludedSG = [
    'tiendm@mastsaigon.com',
    'trungtt@mastsaigon.com',
    'ketoan@mastsaigon.com',
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
  const nvkdList = ['nhannt@mast.com.vn','truongdq@mast.com.vn','tuyetns@mast.com.vn',
                    'hotd@mastsaigon.com','vypt@mastsaigon.com','vuvh@mastsaigon.com'];
  for (const email of nvkdList) {
    const dl = getMonthlyDeadlineNV(testDate, email);
    const ok = dl.getDate() === 3;
    if (ok) pass++; else fail++;
    Logger.log((ok ? '✓ PASS' : '✗ FAIL') + ' | ' + email + ' → deadline tháng ngày: ' + dl.getDate() + ' (kỳ vọng: 3)');
  }
  const sonDl = getMonthlyDeadlineNV(testDate, 'sonpt@mast.com.vn');
  const sonOk = sonDl.getDate() === 5;
  if (sonOk) pass++; else fail++;
  Logger.log((sonOk ? '✓ PASS' : '✗ FAIL') + ' | sonpt@mast.com.vn (TBP KD SG) → deadline tháng ngày: ' + sonDl.getDate() + ' (kỳ vọng: 5)');

  const gdDl = getMonthlyDeadlineNV(testDate, 'nhanntt@mastsaigon.com');
  const gdOk = gdDl.getDate() === 28;
  if (gdOk) pass++; else fail++;
  Logger.log((gdOk ? '✓ PASS' : '✗ FAIL') + ' | nhanntt@mastsaigon.com (GĐ SG) → deadline tháng ngày: ' + gdDl.getDate() + ' (kỳ vọng: 28)');

  Logger.log('\n── EXCLUDED EMAILS SG ──');
  for (const email of excludedSG) {
    const excluded = isExcludedEmail(email);
    const ok = excluded === true;
    if (ok) pass++; else fail++;
    Logger.log((ok ? '✓ PASS' : '✗ FAIL') + ' | ' + email + ' bị loại trừ: ' + excluded + ' (kỳ vọng: true)');
  }

  Logger.log('\n── TRUONG_PHONG CHECK ──');
  const tpExpected = {
    'nhanntt@mastsaigon.com': true,
    'anhntt@mastsaigon.com':  true,
    'sonpt@mast.com.vn':      true,
    'toanlv@mastsaigon.com':  true,
    'nhansu@mastsaigon.com':  true,
    'baond@mastsaigon.com':   true,
    'nhannt@mast.com.vn':     false,
    'hotd@mastsaigon.com':    false,
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

function sendCorrectionEmail() {
  const affected = [
    { email: 'nhansu@mastsaigon.com',  name: 'My'       },
    { email: 'nhanntt@mastsaigon.com', name: 'chị Nhàn' },
    { email: 'anhntt@mastsaigon.com',  name: 'chị Anh'  },
    { email: 'sonpt@mast.com.vn',      name: 'anh Sơn'  },
    { email: 'toanlv@mastsaigon.com',  name: 'anh Toàn' },
    { email: 'baond@mastsaigon.com',   name: 'anh Bảo'  },
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
