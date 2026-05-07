# theo-doi-bao-cao

Hệ thống theo dõi báo cáo nội bộ MAST (Google Apps Script v2.2).

## Mục đích
Tự động hóa việc theo dõi nộp báo cáo TUẦN/THÁNG của nhân viên & trưởng phòng MAST PRO + MAST SG: quét Gmail, ghi nhật ký vào Google Sheet, lưu file lên Drive, nhắc nhở qua email theo deadline, gửi báo cáo tổng hợp cho quản lý.

## Cấu trúc
- `Code.gs` — toàn bộ logic Apps Script (CONFIG, triggers, xử lý email, nhắc nhở, báo cáo).

## Triển khai
1. Tạo Apps Script project (script.google.com), dán nội dung `Code.gs`.
2. Cập nhật `CONFIG.SHEET_ID` (ID Google Sheet) và các email TP, deadline nếu cần.
3. Chạy `migrateToV2()` MỘT LẦN để tạo sheet `Nhật Ký Báo Cáo`.
4. Chạy `setupTriggers()` MỘT LẦN để cài 7 trigger thời gian.
5. (Tuỳ chọn) Chạy `rescanAllEmails()` để nạp lại 30 ngày email gần nhất.
6. Verify: chạy `testMastSGConfig()` và `checkSystemStatus()`.

## Deadline
- BC TUẦN: NV 23:59 Thứ 2 · TP 12:00 trưa Thứ 2.
- BC THÁNG: NV KD ngày 3 · TP (KH+KPI) ngày 28 · GĐ SG (BC tổng kết) ngày 9 · TBP KD SG ngày 5.

## Trigger
| Khi nào | Hàm | Mục đích |
|---|---|---|
| Mỗi 10 phút | `processEmails` | Quét Gmail, ghi nhật ký |
| T2 09:00 | `sendMondayReminder` | Nhắc NV + TP nộp BC tuần |
| T3 09:00 | `sendTuesdayLateReminder` | Nhắc mạnh ai trễ + CC quản lý |
| T4 08:00 | `sendWeeklyReport` | Báo cáo tổng hợp tuần cho quản lý |
| Mùng 1 09:00 | `sendMonthlyReminder_NV` | Nhắc NV KD nộp BC tháng |
| Ngày 7 09:00 | `sendMonthlySummaryReminder_GD` | Nhắc GĐ SG nộp BC tổng kết |
| Ngày 26 09:00 | `sendMonthlyReminder_TP` | Nhắc TP nộp KH + KPI |