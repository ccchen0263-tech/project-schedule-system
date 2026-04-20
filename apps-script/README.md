# Shuxin Appointment System (Google Sheets + Apps Script)

此資料夾提供可直接部署到「空白 Google Sheet」的初始化與核心預約腳本。

## 目標
- 把空白試算表建立成新版正式檔案結構。
- 建立必要分頁與欄位。
- 預載狀態列舉與基礎服務/診次。
- 提供可直接執行的 setup / 初始化腳本。

## 檔案說明
- `AppConfig.gs`：分頁、欄位、列舉、預設資料設定。
- `Setup.gs`：`setupWorkbook()` 初始化入口、驗證規則、種子資料。
- `AppointmentService.gs`：建立預約、衝突檢查、稽核日誌。
- `ImportMigration.gs`：舊資料欄位映射匯入、整體檢查錯與報表輸出。

## 部署步驟（對應你的空白 Google Sheet）
1. 開啟你提供的 Google Sheet。
2. 進入 `擴充功能 > Apps Script`。
3. 新增 4 個 `.gs` 檔，分別貼上本目錄完整內容。
4. 儲存後，執行 `setupWorkbook()`。
5. 回到試算表，確認已建立以下分頁：
   - Clients
   - Staff
   - Sessions
   - ServiceTypes
   - Appointments
   - AppointmentSegments
   - Closures
   - Waitlist
   - AuditLog
   - Enums
   - ValidationReport

## 初始化後可直接測試
可在 Apps Script 執行以下範例：

```javascript
function testCreateAppointment() {
  var result = createAppointment({
    client_id: 'client_001',
    appointment_date: '2026-04-22',
    source: 'phone',
    status: 'booked',
    notes: '首次測試',
    created_by: 'frontdesk@clinic.com',
    segments: [
      {
        service_type_id: 'svc_initial_consult',
        role: 'consultant',
        staff_id: 'staff_consult_001',
        session_id: 'session_1',
        start_at: '2026-04-22T09:00:00+08:00',
        end_at: '2026-04-22T09:30:00+08:00',
        duration_min: 30,
        segment_status: 'booked'
      }
    ]
  });

  Logger.log(result);
}
```

## 檢查錯（建議）
- status 是否在 `Enums` 的 `appointment_status` 白名單。
- 同 staff / 同時段是否撞單。
- 同 session / 同時段是否超過 capacity。
- `AppointmentSegments.appointment_id` 是否能對應 `Appointments.appointment_id`。

## 注意
- `resetWorkbookDangerous()` 會清空所有資料，只建議在測試檔使用。
- 若要對接你的舊專案修正版資料，可先用 setup 建架構，再批次匯入舊資料。


## 匯入與整合舊專案資料
可使用 `importRowsWithMapping(sheetName, sourceHeaders, sourceRows, headerMap)` 將舊欄位映射到新結構。

完成匯入後，執行 `verifyWorkbookIntegrity()`：
- 會檢查缺少分頁 / 缺少欄位
- 檢查狀態白名單
- 檢查孤兒 segment
- 檢查 staff 衝突 / session 超容量
- 結果輸出到 `ValidationReport`
