/**
 * ==============================================================================
 * Google Apps Script (API แบบตอบรับการเข้าร่วมการประชุม)
 * ==============================================================================
 * การเรียงลำดับคอลัมน์ใน Google Sheets ตรงตามลำดับของฟอร์ม HTML 100%:
 * 
 * 1. วัน-เวลาที่ตอบ (Timestamp)
 * 2. ชื่อหน่วยงาน (agency)
 * 3. คำนำหน้า (prefix)
 * 4. ชื่อ-สกุล (fullName)
 * 5. ตำแหน่ง (position)
 * 6. เบอร์โทรศัพท์ (phone)
 * 7. อีเมล (email)
 * 8. แบบตอบรับการเข้าร่วมการประชุม (attendanceStatus)
 * 9. การยินยอม PDPA (pdpaConsent)
 * ==============================================================================
 */

// 🔑 API Key ที่ตรงกับใน index.html
const API_KEY = "ENV_OCC_HEALTH_2026_DDC";

// 📑 ชื่อชีตที่ต้องการบันทึกข้อมูล (หากไม่มีระบบจะใช้ชีตแรก)
const SHEET_NAME = "แบบตอบรับการประชุม";

/**
 * ฟังก์ชันรับคำขอ POST จากแบบฟอร์ม HTML
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  // ล็อกไม่เกิน 30 วินาที เพื่อป้องกันบันทึกข้อมูลซ้อนทับกันเมื่อมีผู้ส่งพร้อมกัน
  lock.tryLock(30000);

  try {
    const params = e.parameter;

    // 1. ตรวจสอบ API Key เพื่อความปลอดภัย
    if (!params.apiKey || params.apiKey !== API_KEY) {
      return createJsonResponse({
        status: "error",
        message: "Unauthorized: API Key ไม่ถูกต้องหรือไม่ได้รับอนุญาต"
      });
    }

    // 2. ตรวจสอบ Honeypot Trap (หากมีค่าแสดงว่าเป็น Bot)
    if (params.extra_company_url && params.extra_company_url.trim() !== "") {
      return createJsonResponse({
        status: "success",
        message: "บันทึกข้อมูลสำเร็จ"
      });
    }

    // 3. เข้าถึง Google Sheets
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.getSheets()[0];
    }

    // 4. ตรวจสอบและสร้างหัวตารางอัตโนมัติ (หากยังไม่มีข้อมูลในแถวที่ 1)
    if (sheet.getLastRow() === 0) {
      const headers = [
        "วัน-เวลาที่ตอบ",
        "ชื่อหน่วยงาน",
        "คำนำหน้า",
        "ชื่อ-สกุล",
        "ตำแหน่ง",
        "เบอร์โทรศัพท์ (มือถือ)",
        "อีเมล",
        "แบบตอบรับการเข้าร่วมการประชุม",
        "การยินยอม PDPA"
      ];
      sheet.appendRow(headers);

      // จัดรูปแบบหัวตาราง (แถวที่ 1) ให้สวยงาม
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground("#0284c7");
      headerRange.setFontColor("#ffffff");
      headerRange.setFontWeight("bold");
      headerRange.setFontFamily("Sarabun");
      sheet.setFrozenRows(1);
    }

    // 5. เตรียมข้อมูลเรียงตามลำดับของฟอร์ม HTML
    const timestamp = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
    const agency = (params.agency || "").trim();
    const prefix = (params.prefix || "").trim();
    const fullName = (params.fullName || "").trim();
    const position = (params.position || "").trim();
    
    // จัดรูปแบบเบอร์โทรศัพท์ให้เป็นข้อความ (ป้องกันเลข 0 นำหน้าหาย)
    let phone = (params.phone || "").trim().replace(/\D/g, "");
    if (phone) {
      phone = "'" + phone;
    }
    
    const email = (params.email || "").trim();
    const attendanceStatus = (params.attendanceStatus || "").trim();
    const pdpaConsent = (params.pdpaConsent || "").trim();

    // 6. แถวข้อมูลเรียงตามลำดับคอลัมน์ใน HTML
    const rowData = [
      timestamp,        // คอลัมน์ A: วัน-เวลาที่ตอบ
      agency,           // คอลัมน์ B: ชื่อหน่วยงาน
      prefix,           // คอลัมน์ C: คำนำหน้า
      fullName,         // คอลัมน์ D: ชื่อ-สกุล
      position,         // คอลัมน์ E: ตำแหน่ง
      phone,            // คอลัมน์ F: เบอร์โทรศัพท์ (มือถือ)
      email,            // คอลัมน์ G: อีเมล
      attendanceStatus, // คอลัมน์ H: แบบตอบรับการเข้าร่วมการประชุม
      pdpaConsent       // คอลัมน์ I: การยินยอม PDPA
    ];

    // 7. บันทึกข้อมูลลงแถวใหม่
    sheet.appendRow(rowData);

    return createJsonResponse({
      status: "success",
      message: "บันทึกข้อมูลเรียบร้อยแล้ว",
      data: {
        timestamp: timestamp,
        name: prefix + fullName,
        agency: agency,
        status: attendanceStatus
      }
    });

  } catch (error) {
    return createJsonResponse({
      status: "error",
      message: "เกิดข้อผิดพลาดในการบันทึกข้อมูล: " + error.toString()
    });

  } finally {
    // ปลดล็อกการทำงานเสมอ
    lock.releaseLock();
  }
}

/**
 * ฟังก์ชันรับคำขอ GET สำหรับทดสอบสถานะระบบ (Health Check)
 */
function doGet(e) {
  return createJsonResponse({
    status: "online",
    message: "API แบบตอบรับการประชุมพร้อมใช้งาน",
    timestamp: Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss")
  });
}

/**
 * ฟังก์ชันช่วยแปลง Object เป็น JSON Output
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
