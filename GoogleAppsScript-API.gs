/**************************************************************************************
 *  สะพานเชื่อม Google Sheets สำหรับไฟล์ stroke-training-app.html
 *  (เวอร์ชันรวม: ล็อกอินพยาบาล + ข้อมูล CID/การวินิจฉัย + ค้นหาผู้ป่วยข้ามเครื่อง
 *   + คะแนน 0-100 + ผู้บันทึก)
 *  --------------------------------------------------------------------------------
 *  วิธีใช้ (ทำครั้งเดียว):
 *   1) เปิด Google Sheet ที่ต้องการเก็บข้อมูล → เมนู ส่วนขยาย (Extensions) → Apps Script
 *   2) ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดนี้ลงไป → กดบันทึก (Save)
 *   3) กด Deploy → New deployment → เลือกชนิด "Web app"
 *        - Execute as: Me (ตัวคุณเอง)
 *        - Who has access: Anyone   ← สำคัญ ต้องเลือกอันนี้
 *   4) กด Deploy แล้วคัดลอก "Web app URL" (ลงท้ายด้วย /exec)
 *   5) นำ URL ไปวางในไฟล์ stroke-training-app.html ที่บรรทัด:
 *        const GOOGLE_SCRIPT_URL = "วาง URL ตรงนี้";
 *
 *  ระบบจะสร้าง 2 ชีตอัตโนมัติ:
 *     - "รายชื่อผู้ป่วย"    : เก็บทะเบียนผู้ป่วย (พยาบาลเพิ่ม/แก้ไข → คนไข้ค้นหาเจอทุกเครื่อง)
 *     - "ข้อมูลการประเมิน" : เก็บผลการฝึกแต่ละครั้ง
 *
 *  รับส่งข้อมูลแบบ JSONP จึงไม่ติดปัญหา CORS และเปิดไฟล์ HTML จากที่ไหนก็ได้
 **************************************************************************************/

var ASSESS_SHEET  = "ข้อมูลการประเมิน";
var PATIENT_SHEET = "รายชื่อผู้ป่วย";

var ASSESS_HEADERS  = ["วัน-เวลา","รหัสผู้ป่วย","ชื่อผู้ป่วย","HN","CID","การวินิจฉัย (Dx)","หัวข้อการฝึก","ผลการประเมิน","คะแนน (0-2)","ผู้บันทึก","หมายเหตุ"];
var PATIENT_HEADERS = ["รหัส","ชื่อ-สกุล","อายุ","HN","CID","การวินิจฉัย (Dx)"];

// ผู้ป่วยเริ่มต้น (ใช้ seed ชีตครั้งแรกเท่านั้น)
var SEED_PATIENTS = [
  { id:"P001", name:"นายสมชาย ใจดี",      age:65, HN:"69-0001", cid:"1100200300401", dx:"อัมพฤกษ์ครึ่งซีกขวา (Right hemiparesis)" },
  { id:"P002", name:"นางสมศรี มีสุข",      age:58, HN:"69-0002", cid:"1100200300502", dx:"เส้นเลือดสมองตีบ (Ischemic stroke)" },
  { id:"P003", name:"นายสมศักดิ์ รักเรียน", age:70, HN:"69-0003", cid:"1100200300603", dx:"ภาวะพูดไม่ชัด (Dysarthria)" }
];

function doGet(e) {
  var callback = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : "callback";
  var action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action).toLowerCase() : "";
  var result;
  try {
    if      (action === "save")          result = saveAssessment_(e.parameter);
    else if (action === "getpatients")   result = { status:"success", patients: getPatients_() };
    else if (action === "savepatient")   result = savePatient_(e.parameter);
    else if (action === "deletepatient") result = deletePatient_(e.parameter);
    else if (action === "gethistory")    result = { status:"success", rows: getHistory_() };
    else if (action === "ping")          result = { status:"success", message:"พร้อมใช้งาน" };
    else result = { status:"error", message:"ไม่รู้จักคำสั่ง (action): " + action };
  } catch (err) {
    result = { status:"error", message: err.toString() };
  }
  return ContentService
    .createTextOutput(callback + "(" + JSON.stringify(result) + ")")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/* ============================== ผลการประเมิน ============================== */
function saveAssessment_(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ASSESS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ASSESS_SHEET);
    sheet.appendRow(ASSESS_HEADERS);
    sheet.getRange(1, 1, 1, ASSESS_HEADERS.length).setFontWeight("bold");
    sheet.getRange("E:E").setNumberFormat("@");   // CID เป็นข้อความ
  }
  sheet.appendRow([
    new Date(),
    p.patientId   || "",
    p.patientName || "",
    p.hn          || "",
    p.cid         || "",
    p.dx          || "",
    p.topic       || "",
    p.score       || "",
    (p.scoreValue !== undefined && p.scoreValue !== "") ? Number(p.scoreValue) : "",
    p.by          || "",
    p.note        || ""
  ]);
  return { status:"success", message:"บันทึกข้อมูลสำเร็จ" };
}

/* ============================== ทะเบียนผู้ป่วย ============================== */
function ensurePatientSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PATIENT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PATIENT_SHEET);
    sheet.appendRow(PATIENT_HEADERS);
    sheet.getRange(1, 1, 1, PATIENT_HEADERS.length).setFontWeight("bold");
    sheet.getRange("A:A").setNumberFormat("@");   // รหัสเป็นข้อความ
    sheet.getRange("E:E").setNumberFormat("@");   // CID เป็นข้อความ
    SEED_PATIENTS.forEach(function(p){ sheet.appendRow([p.id, p.name, p.age, p.HN, p.cid, p.dx]); });
  }
  return sheet;
}

function getPatients_() {
  var sheet = ensurePatientSheet_();
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[0] && !r[1]) continue;
    out.push({ id:String(r[0]), name:String(r[1]), age:r[2], HN:String(r[3]), cid:String(r[4]), dx:String(r[5]) });
  }
  return out;
}

function savePatient_(p) {
  var sheet = ensurePatientSheet_();
  var values = sheet.getDataRange().getValues();
  var row = [String(p.id||""), p.name||"", p.age||"", p.hn||"", String(p.cid||""), p.dx||""];
  var found = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(p.id)) { found = i + 1; break; }
  }
  if (found > 0) sheet.getRange(found, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
  return { status:"success", message:"บันทึกผู้ป่วยสำเร็จ" };
}

function deletePatient_(p) {
  var sheet = ensurePatientSheet_();
  var values = sheet.getDataRange().getValues();
  for (var i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0]) === String(p.id)) { sheet.deleteRow(i + 1); }
  }
  return { status:"success", message:"ลบผู้ป่วยแล้ว" };
}

/* ============================== อ่านประวัติ (เผื่อใช้) ============================== */
function getHistory_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ASSESS_SHEET);
  if (!sheet) return [];
  return sheet.getDataRange().getValues();
}
