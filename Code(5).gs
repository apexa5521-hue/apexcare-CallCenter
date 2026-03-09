// ============================================================
// APEXCARE Call Center Quality & Incentive System
// Code.gs - Google Apps Script Backend
// ============================================================

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const DRIVE_FOLDER_NAME = "APEXCARE_QualityDocs";

// ---- Sheet Names ----
const SHEETS = {
  LOG: "سجل_المخالفات",
  SCORES: "الأرصدة",
  SUPERVISOR: "تقييم_المشرفة",
  SETTINGS: "الإعدادات",
  WEEKLY: "التقرير_الأسبوعي"
};

// ---- Employee Config ----
const EMPLOYEES = {
  "مرح":    { track: "calls" },
  "هيلة":   { track: "calls" },
  "جنان":   { track: "dual" },
  "أسيل":   { track: "messages" },
  "حور":    { track: "messages" },
  "سارة":   { track: "messages" },
  "طيف":    { track: "messages" },
  "روان":   { track: "messages" },
  "تهاني":  { track: "dual" }
};

// ---- Criteria Points (per track) ----
const CRITERIA = {
  calls: {
    "المكالمات": {
      "الصياغة التعريفية في بداية المكالمة": 14,
      "التحقق من وجود ملف للمراجع": 11,
      "الرد على جميع المكالمات الواردة": 10,
      "تقديم خدمة مكتملة حتى الإغلاق": 8,
      "استخدام ألفاظ مهنية مناسبة": 6,
      "إقفال المكالمة بعبارة ختامية": 6
    },
    "فتح الملف": {
      "إدخال الاسم الرباعي بشكل صحيح": 7,
      "إدخال رقم الهوية بشكل صحيح": 6,
      "إدخال تاريخ الميلاد بشكل صحيح": 6,
      "إدخال العنوان بشكل صحيح": 4,
      "التحري بالدقة أثناء إدخال البيانات": 2
    },
    "المواعيد": {
      "تعويض المواعيد الملغاة": 6,
      "تأكيد المواعيد مع المراجعين": 5,
      "تعبئة المواعيد في حال توفر موعد بديل": 5,
      "ترتيب جدول الدكتور (مواعيد متتالية)": 4
    }
  },
  messages: {
    "الرسائل": {
      "الرد على الرسائل خلال 10 دقائق": 14,
      "متابعة المحادثات وعدم ترك محادثة مفتوحة": 11,
      "التحقق من وجود ملف للمراجع": 10,
      "التفاعل مع المراجع خلال دقيقة": 8,
      "عبارة ختامية مهذبة في آخر رد": 6,
      "الالتزام بالاختصارات المعتمدة": 6
    },
    "فتح الملف": {
      "إدخال الاسم الرباعي بشكل صحيح": 7,
      "إدخال رقم الهوية بشكل صحيح": 6,
      "إدخال تاريخ الميلاد بشكل صحيح": 6,
      "إدخال العنوان بشكل صحيح": 4,
      "التحري بالدقة أثناء إدخال البيانات": 2
    },
    "المواعيد": {
      "تعويض المواعيد الملغاة": 6,
      "تأكيد المواعيد مع المراجعين": 5,
      "تعبئة المواعيد في حال توفر موعد بديل": 5,
      "ترتيب جدول الدكتور (مواعيد متتالية)": 4
    }
  },
  dual: {
    "المكالمات": {
      "الصياغة التعريفية في بداية المكالمة": 8,
      "التحقق من وجود ملف للمراجع": 6,
      "الرد على جميع المكالمات الواردة": 5,
      "تقديم خدمة مكتملة حتى الإغلاق": 5,
      "استخدام ألفاظ مهنية مناسبة": 3,
      "إقفال المكالمة بعبارة ختامية": 3
    },
    "الرسائل": {
      "الرد على الرسائل خلال 10 دقائق": 8,
      "متابعة المحادثات وعدم ترك محادثة مفتوحة": 5,
      "التحقق من وجود ملف للمراجع": 5,
      "التفاعل مع المراجع خلال دقيقة": 4,
      "عبارة ختامية مهذبة في آخر رد": 4,
      "الالتزام بالاختصارات المعتمدة": 4
    },
    "فتح الملف": {
      "إدخال الاسم الرباعي بشكل صحيح": 7,
      "إدخال رقم الهوية بشكل صحيح": 6,
      "إدخال تاريخ الميلاد بشكل صحيح": 6,
      "إدخال العنوان بشكل صحيح": 4,
      "التحري بالدقة أثناء إدخال البيانات": 2
    },
    "المواعيد": {
      "تعويض المواعيد الملغاة": 5,
      "تأكيد المواعيد مع المراجعين": 4,
      "تعبئة المواعيد في حال توفر موعد بديل": 3,
      "ترتيب جدول الدكتور (مواعيد متتالية)": 3
    }
  }
};

// ---- Supervisor Criteria ----
const SUPERVISOR_CRITERIA = {
  "المتابعة اليومية للجدول ورصد التقصير": null,
  "الالتزام برصد المخالفات وتوثيقها بالصور": null,
  "التدخل لمعالجة تكرار المخالفات ومتابعة المحادثات المفتوحة": null,
  "الالتزام بتنفيذ المهام الموكلة (جرد، متابعة، تقارير)": null,
  "معالجة التقصير في عدد الموظفات واستكمال التغطية": null,
  "متابعة جودة المكالمات (اتصال يومي على 7078)": null
};

// كل 9 نقاط للمشرفة = 5 ريال
function rahafBonus(pts) {
  return Math.floor(pts / 9) * 5;
}

// ============================================================
// doGet - يستقبل جميع الطلبات عبر payload parameter
// ============================================================
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};

  if (params.payload) {
    let result;
    try {
      const data = JSON.parse(decodeURIComponent(params.payload));
      result = handleAction(data);
    } catch (err) {
      result = { success: false, error: err.toString() };
    }
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // رد بسيط للتأكد من أن الـ API يعمل
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, message: "APEXCARE API يعمل ✅" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// doPost - احتياطي
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const result = handleAction(data);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// handleAction - مشترك بين doGet و doPost
// ============================================================
function handleAction(data) {
  switch (data.action) {
    case "recordViolation":            return recordViolation(data);
    case "recordSupervisorEvaluation": return recordSupervisorEvaluation(data);
    case "uploadImage":                return uploadImage(data.base64Data, data.fileName);
    case "resetMonthlyScores":         return resetMonthlyScores();
    case "getAllData":                  return getAllData();
    case "getWeeklyReport":            return getWeeklyReport();
    case "initSheets":                 return initSheets();
    default: return { success: false, error: "action غير معروف: " + data.action };
  }
}

// ============================================================
// Initialize Sheets
// ============================================================
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // LOG sheet
  let log = ss.getSheetByName(SHEETS.LOG);
  if (!log) {
    log = ss.insertSheet(SHEETS.LOG);
    log.appendRow(["التاريخ", "الوقت", "المشرفة", "الموظفة", "المسار", "الفئة", "المعيار", "نقاط المعيار", "خصم النقاط", "رقم المخالفة", "رابط الصورة", "ملاحظات"]);
    log.getRange(1, 1, 1, 12).setBackground("#1a1a2e").setFontColor("#ffffff").setFontWeight("bold");
  }
  
  // SCORES sheet
  let scores = ss.getSheetByName(SHEETS.SCORES);
  if (!scores) {
    scores = ss.insertSheet(SHEETS.SCORES);
    scores.appendRow(["الموظفة", "المسار", "الرصيد الحالي", "النقاط المخصومة", "الحافز المستحق (ريال)", "آخر تحديث"]);
    scores.getRange(1, 1, 1, 6).setBackground("#1a1a2e").setFontColor("#ffffff").setFontWeight("bold");
    
    // Init employee rows
    Object.entries(EMPLOYEES).forEach(([name, info]) => {
      const trackLabel = info.track === "calls" ? "مكالمات" : info.track === "messages" ? "رسائل" : "مشترك";
      scores.appendRow([name, trackLabel, 100, 0, 500, new Date()]);
    });
  }
  
  // SUPERVISOR sheet
  let sup = ss.getSheetByName(SHEETS.SUPERVISOR);
  if (!sup) {
    sup = ss.insertSheet(SHEETS.SUPERVISOR);
    sup.setRightToLeft(true);
    sup.appendRow(["التاريخ", "المعيار", "نوع التقييم", "الخصم/الكسب", "رصيد رهف", "ملاحظات"]);
    sup.getRange(1, 1, 1, 6).setBackground("#1a1a2e").setFontColor("#ffffff").setFontWeight("bold")
       .setHorizontalAlignment("center");
    // عرض الأعمدة
    sup.setColumnWidth(1, 120); // التاريخ
    sup.setColumnWidth(2, 280); // المعيار
    sup.setColumnWidth(3, 140); // نوع التقييم
    sup.setColumnWidth(4, 120); // الخصم/الكسب
    sup.setColumnWidth(5, 120); // رصيد رهف
    sup.setColumnWidth(6, 200); // ملاحظات
    // صف البيانات الأولي
    sup.appendRow(["إجمالي", "رصيد المشرفة", "رهف", 0, 0, ""]);
    sup.getRange(2, 1, 1, 6).setHorizontalAlignment("center");
  }
  
  return { success: true, message: "تم تهيئة الشيتات بنجاح" };
}

// ============================================================
// Get all data for frontend
// ============================================================
function getAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get scores
  const scoresSheet = ss.getSheetByName(SHEETS.SCORES);
  let employeeScores = {};
  if (scoresSheet && scoresSheet.getLastRow() > 1) {
    const data = scoresSheet.getRange(2, 1, scoresSheet.getLastRow() - 1, 6).getValues();
    data.forEach(row => {
      if (row[0]) {
        employeeScores[row[0]] = {
          track: row[1],
          balance: row[2],
          deducted: row[3],
          bonus: row[4],
          lastUpdate: row[5]
        };
      }
    });
  }
  
  // Get log for violation counts
  const logSheet = ss.getSheetByName(SHEETS.LOG);
  let violationCounts = {}; // {employee_criterion: count}
  if (logSheet && logSheet.getLastRow() > 1) {
    const logs = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 12).getValues();
    logs.forEach(row => {
      if (row[3] && row[6]) {
        const key = `${row[3]}__${row[6]}`;
        violationCounts[key] = (violationCounts[key] || 0) + 1;
      }
    });
  }
  
  // Get supervisor score
  const supSheet = ss.getSheetByName(SHEETS.SUPERVISOR);
  let supervisorScore = 0;
  let rahafBalance = 0;
  if (supSheet && supSheet.getLastRow() > 1) {
    const supData = supSheet.getRange(2, 1, supSheet.getLastRow() - 1, 6).getValues();
    supData.forEach(row => {
      if (row[0] === "إجمالي") {
        supervisorScore = row[3] || 0;
        rahafBalance = row[4] || 0;
      }
    });
  }
  
  return {
    success: true,
    employees: EMPLOYEES,
    criteria: CRITERIA,
    supervisorCriteria: SUPERVISOR_CRITERIA,
    employeeScores,
    violationCounts,
    supervisorScore,
    rahafBalance
  };
}

// ============================================================
// Record a violation
// ============================================================
function recordViolation(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    initSheets();
    
    const { employeeName, criterion, category, imageUrl, notes } = data;
    const empInfo = EMPLOYEES[employeeName];
    if (!empInfo) return { success: false, error: "الموظفة غير موجودة" };
    
    const track = empInfo.track;
    const criteriaPoints = CRITERIA[track][category][criterion];
    if (!criteriaPoints) return { success: false, error: "المعيار غير موجود" };
    
    // Calculate violation number for this employee+criterion
    const logSheet = ss.getSheetByName(SHEETS.LOG);
    let violationNum = 1;
    if (logSheet && logSheet.getLastRow() > 1) {
      const logs = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 12).getValues();
      logs.forEach(row => {
        if (row[3] === employeeName && row[6] === criterion) violationNum++;
      });
    }
    
    // Calculate deduction: split criteriaPoints into 3 parts
    const deductPerViolation = Math.ceil(criteriaPoints / 3);
    let deduction = 0;
    let escalation = null;
    
    if (violationNum <= 3) {
      deduction = deductPerViolation;
    } else {
      // Points exhausted - escalation
      deduction = 0;
      if (violationNum === 4) escalation = "إنذار شفهي";
      else if (violationNum === 5) escalation = "إنذار كتابي";
      else escalation = "رفع للإدارة";
    }
    
    // Log the violation
    const now = new Date();
    logSheet.appendRow([
      Utilities.formatDate(now, "Asia/Riyadh", "yyyy-MM-dd"),
      Utilities.formatDate(now, "Asia/Riyadh", "HH:mm:ss"),
      "رهف",
      employeeName,
      track === "calls" ? "مكالمات" : track === "messages" ? "رسائل" : "مشترك",
      category,
      criterion,
      criteriaPoints,
      deduction,
      violationNum,
      imageUrl || "",
      notes || ""
    ]);
    
    // Update employee score
    updateEmployeeScore(employeeName, deduction);
    
    // Update supervisor score (+2 per correct observation)
    updateSupervisorScore(2);
    
    return {
      success: true,
      deduction,
      violationNum,
      escalation,
      message: escalation ? `⚠️ ${escalation}` : `تم خصم ${deduction} نقطة`
    };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================
// Update employee score in SCORES sheet
// ============================================================
function updateEmployeeScore(employeeName, deduction) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scoresSheet = ss.getSheetByName(SHEETS.SCORES);
  if (!scoresSheet) return;
  
  const data = scoresSheet.getRange(2, 1, scoresSheet.getLastRow() - 1, 6).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === employeeName) {
      const rowNum = i + 2;
      const currentBalance = scoresSheet.getRange(rowNum, 3).getValue();
      const currentDeducted = scoresSheet.getRange(rowNum, 4).getValue();
      const newBalance = Math.max(0, currentBalance - deduction);
      const newDeducted = currentDeducted + deduction;
      const bonus = newBalance * 5;
      scoresSheet.getRange(rowNum, 3).setValue(newBalance);
      scoresSheet.getRange(rowNum, 4).setValue(newDeducted);
      scoresSheet.getRange(rowNum, 5).setValue(bonus);
      scoresSheet.getRange(rowNum, 6).setValue(new Date());
      break;
    }
  }
}

// ============================================================
// Update supervisor (Rahaf) score
// ============================================================
function updateSupervisorScore(points) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const supSheet = ss.getSheetByName(SHEETS.SUPERVISOR);
  if (!supSheet) return;
  
  const data = supSheet.getRange(2, 1, supSheet.getLastRow() - 1, 6).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === "إجمالي") {
      const rowNum = i + 2;
      const currentScore = supSheet.getRange(rowNum, 4).getValue() || 0;
      const maxScore = Object.keys(EMPLOYEES).length * 100; // 9 × 100 = 900
      const newScore = Math.min(maxScore, currentScore + points);
      supSheet.getRange(rowNum, 4).setValue(newScore);
      supSheet.getRange(rowNum, 6).setValue(new Date());
      break;
    }
  }
}

// ============================================================
// Record supervisor (Rahaf) evaluation by manager
// ============================================================
function recordSupervisorEvaluation(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    initSheets();
    
    const { criterion, evalType, deduction, notes } = data;
    
    const supSheet = ss.getSheetByName(SHEETS.SUPERVISOR);
    const now = new Date();
    
    supSheet.appendRow([
      Utilities.formatDate(now, "Asia/Riyadh", "yyyy-MM-dd"),
      criterion,
      evalType,
      -deduction,
      "",
      notes || ""
    ]);
    
    // Update rahaf balance
    const supData = supSheet.getRange(2, 1, supSheet.getLastRow() - 1, 6).getValues();
    for (let i = 0; i < supData.length; i++) {
      if (supData[i][0] === "إجمالي") {
        const rowNum = i + 2;
        const currentBalance = supSheet.getRange(rowNum, 5).getValue() || 0;
        const newBalance = Math.max(0, currentBalance - deduction);
        supSheet.getRange(rowNum, 5).setValue(newBalance);
        supSheet.getRange(rowNum, 6).setValue(now);
        break;
      }
    }
    
    return { success: true, message: `تم خصم ${deduction} نقطة من رهف` };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================
// Upload image to Drive
// ============================================================
function uploadImage(base64Data, fileName) {
  try {
    const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
    let folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
    }
    
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data.split(",")[1]), "image/jpeg", fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { success: true, url: file.getUrl(), id: file.getId() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================
// Generate weekly report data
// ============================================================
function getWeeklyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scoresSheet = ss.getSheetByName(SHEETS.SCORES);
  const logSheet = ss.getSheetByName(SHEETS.LOG);
  
  let report = {
    generatedAt: new Date().toLocaleDateString("ar-SA"),
    employees: [],
    totalBonus: 0,
    topPerformer: "",
    violationSummary: {}
  };
  
  if (scoresSheet && scoresSheet.getLastRow() > 1) {
    const data = scoresSheet.getRange(2, 1, scoresSheet.getLastRow() - 1, 6).getValues();
    data.forEach(row => {
      if (row[0]) {
        const emp = {
          name: row[0],
          track: row[1],
          balance: row[2],
          deducted: row[3],
          bonus: row[4]
        };
        report.employees.push(emp);
        report.totalBonus += emp.bonus;
      }
    });
    
    if (report.employees.length > 0) {
      const top = report.employees.reduce((a, b) => a.balance > b.balance ? a : b);
      report.topPerformer = top.name;
    }
  }
  
  // Get this week's violations
  if (logSheet && logSheet.getLastRow() > 1) {
    const logs = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 12).getValues();
    const oneWeekAgo = new Date();
    oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
    
    logs.forEach(row => {
      if (row[3] && new Date(row[0]) >= oneWeekAgo) {
        const emp = row[3];
        if (!report.violationSummary[emp]) report.violationSummary[emp] = 0;
        report.violationSummary[emp]++;
      }
    });
  }
  
  return { success: true, report };
}

// ============================================================
// Reset monthly scores (run at start of month)
// ============================================================
function resetMonthlyScores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scoresSheet = ss.getSheetByName(SHEETS.SCORES);
  if (!scoresSheet || scoresSheet.getLastRow() <= 1) return;
  
  const data = scoresSheet.getRange(2, 1, scoresSheet.getLastRow() - 1, 6).getValues();
  data.forEach((row, i) => {
    if (row[0]) {
      scoresSheet.getRange(i + 2, 3).setValue(100);
      scoresSheet.getRange(i + 2, 4).setValue(0);
      scoresSheet.getRange(i + 2, 5).setValue(500);
      scoresSheet.getRange(i + 2, 6).setValue(new Date());
    }
  });
  
  // Reset supervisor score
  const supSheet = ss.getSheetByName(SHEETS.SUPERVISOR);
  if (supSheet && supSheet.getLastRow() > 1) {
    const supData = supSheet.getRange(2, 1, supSheet.getLastRow() - 1, 6).getValues();
    supData.forEach((row, i) => {
      if (row[0] === "إجمالي") {
        supSheet.getRange(i + 2, 4).setValue(0);
        supSheet.getRange(i + 2, 5).setValue(0); // رهف تبدأ من 0 كل شهر
      }
    });
  }
  
  return { success: true, message: "تم إعادة تعيين الأرصدة الشهرية" };
}

// ============================================================
// إصلاح شريحة تقييم_المشرفة — شغّلها مرة واحدة فقط
// ============================================================
function fixSupervisorSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sup = ss.getSheetByName(SHEETS.SUPERVISOR);

  // احذف القديمة وأنشئ جديدة نظيفة
  if (sup) ss.deleteSheet(sup);
  sup = ss.insertSheet(SHEETS.SUPERVISOR);
  sup.setRightToLeft(true);

  // ── عرض الأعمدة ──
  sup.setColumnWidth(1, 120); // التاريخ
  sup.setColumnWidth(2, 280); // المعيار
  sup.setColumnWidth(3, 140); // نوع التقييم
  sup.setColumnWidth(4, 120); // الخصم/الكسب
  sup.setColumnWidth(5, 120); // رصيد رهف
  sup.setColumnWidth(6, 200); // ملاحظات

  // ── رأس الجدول ──
  sup.setRowHeight(1, 36);
  const headers = ["التاريخ", "المعيار", "نوع التقييم", "الخصم/الكسب", "رصيد رهف", "ملاحظات"];
  const headerColors = ["#D4A843", "#22D3EE", "#A78BFA", "#EF4444", "#10B981", "#64748B"];
  headers.forEach((h, i) => {
    const cell = sup.getRange(1, i + 1);
    cell.setValue(h);
    cell.setBackground("#111827");
    cell.setFontColor(headerColors[i]);
    cell.setFontSize(11);
    cell.setFontWeight("bold");
    cell.setFontFamily("Cairo");
    cell.setHorizontalAlignment("center");
    cell.setVerticalAlignment("middle");
    cell.setBorder(false, false, true, false, false, false, "#D4A843",
                   SpreadsheetApp.BorderStyle.MEDIUM);
  });

  // ── صف رصيد رهف الأولي ──
  sup.setRowHeight(2, 28);
  const initData = ["إجمالي", "رصيد المشرفة", "رهف", 0, 0, "يُحدَّث تلقائياً من واجهة المدير"];
  initData.forEach((val, i) => {
    const cell = sup.getRange(2, i + 1);
    cell.setValue(val);
    cell.setBackground("#161D2E");
    cell.setFontColor(i === 4 ? "#A78BFA" : i === 3 ? "#D4A843" : "#E2E8F0");
    cell.setFontSize(11);
    cell.setFontWeight(i === 4 ? "bold" : "normal");
    cell.setFontFamily("Cairo");
    cell.setHorizontalAlignment("center");
    cell.setVerticalAlignment("middle");
  });

  // تجميد الصف الأول
  sup.setFrozenRows(1);
  sup.setTabColor("#8B5CF6");
  sup.setHiddenGridlines(false);

  SpreadsheetApp.flush();
  return { success: true, message: "✅ تم إصلاح شريحة تقييم_المشرفة" };
}

// ============================================================
// SETUP DASHBOARD — ينشئ جميع الشرائح بالتنسيق الكامل
// شغّل هذه الدالة مرة واحدة بعد initSheets()
// ============================================================
function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── تأكد من وجود الشيتات الأساسية أولاً ──
  initSheets();

  // ── إنشاء / إعادة بناء شريحة لوحة الأداء ──
  _buildDashboardSheet(ss);

  // ── إنشاء / إعادة بناء شريحة تذكير المعايير ──
  _buildCriteriaSheet(ss);

  // ── ترتيب الشرائح ──
  _reorderSheets(ss);

  SpreadsheetApp.flush();
  return { success: true, message: "✅ تم إنشاء لوحة الأداء وشريحة المعايير بنجاح" };
}

// ────────────────────────────────────────────────────────────
// لوحة الأداء الرئيسية
// ────────────────────────────────────────────────────────────
function _buildDashboardSheet(ss) {
  const SHEET_NAME = "📊 لوحة الأداء";

  // احذف القديمة إن وُجدت وأنشئ جديدة
  let old = ss.getSheetByName(SHEET_NAME);
  if (old) ss.deleteSheet(old);
  const ws = ss.insertSheet(SHEET_NAME);
  ws.setRightToLeft(true);

  // ── أبعاد الأعمدة ──
  ws.setColumnWidth(1, 30);   // هامش
  ws.setColumnWidth(2, 170);  // اسم
  ws.setColumnWidth(3, 120);  // مسار
  ws.setColumnWidth(4, 100);  // رصيد
  ws.setColumnWidth(5, 110);  // مخصوم
  ws.setColumnWidth(6, 120);  // حافز
  ws.setColumnWidth(7, 160);  // شريط
  ws.setColumnWidth(8, 130);  // تقييم
  ws.setColumnWidth(9, 110);  // حالة
  ws.setColumnWidth(10, 30);  // هامش

  // إخفاء خطوط الشبكة
  ws.setHiddenGridlines(true);

  // ── تلوين خلفية كاملة ──
  ws.getRange(1, 1, 60, 10).setBackground("#0A0E1A");

  // ── صف العنوان (row 2) ──
  ws.setRowHeight(1, 12);
  ws.setRowHeight(2, 65);
  ws.setRowHeight(3, 12);

  const headerRange = ws.getRange(2, 2, 1, 8);
  headerRange.merge();
  headerRange.setValue("🏆  لوحة أداء مركز الاتصال  |  APEXCARE");
  headerRange.setBackground("#111827");
  headerRange.setFontColor("#F0C96A");
  headerRange.setFontSize(20);
  headerRange.setFontWeight("bold");
  headerRange.setFontFamily("Cairo");
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");

  // ── صف الوصف (row 4) ──
  ws.setRowHeight(4, 30);
  const subRange = ws.getRange(4, 2, 1, 8);
  subRange.merge();
  subRange.setValue("نظام نقاط التحفيز الشهري — كل نقطة = 5 ريال  ✨");
  subRange.setBackground("#0A0E1A");
  subRange.setFontColor("#64748B");
  subRange.setFontSize(11);
  subRange.setFontFamily("Cairo");
  subRange.setHorizontalAlignment("center");
  subRange.setVerticalAlignment("middle");

  ws.setRowHeight(5, 10);

  // ── رؤوس الأعمدة (row 6) ──
  ws.setRowHeight(6, 36);
  const colHeaders = ["#", "اسم الموظفة", "المسار", "الرصيد\n(من 100)",
                      "المخصوم", "الحافز (ريال)", "شريط الأداء", "التقييم", "الحالة"];
  colHeaders.forEach((h, i) => {
    const cell = ws.getRange(6, i + 2);
    cell.setValue(h);
    cell.setBackground("#111827");
    cell.setFontColor("#D4A843");
    cell.setFontSize(10);
    cell.setFontWeight("bold");
    cell.setFontFamily("Cairo");
    cell.setHorizontalAlignment("center");
    cell.setVerticalAlignment("middle");
    cell.setWrap(true);
    cell.setBorder(false, false, true, false, false, false, "#D4A843", SpreadsheetApp.BorderStyle.MEDIUM);
  });

  // ── بيانات الموظفات ──
  const employees = [
    ["مرح",   "📞 مكالمات"],
    ["هيلة",  "📞 مكالمات"],
    ["تهاني", "🔀 مشترك"],
    ["جنان",  "🔀 مشترك"],
    ["أسيل",  "💬 رسائل"],
    ["حور",   "💬 رسائل"],
    ["سارة",  "💬 رسائل"],
    ["طيف",   "💬 رسائل"],
    ["روان",  "💬 رسائل"],
  ];

  const rankIcons  = ["🥇", "🥈", "🥉", "4", "5", "6", "7", "8", "9"];
  const rankColors = ["#FFD700", "#C0C0C0", "#CD7F32", "#E2E8F0", "#E2E8F0", "#E2E8F0", "#E2E8F0", "#E2E8F0", "#E2E8F0"];
  const rowBgs     = ["#1A1500", "#141414", "#140E00", "#161D2E", "#111827", "#161D2E", "#111827", "#161D2E", "#111827"];
  const START_ROW  = 7;
  const SCORES_REF = "الأرصدة"; // اسم شيت الأرصدة

  employees.forEach(([name, track], idx) => {
    const row   = START_ROW + idx;
    const bg    = rowBgs[idx];
    const rclr  = rankColors[idx];

    ws.setRowHeight(row, 32);

    // تلوين الصف
    ws.getRange(row, 1, 1, 10).setBackground(bg);

    // col 2: اسم + رتبة
    const nameCell = ws.getRange(row, 2);
    nameCell.setValue(`${rankIcons[idx]}  ${name}`);
    nameCell.setBackground(bg).setFontColor(rclr).setFontSize(12)
            .setFontWeight("bold").setFontFamily("Cairo")
            .setHorizontalAlignment("right").setVerticalAlignment("middle");

    // col 3: المسار
    const tClr = track.includes("مكالمات") ? "#22D3EE" :
                 track.includes("رسائل")   ? "#A78BFA" : "#F0C96A";
    ws.getRange(row, 3).setValue(track).setBackground(bg).setFontColor(tClr)
      .setFontSize(10).setFontFamily("Cairo")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

    // col 4: الرصيد — formula من شيت الأرصدة
    const balCell = ws.getRange(row, 4);
    balCell.setFormula(`=IFERROR(VLOOKUP("${name}",الأرصدة!$A:$C,3,0),100)`);
    balCell.setBackground(bg).setFontColor("#10B981").setFontSize(14)
           .setFontWeight("bold").setFontFamily("Cairo")
           .setHorizontalAlignment("center").setVerticalAlignment("middle")
           .setNumberFormat("0");

    // col 5: المخصوم
    const dedCell = ws.getRange(row, 5);
    dedCell.setFormula(`=IFERROR(VLOOKUP("${name}",الأرصدة!$A:$D,4,0),0)`);
    dedCell.setBackground(bg).setFontColor("#EF4444").setFontSize(11)
           .setFontFamily("Cairo")
           .setHorizontalAlignment("center").setVerticalAlignment("middle")
           .setNumberFormat("0");

    // col 6: الحافز
    const bonCell = ws.getRange(row, 6);
    bonCell.setFormula(`=D${row}*5`);
    bonCell.setBackground(bg).setFontColor("#F0C96A").setFontSize(12)
           .setFontWeight("bold").setFontFamily("Cairo")
           .setHorizontalAlignment("center").setVerticalAlignment("middle")
           .setNumberFormat('"﷼ "#,##0');

    // col 7: شريط بصري
    const barCell = ws.getRange(row, 7);
    barCell.setFormula(`=REPT("█",ROUND(D${row}/10,0))&REPT("░",10-ROUND(D${row}/10,0))`);
    const barClr = idx < 3 ? "#10B981" : idx < 6 ? "#0EA5B0" : "#F59E0B";
    barCell.setBackground(bg).setFontColor(barClr).setFontSize(11)
           .setFontFamily("Courier New")
           .setHorizontalAlignment("center").setVerticalAlignment("middle");

    // col 8: التقييم
    const ratCell = ws.getRange(row, 8);
    ratCell.setFormula(`=IF(D${row}>=90,"⭐ ممتاز",IF(D${row}>=75,"✅ جيد جداً",IF(D${row}>=60,"👍 جيد","⚠️ يحتاج تحسين")))`);
    ratCell.setBackground(bg).setFontColor("#E2E8F0").setFontSize(10)
           .setFontFamily("Cairo")
           .setHorizontalAlignment("center").setVerticalAlignment("middle");

    // col 9: الحالة
    const stCell = ws.getRange(row, 9);
    stCell.setFormula(`=IF(D${row}>=80,"✅ آمن",IF(D${row}>=60,"⡷ تنبيه","🚨 خطر"))`);
    stCell.setBackground(bg).setFontColor("#E2E8F0").setFontSize(10)
          .setFontFamily("Cairo")
          .setHorizontalAlignment("center").setVerticalAlignment("middle");

    // حد سفلي خفيف
    ws.getRange(row, 2, 1, 8)
      .setBorder(false, false, true, false, false, false, "#1E2D4A", SpreadsheetApp.BorderStyle.SOLID);
  });

  const LAST_DATA_ROW = START_ROW + employees.length - 1;

  // ── صف الإجماليات ──
  ws.setRowHeight(LAST_DATA_ROW + 1, 10);
  const totRow = LAST_DATA_ROW + 2;
  ws.setRowHeight(totRow, 38);
  ws.getRange(totRow, 1, 1, 10).setBackground("#0A0E1A");

  const totLabel = ws.getRange(totRow, 2, 1, 2);
  totLabel.merge().setValue("📊 الإجماليات");
  totLabel.setBackground("#111827").setFontColor("#D4A843").setFontSize(11)
          .setFontWeight("bold").setFontFamily("Cairo")
          .setHorizontalAlignment("center").setVerticalAlignment("middle");

  const avgCell = ws.getRange(totRow, 4);
  avgCell.setFormula(`=AVERAGE(D${START_ROW}:D${LAST_DATA_ROW})`);
  avgCell.setBackground("#111827").setFontColor("#10B981").setFontSize(13)
         .setFontWeight("bold").setFontFamily("Cairo")
         .setHorizontalAlignment("center").setVerticalAlignment("middle")
         .setNumberFormat("0.0")
         .setBorder(true, false, false, false, false, false, "#D4A843", SpreadsheetApp.BorderStyle.MEDIUM);

  const totBonus = ws.getRange(totRow, 6);
  totBonus.setFormula(`=SUM(F${START_ROW}:F${LAST_DATA_ROW})`);
  totBonus.setBackground("#111827").setFontColor("#F0C96A").setFontSize(14)
          .setFontWeight("bold").setFontFamily("Cairo")
          .setHorizontalAlignment("center").setVerticalAlignment("middle")
          .setNumberFormat('"﷼ "#,##0')
          .setBorder(true, false, false, false, false, false, "#D4A843", SpreadsheetApp.BorderStyle.MEDIUM);

  const topCell = ws.getRange(totRow, 7, 1, 3);
  topCell.merge();
  topCell.setFormula(`=CONCATENATE("🏆 أعلى أداء: ",INDEX(B${START_ROW}:B${LAST_DATA_ROW},MATCH(MAX(D${START_ROW}:D${LAST_DATA_ROW}),D${START_ROW}:D${LAST_DATA_ROW},0)))`);
  topCell.setBackground("#111827").setFontColor("#F0C96A").setFontSize(11)
         .setFontWeight("bold").setFontFamily("Cairo")
         .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // ── قسم المشرفة رهف ──
  const supTitle = totRow + 2;
  ws.setRowHeight(supTitle, 36);
  ws.getRange(supTitle, 1, 1, 10).setBackground("#0A0E1A");

  const supHeader = ws.getRange(supTitle, 2, 1, 8);
  supHeader.merge().setValue("👑  المشرفة رهف — نقاط المتابعة والإشراف");
  supHeader.setBackground("#111827").setFontColor("#F0C96A").setFontSize(13)
           .setFontWeight("bold").setFontFamily("Cairo")
           .setHorizontalAlignment("center").setVerticalAlignment("middle")
           .setBorder(false, false, true, false, false, false, "#8B5CF6", SpreadsheetApp.BorderStyle.MEDIUM);

  const supData = supTitle + 1;
  ws.setRowHeight(supData, 32);
  ws.getRange(supData, 1, 1, 10).setBackground("#161D2E");

  ws.getRange(supData, 2).setValue("رهف").setBackground("#161D2E")
    .setFontColor("#A78BFA").setFontSize(13).setFontWeight("bold")
    .setFontFamily("Cairo").setHorizontalAlignment("right").setVerticalAlignment("middle");

  ws.getRange(supData, 3).setValue("🎯 مشرفة").setBackground("#161D2E")
    .setFontColor("#8B5CF6").setFontSize(10).setFontFamily("Cairo")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  const rBalCell = ws.getRange(supData, 4);
  rBalCell.setFormula(`=IFERROR(VLOOKUP("إجمالي",تقييم_المشرفة!$A:$E,5,0),0)`);
  rBalCell.setBackground("#161D2E").setFontColor("#A78BFA").setFontSize(14)
          .setFontWeight("bold").setFontFamily("Cairo")
          .setHorizontalAlignment("center").setVerticalAlignment("middle")
          .setNumberFormat("0");

  const rPtsCell = ws.getRange(supData, 5);
  rPtsCell.setFormula(`=IFERROR(VLOOKUP("إجمالي",تقييم_المشرفة!$A:$D,4,0),0)`);
  rPtsCell.setBackground("#161D2E").setFontColor("#D4A843").setFontSize(12)
          .setFontFamily("Cairo")
          .setHorizontalAlignment("center").setVerticalAlignment("middle");

  const rBonCell = ws.getRange(supData, 6);
  rBonCell.setFormula(`=D${supData}*5`);
  rBonCell.setBackground("#161D2E").setFontColor("#F0C96A").setFontSize(13)
          .setFontWeight("bold").setFontFamily("Cairo")
          .setHorizontalAlignment("center").setVerticalAlignment("middle")
          .setNumberFormat('"﷼ "#,##0');

  const rBarCell = ws.getRange(supData, 7);
  rBarCell.setFormula(`=REPT("█",ROUND(D${supData}/10,0))&REPT("░",10-ROUND(D${supData}/10,0))`);
  rBarCell.setBackground("#161D2E").setFontColor("#8B5CF6").setFontSize(11)
          .setFontFamily("Courier New")
          .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // ── Conditional Formatting على عمود الرصيد ──
  const balRange = ws.getRange(`D${START_ROW}:D${LAST_DATA_ROW}`);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#10B981", SpreadsheetApp.InterpolationType.NUMBER, "100")
    .setGradientMidpointWithValue("#F59E0B", SpreadsheetApp.InterpolationType.NUMBER, "70")
    .setGradientMinpointWithValue("#EF4444", SpreadsheetApp.InterpolationType.NUMBER, "0")
    .setRanges([balRange])
    .build();
  ws.setConditionalFormatRules([rule]);

  // تجميد الصفوف فقط (الأعمدة تتعارض مع الخلايا المدمجة)
  ws.setFrozenRows(6);

  // إخفاء الأعمدة الهامشية
  ws.hideColumn(ws.getRange("A:A"));
  ws.hideColumn(ws.getRange("J:J"));

  // لون التبويب
  ws.setTabColor("#D4A843");
}

// ────────────────────────────────────────────────────────────
// شريحة تذكير المعايير التشجيعية
// ────────────────────────────────────────────────────────────
function _buildCriteriaSheet(ss) {
  const SHEET_NAME = "💡 تذكير المعايير";

  let old = ss.getSheetByName(SHEET_NAME);
  if (old) ss.deleteSheet(old);
  const ws = ss.insertSheet(SHEET_NAME);
  ws.setRightToLeft(true);
  ws.setHiddenGridlines(true);

  ws.setColumnWidth(1, 20);
  ws.setColumnWidth(2, 45);
  ws.setColumnWidth(3, 380);
  ws.setColumnWidth(4, 90);
  ws.setColumnWidth(5, 20);

  // خلفية
  ws.getRange(1, 1, 130, 5).setBackground("#0A0E1A");

  // ── العنوان الرئيسي ──
  ws.setRowHeight(1, 12);
  ws.setRowHeight(2, 60);
  const title = ws.getRange(2, 2, 1, 3);
  title.merge().setValue("💡  تذكير بمعايير الجودة — طريقك للـ 100 نقطة!");
  title.setBackground("#111827").setFontColor("#F0C96A").setFontSize(18)
       .setFontWeight("bold").setFontFamily("Cairo")
       .setHorizontalAlignment("center").setVerticalAlignment("middle");

  ws.setRowHeight(3, 30);
  const sub = ws.getRange(3, 2, 1, 3);
  sub.merge().setValue("كل نقطة تحافظ عليها = 5 ريال في جيبك 💰  |  الاحتراف عادة وليس صدفة ⭐");
  sub.setBackground("#0A0E1A").setFontColor("#22D3EE").setFontSize(11)
     .setFontFamily("Cairo")
     .setHorizontalAlignment("center").setVerticalAlignment("middle");

  const sections = [
    {
      title: "📞 المكالمات — 55 نقطة",
      headerBg: "#0EA5B0", textColor: "#22D3EE", rowBg: "#061A1C",
      items: [
        ["الصياغة التعريفية عند بداية كل مكالمة",          14, "«حياك الله، معك نورة من عيادات أبكس كير، كيف أقدر أخدمك؟» 🎤"],
        ["التحقق من ملف المراجع برقم الجوال فوراً",          11, "ابحث قبل أن تسأل — احترم وقت المراجع ⏱️"],
        ["الرد على جميع المكالمات خلال أوقات العمل",         10, "كل مكالمة فائتة = فرصة ضائعة للعيادة 📵"],
        ["تقديم خدمة مكتملة حتى الإغلاق",                   8,  "لا تترك المراجع يتساءل — أغلق الطلب بوضوح 🎯"],
        ["استخدام ألفاظ مهنية (تفضلي، أبشري، سمي)",         6,  "اللغة المهنية تعكس صورة العيادة 🌟"],
        ["إقفال المكالمة بعبارة ختامية مهذبة",               6,  "«شكراً لتواصلك معنا، نتمنى لك يوماً سعيداً» 🌸"],
      ]
    },
    {
      title: "💬 الرسائل — 55 نقطة",
      headerBg: "#8B5CF6", textColor: "#A78BFA", rowBg: "#0E0618",
      items: [
        ["الرد على الرسائل خلال 10 دقائق كحد أقصى",         14, "السرعة في الرد = رضا المريض = سمعة العيادة ⚡"],
        ["متابعة المحادثات وعدم تركها مفتوحة",               11, "محادثة مفتوحة = مراجع غير مخدوم — تابع دائماً 👁️"],
        ["التحقق من ملف المراجع فور فتح المحادثة",           10, "ابدأ بالبحث برقم الجوال — قبل أي سؤال 🔍"],
        ["التفاعل مع المراجع خلال دقيقة من الفتح",           8,  "أظهر حضورك — المراجع ينتظرك الآن 💬"],
        ["إنهاء المحادثة بعبارة ختامية مهذبة",               6,  "الختام الجميل يُكمل تجربة المراجع 🌺"],
        ["الالتزام بالاختصارات المعتمدة في الرسائل",          6,  "التوحيد في اللغة يعكس احترافية الفريق 📋"],
      ]
    },
    {
      title: "📁 فتح الملف — 25 نقطة",
      headerBg: "#B45309", textColor: "#F59E0B", rowBg: "#1A1000",
      items: [
        ["إدخال الاسم الرباعي بشكل صحيح",                    7, "الاسم الكامل أساس كل شيء — لا تتساهل في الدقة 📝"],
        ["إدخال رقم الهوية بشكل صحيح",                       6, "تحقق مرتين قبل الحفظ — الخطأ يكلف وقتاً 🔢"],
        ["إدخال تاريخ الميلاد بشكل صحيح",                    6, "بيانات دقيقة = ملف طبي موثوق 📅"],
        ["إدخال العنوان بشكل صحيح",                           4, "العنوان الصحيح يسهّل التواصل مستقبلاً 🗺️"],
        ["التحري بالدقة أثناء إدخال جميع البيانات",           2, "الدقة عادة — اجعلها جزءاً من شخصيتك ✨"],
      ]
    },
    {
      title: "📅 المواعيد — 20 نقطة",
      headerBg: "#047857", textColor: "#10B981", rowBg: "#031A0E",
      items: [
        ["تعويض المواعيد الملغاة عند توفر بديل",              6, "كل موعد ملغى فرصة لمراجع آخر — لا تتركها 🔄"],
        ["تأكيد المواعيد مع المراجعين",                       5, "التأكيد يقلل الغياب ويحترم وقت الطبيب ☑️"],
        ["تعبئة المواعيد عند توفر بديل",                      5, "الجدول الممتلئ = عيادة ناجحة 📈"],
        ["ترتيب جدول الدكتور (مواعيد متتالية)",               4, "الجدول المنظم يعكس كفاءتك الإدارية 🗓️"],
      ]
    }
  ];

  let curRow = 5;

  sections.forEach(sec => {
    curRow++; // فراغ
    ws.setRowHeight(curRow, 8);
    curRow++;

    // عنوان القسم
    ws.setRowHeight(curRow, 38);
    const secHeader = ws.getRange(curRow, 2, 1, 3);
    secHeader.merge().setValue(sec.title);
    secHeader.setBackground(sec.headerBg).setFontColor("#FFFFFF")
             .setFontSize(13).setFontWeight("bold").setFontFamily("Cairo")
             .setHorizontalAlignment("right").setVerticalAlignment("middle");
    curRow++;

    // رؤوس أعمدة القسم
    ws.setRowHeight(curRow, 26);
    ws.getRange(curRow, 2).setValue("#").setBackground("#111827")
      .setFontColor("#64748B").setFontSize(9).setFontWeight("bold")
      .setFontFamily("Cairo").setHorizontalAlignment("center").setVerticalAlignment("middle");
    ws.getRange(curRow, 3).setValue("المعيار").setBackground("#111827")
      .setFontColor("#D4A843").setFontSize(10).setFontWeight("bold")
      .setFontFamily("Cairo").setHorizontalAlignment("right").setVerticalAlignment("middle");
    ws.getRange(curRow, 4).setValue("النقاط").setBackground("#111827")
      .setFontColor("#22D3EE").setFontSize(10).setFontWeight("bold")
      .setFontFamily("Cairo").setHorizontalAlignment("center").setVerticalAlignment("middle");
    curRow++;

    sec.items.forEach(([criterion, pts, tip], idx) => {
      // صف المعيار
      ws.setRowHeight(curRow, 26);
      ws.getRange(curRow, 1, 1, 5).setBackground(sec.rowBg);
      ws.getRange(curRow, 2).setValue(`${idx + 1}`).setBackground(sec.rowBg)
        .setFontColor(sec.textColor).setFontSize(10).setFontWeight("bold")
        .setFontFamily("Cairo").setHorizontalAlignment("center").setVerticalAlignment("middle");
      ws.getRange(curRow, 3).setValue(`  ✅ ${criterion}`).setBackground(sec.rowBg)
        .setFontColor("#E2E8F0").setFontSize(10).setFontFamily("Cairo")
        .setHorizontalAlignment("right").setVerticalAlignment("middle");
      ws.getRange(curRow, 4).setValue(pts).setBackground(sec.rowBg)
        .setFontColor("#F0C96A").setFontSize(12).setFontWeight("bold")
        .setFontFamily("Cairo").setHorizontalAlignment("center").setVerticalAlignment("middle");
      ws.getRange(curRow, 2, 1, 3)
        .setBorder(false, false, true, false, false, false, "#1E2D4A", SpreadsheetApp.BorderStyle.SOLID);
      curRow++;

      // صف التلميح
      ws.setRowHeight(curRow, 20);
      ws.getRange(curRow, 1, 1, 5).setBackground("#0A0E1A");
      const tipCell = ws.getRange(curRow, 3, 1, 2);
      tipCell.merge().setValue(`    💡 ${tip}`).setBackground("#0A0E1A")
             .setFontColor("#64748B").setFontSize(9).setFontFamily("Cairo")
             .setHorizontalAlignment("right").setVerticalAlignment("middle")
             .setFontStyle("italic");
      curRow++;
    });
  });

  // ── بانر تحفيزي ختامي ──
  curRow += 2;
  ws.setRowHeight(curRow, 50);
  const motivBanner = ws.getRange(curRow, 2, 1, 3);
  motivBanner.merge();
  motivBanner.setValue("🌟  أنتِ قادرة على الـ 100 نقطة!  كل التزام بمعيار = استثمار في راتبك ومستقبلك  🌟");
  motivBanner.setBackground("#D4A843").setFontColor("#0A0E1A")
             .setFontSize(13).setFontWeight("bold").setFontFamily("Cairo")
             .setHorizontalAlignment("center").setVerticalAlignment("middle")
             .setWrap(true);
  curRow++;

  ws.setRowHeight(curRow, 34);
  const subBanner = ws.getRange(curRow, 2, 1, 3);
  subBanner.merge().setValue("الاحتراف ليس خياراً — هو هويتك في APEXCARE  💎");
  subBanner.setBackground("#111827").setFontColor("#22D3EE")
           .setFontSize(11).setFontFamily("Cairo")
           .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // إخفاء الأعمدة الهامشية
  ws.hideColumn(ws.getRange("A:A"));
  ws.hideColumn(ws.getRange("E:E"));

  ws.setTabColor("#0EA5B0");
}

// ────────────────────────────────────────────────────────────
// ترتيب الشرائح
// ────────────────────────────────────────────────────────────
function _reorderSheets(ss) {
  const order = [
    "📊 لوحة الأداء",
    "💡 تذكير المعايير",
    SHEETS.SCORES,
    SHEETS.SUPERVISOR,
    SHEETS.LOG
  ];

  order.forEach((name, idx) => {
    const sheet = ss.getSheetByName(name);
    if (sheet) ss.setActiveSheet(sheet) && ss.moveActiveSheet(idx + 1);
  });
}
