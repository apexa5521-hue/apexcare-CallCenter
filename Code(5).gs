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
  "جنان":   { track: "messages" },
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

// ============================================================
// doGet - Serve the HTML interface
// ============================================================
function doGet(e) {
  const page = e.parameter.page || "supervisor";
  const template = HtmlService.createTemplateFromFile("index");
  template.page = page;
  return template.evaluate()
    .setTitle("APEXCARE - نظام الجودة")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
    sup.appendRow(["التاريخ", "المعيار", "نوع التقييم", "الخصم/الكسب", "رصيد رهف", "ملاحظات"]);
    sup.getRange(1, 1, 1, 6).setBackground("#1a1a2e").setFontColor("#ffffff").setFontWeight("bold");
    // Init rahaf score row
    sup.appendRow(["إجمالي", "رصيد المشرفة", "رهف", 0, 0, new Date()]);
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
  let rahafBalance = 100;
  if (supSheet && supSheet.getLastRow() > 1) {
    const supData = supSheet.getRange(2, 1, supSheet.getLastRow() - 1, 6).getValues();
    supData.forEach(row => {
      if (row[0] === "إجمالي") {
        supervisorScore = row[3] || 0;
        rahafBalance = row[4] || 100;
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
      const maxScore = Object.keys(EMPLOYEES).length * 100; // 900
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
        const currentBalance = supSheet.getRange(rowNum, 5).getValue() || 100;
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
        supSheet.getRange(i + 2, 5).setValue(100);
      }
    });
  }
  
  return { success: true, message: "تم إعادة تعيين الأرصدة الشهرية" };
}
