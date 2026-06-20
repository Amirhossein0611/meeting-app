/* =========================================================
   MEETING APP BACKEND - GOOGLE APPS SCRIPT
   Features:
   - Receive meeting data from Vercel
   - Save meeting to Google Sheets
   - Generate AI summary using OpenAI
   - Generate PDF report
   - Save tasks
========================================================= */


/* =========================
   CONFIG
========================= */

const CONFIG = {
  MEETINGS_SHEET: "Meetings",
  TASKS_SHEET: "Tasks",
  TIMEZONE: "Asia/Tehran",

  // اگر می‌خواهی کلید را مستقیم داخل کد بگذاری، اینجا بگذار
  // حتماً داخل کوتیشن باشد
  OPENAI_API_KEY: "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",

  OPENAI_MODEL: "gpt-4o-mini"
};

const MEETING_HEADERS = [
  "Meeting ID",
  "Created At",
  "Title",
  "Summary",
  "Transcript",
  "PDF URL",
  "PDF Download URL",
  "PDF File ID"
];

const TASK_HEADERS = [
  "Task ID",
  "Meeting ID",
  "Title",
  "Owner",
  "Priority",
  "Start Date",
  "End Date",
  "Progress",
  "Status"
];


/* =========================
   HTTP ENTRY POINTS
========================= */

function doGet(e) {
  ensureSheets_();

  return ContentService
    .createTextOutput("Meeting backend is running successfully.")
    .setMimeType(ContentService.MimeType.TEXT);
}


function doPost(e) {
  try {
    ensureSheets_();

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("هیچ داده‌ای از سمت فرانت‌اند دریافت نشد.");
    }

    const data = JSON.parse(e.postData.contents);
    const result = saveMeeting(data);

    return jsonResponse_(result);

  } catch (err) {
    return jsonResponse_({
      success: false,
      error: err.message
    });
  }
}


/* =========================
   MAIN SAVE FUNCTION
========================= */

function saveMeeting(data) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    ensureSheets_();

    if (!data) {
      throw new Error("داده جلسه خالی است.");
    }

    const title = String(data.title || "").trim();
    const transcript = String(data.transcript || "").trim();
    let summary = String(data.summary || "").trim();

    const tasks = Array.isArray(data.tasks) ? data.tasks : [];

    if (!title) {
      throw new Error("عنوان جلسه الزامی است.");
    }

    // اگر خلاصه از فرانت نیامده باشد، با AI تولید می‌شود
    if (!summary && transcript) {
      summary = summarizeText(transcript);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const meetingSheet = ss.getSheetByName(CONFIG.MEETINGS_SHEET);
    const taskSheet = ss.getSheetByName(CONFIG.TASKS_SHEET);

    const meetingId = Utilities.getUuid();
    const createdAt = new Date();

    appendRowByHeaders_(meetingSheet, MEETING_HEADERS, {
      "Meeting ID": meetingId,
      "Created At": createdAt,
      "Title": title,
      "Summary": summary,
      "Transcript": transcript,
      "PDF URL": "",
      "PDF Download URL": "",
      "PDF File ID": ""
    });

    // ذخیره تسک‌ها
    tasks.forEach(function(task) {
      const taskTitle = String(task.title || "").trim();

      if (!taskTitle) return;

      appendRowByHeaders_(taskSheet, TASK_HEADERS, {
        "Task ID": Utilities.getUuid(),
        "Meeting ID": meetingId,
        "Title": taskTitle,
        "Owner": task.owner || "",
        "Priority": task.priority || "متوسط",
        "Start Date": task.start || task.startDate || "",
        "End Date": task.end || task.endDate || "",
        "Progress": task.progress || 0,
        "Status": task.status || "شروع نشده"
      });
    });

    // ساخت PDF
    const pdfInfo = generateMeetingPDF(meetingId);

    // ذخیره اطلاعات PDF در شیت
    updateMeetingPdfInfo_(meetingId, pdfInfo);

    return {
      success: true,
      message: "جلسه با موفقیت ذخیره شد.",
      meetingId: meetingId,
      summary: summary,
      pdfUrl: pdfInfo.viewUrl,
      downloadUrl: pdfInfo.downloadUrl,
      fileId: pdfInfo.fileId
    };

  } catch (err) {
    return {
      success: false,
      error: err.message
    };

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}


/* =========================
   AI SUMMARY
========================= */

function summarizeText(text) {
  if (!text) return "";

  const OPENAI_API_KEY = CONFIG.OPENAI_API_KEY;

  if (!OPENAI_API_KEY || OPENAI_API_KEY === "PUT_YOUR_OPENAI_API_KEY_HERE") {
    throw new Error("کلید OpenAI تنظیم نشده است. مقدار OPENAI_API_KEY را در CONFIG وارد کن.");
  }

  const url = "https://api.openai.com/v1/chat/completions";

  const payload = {
    model: CONFIG.OPENAI_MODEL,
    messages: [
      {
        role: "system",
        content:
          "تو یک منشی حرفه‌ای جلسات هستی. " +
          "متن جلسه را به زبان فارسی، رسمی، دقیق و ساختاریافته خلاصه کن. " +
          "خروجی را دقیقاً با این ساختار تولید کن:\n\n" +
          "۱- خلاصه مدیریتی\n" +
          "۲- نکات کلیدی\n" +
          "۳- تصمیمات جلسه\n" +
          "۴- اقدامات پیشنهادی\n" +
          "۵- پیگیری‌های لازم\n\n" +
          "اگر تصمیم، اقدام یا پیگیری مشخصی در متن وجود نداشت، بنویس: مورد مشخصی ذکر نشده است."
      },
      {
        role: "user",
        content: text
      }
    ],
    temperature: 0.3
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + OPENAI_API_KEY
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (statusCode !== 200) {
    throw new Error("خطا در ارتباط با OpenAI: " + responseText);
  }

  const result = JSON.parse(responseText);

  if (
    !result.choices ||
    !result.choices[0] ||
    !result.choices[0].message ||
    !result.choices[0].message.content
  ) {
    throw new Error("پاسخ OpenAI نامعتبر است: " + responseText);
  }

  return result.choices[0].message.content.trim();
}


/* =========================
   PDF GENERATION
========================= */

function generateMeetingPDF(meetingId) {
  ensureSheets_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meetingSheet = ss.getSheetByName(CONFIG.MEETINGS_SHEET);
  const taskSheet = ss.getSheetByName(CONFIG.TASKS_SHEET);

  const meetingMap = getHeaderMap_(meetingSheet);
  const taskMap = getHeaderMap_(taskSheet);

  const meetingData = meetingSheet.getDataRange().getValues();
  const taskData = taskSheet.getDataRange().getValues();

  let meetingRow = null;

  for (let i = 1; i < meetingData.length; i++) {
    const row = meetingData[i];
    if (String(row[meetingMap["Meeting ID"] - 1]) === String(meetingId)) {
      meetingRow = row;
      break;
    }
  }

  if (!meetingRow) {
    throw new Error("جلسه برای ساخت PDF پیدا نشد.");
  }

  const relatedTasks = [];

  for (let j = 1; j < taskData.length; j++) {
    const row = taskData[j];
    if (String(row[taskMap["Meeting ID"] - 1]) === String(meetingId)) {
      relatedTasks.push(row);
    }
  }

  const title = meetingRow[meetingMap["Title"] - 1] || "بدون عنوان";
  const createdAt = meetingRow[meetingMap["Created At"] - 1] || "";
  const summary = meetingRow[meetingMap["Summary"] - 1] || "";
  const transcript = meetingRow[meetingMap["Transcript"] - 1] || "";

  const formattedDate = formatDateTime_(createdAt);
  const safeTitle = sanitizeFileName_(title);

  const docName = "صورتجلسه - " + safeTitle;
  const doc = DocumentApp.create(docName);
  const body = doc.getBody();

  body.clear();

  // عنوان اصلی
  const mainTitle = body.appendParagraph("صورتجلسه");
  mainTitle
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph("");

  // اطلاعات جلسه
  body.appendParagraph("اطلاعات جلسه")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const infoTable = body.appendTable([
    ["عنوان جلسه", String(title)],
    ["تاریخ ثبت", String(formattedDate)],
    ["شناسه جلسه", String(meetingId)]
  ]);

  styleTable_(infoTable);

  body.appendParagraph("");

  // خلاصه
  body.appendParagraph("خلاصه جلسه")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  if (summary) {
    appendMultilineParagraph_(body, summary);
  } else {
    body.appendParagraph("خلاصه‌ای ثبت نشده است.");
  }

  body.appendParagraph("");

  // تسک‌ها
  body.appendParagraph("مصوبات و اقدامات")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  if (relatedTasks.length > 0) {
    const taskRows = [
      ["عنوان", "مسئول", "اولویت", "شروع", "پایان", "پیشرفت", "وضعیت"]
    ];

    relatedTasks.forEach(function(row) {
      taskRows.push([
        String(row[taskMap["Title"] - 1] || ""),
        String(row[taskMap["Owner"] - 1] || ""),
        String(row[taskMap["Priority"] - 1] || ""),
        String(row[taskMap["Start Date"] - 1] || ""),
        String(row[taskMap["End Date"] - 1] || ""),
        String(row[taskMap["Progress"] - 1] || ""),
        String(row[taskMap["Status"] - 1] || "")
      ]);
    });

    const taskTable = body.appendTable(taskRows);
    styleTable_(taskTable);

  } else {
    body.appendParagraph("موردی ثبت نشده است.");
  }

  body.appendParagraph("");

  // متن کامل جلسه
  body.appendParagraph("متن کامل جلسه")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  if (transcript) {
    appendMultilineParagraph_(body, transcript);
  } else {
    body.appendParagraph("متن کامل جلسه ثبت نشده است.");
  }

  body.appendParagraph("");

  const footer = body.appendParagraph("این فایل به‌صورت خودکار توسط سامانه مدیریت جلسات تولید شده است.");
  footer
    .setFontSize(9)
    .setItalic(true)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  doc.saveAndClose();

  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = docFile.getAs(MimeType.PDF).setName(docName + ".pdf");

  const pdfFile = DriveApp.createFile(pdfBlob);

  // قابل مشاهده با لینک
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // حذف فایل Google Docs واسط، فقط PDF بماند
  docFile.setTrashed(true);

  return {
    fileId: pdfFile.getId(),
    viewUrl: pdfFile.getUrl(),
    downloadUrl: "https://drive.google.com/uc?export=download&id=" + pdfFile.getId()
  };
}


/* =========================
   UPDATE PDF INFO IN SHEET
========================= */

function updateMeetingPdfInfo_(meetingId, pdfInfo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.MEETINGS_SHEET);

  if (!sheet) {
    throw new Error("Sheet جلسات پیدا نشد.");
  }

  const map = getHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const idColumn = map["Meeting ID"];
  const ids = sheet.getRange(2, idColumn, lastRow - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(meetingId)) {
      const rowNumber = i + 2;

      sheet.getRange(rowNumber, map["PDF URL"]).setValue(pdfInfo.viewUrl);
      sheet.getRange(rowNumber, map["PDF Download URL"]).setValue(pdfInfo.downloadUrl);
      sheet.getRange(rowNumber, map["PDF File ID"]).setValue(pdfInfo.fileId);

      return;
    }
  }
}


/* =========================
   SHEET SETUP
========================= */

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let meetingSheet = ss.getSheetByName(CONFIG.MEETINGS_SHEET);

  if (!meetingSheet) {
    meetingSheet = ss.insertSheet(CONFIG.MEETINGS_SHEET);
    meetingSheet.appendRow(MEETING_HEADERS);
    meetingSheet.setFrozenRows(1);
  } else {
    ensureHeaders_(meetingSheet, MEETING_HEADERS);
  }

  let taskSheet = ss.getSheetByName(CONFIG.TASKS_SHEET);

  if (!taskSheet) {
    taskSheet = ss.insertSheet(CONFIG.TASKS_SHEET);
    taskSheet.appendRow(TASK_HEADERS);
    taskSheet.setFrozenRows(1);
  } else {
    ensureHeaders_(taskSheet, TASK_HEADERS);
  }
}


function ensureHeaders_(sheet, requiredHeaders) {
  const lastColumn = sheet.getLastColumn();

  // اگر شیت کاملاً خالی باشد
  if (lastColumn === 0 || sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    return;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  requiredHeaders.forEach(function(header) {
    if (currentHeaders.indexOf(header) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
    }
  });

  sheet.setFrozenRows(1);
}


function getHeaderMap_(sheet) {
  const lastColumn = sheet.getLastColumn();

  if (lastColumn === 0) {
    throw new Error("هیچ ستونی در Sheet وجود ندارد.");
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const map = {};

  headers.forEach(function(header, index) {
    map[String(header).trim()] = index + 1;
  });

  return map;
}


function appendRowByHeaders_(sheet, headers, obj) {
  ensureHeaders_(sheet, headers);

  const map = getHeaderMap_(sheet);
  const row = new Array(sheet.getLastColumn()).fill("");

  Object.keys(obj).forEach(function(key) {
    if (map[key]) {
      row[map[key] - 1] = obj[key];
    }
  });

  sheet.appendRow(row);
}


/* =========================
   HELPER FUNCTIONS
========================= */

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


function formatDateTime_(value) {
  if (!value) return "";

  let date;

  if (value instanceof Date) {
    date = value;
  } else {
    date = new Date(value);
  }

  if (isNaN(date.getTime())) {
    return String(value);
  }

  return Utilities.formatDate(date, CONFIG.TIMEZONE, "yyyy/MM/dd HH:mm");
}


function sanitizeFileName_(name) {
  return String(name || "meeting")
    .replace(/[\\\/:*?"<>|#%{}~&]/g, "-")
    .substring(0, 80);
}


function appendMultilineParagraph_(body, text) {
  const lines = String(text || "").split(/\n+/);

  lines.forEach(function(line) {
    const cleanLine = line.trim();

    if (cleanLine) {
      body.appendParagraph(cleanLine);
    }
  });
}


function styleTable_(table) {
  table.setBorderWidth(1);

  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);

    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      const text = cell.editAsText();

      text.setFontFamily("Arial");
      text.setFontSize(10);

      if (r === 0) {
        text.setBold(true);
        cell.setBackgroundColor("#E8F0FE");
      }
    }
  }
}


/* =========================
   OPTIONAL TEST FUNCTION
   این تابع را دستی Run کن برای تست
========================= */

function testSaveMeeting() {
  const sampleData = {
    title: "جلسه تست",
    transcript: "در این جلسه درباره برنامه توسعه محصول، زمان‌بندی انتشار نسخه جدید و تقسیم وظایف بین اعضای تیم صحبت شد. تصمیم گرفته شد نسخه اولیه تا پایان ماه آماده شود.",
    tasks: [
      {
        title: "آماده‌سازی نسخه اولیه محصول",
        owner: "علی",
        priority: "بالا",
        start: "2026/06/20",
        end: "2026/06/30"
      },
      {
        title: "بررسی نیازمندی‌های مشتری",
        owner: "سارا",
        priority: "متوسط",
        start: "2026/06/21",
        end: "2026/06/25"
      }
    ]
  };

  const result = saveMeeting(sampleData);
  Logger.log(result);
}
