function doGet() {
  ensureSheets_();

  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle("Meeting System")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================
   CONFIG
========================= */

const CONFIG = {
  MEETINGS_SHEET: "Meetings",
  TASKS_SHEET: "Tasks",
  TIMEZONE: Session.getScriptTimeZone() || "Asia/Tehran"
};

const MEETING_HEADERS = [
  "Meeting ID",
  "Created At",
  "Title",
  "Summary",
  "PDF URL",
  "PDF Download URL",
  "PDF File ID",
  "Transcript"
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
   ENSURE SHEETS
========================= */

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let meetingSheet = ss.getSheetByName(CONFIG.MEETINGS_SHEET);

  if (!meetingSheet) {
    meetingSheet = ss.insertSheet(CONFIG.MEETINGS_SHEET);
    meetingSheet.appendRow(MEETING_HEADERS);
    meetingSheet.setFrozenRows(1);
  } else {
    ensureMeetingHeaders_(meetingSheet);
  }

  let taskSheet = ss.getSheetByName(CONFIG.TASKS_SHEET);

  if (!taskSheet) {
    taskSheet = ss.insertSheet(CONFIG.TASKS_SHEET);
    taskSheet.appendRow(TASK_HEADERS);
    taskSheet.setFrozenRows(1);
  } else {
    ensureTaskHeaders_(taskSheet);
  }
}

function ensureMeetingHeaders_(sheet) {
  ensureHeaders_(sheet, MEETING_HEADERS);
}

function ensureTaskHeaders_(sheet) {
  ensureHeaders_(sheet, TASK_HEADERS);
}

function ensureHeaders_(sheet, requiredHeaders) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    return;
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const currentHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  requiredHeaders.forEach(header => {
    if (currentHeaders.indexOf(header) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      currentHeaders.push(header);
    }
  });

  sheet.setFrozenRows(1);
}

function getHeaderMap_(sheet) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  const map = {};

  headers.forEach((header, index) => {
    if (header) {
      map[String(header).trim()] = index + 1;
    }
  });

  return map;
}

function getCellByHeader_(row, headerMap, headerName) {
  const col = headerMap[headerName];

  if (!col) {
    return "";
  }

  return row[col - 1];
}

function appendRowByHeaders_(sheet, requiredHeaders, rowObject) {
  ensureHeaders_(sheet, requiredHeaders);

  const headerMap = getHeaderMap_(sheet);
  const lastColumn = sheet.getLastColumn();
  const row = new Array(lastColumn).fill("");

  Object.keys(rowObject).forEach(key => {
    const col = headerMap[key];

    if (col) {
      row[col - 1] = rowObject[key];
    }
  });

  sheet.appendRow(row);
}

/* =========================
   SAVE MEETING + TASKS
========================= */

function saveMeeting(data) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    ensureSheets_();

    if (!data) {
      throw new Error("داده‌ای برای ثبت دریافت نشد.");
    }

    const title = String(data.title || "").trim();
    const transcript = String(data.transcript || "").trim();
    let summary = String(data.summary || "").trim();

if (!summary && transcript) {
  summary = summarizeText(transcript);
}

    const tasks = Array.isArray(data.tasks) ? data.tasks : [];

    if (!title) {
      throw new Error("عنوان جلسه الزامی است.");
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
      "PDF URL": "",
      "PDF Download URL": "",
      "PDF File ID": "",
      "Transcript": transcript
    });

    tasks.forEach(task => {
      const taskTitle = String(task.title || "").trim();

      if (!taskTitle) return;

      appendRowByHeaders_(taskSheet, TASK_HEADERS, {
        "Task ID": Utilities.getUuid(),
        "Meeting ID": meetingId,
        "Title": taskTitle,
        "Owner": String(task.owner || "").trim(),
        "Priority": String(task.priority || "متوسط").trim(),
        "Start Date": String(task.start || "").trim(),
        "End Date": String(task.end || "").trim(),
        "Progress": 0,
        "Status": "شروع نشده"
      });
    });

    const pdfInfo = generateMeetingPDF(meetingId);

    updateMeetingPdfInfo_(meetingId, pdfInfo);

    return {
      success: true,
      message: "جلسه با موفقیت ثبت شد و PDF ساخته شد.",
      meetingId: meetingId,
      pdfUrl: pdfInfo.viewUrl,
      downloadUrl: pdfInfo.downloadUrl
    };

  } catch (err) {
    return {
      success: false,
      error: err && err.message ? err.message : String(err)
    };

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

/* =========================
   UPDATE PDF INFO IN SHEET
========================= */

function updateMeetingPdfInfo_(meetingId, pdfInfo) {
  ensureSheets_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meetingSheet = ss.getSheetByName(CONFIG.MEETINGS_SHEET);

  const lastRow = meetingSheet.getLastRow();

  if (lastRow < 2) return;

  const headerMap = getHeaderMap_(meetingSheet);

  const meetingIdCol = headerMap["Meeting ID"];
  const pdfUrlCol = headerMap["PDF URL"];
  const pdfDownloadCol = headerMap["PDF Download URL"];
  const pdfFileIdCol = headerMap["PDF File ID"];

  if (!meetingIdCol) return;

  const ids = meetingSheet.getRange(2, meetingIdCol, lastRow - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(meetingId)) {
      const row = i + 2;

      if (pdfUrlCol) {
        meetingSheet.getRange(row, pdfUrlCol).setValue(pdfInfo.viewUrl || "");
      }

      if (pdfDownloadCol) {
        meetingSheet.getRange(row, pdfDownloadCol).setValue(pdfInfo.downloadUrl || "");
      }

      if (pdfFileIdCol) {
        meetingSheet.getRange(row, pdfFileIdCol).setValue(pdfInfo.fileId || "");
      }

      return;
    }
  }
}

/* =========================
   PDF GENERATOR
========================= */

function generateMeetingPDF(meetingId) {
  ensureSheets_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meetingSheet = ss.getSheetByName(CONFIG.MEETINGS_SHEET);
  const taskSheet = ss.getSheetByName(CONFIG.TASKS_SHEET);

  const meetings = meetingSheet.getDataRange().getValues();
  const tasks = taskSheet.getDataRange().getValues();

  const meetingHeaderMap = getHeaderMap_(meetingSheet);
  const taskHeaderMap = getHeaderMap_(taskSheet);

  let meeting = null;
  const relatedTasks = [];

  for (let i = 1; i < meetings.length; i++) {
    const rowMeetingId = getCellByHeader_(meetings[i], meetingHeaderMap, "Meeting ID");

    if (String(rowMeetingId) === String(meetingId)) {
      meeting = meetings[i];
      break;
    }
  }

  if (!meeting) {
    throw new Error("جلسه برای تولید PDF پیدا نشد.");
  }

  for (let i = 1; i < tasks.length; i++) {
    const rowMeetingId = getCellByHeader_(tasks[i], taskHeaderMap, "Meeting ID");

    if (String(rowMeetingId) === String(meetingId)) {
      relatedTasks.push(tasks[i]);
    }
  }

  const meetingTitle = String(getCellByHeader_(meeting, meetingHeaderMap, "Title") || "");
  const meetingSummary = String(getCellByHeader_(meeting, meetingHeaderMap, "Summary") || "");
  const meetingTranscript = String(getCellByHeader_(meeting, meetingHeaderMap, "Transcript") || "");
  const meetingCreatedAt = getCellByHeader_(meeting, meetingHeaderMap, "Created At");

  const title = sanitizeFileName_(meetingTitle || "Meeting");
  const createdAtText = formatDateTime_(meetingCreatedAt);

  const doc = DocumentApp.create("صورتجلسه - " + title);
  const body = doc.getBody();

  body.clear();

  const mainTitle = body.appendParagraph("صورتجلسه");
  mainTitle
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph("");

  const infoTable = body.appendTable([
    ["عنوان جلسه", meetingTitle],
    ["تاریخ ثبت", createdAtText],
    ["شناسه جلسه", String(meetingId)]
  ]);

  styleTable_(infoTable);

  body.appendParagraph("");

  const summaryTitle = body.appendParagraph("خلاصه جلسه");
  summaryTitle
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  const summaryParagraph = body.appendParagraph(meetingSummary || "—");
  summaryParagraph.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  if (meetingTranscript) {
    body.appendParagraph("");

    const transcriptTitle = body.appendParagraph("متن تبدیل‌شده از گفتار");
    transcriptTitle
      .setHeading(DocumentApp.ParagraphHeading.HEADING2)
      .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

    const transcriptParagraph = body.appendParagraph(meetingTranscript);
    transcriptParagraph
      .setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
      .setLineSpacing(1.5);
  }

  body.appendParagraph("");

  const tasksTitle = body.appendParagraph("مصوبات و تسک‌ها");
  tasksTitle
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  if (relatedTasks.length > 0) {
    const tableData = [
      ["ردیف", "عنوان تسک", "مسئول", "اولویت", "شروع", "پایان", "وضعیت"]
    ];

    relatedTasks.forEach((t, index) => {
      tableData.push([
        String(index + 1),
        String(getCellByHeader_(t, taskHeaderMap, "Title") || ""),
        String(getCellByHeader_(t, taskHeaderMap, "Owner") || ""),
        String(getCellByHeader_(t, taskHeaderMap, "Priority") || "متوسط"),
        String(getCellByHeader_(t, taskHeaderMap, "Start Date") || ""),
        String(getCellByHeader_(t, taskHeaderMap, "End Date") || ""),
        String(getCellByHeader_(t, taskHeaderMap, "Status") || "شروع نشده")
      ]);
    });

    const taskTable = body.appendTable(tableData);
    styleTable_(taskTable);
  } else {
    body.appendParagraph("تسکی برای این جلسه ثبت نشده است.")
      .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  }

  body.appendParagraph("");
  body.appendParagraph("");

  const signTitle = body.appendParagraph("امضاها");
  signTitle
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  const signTable = body.appendTable([
    ["نام و سمت", "امضا", "تاریخ"],
    ["", "", ""],
    ["", "", ""]
  ]);

  styleTable_(signTable);

  body.appendParagraph("");

  const footer = body.appendParagraph("این سند به صورت خودکار توسط سیستم جلسات تولید شده است.");
  footer
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setFontSize(9)
    .setForegroundColor("#666666");

  doc.saveAndClose();

  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = docFile.getAs(MimeType.PDF);

  const pdfName = "صورتجلسه - " + title + ".pdf";
  pdfBlob.setName(pdfName);

  const pdfFile = DriveApp.createFile(pdfBlob);

  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  docFile.setTrashed(true);

  return {
    fileId: pdfFile.getId(),
    viewUrl: pdfFile.getUrl(),
    downloadUrl: "https://drive.google.com/uc?export=download&id=" + pdfFile.getId()
  };
}

/* =========================
   GET MEETINGS
========================= */

function getMeetings() {
  try {
    ensureSheets_();

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MEETINGS_SHEET);
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return [];
    }

    const headerMap = getHeaderMap_(sheet);

    return data.slice(1)
      .filter(r => getCellByHeader_(r, headerMap, "Meeting ID"))
      .reverse()
      .map(r => ({
        id: String(getCellByHeader_(r, headerMap, "Meeting ID") || ""),
        date: formatDateTime_(getCellByHeader_(r, headerMap, "Created At")),
        title: String(getCellByHeader_(r, headerMap, "Title") || ""),
        summary: String(getCellByHeader_(r, headerMap, "Summary") || ""),
        transcript: String(getCellByHeader_(r, headerMap, "Transcript") || ""),
        pdfUrl: String(getCellByHeader_(r, headerMap, "PDF URL") || ""),
        downloadUrl: String(getCellByHeader_(r, headerMap, "PDF Download URL") || "")
      }));

  } catch (err) {
    throw new Error(err && err.message ? err.message : String(err));
  }
}

/* =========================
   GET TASKS
========================= */

function getTasks(meetingId) {
  try {
    ensureSheets_();

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TASKS_SHEET);
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return [];
    }

    const headerMap = getHeaderMap_(sheet);

    return data.slice(1)
      .filter(r => String(getCellByHeader_(r, headerMap, "Meeting ID")) === String(meetingId))
      .map(r => ({
        id: String(getCellByHeader_(r, headerMap, "Task ID") || ""),
        title: String(getCellByHeader_(r, headerMap, "Title") || ""),
        owner: String(getCellByHeader_(r, headerMap, "Owner") || ""),
        priority: String(getCellByHeader_(r, headerMap, "Priority") || ""),
        start: String(getCellByHeader_(r, headerMap, "Start Date") || ""),
        end: String(getCellByHeader_(r, headerMap, "End Date") || ""),
        progress: Number(getCellByHeader_(r, headerMap, "Progress") || 0),
        status: String(getCellByHeader_(r, headerMap, "Status") || "شروع نشده")
      }));

  } catch (err) {
    throw new Error(err && err.message ? err.message : String(err));
  }
}

/* =========================
   UPDATE TASK
========================= */

function updateTaskProgress(taskId, progress) {
  try {
    ensureSheets_();

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.TASKS_SHEET);
    const data = sheet.getDataRange().getValues();

    const headerMap = getHeaderMap_(sheet);

    const taskIdCol = headerMap["Task ID"];
    const progressCol = headerMap["Progress"];
    const statusCol = headerMap["Status"];

    if (!taskIdCol || !progressCol || !statusCol) {
      throw new Error("ستون‌های لازم برای به‌روزرسانی تسک پیدا نشد.");
    }

    const numericProgress = Math.max(0, Math.min(100, Number(progress) || 0));

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][taskIdCol - 1]) === String(taskId)) {
        sheet.getRange(i + 1, progressCol).setValue(numericProgress);

        const status =
          numericProgress === 0 ? "شروع نشده" :
          numericProgress === 100 ? "انجام شده" :
          "درحال انجام";

        sheet.getRange(i + 1, statusCol).setValue(status);

        return {
          success: true,
          status: status,
          progress: numericProgress
        };
      }
    }

    return {
      success: false,
      error: "تسک پیدا نشد."
    };

  } catch (err) {
    return {
      success: false,
      error: err && err.message ? err.message : String(err)
    };
  }
}

/* =========================
   DASHBOARD
========================= */

function getTasksDashboard() {
  try {
    ensureSheets_();

    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.TASKS_SHEET);

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return [];
    }

    const headerMap = getHeaderMap_(sheet);

    return data.slice(1).map(r => ({
      title: String(getCellByHeader_(r, headerMap, "Title") || ""),
      progress: Number(getCellByHeader_(r, headerMap, "Progress") || 0)
    }));

  } catch (err) {
    throw new Error(err && err.message ? err.message : String(err));
  }
}

/* =========================
   HELPERS
========================= */

function formatDateTime_(value) {
  if (!value) return "";

  try {
    const date = value instanceof Date ? value : new Date(value);
    return Utilities.formatDate(date, CONFIG.TIMEZONE, "yyyy/MM/dd HH:mm");
  } catch (e) {
    return String(value);
  }
}

function sanitizeFileName_(name) {
  return String(name || "Meeting")
    .replace(/[\\\/:*?"<>|#%{}~&]/g, "-")
    .substring(0, 80);
}

function styleTable_(table) {
  table.setBorderWidth(1);

  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);

    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);

      cell.setPaddingTop(6);
      cell.setPaddingBottom(6);
      cell.setPaddingLeft(6);
      cell.setPaddingRight(6);

      const text = cell.editAsText();
      text.setFontFamily("Arial");
      text.setFontSize(10);

      if (r === 0) {
        text.setBold(true);
        cell.setBackgroundColor("#eeeeee");
      }
    }
  }
}
function summarizeText(text) {
  const OPENAI_API_KEY = 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;

  const url = "https://api.openai.com/v1/chat/completions";

  const payload = {
    model: "gpt-4o-mini",
    messages: [
      {
        role: "system",
        content: "تو یک منشی حرفه‌ای جلسات هستی. متن جلسه را به صورت رسمی و ساختاریافته خلاصه کن."
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
  const result = JSON.parse(response.getContentText());

  return result.choices[0].message.content;
}
function summarizeText(text) {
  if (!text) return "";

  const OPENAI_API_KEY = "YOUR_OPENAI_API_KEY_HERE"; // ← کلیدت را اینجا بگذار

  const url = "https://api.openai.com/v1/chat/completions";

  const payload = {
    model: "gpt-4o-mini",
    messages: [
      {
        role: "system",
        content: `
تو یک منشی حرفه‌ای جلسات هستی.
متن جلسه را به صورت ساختاریافته تولید کن با قالب زیر:

۱- خلاصه مدیریتی
۲- نکات کلیدی
۳- تصمیمات
۴- اقدامات پیشنهادی
        `
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
  const code = response.getResponseCode();

  if (code !== 200) {
    throw new Error("خطا از OpenAI: " + response.getContentText());
  }

  const result = JSON.parse(response.getContentText());

  return result.choices[0].message.content.trim();
}
