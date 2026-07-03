// ============================================================
// CONFIGURATION
// ============================================================
const CONFIG = {
  API_KEY: "YOUR_VIROHAN_API_KEY",
  OCR_ENDPOINT: "https://api-gpt.virohan.com/api/v1/ocr",
  LLM_ENDPOINT: "https://api-gpt.virohan.com/api/v1/llm/responses",
  SHEET_NAME: "Sheet1",               // ← Change to your actual sheet tab name
  SUMMARY_SHEET_NAME: "Campus Summary",
  LINK_COL: 9,                        // Column H — marksheet_direct_link
  BOARD_TYPE_COL: 13,                 // Column L — board_type
  OCR_STATUS_COL: 14,                 // Column M — board_ocr_status
  START_ROW: 2,
  MAX_RUNTIME_MS: 4.5 * 60 * 1000,   // Stop at 4.5 mins to stay under 6-min limit
  TRIGGER_INTERVAL_MINS: 5,           // Restart every 5 minutes
  BATCH_PAUSE_MS: 500,                // Small pause between rows
};


// ============================================================
// MAIN ENTRY — Called automatically by trigger, or manually
// ============================================================
function processMarksheets() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const sheet    = ss.getSheetByName(CONFIG.SHEET_NAME);
  const startTime = Date.now();

  if (!sheet) {
    Logger.log(`❌ Sheet "${CONFIG.SHEET_NAME}" not found!`);
    deleteTriggers_(); // Stop triggers if sheet missing
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) {
    Logger.log("No data rows found.");
    deleteTriggers_();
    return;
  }

  let processed = 0;
  let skipped   = 0;
  let errors    = 0;
  let allDone   = true; // Assume done unless we find unprocessed rows

  for (let row = CONFIG.START_ROW; row <= lastRow; row++) {

    // ⏱️ Check time — stop before hitting 6-min limit
    if (Date.now() - startTime > CONFIG.MAX_RUNTIME_MS) {
      Logger.log(`⏱️ Approaching time limit at row ${row}. Trigger will restart soon.`);
      allDone = false;
      break;
    }

    const boardTypeCell = sheet.getRange(row, CONFIG.BOARD_TYPE_COL);
    const ocrStatusCell = sheet.getRange(row, CONFIG.OCR_STATUS_COL);
    const linkCell      = sheet.getRange(row, CONFIG.LINK_COL);

    const existingStatus = String(ocrStatusCell.getValue()).trim();
    const existingBoard  = String(boardTypeCell.getValue()).trim();

    // ✅ Skip already successfully processed rows
    if (
      existingBoard !== "" &&
      existingStatus !== "" &&
      !existingStatus.startsWith("ERROR") &&
      !existingStatus.startsWith("SKIPPED")
    ) {
      skipped++;
      continue;
    }

    const fileUrl = String(linkCell.getValue()).trim();

    if (!fileUrl || fileUrl === "" || fileUrl.toLowerCase() === "undefined") {
      ocrStatusCell.setValue("SKIPPED: No URL");
      skipped++;
      continue;
    }

    Logger.log(`Row ${row}: Processing ${fileUrl.substring(0, 80)}...`);

    try {
      const ocrText = runOCRFromUrl_(fileUrl);

      if (!ocrText || ocrText.trim() === "") {
        ocrStatusCell.setValue("ERROR: OCR returned empty text");
        errors++;
        continue;
      }

      const boardResult = classifyBoard_(ocrText);

      if (!boardResult) {
        ocrStatusCell.setValue("ERROR: LLM classification failed");
        errors++;
        continue;
      }

      boardTypeCell.setValue(boardResult.board_name);
      ocrStatusCell.setValue(
        boardResult.is_english ? "✅ English Board" : "🌐 Regional/State Board"
      );

      processed++;
      SpreadsheetApp.flush();
      Utilities.sleep(CONFIG.BATCH_PAUSE_MS);

    } catch (e) {
      Logger.log(`Row ${row} error: ${e.message}`);
      ocrStatusCell.setValue(`ERROR: ${e.message}`);
      errors++;
    }
  }

  Logger.log(`✅ Batch done — Processed: ${processed} | Skipped: ${skipped} | Errors: ${errors}`);

  if (allDone) {
    // 🎉 All rows complete!
    Logger.log("🎉 ALL ROWS PROCESSED! Stopping trigger and building summary...");
    deleteTriggers_();
    buildCampusSummary_();
    sendCompletionEmail_(processed, skipped, errors);
  }
  // If not allDone — trigger will auto-fire again in 5 mins, no action needed
}


// ============================================================
// START — Call this ONCE to kick off the whole process
// ============================================================
function startProcessing() {
  // Clear any existing triggers first
  deleteTriggers_();

  // Set up a repeating 5-minute trigger
  ScriptApp.newTrigger("processMarksheets")
    .timeBased()
    .everyMinutes(CONFIG.TRIGGER_INTERVAL_MINS)
    .create();

  Logger.log(`✅ Trigger set! processMarksheets() will run every ${CONFIG.TRIGGER_INTERVAL_MINS} mins automatically.`);
  Logger.log("Your system does NOT need to stay on. Google will handle it. ✅");

  // Run immediately for the first batch
  processMarksheets();
}


// ============================================================
// STOP — Call this to cancel everything mid-way if needed
// ============================================================
function stopProcessing() {
  deleteTriggers_();
  SpreadsheetApp.getUi().alert("🛑 Processing stopped. Run startProcessing() to resume.");
  Logger.log("🛑 All triggers deleted. Processing stopped.");
}


// ============================================================
// DELETE all existing triggers for this script
// ============================================================
function deleteTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "processMarksheets") {
      ScriptApp.deleteTrigger(t);
      Logger.log("🗑️ Deleted trigger: " + t.getUniqueId());
    }
  });
}


// ============================================================
// SEND COMPLETION EMAIL to the sheet owner
// ============================================================
function sendCompletionEmail_(processed, skipped, errors) {
  try {
    const email   = Session.getActiveUser().getEmail();
    const subject = "✅ Virohan Board OCR — Processing Complete!";
    const body    = `
Hi,

Your marksheet board detection job has completed successfully!

📊 Summary:
- ✅ Processed: ${processed} rows
- ⏭️ Skipped (already done / no URL): ${skipped} rows  
- ❌ Errors: ${errors} rows

The "Campus Summary" sheet has been updated with campus-wise breakdowns.

Open your sheet to view results:
${SpreadsheetApp.getActiveSpreadsheet().getUrl()}

— Virohan GPT
    `.trim();

    MailApp.sendEmail(email, subject, body);
    Logger.log(`📧 Completion email sent to ${email}`);
  } catch (e) {
    Logger.log(`Email send failed: ${e.message}`);
  }
}


// ============================================================
// OCR — Pass URL directly to Virohan OCR API
// ============================================================
function runOCRFromUrl_(fileUrl) {
  try {
    const options = {
      method: "POST",
      headers: {
        "X-API-Key": CONFIG.API_KEY,
        "User-Agent": "VirohanAPIClient/1.0",
        "X-Client": "Virohan-Gpt"
      },
      payload: { file_url: fileUrl },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(CONFIG.OCR_ENDPOINT, options);
    const code     = response.getResponseCode();
    const body     = JSON.parse(response.getContentText());

    if (!body.success) {
      Logger.log(`OCR API error [${code}]: ${body.message}`);
      return null;
    }

    return body.data.text || null;

  } catch (e) {
    Logger.log(`runOCRFromUrl_ error: ${e.message}`);
    return null;
  }
}


// ============================================================
// LLM — Classify board from OCR text
// ============================================================
function classifyBoard_(ocrText) {
  try {
    const prompt = `
You are analyzing the OCR text of a student's Class 12th marksheet from India.

Your job is to:
1. Identify the name of the Board (e.g., CBSE, ICSE, UP Board, Bihar Board, MP Board, Rajasthan Board, etc.)
2. Determine if this is an English-medium national board (CBSE or ICSE/ISC) — return is_english: true
   OR a regional/state board (any other board) — return is_english: false

Respond ONLY in this exact JSON format (no extra text, no markdown, no code blocks):
{"board_name": "BOARD NAME HERE", "is_english": true}

If you cannot determine the board, respond:
{"board_name": "Unknown", "is_english": false}

OCR Text (first 3000 chars):
${ocrText.substring(0, 3000)}
    `.trim();

    const options = {
      method: "POST",
      headers: {
        "X-API-Key": CONFIG.API_KEY,
        "User-Agent": "VirohanAPIClient/1.0",
        "X-Client": "Virohan-Gpt",
        "Content-Type": "application/json"
      },
      payload: JSON.stringify({
        prompt: prompt,
        system_prompt: `You are a precise Indian education board document classifier. Output only valid raw JSON. No markdown, no explanation.`
      }),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(CONFIG.LLM_ENDPOINT, options);
    const body     = JSON.parse(response.getContentText());

    if (!body.success) {
      Logger.log(`LLM API error: ${body.message}`);
      return null;
    }

    const rawText   = body.data.text.trim();
    const jsonMatch = rawText.match(/\{[\s\S]*?\}/);
    if (!jsonMatch) return { board_name: "Parse Error", is_english: false };

    return JSON.parse(jsonMatch[0]);

  } catch (e) {
    Logger.log(`classifyBoard_ error: ${e.message}`);
    return null;
  }
}


// ============================================================
// CAMPUS SUMMARY SHEET BUILDER
// ============================================================
function buildCampusSummary_() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  let summarySheet = ss.getSheetByName(CONFIG.SUMMARY_SHEET_NAME);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(CONFIG.SUMMARY_SHEET_NAME);
  } else {
    summarySheet.clearContents();
    summarySheet.clearFormats();
  }

  const lastRow = dataSheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) return;

  const data = dataSheet
    .getRange(CONFIG.START_ROW, 1, lastRow - CONFIG.START_ROW + 1, CONFIG.OCR_STATUS_COL)
    .getValues();

  const summary = {};

  data.forEach(row => {
    const campus      = String(row[2]  || "").trim();
    const programType = String(row[4]  || "").trim();
    const ocrStatus   = String(row[12] || "").trim();

    if (!campus || !programType) return;

    let progKey = "Other";
    const pt = programType.toLowerCase();
    if      (pt.includes("b.sc")  || pt.includes("bsc"))  progKey = "B.Sc";
    else if (pt.includes("b.voc") || pt.includes("bvoc")) progKey = "B.Voc";

    if (!summary[campus])          summary[campus] = {};
    if (!summary[campus][progKey]) summary[campus][progKey] = { total: 0, english: 0, regional: 0, pending: 0 };

    summary[campus][progKey].total++;

    if      (ocrStatus.includes("English Board")) summary[campus][progKey].english++;
    else if (ocrStatus.includes("Regional"))      summary[campus][progKey].regional++;
    else                                           summary[campus][progKey].pending++;
  });

  const headers = [
    "Campus", "Program Type", "Total Students",
    "English Board Count", "English Board %",
    "Regional/State Board Count", "Regional/State Board %",
    "Pending / Error / Not Processed"
  ];

  summarySheet.appendRow(headers);
  const headerRange = summarySheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

  Object.keys(summary).sort().forEach(campus => {
    Object.keys(summary[campus]).sort().forEach(prog => {
      const d      = summary[campus][prog];
      const engPct = d.total > 0 ? ((d.english  / d.total) * 100).toFixed(1) + "%" : "0%";
      const regPct = d.total > 0 ? ((d.regional / d.total) * 100).toFixed(1) + "%" : "0%";
      summarySheet.appendRow([campus, prog, d.total, d.english, engPct, d.regional, regPct, d.pending]);
    });
  });

  summarySheet.autoResizeColumns(1, headers.length);
  summarySheet.setFrozenRows(1);

  const totalRows = summarySheet.getLastRow();
  for (let r = 2; r <= totalRows; r++) {
    summarySheet.getRange(r, 1, 1, headers.length)
      .setBackground(r % 2 === 0 ? "#f0f4ff" : "#ffffff");
  }

  Logger.log("✅ Campus Summary sheet updated.");
}


// ============================================================
// TEST FUNCTION — Run this first!
// ============================================================
function testScript() {
  Logger.log("========== VIROHAN BOARD OCR TEST ==========");

  const testUrl = "https://media.virohan.com/students-doc/343420-12th-marksheet-2022-12-01 06:54:51.991154.png";
  Logger.log("TEST 1: OCR via file_url...");
  const ocrText = runOCRFromUrl_(testUrl);
  Logger.log(ocrText ? `✅ OCR OK — "${ocrText.substring(0, 200)}"` : "⚠️ OCR empty/failed");

  Logger.log("TEST 2: LLM — CBSE mock...");
  const r1 = classifyBoard_("CENTRAL BOARD OF SECONDARY EDUCATION\nClass XII Examination 2023\nResult: PASS");
  Logger.log(r1 ? `✅ Board: ${r1.board_name}, English: ${r1.is_english}` : "❌ Failed");

  Logger.log("TEST 3: LLM — UP Board mock...");
  const r2 = classifyBoard_("उत्तर प्रदेश माध्यमिक शिक्षा परिषद्\nIntermediate Examination 2022");
  Logger.log(r2 ? `✅ Board: ${r2.board_name}, English: ${r2.is_english}` : "❌ Failed");

  Logger.log("========== TEST COMPLETE ==========");
}


// ============================================================
// UTILITY — Reset board columns to re-process from scratch
// ============================================================
function resetBoardColumns() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) return;
  const numRows = lastRow - CONFIG.START_ROW + 1;
  sheet.getRange(CONFIG.START_ROW, CONFIG.BOARD_TYPE_COL, numRows, 1).clearContent();
  sheet.getRange(CONFIG.START_ROW, CONFIG.OCR_STATUS_COL,  numRows, 1).clearContent();
  SpreadsheetApp.getUi().alert("✅ Reset done! Run startProcessing() to begin.");
}
