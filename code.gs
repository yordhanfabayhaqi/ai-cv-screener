/**
 * AI CV Screener — Google Apps Script Template
 * 100% AI-generated, 100% human-directed.
 * Directed and iterated by Yordhan Fitrians Akhmad B. based on real HR pain points —
 * refined repeatedly for business fit, user experience, and workflow comfort.
 * Published as a generic, reusable template. Configure everything via the ⚙️ Setup menu.
 *
 * Features:
 * - AI-powered CV screening using the Anthropic Claude API
 * - Reads CVs from Google Drive (PDF, DOCX, Google Docs, images)
 * - Job Description knowledge base for role-specific matching
 * - Duplicate detection (same email + same role)
 * - Batch or full-auto screening with background triggers
 * - Color-coded verdicts with reasoning notes
 *
 * License: MIT
 */

// ============================================================
// CONFIG — all values come from Script Properties (⚙️ Setup menu)
// No secrets, IDs, or company-specific values in code.
// ============================================================

var DEFAULTS = {
  BATCH_SIZE: 10,
  AUTO_BATCH_SIZE: 15,
  MODEL: "claude-haiku-4-5-20251001",
  MAX_TOKENS: 800,
  MAX_CV_BYTES: 5 * 1024 * 1024,
  RESPONSE_SHEET: "Form Responses 1",
  JD_SHEET: "Job Desc Knowledge",
  // Column positions (1-indexed). Adjust via Setup if your form has a different layout.
  COL_EMAIL: 3,
  COL_LINKEDIN: 4,
  COL_CV: 5,
  COL_ROLE: 6,
  COL_EXP: 7,
  COL_CUR_SAL: 8,
  COL_EXP_SAL: 9,
  COL_VERDICT: 13,
  COL_NOTES: 14
};

var LABELS = {
  0: "0 - Duplicate",
  1: "1 - Big No",
  2: "2 - Not a Strong Fit",
  3: "3 - Need Consideration",
  4: "4 - Qualified"
};

var COLORS = {
  "0 - Duplicate":          { bg: "#D5D5D5", fg: "#666666" },
  "1 - Big No":             { bg: "#F4CCCC", fg: "#990000" },
  "2 - Not a Strong Fit":   { bg: "#FCE5CD", fg: "#B45F06" },
  "3 - Need Consideration": { bg: "#FFF2CC", fg: "#7F6000" },
  "4 - Qualified":          { bg: "#D9EAD3", fg: "#274E13" }
};

var _batchStartTime = new Date().getTime();

// ============================================================
// CONFIG LAYER — Properties-based, no hardcoded secrets
// ============================================================

/**
 * Reads all configuration from Script Properties.
 * Falls back to DEFAULTS for anything not explicitly set.
 */
function getConfig_() {
  var props = PropertiesService.getScriptProperties().getProperties();
  return {
    API_KEY:        props["API_KEY"] || "",
    MODEL:          props["MODEL"] || DEFAULTS.MODEL,
    MAX_TOKENS:     Number(props["MAX_TOKENS"]) || DEFAULTS.MAX_TOKENS,
    BATCH_SIZE:     Number(props["BATCH_SIZE"]) || DEFAULTS.BATCH_SIZE,
    AUTO_BATCH_SIZE:Number(props["AUTO_BATCH_SIZE"]) || DEFAULTS.AUTO_BATCH_SIZE,
    RESPONSE_SHEET: props["RESPONSE_SHEET"] || DEFAULTS.RESPONSE_SHEET,
    JD_SHEET:       props["JD_SHEET"] || DEFAULTS.JD_SHEET,
    COMPANY_NAME:   props["COMPANY_NAME"] || "Your Company",
    COMPANY_DESC:   props["COMPANY_DESC"] || "A company hiring for various roles.",
    COL: {
      EMAIL:    Number(props["COL_EMAIL"])    || DEFAULTS.COL_EMAIL,
      LINKEDIN: Number(props["COL_LINKEDIN"]) || DEFAULTS.COL_LINKEDIN,
      CV:       Number(props["COL_CV"])       || DEFAULTS.COL_CV,
      ROLE:     Number(props["COL_ROLE"])     || DEFAULTS.COL_ROLE,
      EXP:      Number(props["COL_EXP"])      || DEFAULTS.COL_EXP,
      CUR_SAL:  Number(props["COL_CUR_SAL"]) || DEFAULTS.COL_CUR_SAL,
      EXP_SAL:  Number(props["COL_EXP_SAL"]) || DEFAULTS.COL_EXP_SAL,
      VERDICT:  Number(props["COL_VERDICT"]) || DEFAULTS.COL_VERDICT,
      NOTES:    Number(props["COL_NOTES"])   || DEFAULTS.COL_NOTES
    }
  };
}

/**
 * Saves a key-value pair to Script Properties.
 */
function saveConfig_(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, String(value));
}

// ============================================================
// MENU
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu("🤖 AI CV Screener")
    .addItem("⚙️ Setup", "showSetup")
    .addSeparator()
    .addItem("▶ Screen Batch", "startAIScreening")
    .addItem("🚀 Screen ALL (auto-continue)", "startScreenAll")
    .addItem("⏹ Stop Auto-Screening", "stopScreenAll")
    .addSeparator()
    .addItem("📋 Check JD Coverage", "checkJDCoverage")
    .addItem("📊 View Progress", "viewProgress")
    .addItem("🧹 Clear All Verdicts", "clearVerdicts")
    .addToUi();
}

// ============================================================
// SETUP — interactive config for all keys/IDs
// ============================================================

function showSetup() {
  var ui = SpreadsheetApp.getUi();
  var cfg = getConfig_();

  // Step 1: API Key
  var curKey = cfg.API_KEY ? cfg.API_KEY.substring(0, 10) + "..." + cfg.API_KEY.slice(-4) : "(not set)";
  var r1 = ui.prompt("⚙️ Step 1/4 — Claude API Key",
    "Current: " + curKey + "\n\nPaste your Anthropic API key (sk-ant-...).\nLeave blank to keep current.",
    ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var newKey = r1.getResponseText().trim();
  if (newKey) {
    if (!newKey.startsWith("sk-ant-")) { ui.alert("❌ Invalid key format."); return; }
    saveConfig_("API_KEY", newKey);
  }

  // Step 2: Company info
  var r2 = ui.prompt("⚙️ Step 2/4 — Company Name",
    "Current: " + cfg.COMPANY_NAME + "\n\nUsed in AI screening context.\nLeave blank to keep current.",
    ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var newCompany = r2.getResponseText().trim();
  if (newCompany) saveConfig_("COMPANY_NAME", newCompany);

  var r2b = ui.prompt("⚙️ Step 2/4 — Company Description (1-2 sentences)",
    "Current: " + cfg.COMPANY_DESC + "\n\nBrief description for AI context.\nLeave blank to keep current.",
    ui.ButtonSet.OK_CANCEL);
  if (r2b.getSelectedButton() !== ui.Button.OK) return;
  var newDesc = r2b.getResponseText().trim();
  if (newDesc) saveConfig_("COMPANY_DESC", newDesc);

  // Step 3: Sheet names
  var r3 = ui.prompt("⚙️ Step 3/4 — Response Sheet Name",
    "Current: " + cfg.RESPONSE_SHEET + "\n\nThe sheet tab containing candidate form responses.\nLeave blank to keep current.",
    ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK) return;
  var newSheet = r3.getResponseText().trim();
  if (newSheet) saveConfig_("RESPONSE_SHEET", newSheet);

  var r3b = ui.prompt("⚙️ Step 3/4 — JD Knowledge Sheet Name",
    "Current: " + cfg.JD_SHEET + "\n\nThe sheet tab with job descriptions (Col A=Title, Col B=JD text).\nLeave blank to keep current.",
    ui.ButtonSet.OK_CANCEL);
  if (r3b.getSelectedButton() !== ui.Button.OK) return;
  var newJD = r3b.getResponseText().trim();
  if (newJD) saveConfig_("JD_SHEET", newJD);

  // Step 4: Column mapping
  var colMsg = "Current column positions (1-indexed):\n" +
    "  Email=" + cfg.COL.EMAIL + "  LinkedIn=" + cfg.COL.LINKEDIN +
    "  CV=" + cfg.COL.CV + "  Role=" + cfg.COL.ROLE +
    "\n  Experience=" + cfg.COL.EXP + "  Current Salary=" + cfg.COL.CUR_SAL +
    "  Expected Salary=" + cfg.COL.EXP_SAL +
    "\n  Verdict Output=" + cfg.COL.VERDICT + "  Notes Output=" + cfg.COL.NOTES +
    "\n\nChange? Enter as comma-separated (all 9 values) or leave blank to keep current." +
    "\nOrder: Email,LinkedIn,CV,Role,Exp,CurSal,ExpSal,Verdict,Notes";
  var r4 = ui.prompt("⚙️ Step 4/4 — Column Mapping", colMsg, ui.ButtonSet.OK_CANCEL);
  if (r4.getSelectedButton() !== ui.Button.OK) return;
  var colInput = r4.getResponseText().trim();
  if (colInput) {
    var parts = colInput.split(",");
    if (parts.length === 9) {
      var colKeys = ["COL_EMAIL","COL_LINKEDIN","COL_CV","COL_ROLE","COL_EXP","COL_CUR_SAL","COL_EXP_SAL","COL_VERDICT","COL_NOTES"];
      for (var i = 0; i < 9; i++) {
        var val = parseInt(parts[i].trim(), 10);
        if (val > 0) saveConfig_(colKeys[i], val);
      }
    } else {
      ui.alert("Expected 9 comma-separated numbers. Columns not changed.");
    }
  }

  ui.alert("✅ Setup complete!\n\nReload the spreadsheet for changes to take effect.");
}

// ============================================================
// JOB DESCRIPTION KNOWLEDGE BASE
// ============================================================

function loadJDKnowledge_(ss, cfg) {
  var jdSheet = ss.getSheetByName(cfg.JD_SHEET);
  if (!jdSheet) return { map: {}, found: false };

  var lastRow = jdSheet.getLastRow();
  if (lastRow < 2) return { map: {}, found: true, empty: true };

  var data = jdSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var map = {};

  for (var i = 0; i < data.length; i++) {
    var title = String(data[i][0] || "").trim();
    var jdText = String(data[i][1] || "").trim();
    if (!title || !jdText) continue;
    map[title.toLowerCase()] = { title: title, jd: jdText };
  }

  return { map: map, found: true, empty: Object.keys(map).length === 0 };
}

function matchJD_(jdMap, roleApplied) {
  if (!roleApplied || !jdMap) return null;
  var key = roleApplied.toLowerCase().trim();

  if (jdMap[key]) return jdMap[key];

  var jdKeys = Object.keys(jdMap);
  for (var i = 0; i < jdKeys.length; i++) {
    if (key.indexOf(jdKeys[i]) !== -1 || jdKeys[i].indexOf(key) !== -1) {
      return jdMap[jdKeys[i]];
    }
  }
  return null;
}

function checkJDCoverage() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = getConfig_();

  var jdData = loadJDKnowledge_(ss, cfg);
  if (!jdData.found) {
    ui.alert("❌ Sheet '" + cfg.JD_SHEET + "' not found.\n\nCreate it with:\n  Column A: Role Title\n  Column B: Job Description / Ad text");
    return;
  }
  if (jdData.empty) {
    ui.alert("⚠️ Sheet '" + cfg.JD_SHEET + "' exists but has no data.\n\nAdd role titles in Column A and JD text in Column B (starting row 2).");
    return;
  }

  var sheet = ss.getSheetByName(cfg.RESPONSE_SHEET);
  if (!sheet || sheet.getLastRow() < 2) { ui.alert("No candidate data."); return; }

  var roles = sheet.getRange(2, cfg.COL.ROLE, sheet.getLastRow() - 1, 1).getValues();
  var uniqueRoles = {};
  for (var i = 0; i < roles.length; i++) {
    var r = str_(roles[i][0]);
    if (r) uniqueRoles[r] = (uniqueRoles[r] || 0) + 1;
  }

  var covered = [], missing = [];
  var roleNames = Object.keys(uniqueRoles).sort();
  for (var i = 0; i < roleNames.length; i++) {
    var role = roleNames[i];
    var jd = matchJD_(jdData.map, role);
    var count = uniqueRoles[role];
    if (jd) {
      covered.push("✅ " + role + " (" + count + ") → \"" + jd.title + "\"");
    } else {
      missing.push("❌ " + role + " (" + count + ") → NO JD");
    }
  }

  var msg = "JDs loaded: " + Object.keys(jdData.map).length + "\n\n";
  if (covered.length > 0) msg += "--- MATCHED ---\n" + covered.join("\n") + "\n\n";
  if (missing.length > 0) msg += "--- MISSING (generic knowledge used) ---\n" + missing.join("\n");
  ui.alert("📋 JD Coverage Report", msg, ui.ButtonSet.OK);
}

// ============================================================
// SHARED HELPERS
// ============================================================

function markDuplicates_(sheet, allData, cfg) {
  var seen = {}, count = 0;
  for (var i = 0; i < allData.length; i++) {
    if (str_(allData[i][cfg.COL.VERDICT - 1]) !== "") continue;
    var email = str_(allData[i][cfg.COL.EMAIL - 1]).toLowerCase();
    var role  = str_(allData[i][cfg.COL.ROLE - 1]).toLowerCase();
    if (!email) continue;
    var key = email + "|" + role;
    if (seen[key]) {
      writeVerdict_(sheet, i + 2, LABELS[0], "Duplicate (same email + role). First entry at row " + seen[key] + ".", cfg);
      count++;
    } else {
      seen[key] = i + 2;
    }
  }
  return count;
}

function getUnscreened_(sheet, cfg) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var maxCol = Math.max(cfg.COL.NOTES, sheet.getLastColumn());
  var allData = sheet.getRange(2, 1, lastRow - 1, maxCol).getValues();
  var queue = [];
  for (var i = 0; i < allData.length; i++) {
    if (str_(allData[i][cfg.COL.VERDICT - 1]) === "") {
      queue.push({ row: i + 2, data: allData[i] });
    }
  }
  return queue;
}

function ensureHeaders_(sheet, cfg) {
  if (sheet.getRange(1, cfg.COL.VERDICT).getValue() !== "🤖 Verdict") {
    sheet.getRange(1, cfg.COL.VERDICT).setValue("🤖 Verdict");
    sheet.getRange(1, cfg.COL.NOTES).setValue("🤖 AI Screening Notes");
    sheet.getRange(1, cfg.COL.VERDICT, 1, 2).setFontWeight("bold").setBackground("#444444").setFontColor("#FFFFFF");
    sheet.setColumnWidth(cfg.COL.VERDICT, 180);
    sheet.setColumnWidth(cfg.COL.NOTES, 500);
  }
}

function processBatch_(apiKey, sheet, batch, jdMap, cfg) {
  var done = 0, errs = 0;
  for (var b = 0; b < batch.length; b++) {
    var elapsed = (new Date().getTime() - _batchStartTime) / 1000;
    if (elapsed > 300) {
      Logger.log("Time limit at candidate " + (b + 1) + ". Stopping batch.");
      break;
    }
    try {
      var d = batch[b].data;
      var email    = str_(d[cfg.COL.EMAIL - 1]);
      var linkedin = str_(d[cfg.COL.LINKEDIN - 1]);
      var cvLink   = str_(d[cfg.COL.CV - 1]);
      var role     = str_(d[cfg.COL.ROLE - 1]);
      var exp      = Number(d[cfg.COL.EXP - 1]) || 0;
      var curSal   = Number(d[cfg.COL.CUR_SAL - 1]) || 0;
      var expSal   = Number(d[cfg.COL.EXP_SAL - 1]) || 0;

      if (!email) {
        writeVerdict_(sheet, batch[b].row, LABELS[1], "No email provided.", cfg);
        done++; continue;
      }

      var cv = readCV_(cvLink);
      var jd = matchJD_(jdMap, role);
      var result = callClaude_(apiKey, email, linkedin, cvLink, cv, role, exp, curSal, expSal, jd, cfg);
      writeVerdict_(sheet, batch[b].row, result.verdict, result.notes, cfg);
      done++;
    } catch (e) {
      errs++;
      writeVerdict_(sheet, batch[b].row, "ERROR", "Script: " + e.message, cfg);
    }
    if ((b + 1) % 3 === 0) SpreadsheetApp.flush();
  }
  SpreadsheetApp.flush();
  return { done: done, errors: errs };
}

// ============================================================
// MODE 1: MANUAL BATCH
// ============================================================

function startAIScreening() {
  _batchStartTime = new Date().getTime();
  var ui = SpreadsheetApp.getUi();
  var cfg = getConfig_();
  if (!cfg.API_KEY) { ui.alert("❌ No API key. Run ⚙️ Setup first."); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(cfg.RESPONSE_SHEET);
  if (!sheet) { ui.alert("❌ Sheet '" + cfg.RESPONSE_SHEET + "' not found."); return; }
  if (sheet.getLastRow() < 2) { ui.alert("❌ No data."); return; }

  ensureHeaders_(sheet, cfg);

  var jdData = loadJDKnowledge_(ss, cfg);
  var jdStatus = jdData.found ? (jdData.empty ? "⚠️ empty" : "✅ " + Object.keys(jdData.map).length + " JDs") : "⚠️ not found";

  var maxCol = Math.max(cfg.COL.NOTES, sheet.getLastColumn());
  var allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, maxCol).getValues();
  var dupsMarked = markDuplicates_(sheet, allData, cfg);

  var queue = getUnscreened_(sheet, cfg);
  if (queue.length === 0) {
    ui.alert("✅ All candidates screened." + (dupsMarked ? "\n(" + dupsMarked + " duplicates marked)" : ""));
    return;
  }

  var batch = queue.slice(0, cfg.BATCH_SIZE);
  var ok = ui.alert("🤖 AI CV Screener",
    "JD Knowledge: " + jdStatus +
    "\nDuplicates marked: " + dupsMarked +
    "\nUnscreened: " + queue.length +
    "\nThis batch: " + batch.length +
    "\n\nProceed?", ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  var result = processBatch_(cfg.API_KEY, sheet, batch, jdData.map, cfg);
  var left = queue.length - result.done;

  ui.alert("🤖 Batch Complete",
    "Processed: " + result.done + "\nErrors: " + result.errors + "\nRemaining: " + left +
    (left > 0 ? "\n\nClick ▶ again to continue." : "\n\n🎉 All done!"), ui.ButtonSet.OK);
}

// ============================================================
// MODE 2: SCREEN ALL (auto-continue with triggers)
// ============================================================

function startScreenAll() {
  var ui = SpreadsheetApp.getUi();
  var cfg = getConfig_();
  if (!cfg.API_KEY) { ui.alert("❌ No API key. Run ⚙️ Setup first."); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(cfg.RESPONSE_SHEET);
  if (!sheet) { ui.alert("❌ Sheet not found."); return; }
  if (sheet.getLastRow() < 2) { ui.alert("❌ No data."); return; }

  ensureHeaders_(sheet, cfg);

  var jdData = loadJDKnowledge_(ss, cfg);
  var jdStatus = jdData.found ? (jdData.empty ? "⚠️ empty" : "✅ " + Object.keys(jdData.map).length + " JDs") : "⚠️ not found";

  var maxCol = Math.max(cfg.COL.NOTES, sheet.getLastColumn());
  var allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, maxCol).getValues();
  var dupsMarked = markDuplicates_(sheet, allData, cfg);
  SpreadsheetApp.flush();

  var queue = getUnscreened_(sheet, cfg);
  if (queue.length === 0) {
    ui.alert("✅ All candidates already screened." + (dupsMarked ? "\n(" + dupsMarked + " duplicates)" : ""));
    return;
  }

  var estCost = (queue.length * 0.006).toFixed(2);
  var estMin = Math.ceil(queue.length / cfg.AUTO_BATCH_SIZE) * 1.5;

  var ok = ui.alert("🚀 Screen ALL Candidates",
    "JD Knowledge: " + jdStatus +
    "\nDuplicates marked: " + dupsMarked +
    "\nTo screen: " + queue.length +
    "\n\nBatch: " + cfg.AUTO_BATCH_SIZE + " per cycle" +
    "\nEstimated: ~" + Math.round(estMin) + " min, ~$" + estCost +
    "\n\nProceed?", ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  PropertiesService.getScriptProperties().setProperty("JD_CACHE", JSON.stringify(jdData.map));
  PropertiesService.getScriptProperties().setProperty("AUTO_SCREEN_RUNNING", "true");
  clearAutoTriggers_();

  runAutoBatch_();

  ui.alert("🚀 Started!\n\nJD: " + jdStatus + "\nCheck 📊 View Progress anytime.\nUse ⏹ Stop to cancel.");
}

function runAutoBatch_() {
  _batchStartTime = new Date().getTime();

  if (PropertiesService.getScriptProperties().getProperty("AUTO_SCREEN_RUNNING") !== "true") {
    clearAutoTriggers_();
    return;
  }

  var cfg = getConfig_();
  if (!cfg.API_KEY) {
    PropertiesService.getScriptProperties().setProperty("AUTO_SCREEN_RUNNING", "false");
    clearAutoTriggers_();
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    var ssId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
    if (ssId) { ss = SpreadsheetApp.openById(ssId); }
    else {
      PropertiesService.getScriptProperties().setProperty("AUTO_SCREEN_RUNNING", "false");
      clearAutoTriggers_();
      return;
    }
  } else {
    PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", ss.getId());
  }

  var sheet = ss.getSheetByName(cfg.RESPONSE_SHEET);
  if (!sheet) {
    PropertiesService.getScriptProperties().setProperty("AUTO_SCREEN_RUNNING", "false");
    clearAutoTriggers_();
    return;
  }

  ensureHeaders_(sheet, cfg);

  var jdMap = {};
  try {
    var cached = PropertiesService.getScriptProperties().getProperty("JD_CACHE");
    if (cached) jdMap = JSON.parse(cached);
  } catch (e) {}
  if (Object.keys(jdMap).length === 0) {
    var jdData = loadJDKnowledge_(ss, cfg);
    jdMap = jdData.map || {};
    if (Object.keys(jdMap).length > 0) {
      PropertiesService.getScriptProperties().setProperty("JD_CACHE", JSON.stringify(jdMap));
    }
  }

  var queue = getUnscreened_(sheet, cfg);
  if (queue.length === 0) {
    PropertiesService.getScriptProperties().setProperty("AUTO_SCREEN_RUNNING", "false");
    PropertiesService.getScriptProperties().deleteProperty("JD_CACHE");
    clearAutoTriggers_();
    Logger.log("All candidates screened.");
    try {
      var email = Session.getActiveUser().getEmail();
      if (email) {
        MailApp.sendEmail(email, "AI CV Screener — Screening Complete",
          "All candidates have been screened.\nOpen your spreadsheet to review the results.");
      }
    } catch (e) {}
    return;
  }

  var batch = queue.slice(0, cfg.AUTO_BATCH_SIZE);
  Logger.log("Auto-batch: " + batch.length + " of " + queue.length + " remaining.");

  var result = processBatch_(cfg.API_KEY, sheet, batch, jdMap, cfg);
  Logger.log("Done: " + result.done + ", Errors: " + result.errors);

  var remaining = queue.length - result.done;
  if (remaining > 0 && PropertiesService.getScriptProperties().getProperty("AUTO_SCREEN_RUNNING") === "true") {
    clearAutoTriggers_();
    ScriptApp.newTrigger("runAutoBatch_").timeBased().after(60 * 1000).create();
    Logger.log("Next batch in 1 min. " + remaining + " left.");
  } else {
    PropertiesService.getScriptProperties().setProperty("AUTO_SCREEN_RUNNING", "false");
    PropertiesService.getScriptProperties().deleteProperty("JD_CACHE");
    clearAutoTriggers_();
  }
}

function stopScreenAll() {
  PropertiesService.getScriptProperties().setProperty("AUTO_SCREEN_RUNNING", "false");
  PropertiesService.getScriptProperties().deleteProperty("JD_CACHE");
  clearAutoTriggers_();
  SpreadsheetApp.getUi().alert("⏹ Stopped.\n\nAll screened results preserved.\nResume anytime with 🚀 Screen ALL.");
}

function clearAutoTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runAutoBatch_") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// ============================================================
// READ CV FROM GOOGLE DRIVE
// ============================================================

function readCV_(url) {
  var out = { ok: false, text: "", b64: "", mime: "", err: "" };
  if (!url) { out.err = "No CV link"; return out; }

  var fid = extractId_(url);
  if (!fid) { out.err = "Cannot parse Drive ID from URL"; return out; }

  try {
    var file = DriveApp.getFileById(fid);
    var mime = file.getMimeType();
    var size = file.getSize();
    out.mime = mime;

    if (size > DEFAULTS.MAX_CV_BYTES) {
      out.err = "File too large (" + Math.round(size / 1024 / 1024) + "MB)";
      return out;
    }

    if (mime === "application/vnd.google-apps.document") {
      out.text = DocumentApp.openById(fid).getBody().getText();
      out.ok = true; return out;
    }
    if (mime === "application/pdf") {
      out.b64 = Utilities.base64Encode(file.getBlob().getBytes());
      out.ok = true; return out;
    }
    if (mime.indexOf("wordprocessingml") !== -1 || mime === "application/msword") {
      try {
        out.b64 = Utilities.base64Encode(file.getAs("application/pdf").getBytes());
        out.mime = "application/pdf"; out.ok = true;
      } catch (e) { out.err = "DOCX convert failed: " + e.message; }
      return out;
    }
    if (mime.indexOf("image/") === 0) {
      out.b64 = Utilities.base64Encode(file.getBlob().getBytes());
      out.ok = true; return out;
    }
    try {
      out.text = file.getBlob().getDataAsString(); out.ok = true;
    } catch (e) { out.err = "Unsupported type: " + mime; }
    return out;

  } catch (e) {
    out.err = e.message.indexOf("Access denied") !== -1 ? "Access denied (file not shared)" :
              e.message.indexOf("not found") !== -1 ? "File not found or deleted" :
              "Drive error: " + e.message;
    return out;
  }
}

function extractId_(url) {
  if (!url) return null;
  var m = String(url).match(/[?&]id=([a-zA-Z0-9_-]+)/) ||
          String(url).match(/\/file\/d\/([a-zA-Z0-9_-]+)/) ||
          String(url).match(/\/document\/d\/([a-zA-Z0-9_-]+)/) ||
          String(url).match(/^([a-zA-Z0-9_-]{20,})$/);
  return m ? m[1] : null;
}

// ============================================================
// CLAUDE API — AI SCREENING
// ============================================================

function callClaude_(apiKey, email, linkedin, cvLink, cv, role, exp, curSal, expSal, jd, cfg) {
  // --- SYSTEM PROMPT ---
  // NOTE: The salary benchmarks below are EXAMPLE ranges only.
  // Replace them with ranges appropriate to your market, currency, and role levels.
  var system = [
    "You are a Senior Recruiter at " + cfg.COMPANY_NAME + ".",
    cfg.COMPANY_DESC,
    "",
    "VERDICT SCALE:",
    "4 = Qualified — matches JD requirements, relevant CV, reasonable salary. Worth interviewing.",
    "3 = Need Consideration — partial JD match or some gaps, but potentially viable.",
    "2 = Not a Strong Fit — significant mismatch vs JD, experience, skills, or salary.",
    "1 = Big No — clearly unqualified vs JD, absurd salary, or major red flags.",
    "",
    "EVALUATION PRIORITY (in order):",
    "",
    "1. JD MATCH (most important if JD provided):",
    "   - Compare CV against the SPECIFIC job description.",
    "   - Check qualifications, scope coverage, required tools/certifications.",
    "   - Flag anomalies: background completely unrelated to role, overqualified for entry-level,",
    "     or skills mismatch (e.g. frontend-only CV for fullstack JD).",
    "   - If no JD provided, use general role knowledge.",
    "",
    "2. EXPERIENCE — years vs role expectation.",
    "   Intern: 0yr fine. Entry: 1-2yr. Mid-level: 2-5yr. Senior: 5+yr.",
    "",
    // EXAMPLE salary benchmarks — replace with your own market data.
    // These are deliberately generic and do not represent any real company's ranges.
    "3. SALARY — check against market benchmarks for the role level.",
    "   (Use your knowledge of the relevant job market. The candidate's currency and",
    "   location should inform whether the ask is reasonable, high, or low.)",
    "   Extremely high asks for junior roles = red flag. Zero or obvious typo = flag but don't auto-reject.",
    "",
    "4. CV QUALITY — skills, progression, completeness.",
    "   If CV unreadable, evaluate with available data and note the gap.",
    "",
    "5. LINKEDIN — you cannot read the profile. Note if URL present/absent.",
    "",
    "RESPOND ONLY with this JSON (no backticks, no preamble):",
    "{\"verdict\": <1-4>, \"notes\": \"<2-4 sentences. Start with JD match if JD was provided.>\"}"
  ].join("\n");

  // --- USER MESSAGE ---
  var content = [];

  if (cv.ok && cv.b64) {
    var mediaType = cv.mime || "application/pdf";
    if (mediaType.indexOf("image/") === 0) {
      content.push({ type: "image", source: { type: "base64", media_type: mediaType, data: cv.b64 } });
    } else {
      content.push({ type: "document", source: { type: "base64", media_type: "application/pdf", data: cv.b64 } });
    }
  }

  var txt = [
    "CANDIDATE:",
    "Email: " + email,
    "Role applied: " + role,
    "Experience: " + exp + " years",
    "Current salary: " + fmtNum_(curSal),
    "Expected salary: " + fmtNum_(expSal),
    "LinkedIn: " + (linkedin || "(none)"),
    ""
  ];

  if (jd) {
    var jdText = jd.jd;
    if (jdText.length > 4000) jdText = jdText.substring(0, 4000) + "\n[...truncated...]";
    txt.push("===== JOB DESCRIPTION =====");
    txt.push("Title: " + jd.title);
    txt.push(jdText);
    txt.push("===== END JD =====");
    txt.push("");
    txt.push("Screen this candidate AGAINST the JD above.");
  } else {
    txt.push("(No specific JD on file for this role. Use general role knowledge.)");
  }

  txt.push("");

  if (cv.ok && cv.text) {
    var t = cv.text.length > 8000 ? cv.text.substring(0, 8000) + "\n[...truncated...]" : cv.text;
    txt.push("CV TEXT:\n" + t);
  } else if (cv.ok && cv.b64) {
    txt.push("CV DOCUMENT: attached above. Read and evaluate it.");
  } else {
    txt.push("CV: UNREADABLE — " + (cv.err || "unknown error") + ". Evaluate with available data.");
  }
  txt.push("\nVerdict?");

  content.push({ type: "text", text: txt.join("\n") });

  var resp = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: { "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
    payload: JSON.stringify({
      model: cfg.MODEL,
      max_tokens: cfg.MAX_TOKENS,
      system: system,
      messages: [{ role: "user", content: content }]
    }),
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  var body = resp.getContentText();

  if (code !== 200) {
    var msg = "API " + code;
    try { msg += ": " + JSON.parse(body).error.message; } catch (e) { msg += ": " + body.substring(0, 200); }
    return { verdict: "ERROR", notes: msg };
  }

  var data = JSON.parse(body);
  var aiText = "";
  for (var i = 0; i < data.content.length; i++) {
    if (data.content[i].type === "text") aiText += data.content[i].text;
  }
  return parseResponse_(aiText);
}

function parseResponse_(text) {
  try {
    var clean = text.replace(/```json\s*/g, "").replace(/```\s*/g, "").trim();
    var j = JSON.parse(clean);
    var v = Number(j.verdict);
    if (v < 1 || v > 4 || isNaN(v)) return { verdict: "ERROR", notes: "Bad verdict: " + text.substring(0, 300) };
    return { verdict: LABELS[v], notes: String(j.notes || "") };
  } catch (e) {
    var vm = text.match(/"verdict"\s*:\s*(\d)/);
    var nm = text.match(/"notes"\s*:\s*"([^"]+)"/);
    if (vm) return { verdict: LABELS[Number(vm[1])] || "ERROR", notes: nm ? nm[1] : "Parsed from malformed JSON" };
    return { verdict: "ERROR", notes: "Parse failed: " + text.substring(0, 300) };
  }
}

// ============================================================
// WRITE VERDICT + COLOR
// ============================================================

function writeVerdict_(sheet, row, verdict, notes, cfg) {
  sheet.getRange(row, cfg.COL.VERDICT).setValue(verdict);
  sheet.getRange(row, cfg.COL.NOTES).setValue(notes);
  var c = COLORS[verdict];
  if (c) {
    sheet.getRange(row, cfg.COL.VERDICT).setBackground(c.bg).setFontColor(c.fg);
  } else {
    sheet.getRange(row, cfg.COL.VERDICT).setBackground("#EA4335").setFontColor("#FFFFFF");
  }
}

// ============================================================
// PROGRESS
// ============================================================

function viewProgress() {
  var cfg = getConfig_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.RESPONSE_SHEET);
  if (!sheet) return;
  var last = sheet.getLastRow();
  if (last < 2) return;

  var vals = sheet.getRange(2, cfg.COL.VERDICT, last - 1, 1).getValues();
  var total = vals.length, screened = 0, unscreened = 0, errs = 0;
  var counts = {};

  for (var i = 0; i < vals.length; i++) {
    var v = str_(vals[i][0]);
    if (v === "") { unscreened++; }
    else if (v === "ERROR") { errs++; }
    else { screened++; counts[v] = (counts[v] || 0) + 1; }
  }

  var running = PropertiesService.getScriptProperties().getProperty("AUTO_SCREEN_RUNNING") === "true";
  var msg = "Total: " + total + "\nScreened: " + screened + "\nUnscreened: " + unscreened + "\nErrors: " + errs +
    "\n\nAuto-screening: " + (running ? "🟢 RUNNING" : "⚪ Stopped") + "\n";
  var keys = Object.keys(counts).sort();
  for (var k = 0; k < keys.length; k++) msg += "\n" + keys[k] + ": " + counts[keys[k]];

  SpreadsheetApp.getUi().alert("📊 Progress", msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================
// CLEAR VERDICTS
// ============================================================

function clearVerdicts() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert("Clear ALL verdicts?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  PropertiesService.getScriptProperties().setProperty("AUTO_SCREEN_RUNNING", "false");
  PropertiesService.getScriptProperties().deleteProperty("JD_CACHE");
  clearAutoTriggers_();
  var cfg = getConfig_();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.RESPONSE_SHEET);
  if (!sheet) return;
  sheet.getRange(1, cfg.COL.VERDICT, sheet.getLastRow(), 2).clear();
}

// ============================================================
// UTILITIES
// ============================================================

function str_(v) {
  var s = String(v || "").trim();
  return (s === "undefined" || s === "nan" || s === "null") ? "" : s;
}

function fmtNum_(n) {
  return (Number(n) || 0).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}
