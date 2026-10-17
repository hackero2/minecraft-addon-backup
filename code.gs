/**
 * Minecraft Addon - Google Apps Script backend
 *
 * Setup:
 * 1. Create/open the Google Sheet named "Minecraft Addon Database".
 * 2. Paste its spreadsheet ID into SPREADSHEET_ID below.
 * 3. Paste your Google Drive folder ID into ADDON_FOLDER_ID below.
 * 4. Deploy as Web app: Execute as Me; access Anyone.
 *
 * API:
 * GET  ?action=getAddons
 * GET  ?action=health
 * POST JSON { action:"uploadAddon", name, description, filename, mimeType, fileData, author }
 * POST JSON { action:"deleteAddon", id, author }
 *
 * fileData must be a base64 string (raw base64 or data URL).
 */

const SPREADSHEET_ID = "1u3lkhKM2ngJd5bO8zRDeuttWX0bBf0yl3acsQ8TCGgg";
const ADDON_FOLDER_ID = "15qw84CkLrKmyvQOxkxjk0oqYKrWLa2VL";
const SHEET_NAME = "Addons";
const MAX_FILE_BYTES = 8 * 1024 * 1024; // 8 MB

const HEADERS = [
  "ID", "Name", "Description", "Filename", "MimeType",
  "DriveFileId", "DownloadUrl", "Author", "Size", "CreatedAt"
];

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || "getAddons";

    if (action === "health") {
      return jsonOutput_({ success: true, message: "Minecraft Addon API is running." });
    }

    if (action === "getAddons") {
      return jsonOutput_({ success: true, addons: getAddons_() });
    }

    return jsonOutput_({ success: false, error: "Unknown action." });
  } catch (err) {
    return jsonOutput_({ success: false, error: err.message || String(err) });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOutput_({ success: false, error: "Missing JSON request body." });
    }

    const body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case "uploadAddon":
        return jsonOutput_(uploadAddon_(body));
      case "deleteAddon":
        return jsonOutput_(deleteAddon_(body));
      default:
        return jsonOutput_({ success: false, error: "Unknown action." });
    }
  } catch (err) {
    return jsonOutput_({ success: false, error: err.message || String(err) });
  }
}

function setupDatabase() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.autoResizeColumns(1, HEADERS.length);
  } else {
    const existing = sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0];
    const matches = HEADERS.every((h, i) => existing[i] === h);
    if (!matches) {
      throw new Error(
        'The "' + SHEET_NAME + '" sheet has unexpected headers. ' +
        "Please rename/create a sheet named Addons with these headers: " + HEADERS.join(", ")
      );
    }
  }

  // Validate that the Drive folder is accessible to the script owner.
  DriveApp.getFolderById(ADDON_FOLDER_ID).getName();
  return "Setup complete. Sheet and Drive folder are accessible.";
}

function getAddons_() {
  const sheet = getAddonSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return rows
    .filter(row => row[0] !== "" && row[0] !== null)
    .map(row => ({
      id: String(row[0]),
      name: String(row[1] || ""),
      description: String(row[2] || ""),
      filename: String(row[3] || ""),
      mimeType: String(row[4] || ""),
      driveFileId: String(row[5] || ""),
      downloadUrl: String(row[6] || ""),
      author: String(row[7] || ""),
      size: Number(row[8] || 0),
      createdAt: row[9] instanceof Date
        ? row[9].toISOString()
        : String(row[9] || "")
    }))
    .reverse();
}

function uploadAddon_(body) {
  const name = cleanText_(body.name, 120);
  const description = cleanText_(body.description, 1000);
  const filename = cleanFilename_(body.filename);
  const author = cleanText_(body.author, 100);
  let fileData = String(body.fileData || "").trim();

  if (!name) throw new Error("Please provide an addon name.");
  if (!filename) throw new Error("Please provide a filename.");
  if (!author) throw new Error("Please provide an uploader name.");
  if (!fileData) throw new Error("No file data received.");

  // Accept either raw base64 or a data URL like data:...;base64,AAAA
  const commaIndex = fileData.indexOf(",");
  if (fileData.indexOf("base64,") !== -1 && commaIndex !== -1) {
    fileData = fileData.substring(commaIndex + 1);
  }

  const extensionMatch = filename.toLowerCase().match(/\.(mcaddon|mcpack|mcworld|zip)$/);
  if (!extensionMatch) {
    throw new Error("Unsupported file type. Upload .mcaddon, .mcpack, .mcworld, or .zip.");
  }

  const bytes = Utilities.base64Decode(fileData);
  if (!bytes.length) throw new Error("The uploaded file is empty.");
  if (bytes.length > MAX_FILE_BYTES) {
    throw new Error("File is too large. Maximum upload size is 8 MB.");
  }

  const mimeType = cleanText_(body.mimeType, 150) || "application/octet-stream";
  const blob = Utilities.newBlob(bytes, mimeType, filename);
  const folder = DriveApp.getFolderById(ADDON_FOLDER_ID);
  const driveFile = folder.createFile(blob);
  driveFile.setDescription("Minecraft Addon: " + name + (description ? "\n" + description : ""));

  // Anyone-with-link can download the file. Use only for files intended to be public.
  driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const id = Utilities.getUuid();
  const downloadUrl = "https://drive.google.com/uc?export=download&id=" + driveFile.getId();
  const sheet = getAddonSheet_();
  sheet.appendRow([
    id, name, description, filename, mimeType,
    driveFile.getId(), downloadUrl, author, bytes.length, new Date()
  ]);

  return {
    success: true,
    message: "Addon uploaded successfully.",
    addon: {
      id: id,
      name: name,
      description: description,
      filename: filename,
      mimeType: mimeType,
      driveFileId: driveFile.getId(),
      downloadUrl: downloadUrl,
      author: author,
      size: bytes.length,
      createdAt: new Date().toISOString()
    }
  };
}

function deleteAddon_(body) {
  const id = cleanText_(body.id, 100);
  const author = cleanText_(body.author, 100);
  if (!id || !author) throw new Error("Missing addon ID or uploader name.");

  const sheet = getAddonSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("Addon not found.");

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (String(row[0]) === id) {
      // This is a basic ownership check, not secure authentication.
      if (String(row[7]) !== author) {
        throw new Error("You can only delete addons uploaded under the same author name.");
      }

      const driveFileId = String(row[5] || "");
      if (driveFileId) {
        try {
          DriveApp.getFileById(driveFileId).setTrashed(true);
        } catch (driveError) {
          // Continue removing the database row if the file is already unavailable.
        }
      }
      sheet.deleteRow(i + 2);
      return { success: true, message: "Addon deleted." };
    }
  }

  throw new Error("Addon not found.");
}

function getAddonSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  }
  return sheet;
}

function getSpreadsheet_() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID === "PASTE_YOUR_GOOGLE_SHEET_ID_HERE") {
    throw new Error("Please set SPREADSHEET_ID in Code.gs to your Google Sheet ID.");
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function cleanText_(value, maxLength) {
  return String(value == null ? "" : value).trim().replace(/\u0000/g, "").substring(0, maxLength);
}

function cleanFilename_(value) {
  const name = String(value == null ? "" : value).trim().split(/[\\/]/).pop();
  return name.substring(0, 180).replace(/[\u0000-\u001f]/g, "");
}

function jsonOutput_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
