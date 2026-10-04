// =========================
// 1. Configuration
// =========================
const FOLDER_ID1 = "YOUR_FOLDER_ID1";   // Year folders
const FOLDER_ID2 = "YOUR_FOLDER_ID2";
const FOLDER_ID3 = "YOUR_FOLDER_ID3";
const OUTPUT_SPREADSHEET_ID = "YOUR_OUTPUT_SPREADSHEET_ID"; // Report spreadsheet

// =========================
// 2. Name matching
// =========================

/** Case-insensitive Levenshtein distance. */
function levenshteinDistance(a, b) {
  a = (a || "").toString().toLowerCase();
  b = (b || "").toString().toLowerCase();
  var costs = [];
  for (var i = 0; i <= a.length; i++) {
    costs[i] = [i];
  }
  for (var j = 0; j <= b.length; j++) {
    costs[0][j] = j;
  }
  for (var i = 1; i <= a.length; i++) {
    for (var j = 1; j <= b.length; j++) {
      var cost = a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1;
      costs[i][j] = Math.min(
        costs[i - 1][j] + 1,
        costs[i][j - 1] + 1,
        costs[i - 1][j - 1] + cost
      );
    }
  }
  return costs[a.length][b.length];
}

/** Tokens match if they differ by at most one character. */
function areTokensSimilar(token1, token2) {
  if (typeof token1 !== 'string' || typeof token2 !== 'string') return false;
  token1 = token1.toLowerCase();
  token2 = token2.toLowerCase();
  if (token1 === token2) return true;
  var distance = levenshteinDistance(token1, token2);
  return (distance <= 1);
}

/** Fuzzy name match: every word of inputName must match some word of folderName. */
function isClientNameMatch(inputName, folderName) {
  if (typeof inputName !== 'string' || typeof folderName !== 'string') return false;
  var inputTokens = inputName.split(/\s+/).map(function(token) { return token.trim(); }).filter(function(token) { return token.length > 0; });
  var folderTokens = folderName.split(/\s+/).map(function(token) { return token.trim(); }).filter(function(token) { return token.length > 0; });
  for (var i = 0; i < inputTokens.length; i++) {
    var found = false;
    for (var j = 0; j < folderTokens.length; j++) {
      if (areTokensSimilar(inputTokens[i], folderTokens[j])) {
        found = true;
        break;
      }
    }
    if (!found) return false;
  }
  return true;
}

/** Strips text in parentheses, the "EN" prefix and the word "СРОЧНО". */
function normalizeClientName(name) {
  if (typeof name !== 'string') return "";
  name = name.replace(/\(.*?\)/g, "");
  name = name.replace(/^EN\s+/i, "").replace(/\s*СРОЧНО\s*/i, "");
  return name.trim();
}

// =========================
// 3. Text extraction
// =========================

/** Returns a Slides shape's text, or "" if it has none. */
function getShapeText(shape) {
  try {
    if (shape.getText) {
      var text = shape.getText();
      return text ? text.asString() : "";
    }
  } catch(e) {}
  return "";
}

/**
 * Extracts a file's text and counts non-whitespace characters.
 * Supports Google Docs/Sheets/Slides, MS Office and PDF (Office and PDF via temporary conversion).
 * Returns { text, count }.
 */
function getTextAndCount(file) {
  var mimeType = file.getMimeType();
  var text = "";
  try {
    if (mimeType === MimeType.GOOGLE_DOCS) {
      var doc = DocumentApp.openById(file.getId());
      text = doc.getBody().getText();
    } else if (mimeType === MimeType.GOOGLE_SHEETS) {
      var sheetFile = SpreadsheetApp.openById(file.getId()).getActiveSheet();
      var values = sheetFile.getDataRange().getDisplayValues();
      text = values.flat().join("");
    } else if (mimeType === MimeType.GOOGLE_SLIDES) {
      var presentation = SlidesApp.openById(file.getId());
      var slides = presentation.getSlides();
      var allText = "";
      slides.forEach(function(slide){
        slide.getShapes().forEach(function(shape){
          allText += getShapeText(shape);
        });
      });
      text = allText;
    } else if (mimeType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
               mimeType === "application/msword") {
      var blob = file.getBlob();
      var resource = { title: file.getName() + " (converted)" };
      var tempFile = Drive.Files.insert(resource, blob, {convert: true});
      var tempDoc = DocumentApp.openById(tempFile.id);
      text = tempDoc.getBody().getText();
      DriveApp.getFileById(tempFile.id).setTrashed(true);
    } else if (mimeType === MimeType.MICROSOFT_EXCEL) {
      var blobXls = file.getBlob();
      var tempSheetFile = Drive.Files.insert({mimeType: MimeType.GOOGLE_SHEETS, title: file.getName() + " (converted)"}, blobXls, {convert: true});
      var tempSpreadsheet = SpreadsheetApp.openById(tempSheetFile.id);
      var sheet = tempSpreadsheet.getSheets()[0];
      var values = sheet.getDataRange().getDisplayValues();
      text = values.flat().join("");
      DriveApp.getFileById(tempSheetFile.id).setTrashed(true);
    } else if (mimeType === MimeType.MICROSOFT_POWERPOINT) {
      var blobPpt = file.getBlob();
      var tempSlideFile = Drive.Files.insert({mimeType: MimeType.GOOGLE_SLIDES, title: file.getName() + " (converted)"}, blobPpt, {convert: true});
      var tempPresentation = SlidesApp.openById(tempSlideFile.id);
      var slides = tempPresentation.getSlides();
      var presText = "";
      slides.forEach(function(slide){
        slide.getShapes().forEach(function(shape){
          presText += getShapeText(shape);
        });
      });
      text = presText;
      DriveApp.getFileById(tempSlideFile.id).setTrashed(true);
    } else if (mimeType === "application/pdf") {
      var blobPdf = file.getBlob();
      var resource = { title: file.getName() + " (converted)" };
      var tempFile = Drive.Files.insert(resource, blobPdf, {convert: true});
      var tempDoc = DocumentApp.openById(tempFile.id);
      text = tempDoc.getBody().getText();
      DriveApp.getFileById(tempFile.id).setTrashed(true);
    } else {
      throw new Error("Неподдерживаемый тип файла: " + mimeType);
    }
  } catch(e) {
    Logger.log('Ошибка при извлечении текста из файла "' + (file.getName() || "") + '": ' + e.message);
    text = "";
  }
  var countWithoutSpaces = text.replace(/\s/g, "").length;
  return { text: text, count: countWithoutSpaces };
}

// =========================
// 4. Processing
// =========================

/**
 * Finds a client's documents in year/month/day/client folders.
 * Returns { documents: [[name, document, date, chars], ...], totalChars }.
 */
function processSingleClient(clientName, yearFolderIds) {
  var results = { documents: [], totalChars: 0 };
  for (var i = 0; i < yearFolderIds.length; i++) {
    var folderId = yearFolderIds[i];
    var yearFolder;
    try {
      yearFolder = DriveApp.getFolderById(folderId);
    } catch(e) {
      Logger.log("Ошибка доступа к годовой папке с ID " + folderId + ": " + e.message);
      continue;
    }
    var monthFolders = yearFolder.getFolders();
    while(monthFolders.hasNext()){
      var monthFolder = monthFolders.next();
      var dayFolders = monthFolder.getFolders();
      while(dayFolders.hasNext()){
        var dayFolder = dayFolders.next();
        var dayFolderName = (dayFolder.getName() || "").toString();
        var dayFolderLower = dayFolderName.toLowerCase();
        // Only day folders whose name contains "переведенное"
        if(dayFolderLower.indexOf("переведенное") === -1) {
          continue;
        }
        var clientFolders = dayFolder.getFolders();
        while(clientFolders.hasNext()){
          var clientFolder = clientFolders.next();
          var folderClientName = (clientFolder.getName() || "").toString().trim();
          if (!isClientNameMatch(normalizeClientName(clientName), normalizeClientName(folderClientName))) {
            continue;
          }
          var files = clientFolder.getFiles();
          while(files.hasNext()){
            var file = files.next();
            var docName = (file.getName() || "").toString();
            var result = getTextAndCount(file);
            var text = result.text;
            // Skip documents whose first 1000 characters contain an excluded phrase
            if (text) {
              var snippet = text.substring(0,1000).toUpperCase();
              // Examples; replace with your own phrases, in UPPERCASE
              var forbiddenPhrases = [
                "DRAFT",
                "TEMPLATE",
                "SAMPLE DOCUMENT",
                "INTERNAL USE ONLY"
              ];
              var skip = false;
              for (var k = 0; k < forbiddenPhrases.length; k++){
                if (snippet.indexOf(forbiddenPhrases[k]) !== -1) {
                  skip = true;
                  break;
                }
              }
              if (skip) {
                Logger.log('Пропускается документ "'+ docName +'" клиента "'+ folderClientName +'" из-за запрещенной фразы.');
                continue;
              }
            }
            var charCount = result.count;
            results.documents.push([clientName, docName, dayFolderName, charCount]);
            results.totalChars += charCount;
          }
        }
      }
    }
  }
  results.documents.sort(function(a, b){ return a[2].localeCompare(b[2]); });
  return results;
}

/**
 * Processes clients in batches to stay under the execution time limit.
 * Progress is saved in script properties; a trigger resumes the next batch.
 */
function processClientsBatch() {
  var props = PropertiesService.getScriptProperties();
  var clientListSpreadsheetId = props.getProperty("CLIENT_LIST_SPREADSHEET_ID");
  if (!clientListSpreadsheetId) {
    throw new Error("CLIENT_LIST_SPREADSHEET_ID не задан.");
  }
  
  var outputSpreadsheetId = OUTPUT_SPREADSHEET_ID;
  var currentIndex = parseInt(props.getProperty("CURRENT_INDEX") || "0");
  
  var clientListSS;
  try {
    clientListSS = SpreadsheetApp.openById(clientListSpreadsheetId);
  } catch (e) {
    throw new Error("Ошибка открытия таблицы с ФИО: " + e.message);
  }
  
  var clientSheet = clientListSS.getSheets()[0];
  var clientData = clientSheet.getDataRange().getValues();
  var clientsFromList = [];
  for (var i = 0; i < clientData.length; i++) {
    var name = clientData[i][0];
    if (name && typeof name === 'string' && name.trim() !== "") {
      clientsFromList.push(name.trim());
    }
  }
  var totalClients = clientsFromList.length;
  
  var outputSS;
  try {
    outputSS = SpreadsheetApp.openById(outputSpreadsheetId);
  } catch (e) {
    throw new Error("Ошибка открытия таблицы вывода: " + e.message);
  }
  
  var summarySheet = outputSS.getSheetByName("Summary");
  if (!summarySheet) {
    summarySheet = outputSS.insertSheet("Summary");
    summarySheet.appendRow(["ФИО клиента", "Количество символов", "Ссылка"]);
    summarySheet.getRange(1, 1, 1, summarySheet.getLastColumn()).setFontWeight("bold");
  }
  
  var yearFolderIds = [FOLDER_ID1, FOLDER_ID2, FOLDER_ID3];
  var startTime = new Date().getTime();
  var timeLimit = 5 * 60 * 1000; // 5 minutes
  
  for (var i = currentIndex; i < totalClients; i++) {
    var clientName = clientsFromList[i];
    try {
      var clientResults = processSingleClient(clientName, yearFolderIds);
      // Recreate the client's sheet from scratch
      var normalizedClientName = normalizeClientName(clientName);
      var existingSheet = outputSS.getSheetByName(normalizedClientName);
      if (existingSheet) {
        outputSS.deleteSheet(existingSheet);
      }
      var clientSheet = outputSS.insertSheet(normalizedClientName);
      clientSheet.appendRow(["ФИО клиента", "Документ", "Дата файла", "Количество символов"]);
      clientSheet.getRange(1, 1, 1, clientSheet.getLastColumn()).setFontWeight("bold");
      clientResults.documents.forEach(function(row) {
        clientSheet.appendRow(row);
      });
      // Total row
      clientSheet.appendRow(["", "", "", clientResults.totalChars]);
      
      var sheetGid = clientSheet.getSheetId();
      var linkFormula = '=HYPERLINK("#gid=' + sheetGid + '", "Детали")';
      summarySheet.appendRow([clientName, clientResults.totalChars, linkFormula]);
      
      props.setProperty("CURRENT_INDEX", (i+1).toString());
    } catch (e) {
      Logger.log("Ошибка при обработке клиента " + clientName + ": " + e.message);
      summarySheet.appendRow([clientName, "Ошибка: " + e.message, ""]);
      // Stop so the failed client is retried on the next run
      return;
    }
    if (new Date().getTime() - startTime > timeLimit) {
      Logger.log("Время выполнения достигло лимита, остановка на клиенте " + (i+1) + " из " + totalClients);
      ScriptApp.newTrigger("continueProcessingClients")
               .timeBased()
               .after(60 * 1000)
               .create();
      return;
    }
  }
  props.deleteProperty("CURRENT_INDEX");
  Logger.log("Обработка завершена для всех клиентов.");
  // Retry the last client if its total is 0
  reprocessLastClient();
}

/** Trigger handler: removes its trigger and processes the next batch. */
function continueProcessingClients() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "continueProcessingClients") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  processClientsBatch();
}

/** Reprocesses the last client in "Summary" if its total is 0. */
function reprocessLastClient() {
  var props = PropertiesService.getScriptProperties();
  var yearFolderIds = [FOLDER_ID1, FOLDER_ID2, FOLDER_ID3];
  var outputSS = SpreadsheetApp.openById(OUTPUT_SPREADSHEET_ID);
  var summarySheet = outputSS.getSheetByName("Summary");
  var lastRow = summarySheet.getLastRow();
  if (lastRow < 2) return;
  var rowData = summarySheet.getRange(lastRow, 1, 1, 3).getValues()[0];
  var clientName = rowData[0];
  var currentTotal = rowData[1];
  if (clientName && currentTotal === 0) {
    Logger.log("Переобработка клиента " + clientName + " (итог 0)");
    var clientResults = processSingleClient(clientName, yearFolderIds);
    var normalizedClientName = normalizeClientName(clientName);
    var existingSheet = outputSS.getSheetByName(normalizedClientName);
    if (existingSheet) {
      outputSS.deleteSheet(existingSheet);
    }
    var clientSheet = outputSS.insertSheet(normalizedClientName);
    clientSheet.appendRow(["ФИО клиента", "Документ", "Дата файла", "Количество символов"]);
    clientSheet.getRange(1, 1, 1, clientSheet.getLastColumn()).setFontWeight("bold");
    clientResults.documents.forEach(function(row) {
      clientSheet.appendRow(row);
    });
    clientSheet.appendRow(["", "", "", clientResults.totalChars]);
    var sheetGid = clientSheet.getSheetId();
    var linkFormula = '=HYPERLINK("#gid=' + sheetGid + '", "Детали")';
    summarySheet.getRange(lastRow, 1, 1, 3).setValues([[clientName, clientResults.totalChars, linkFormula]]);
  }
}

/**
 * Saves settings and starts processing.
 * @param {string} clientListSpreadsheetId - spreadsheet with client names in column A
 * @param {boolean} resetIndex - true to start over, false to resume
 */
function initializeClientProcessing(clientListSpreadsheetId, resetIndex) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty("CLIENT_LIST_SPREADSHEET_ID", String(clientListSpreadsheetId));
  props.setProperty("FOLDER_ID1", FOLDER_ID1);
  props.setProperty("FOLDER_ID2", FOLDER_ID2);
  props.setProperty("FOLDER_ID3", FOLDER_ID3);
  props.setProperty("OUTPUT_SPREADSHEET_ID", OUTPUT_SPREADSHEET_ID);
  if (resetIndex) {
    props.setProperty("CURRENT_INDEX", "0");
  }
  processClientsBatch();
}

// =========================
// 5. Entry points
// =========================

/** Start from the first client. */
function testProcessingFromBeginning() {
  var clientListSpreadsheetId = "YOUR_CLIENT_LIST_SPREADSHEET_ID";
  initializeClientProcessing(clientListSpreadsheetId, true);
}

/** Resume where the last run stopped. */
function testProcessingFromPrevious() {
  var clientListSpreadsheetId = "YOUR_CLIENT_LIST_SPREADSHEET_ID";
  initializeClientProcessing(clientListSpreadsheetId, false);
}
