/*******************************************************
 * VPAT PROCESSOR - CONFIGURATION
 *
 * This script extracts VPAT data from Google Docs/DOCX files
 * and populates a Google Sheet with conformance levels and remarks.
 *******************************************************/

/*******************************************************
 * COLUMN CONFIGURATION
 * Update these constants if your sheet column names change
 *******************************************************/
const CONFIG = {
  // Sheet Column Names
  COLUMN_NAMES: {
    CRITERIA: "Criteria",
    CONFORMANCE_LEVEL: "Conformance Level (original)",
    REMARKS: "Remarks and Explanations (original)",
    // Interpreted columns
    CONFORMANCE_INTERPRETED: "Conformance Level (interpreted)",
    WEB_INTERPRETED: "Web (interpreted)",
    ELECTRONIC_DOCS_INTERPRETED: "Electronic Docs (interpreted)",
    SOFTWARE_INTERPRETED: "Software (interpreted)",
    CLOSED_INTERPRETED: "Closed (interpreted)",
    AUTHORING_INTERPRETED: "Authoring Tool (interpreted)",
    AI_COMMENT: "AI Comment",
    NEEDS_REVIEW: "Needs Review",
  },

  // Processing Settings
  DEFAULT_START_ROW: 2, // First row after headers

  // Prompts Sheet Configuration
  PROMPTS_SHEET_NAME: "Prompts",
  PROMPT_NAME_COLUMN: 1, // Column A
  PROMPT_TEXT_COLUMN: 2, // Column B
  INTERPRET_PROMPT_NAME: "INTERPRET_CONFORMANCE",
  CONFIDENCE_LEVEL_NAME: "CONFIDENCE_LEVEL",
  DEFAULT_CONFIDENCE_THRESHOLD: 70, // Default if not set in Prompts sheet

  // AI Provider Sheet Configuration
  AI_PROVIDER_SHEET_NAME: "AI Provider",
  AI_PROVIDER_CELL: "A2", // Cell containing "OpenAI" or "Gemini"
  AI_API_KEY_CELL: "B2", // Cell containing the API key

  // AI Model Configuration
  AI_MODEL: {
    CHATGPT_BASE_URL: "https://ai-gateway.apps.cloud.rt.nyu.edu/v1",
    CHATGPT_MODEL: "@openai-nyu-it-d-5b382a/gpt-4o-mini",
    CHATGPT_API_KEY_PROPERTY: "CHATGPT_API_KEY", // Script property name
    OPENAI_BASE_URL: "https://api.openai.com/v1",
    OPENAI_MODEL: "gpt-4o-mini",
    OPENAI_API_KEY_PROPERTY: "OPENAI_API_KEY",
    GEMINI_BASE_URL: "https://generativelanguage.googleapis.com/v1beta",
    GEMINI_MODEL: "gemini-2.5-flash",
    GEMINI_API_KEY_PROPERTY: "GEMINI_API_KEY",
  },

  // Rate Limiting
  API_DELAY_MS: 1000, // Delay between API calls
  BATCH_SIZE: 5, // Number of rows to process in a single API call (0 = all at once)

  // Valid conformance values
  VALID_CONFORMANCE_VALUES: [
    "Supports",
    "Partially Supports",
    "Does Not Support",
    "Not Applicable",
    "Not Evaluated",
  ],

  // Document Processing
  EXPECTED_TABLE_COLUMNS: 3, // Criteria, Conformance Level, Remarks

  // Quality Checklist Configuration
  QUALITY_CHECKLIST: {
    SHEET_NAME: "Quality Requirements",

    // Column indices (1-based)
    COLUMNS: {
      REQ_ID: 1, // A: Req ID (e.g., E-07)
      QUESTION: 2, // B: Simplified Questions
      AI_GUIDELINES: 3, // C: AI Guidelines (optional)
      RESPONSE_TYPE: 4, // D: Response Type
      CRITERIA: 5, // E: Criteria (number for grouping)
      CRITERIA_NAME: 6, // F: Criteria Name (VPAT section name)
      IMPACT: 7, // G: Impact (1-5 scale)
      TYPE: 8, // H: Type (Essential, etc.)
      AI_RESPONSE: 9, // I: AI Response (answer only based on Response Type)
      ORIGINAL_FROM_VPAT: 10, // J: Original from VPAT (exact quote)
      AI_EXPLANATION: 11, // K: AI Explanation (how AI decided)
    },

    // System prompt name in Prompts sheet
    SYSTEM_PROMPT_NAME: "QUALITY_CHECKLIST_ANALYSIS",

    // Processing settings
    MAX_DOC_LENGTH: 100000, // Maximum document text length to send to AI
    REQUIREMENTS_BATCH_SIZE: 5, // Number of requirements to process per API call (recommended: 5-10)
  },
};

/*******************************************************
 * MENU FUNCTIONS
 *******************************************************/

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("VPAT Processor")
    .addItem("1. Extract from Document", "processVPATDocument")
    .addItem("2. Interpret with AI", "interpretConformanceLevels")
    .addItem("3. Quality Checklist Analysis", "analyzeQualityChecklist")
    .addSeparator()
    .addItem(
      "Run All (Extract + Interpret + Quality)",
      "runFullProcessingWithQuality",
    )
    .addToUi();
}

/*******************************************************
 * MAIN PROCESSING FUNCTION
 *******************************************************/

/**
 * Runs both extraction and interpretation in sequence
 */
function runFullProcessing() {
  processVPATDocument();
  // Small delay between operations
  Utilities.sleep(2000);
  interpretConformanceLevels();
}

/**
 * Runs all three processes: extraction, interpretation, and quality analysis
 */
function runFullProcessingWithQuality() {
  processVPATDocument();
  Utilities.sleep(2000);
  interpretConformanceLevels();
  Utilities.sleep(2000);
  analyzeQualityChecklist();
}

/**
 * Main entry point for VPAT processing (extraction only)
 * Orchestrates the entire workflow
 */
function processVPATDocument() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  try {
    // Step 1: Get user input
    const userConfig = getUserInput(ui);
    if (!userConfig) {
      return; // User cancelled
    }

    // Step 2: Validate sheet structure
    const columnMap = validateAndMapColumns(sheet);

    // Step 3: Determine row range (use all rows if not specified)
    let startRow = userConfig.startRow;
    let endRow = userConfig.endRow;

    if (startRow === null || endRow === null) {
      // Get all rows with data in the criteria column
      const lastRow = sheet.getLastRow();
      startRow = CONFIG.DEFAULT_START_ROW;
      endRow = lastRow;
      showProgress(
        `Processing all criteria (rows ${startRow} to ${endRow})...`,
      );
    }

    // Step 4: Get criteria from sheet
    const criteriaMap = getCriteriaFromSheet(
      sheet,
      columnMap,
      startRow,
      endRow,
    );

    if (criteriaMap.size === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `No criteria found in rows ${startRow} to ${endRow}.`,
        "Error",
        10,
      );
      return;
    }

    // Step 5: Show progress indicator
    showProgress(`Processing ${criteriaMap.size} criteria...`);

    // Step 6: Load document and extract tables
    const documentData = loadDocument(userConfig.fileId);

    // Step 7: Extract VPAT data from tables
    showProgress(
      `Extracting data from ${documentData.tables.length} tables...`,
    );
    const vpatData = extractVPATData(documentData.tables, criteriaMap);

    // Step 9: Write data to sheet
    showProgress(`Writing data to sheet...`);
    const results = writeDataToSheet(sheet, vpatData, columnMap);

    // Step 10: Cleanup and show results
    showProgress(`Cleaning up temporary files...`);
    cleanup(documentData.tempDocId);

    // Use toast instead of blocking alert
    const message = `✓ Complete! Updated ${results.rowsUpdated} of ${results.totalRows} rows`;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, "Success", 10);
  } catch (error) {
    Logger.log(`Error in processVPATDocument: ${error.message}`);
    Logger.log(error.stack);

    // Use toast for errors too
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Error: ${error.message}`,
      "Processing Failed",
      10,
    );
  } finally {
    // Always hide progress, never throw
    try {
      hideProgress();
    } catch (e) {
      // Ignore errors in cleanup
    }
  }
}

/*******************************************************
 * USER INPUT FUNCTIONS
 *******************************************************/

/**
 * Gets processing configuration from user
 * @param {GoogleAppsScript.Base.Ui} ui - The UI object
 * @returns {Object|null} Configuration object or null if cancelled
 */
function getUserInput(ui) {
  const response = ui.prompt(
    "VPAT Processor Configuration",
    "Enter the following separated by commas:\n\n" +
      "1. Google Drive File ID (Doc/DOCX/PDF)\n" +
      "2. Start Row (optional - leave empty for all rows)\n" +
      "3. End Row (optional - leave empty for all rows)\n\n" +
      "Examples:\n" +
      "  1AbCdEfG12345..., 2, 50  (process rows 2-50)\n" +
      "  1AbCdEfG12345...  (process all rows)",
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return null;
  }

  const inputs = response
    .getResponseText()
    .split(",")
    .map((s) => s.trim());

  const fileId = inputs[0];

  // Validate file ID
  if (!fileId) {
    throw new Error("File ID cannot be empty.");
  }

  // If no row numbers provided, return null to indicate "all rows"
  if (inputs.length === 1 || !inputs[1] || !inputs[2]) {
    return { fileId, startRow: null, endRow: null };
  }

  const startRow = parseInt(inputs[1], 10);
  const endRow = parseInt(inputs[2], 10);

  // Validate row numbers if provided
  if (
    isNaN(startRow) ||
    isNaN(endRow) ||
    startRow < CONFIG.DEFAULT_START_ROW ||
    endRow < startRow
  ) {
    throw new Error(
      `Invalid row numbers. Start row must be >= ${CONFIG.DEFAULT_START_ROW} and end row must be >= start row.`,
    );
  }

  return { fileId, startRow, endRow };
}

/*******************************************************
 * SHEET VALIDATION AND COLUMN MAPPING
 *******************************************************/

/**
 * Validates sheet has required columns and creates column index map
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The active sheet
 * @returns {Object} Map of column names to column indices
 */
function validateAndMapColumns(sheet) {
  const headerRow = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  const columnMap = {};

  // Find each required column
  for (const [key, columnName] of Object.entries(CONFIG.COLUMN_NAMES)) {
    const colIndex = headerRow.indexOf(columnName);
    if (colIndex === -1) {
      throw new Error(
        `Required column "${columnName}" not found in sheet headers.`,
      );
    }
    columnMap[key] = colIndex + 1; // Convert to 1-based index
  }

  return columnMap;
}

/*******************************************************
 * CRITERIA EXTRACTION
 *******************************************************/

/**
 * Extracts criteria from sheet and creates a lookup map
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet
 * @param {Object} columnMap - Column index mapping
 * @param {number} startRow - Starting row number
 * @param {number} endRow - Ending row number
 * @returns {Map<string, number>} Map of normalized criteria to row numbers
 */
function getCriteriaFromSheet(sheet, columnMap, startRow, endRow) {
  const criteriaCol = columnMap.CRITERIA;
  const range = sheet.getRange(startRow, criteriaCol, endRow - startRow + 1, 1);
  const values = range.getValues();

  const criteriaMap = new Map();

  Logger.log(`=== GET CRITERIA FROM SHEET DEBUG ===`);
  Logger.log(`Criteria column: ${criteriaCol}`);
  Logger.log(`Row range: ${startRow} to ${endRow}`);

  values.forEach((row, index) => {
    const criteriaText = row[0];
    const normalizedKey = normalizeCriteriaKey(criteriaText);

    Logger.log(
      `Sheet row ${
        startRow + index
      }: "${criteriaText}" -> normalized key: "${normalizedKey}"`,
    );

    if (normalizedKey) {
      const rowNumber = startRow + index;
      criteriaMap.set(normalizedKey, rowNumber);
      Logger.log(`Added to map: "${normalizedKey}" -> row ${rowNumber}`);
    }
  });

  Logger.log(`=== CRITERIA MAP COMPLETE: ${criteriaMap.size} entries ===`);
  return criteriaMap;
}

/**
 * Normalizes criteria text to extract key (e.g., "1.1.1" from "1.1.1 Non-text Content")
 * @param {string} text - The criteria text
 * @returns {string|null} Normalized key or null
 */
function normalizeCriteriaKey(text) {
  if (!text) return null;

  const cleaned = String(text).replace(/\n/g, " ").replace(/\s+/g, " ").trim();

  // Extract numeric pattern like "1.1.1" or "4.1"
  const match = cleaned.match(/^[^\d]*([0-9]+\.[0-9]+(?:\.[0-9]+)*)/);

  return match ? match[1] : null;
}

/*******************************************************
 * DOCUMENT LOADING
 *******************************************************/

/**
 * Loads document from Drive and extracts tables
 * Handles Google Docs, DOCX, and PDF files (auto-converts as needed)
 * @param {string} fileId - Google Drive file ID
 * @returns {Object} Object containing tables array and optional tempDocId
 */
function loadDocument(fileId) {
  let tempDocId = null;

  try {
    const file = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType();

    let docId = fileId;

    // Handle DOCX conversion
    if (
      mimeType ===
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    ) {
      Logger.log(`Converting DOCX to Google Doc: ${file.getName()}`);
      docId = convertDocxToGoogleDoc(file);
      tempDocId = docId;
    }
    // Handle PDF conversion (with OCR)
    else if (mimeType === "application/pdf") {
      Logger.log(`Converting PDF to Google Doc (with OCR): ${file.getName()}`);
      docId = convertPdfToGoogleDoc(file);
      tempDocId = docId;
    } else if (mimeType !== "application/vnd.google-apps.document") {
      throw new Error(
        `Unsupported file type: ${mimeType}. Please use Google Doc, DOCX, or PDF format.`,
      );
    }

    // Open document and extract tables
    const doc = DocumentApp.openById(docId);
    const tables = doc.getBody().getTables();

    if (!tables || tables.length === 0) {
      throw new Error(
        "No tables found in document. Please check document structure.",
      );
    }

    return { tables, tempDocId };
  } catch (error) {
    // Cleanup temp doc if error occurs
    if (tempDocId) {
      cleanupTempDoc(tempDocId);
    }
    throw new Error(`Failed to load document: ${error.message}`);
  }
}

/**
 * Converts DOCX file to Google Doc format
 * @param {GoogleAppsScript.Drive.File} file - The DOCX file
 * @returns {string} ID of converted Google Doc
 */
function convertDocxToGoogleDoc(file) {
  showProgress("Converting DOCX to Google Doc...");
  const blob = file.getBlob();

  const converted = Drive.Files.create(
    {
      name: `[Temp VPAT] ${file.getName()} ${new Date().toISOString()}`,
      mimeType: "application/vnd.google-apps.document",
    },
    blob,
  );

  // Wait a moment for conversion to complete
  Utilities.sleep(2000);
  Logger.log(`DOCX converted to Google Doc with ID: ${converted.id}`);
  return converted.id;
}

/**
 * Converts PDF file to Google Doc format with OCR
 * @param {GoogleAppsScript.Drive.File} file - The PDF file
 * @returns {string} ID of converted Google Doc
 */
function convertPdfToGoogleDoc(file) {
  showProgress("Converting PDF to Google Doc (this may take a moment)...");
  const blob = file.getBlob();

  // Convert PDF to Google Doc with OCR enabled
  const converted = Drive.Files.create(
    {
      name: `[Temp VPAT] ${file.getName()} ${new Date().toISOString()}`,
      mimeType: "application/vnd.google-apps.document",
    },
    blob,
    {
      ocrLanguage: "en", // OCR is automatically applied when converting PDF to Doc
    },
  );

  // Wait longer for PDF conversion (OCR takes time)
  showProgress("Waiting for PDF conversion to complete...");
  Utilities.sleep(5000);
  Logger.log(`PDF converted to Google Doc with ID: ${converted.id}`);
  return converted.id;
}

/*******************************************************
 * VPAT DATA EXTRACTION
 *******************************************************/

/**
 * Extracts VPAT data from document tables
 * @param {GoogleAppsScript.Document.Table[]} tables - Array of tables from document
 * @param {Map<string, number>} criteriaMap - Map of criteria to row numbers
 * @returns {Object} Map of row numbers to VPAT data
 */
function extractVPATData(tables, criteriaMap) {
  const vpatData = {};
  let rowsProcessed = 0;

  Logger.log(`=== EXTRACT VPAT DATA DEBUG ===`);
  Logger.log(`Number of tables: ${tables.length}`);
  Logger.log(`Criteria map size: ${criteriaMap.size}`);
  Logger.log(
    `Criteria keys in map: ${Array.from(criteriaMap.keys()).join(", ")}`,
  );

  for (const table of tables) {
    const numRows = table.getNumRows();
    Logger.log(`Processing table with ${numRows} rows`);

    // Skip header row (start from index 1)
    for (let i = 1; i < numRows; i++) {
      try {
        const row = table.getRow(i);

        // Ensure row has expected number of columns
        if (row.getNumCells() < CONFIG.EXPECTED_TABLE_COLUMNS) {
          Logger.log(
            `Row ${i} has ${row.getNumCells()} cells, expected ${
              CONFIG.EXPECTED_TABLE_COLUMNS
            }, skipping`,
          );
          continue;
        }

        // Extract cell data
        const criteriaText = getCellText(row.getCell(0));
        const conformanceText = getCellText(row.getCell(1));
        const remarksText = getCellText(row.getCell(2));

        // Normalize criteria and find matching row
        const criteriaKey = normalizeCriteriaKey(criteriaText);

        Logger.log(
          `Table row ${i}: criteria="${criteriaText}" -> key="${criteriaKey}"`,
        );

        if (criteriaKey && criteriaMap.has(criteriaKey)) {
          const targetRow = criteriaMap.get(criteriaKey);

          vpatData[targetRow] = {
            conformanceLevel: conformanceText,
            remarks: remarksText,
            originalCriteria: criteriaText,
          };

          Logger.log(
            `✓ Matched! Writing to sheet row ${targetRow}: conformance="${conformanceText}"`,
          );
          rowsProcessed++;
        } else {
          Logger.log(`✗ No match in criteria map for key "${criteriaKey}"`);
        }
      } catch (error) {
        Logger.log(`Error processing table row ${i}: ${error.message}`);
        // Continue processing other rows
      }
    }
  }

  Logger.log(`=== EXTRACTION COMPLETE: ${rowsProcessed} criteria matched ===`);
  return vpatData;
}

/**
 * Safely extracts text from a table cell
 * @param {GoogleAppsScript.Document.TableCell} cell - The table cell
 * @returns {string} Cell text content
 */
function getCellText(cell) {
  try {
    return cell.getText().trim();
  } catch (error) {
    Logger.log(`Error reading cell text: ${error.message}`);
    return "";
  }
}

/*******************************************************
 * SHEET WRITING
 *******************************************************/

/**
 * Writes extracted VPAT data to sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The target sheet
 * @param {Object} vpatData - Map of row numbers to VPAT data
 * @param {Object} columnMap - Column index mapping
 * @returns {Object} Processing results statistics
 */
function writeDataToSheet(sheet, vpatData, columnMap) {
  let rowsUpdated = 0;

  const sortedRows = Object.keys(vpatData)
    .map((r) => parseInt(r, 10))
    .sort((a, b) => a - b);

  Logger.log(`=== WRITE DATA TO SHEET DEBUG ===`);
  Logger.log(`Total rows to write: ${sortedRows.length}`);
  Logger.log(`vpatData keys: ${Object.keys(vpatData).join(", ")}`);
  Logger.log(`Column map CONFORMANCE_LEVEL: ${columnMap.CONFORMANCE_LEVEL}`);
  Logger.log(`Column map REMARKS: ${columnMap.REMARKS}`);

  for (const rowNum of sortedRows) {
    try {
      const data = vpatData[rowNum];
      Logger.log(
        `Writing row ${rowNum}: Conformance="${
          data.conformanceLevel
        }", Remarks="${data.remarks.substring(0, 50)}..."`,
      );

      // Write conformance level
      sheet
        .getRange(rowNum, columnMap.CONFORMANCE_LEVEL)
        .setValue(data.conformanceLevel);

      // Write remarks
      sheet.getRange(rowNum, columnMap.REMARKS).setValue(data.remarks);

      rowsUpdated++;
      Logger.log(`Successfully wrote row ${rowNum}`);
    } catch (error) {
      Logger.log(`Error writing data to row ${rowNum}: ${error.message}`);
      Logger.log(`Error stack: ${error.stack}`);
    }
  }

  Logger.log(`=== WRITE COMPLETE: ${rowsUpdated} rows updated ===`);

  return {
    totalRows: sortedRows.length,
    rowsUpdated: rowsUpdated,
    rowsFailed: sortedRows.length - rowsUpdated,
  };
}

/*******************************************************
 * UTILITY FUNCTIONS
 *******************************************************/

/**
 * Shows progress toast notification
 * @param {string} message - Progress message
 */
function showProgress(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, "VPAT Processor", -1);
}

/**
 * Hides progress toast notification
 */
function hideProgress() {
  SpreadsheetApp.getActiveSpreadsheet().toast("", "VPAT Processor", 1);
}

/**
 * Cleans up temporary document if it exists
 * @param {string|null} tempDocId - Temporary document ID
 */
function cleanup(tempDocId) {
  if (tempDocId) {
    cleanupTempDoc(tempDocId);
  }
}

/**
 * Deletes temporary document from Drive
 * @param {string} docId - Document ID to delete
 */
function cleanupTempDoc(docId) {
  try {
    // Use DriveApp instead of Drive API for more reliable deletion
    const file = DriveApp.getFileById(docId);
    file.setTrashed(true);
    Logger.log(`Cleaned up temporary document: ${docId}`);
  } catch (error) {
    // Don't throw - just log the error
    Logger.log(`Failed to cleanup temp doc ${docId}: ${error.message}`);
  }
}

/**
 * Shows completion message to user
 * @param {GoogleAppsScript.Base.Ui} ui - UI object
 * @param {Object} results - Processing results
 */
function showCompletionMessage(ui, results) {
  const message =
    `Processing Complete!\n\n` +
    `Rows Updated: ${results.rowsUpdated}\n` +
    `Total Processed: ${results.totalRows}`;

  if (results.rowsFailed > 0) {
    ui.alert(
      "Processing Complete (with warnings)",
      `${message}\n\nWarning: ${results.rowsFailed} rows had errors. Check logs for details.`,
      ui.ButtonSet.OK,
    );
  } else {
    ui.alert("Success", message, ui.ButtonSet.OK);
  }
}

/*******************************************************
 * AI INTERPRETATION FUNCTIONALITY
 *******************************************************/

/**
 * Main function to interpret conformance levels using AI
 * Processes only rows that were populated in the last extraction run
 */
function interpretConformanceLevels() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  try {
    // Step 1: Validate columns exist
    const columnMap = validateAndMapColumnsForInterpretation(sheet);

    // Step 2: Get the prompt and confidence threshold from Prompts sheet
    const systemPrompt = getPromptFromSheet(ss);
    const confidenceThreshold = getConfidenceThreshold(ss);

    // Step 3: Get user input for row range
    const userConfig = getRowRangeForInterpretation(ui);
    if (!userConfig) {
      return; // User cancelled
    }

    // Determine row range
    let startRow = userConfig.startRow;
    let endRow = userConfig.endRow;

    if (startRow === null || endRow === null) {
      const lastRow = sheet.getLastRow();
      startRow = CONFIG.DEFAULT_START_ROW;
      endRow = lastRow;
    }

    // Step 4: Find rows with original conformance data (from extraction)
    showProgress(`Analyzing ${endRow - startRow + 1} rows...`);
    const rowsToProcess = findRowsWithConformanceData(
      sheet,
      columnMap,
      startRow,
      endRow,
    );

    if (rowsToProcess.length === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "No rows with conformance data found. Run extraction first.",
        "No Data",
        5,
      );
      return;
    }

    // Step 5: Process each row with AI
    showProgress(`Interpreting ${rowsToProcess.length} rows with AI...`);
    const results = processRowsWithAI(
      sheet,
      columnMap,
      rowsToProcess,
      systemPrompt,
      confidenceThreshold,
    );

    // Step 6: Show completion
    const message = `✓ Interpreted ${results.success} of ${rowsToProcess.length} rows`;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, "Success", 10);
  } catch (error) {
    Logger.log(`Error in interpretConformanceLevels: ${error.message}`);
    Logger.log(error.stack);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Error: ${error.message}`,
      "Interpretation Failed",
      10,
    );
  } finally {
    hideProgress();
  }
}

/**
 * Gets row range from user for interpretation
 */
function getRowRangeForInterpretation(ui) {
  const response = ui.prompt(
    "AI Interpretation Configuration",
    "Enter row range to interpret (optional):\n\n" +
      "Start Row, End Row\n\n" +
      "Examples:\n" +
      "  2, 50  (interpret rows 2-50)\n" +
      "  (leave empty to interpret all rows)",
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return null;
  }

  const input = response.getResponseText().trim();

  // If empty, process all rows
  if (!input) {
    return { startRow: null, endRow: null };
  }

  const inputs = input.split(",").map((s) => s.trim());
  const startRow = parseInt(inputs[0], 10);
  const endRow = parseInt(inputs[1], 10);

  if (
    isNaN(startRow) ||
    isNaN(endRow) ||
    startRow < CONFIG.DEFAULT_START_ROW ||
    endRow < startRow
  ) {
    throw new Error("Invalid row numbers.");
  }

  return { startRow, endRow };
}

/**
 * Validates that interpreted columns exist in the sheet
 */
function validateAndMapColumnsForInterpretation(sheet) {
  const headerRow = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  const columnMap = {};

  // Required columns for interpretation
  const requiredColumns = {
    CRITERIA: CONFIG.COLUMN_NAMES.CRITERIA,
    CONFORMANCE_LEVEL: CONFIG.COLUMN_NAMES.CONFORMANCE_LEVEL,
    REMARKS: CONFIG.COLUMN_NAMES.REMARKS,
    CONFORMANCE_INTERPRETED: CONFIG.COLUMN_NAMES.CONFORMANCE_INTERPRETED,
    WEB_INTERPRETED: CONFIG.COLUMN_NAMES.WEB_INTERPRETED,
    ELECTRONIC_DOCS_INTERPRETED:
      CONFIG.COLUMN_NAMES.ELECTRONIC_DOCS_INTERPRETED,
    SOFTWARE_INTERPRETED: CONFIG.COLUMN_NAMES.SOFTWARE_INTERPRETED,
    CLOSED_INTERPRETED: CONFIG.COLUMN_NAMES.CLOSED_INTERPRETED,
    AUTHORING_INTERPRETED: CONFIG.COLUMN_NAMES.AUTHORING_INTERPRETED,
    AI_COMMENT: CONFIG.COLUMN_NAMES.AI_COMMENT,
    NEEDS_REVIEW: CONFIG.COLUMN_NAMES.NEEDS_REVIEW,
  };

  for (const [key, columnName] of Object.entries(requiredColumns)) {
    const colIndex = headerRow.indexOf(columnName);
    if (colIndex === -1) {
      throw new Error(
        `Required column "${columnName}" not found in sheet headers.`,
      );
    }
    columnMap[key] = colIndex + 1;
  }

  return columnMap;
}

/**
 * Gets the interpretation prompt from the Prompts sheet
 */
function getPromptFromSheet(spreadsheet) {
  try {
    const promptsSheet = spreadsheet.getSheetByName(CONFIG.PROMPTS_SHEET_NAME);

    if (!promptsSheet) {
      throw new Error(
        `"${CONFIG.PROMPTS_SHEET_NAME}" sheet not found. Please create it with columns: "Prompt Name" and "Prompt Text"`,
      );
    }

    const data = promptsSheet.getDataRange().getValues();

    // Find the prompt by name (skip header row)
    for (let i = 1; i < data.length; i++) {
      const promptName = data[i][CONFIG.PROMPT_NAME_COLUMN - 1];
      const promptText = data[i][CONFIG.PROMPT_TEXT_COLUMN - 1];

      if (promptName === CONFIG.INTERPRET_PROMPT_NAME) {
        if (!promptText) {
          throw new Error(`Prompt "${CONFIG.INTERPRET_PROMPT_NAME}" is empty.`);
        }
        return promptText;
      }
    }

    throw new Error(
      `Prompt "${CONFIG.INTERPRET_PROMPT_NAME}" not found in "${CONFIG.PROMPTS_SHEET_NAME}" sheet.`,
    );
  } catch (error) {
    throw new Error(`Failed to load prompt: ${error.message}`);
  }
}

/**
 * Gets the API key from the AI Provider sheet (cell B2)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet
 * @returns {string} API key from cell B2
 */
function getAPIKeyFromSheet(spreadsheet) {
  try {
    const providerSheet = spreadsheet.getSheetByName(
      CONFIG.AI_PROVIDER_SHEET_NAME,
    );

    if (!providerSheet) {
      throw new Error(
        `Sheet "${CONFIG.AI_PROVIDER_SHEET_NAME}" not found. Please create it with API key in cell B2.`,
      );
    }

    const apiKey = providerSheet.getRange(CONFIG.AI_API_KEY_CELL).getValue();

    if (!apiKey || String(apiKey).trim() === "") {
      throw new Error(
        `API key not found in cell ${CONFIG.AI_API_KEY_CELL} of "${CONFIG.AI_PROVIDER_SHEET_NAME}" sheet. Please enter your API key.`,
      );
    }

    return String(apiKey).trim();
  } catch (error) {
    throw new Error(`Failed to read API key from sheet: ${error.message}`);
  }
}

/**
 * Gets the AI provider from the AI Provider sheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet
 * @returns {string} Provider name ("OPENAI" or "GEMINI")
 */
function getAIProvider(spreadsheet) {
  try {
    const providerSheet = spreadsheet.getSheetByName(
      CONFIG.AI_PROVIDER_SHEET_NAME,
    );

    if (!providerSheet) {
      Logger.log(
        `"${CONFIG.AI_PROVIDER_SHEET_NAME}" sheet not found, defaulting to OpenAI`,
      );
      return "OPENAI";
    }

    const providerValue = providerSheet
      .getRange(CONFIG.AI_PROVIDER_CELL)
      .getValue();
    const provider = String(providerValue).trim().toUpperCase();

    // Normalize provider names
    if (provider === "OPENAI" || provider === "OPEN AI") {
      Logger.log("Using OpenAI provider");
      return "OPENAI";
    } else if (provider === "GEMINI") {
      Logger.log("Using Gemini provider");
      return "GEMINI";
    } else {
      Logger.log(`Unknown provider "${providerValue}", defaulting to OpenAI`);
      return "OPENAI";
    }
  } catch (error) {
    Logger.log(
      `Error reading AI provider: ${error.message}, defaulting to OpenAI`,
    );
    return "OPENAI";
  }
}

/**
 * Gets the confidence threshold from the Prompts sheet
 */
function getConfidenceThreshold(spreadsheet) {
  try {
    const promptsSheet = spreadsheet.getSheetByName(CONFIG.PROMPTS_SHEET_NAME);

    if (!promptsSheet) {
      return CONFIG.DEFAULT_CONFIDENCE_THRESHOLD;
    }

    const data = promptsSheet.getDataRange().getValues();

    // Find the confidence level setting (skip header row)
    for (let i = 1; i < data.length; i++) {
      const promptName = data[i][CONFIG.PROMPT_NAME_COLUMN - 1];
      const promptText = data[i][CONFIG.PROMPT_TEXT_COLUMN - 1];

      if (promptName === CONFIG.CONFIDENCE_LEVEL_NAME) {
        const threshold = parseInt(promptText, 10);
        if (!isNaN(threshold) && threshold >= 0 && threshold <= 100) {
          Logger.log(`Using confidence threshold: ${threshold}`);
          return threshold;
        }
      }
    }

    Logger.log(
      `Confidence threshold not found, using default: ${CONFIG.DEFAULT_CONFIDENCE_THRESHOLD}`,
    );
    return CONFIG.DEFAULT_CONFIDENCE_THRESHOLD;
  } catch (error) {
    Logger.log(`Error reading confidence threshold: ${error.message}`);
    return CONFIG.DEFAULT_CONFIDENCE_THRESHOLD;
  }
}

/**
 * Normalizes conformance value to handle AI response variations
 * @param {string} value - Raw value from AI
 * @returns {string} Normalized valid conformance value or original if already valid
 */
function normalizeConformanceValue(value) {
  if (!value) return "";

  const normalized = String(value).trim();

  // Check if already valid (exact match with proper capitalization)
  if (CONFIG.VALID_CONFORMANCE_VALUES.includes(normalized)) {
    return normalized;
  }

  // Normalize common variations (case-insensitive matching)
  const lowerValue = normalized.toLowerCase();

  // Map variations to correct values
  const mappings = {
    // Supports variations
    supports: "Supports",
    support: "Supports",
    yes: "Supports",
    supported: "Supports",

    // Partially Supports variations
    "partially supports": "Partially Supports",
    "partial supports": "Partially Supports",
    "partially support": "Partially Supports",
    "partial support": "Partially Supports",
    partially: "Partially Supports",
    partial: "Partially Supports",
    "supports with exceptions": "Partially Supports",
    "supports with exception": "Partially Supports",
    sometimes: "Partially Supports",

    // Does Not Support variations
    "does not support": "Does Not Support",
    "doesn't support": "Does Not Support",
    "does not supports": "Does Not Support",
    "not supported": "Does Not Support",
    "not support": "Does Not Support",
    no: "Does Not Support",
    fails: "Does Not Support",
    failed: "Does Not Support",

    // Not Applicable variations
    "not applicable": "Not Applicable",
    "n/a": "Not Applicable",
    na: "Not Applicable",
    "not apply": "Not Applicable",

    // Not Evaluated variations
    "not evaluated": "Not Evaluated",
    "not assessed": "Not Evaluated",
    unevaluated: "Not Evaluated",
    unknown: "Not Evaluated",
  };

  if (mappings[lowerValue]) {
    return mappings[lowerValue];
  }

  // If no mapping found, return original value (will be logged as invalid)
  return normalized;
}

/**
 * Finds rows that have conformance data (from extraction)
 */
function findRowsWithConformanceData(sheet, columnMap, startRow, endRow) {
  const conformanceCol = columnMap.CONFORMANCE_LEVEL;
  const numRows = endRow - startRow + 1;

  const conformanceData = sheet
    .getRange(startRow, conformanceCol, numRows, 1)
    .getValues();

  const rowsToProcess = [];

  for (let i = 0; i < conformanceData.length; i++) {
    const value = conformanceData[i][0];
    // Check if cell has data (not empty)
    if (value && String(value).trim() !== "") {
      rowsToProcess.push(startRow + i);
    }
  }

  return rowsToProcess;
}

/**
 * Processes rows with AI to interpret conformance levels
 */
function processRowsWithAI(
  sheet,
  columnMap,
  rowNumbers,
  systemPrompt,
  confidenceThreshold,
) {
  let successCount = 0;
  let errorCount = 0;

  // Determine batch size (0 = process all at once)
  const batchSize = CONFIG.BATCH_SIZE || rowNumbers.length;

  // Process in batches
  for (let i = 0; i < rowNumbers.length; i += batchSize) {
    const batchRows = rowNumbers.slice(i, i + batchSize);

    try {
      // Collect all data for this batch
      const batchData = [];
      for (const rowNum of batchRows) {
        const conformanceLevel = sheet
          .getRange(rowNum, columnMap.CONFORMANCE_LEVEL)
          .getValue();
        const remarks = sheet.getRange(rowNum, columnMap.REMARKS).getValue();
        const criteria = sheet.getRange(rowNum, columnMap.CRITERIA).getValue();

        batchData.push({
          rowNum: rowNum,
          criteria: String(criteria || "").trim(),
          conformanceLevel: String(conformanceLevel || "").trim(),
          remarks: String(remarks || "").trim(),
        });
      }

      // Build batch message
      const userMessage = buildBatchMessage(batchData);

      // Call AI once for entire batch
      showProgress(
        `Processing batch ${Math.floor(i / batchSize) + 1} of ${Math.ceil(
          rowNumbers.length / batchSize,
        )} (${batchData.length} rows)...`,
      );
      const batchInterpretations = callChatGPTForInterpretation(
        systemPrompt,
        userMessage,
      );

      // Write results for each row in the batch
      for (let j = 0; j < batchData.length; j++) {
        const rowNum = batchData[j].rowNum;
        const interpretation = batchInterpretations[j] || {};

        try {
          writeInterpretedValues(
            sheet,
            columnMap,
            rowNum,
            interpretation,
            confidenceThreshold,
          );
          successCount++;
        } catch (writeError) {
          Logger.log(`Error writing row ${rowNum}: ${writeError.message}`);
          errorCount++;
        }
      }

      // Rate limiting between batches
      if (i + batchSize < rowNumbers.length) {
        Utilities.sleep(CONFIG.API_DELAY_MS);
      }
    } catch (error) {
      Logger.log(
        `Error processing batch starting at row ${batchRows[0]}: ${error.message}`,
      );
      errorCount += batchRows.length;
    }
  }

  return { success: successCount, errors: errorCount };
}

/**
 * Builds a batch message with multiple rows
 */
function buildBatchMessage(batchData) {
  let message =
    "Analyze the following VPAT entries and return a JSON array with interpretations:\n\n";

  batchData.forEach((item, index) => {
    message += `Entry ${index}:\n`;
    if (item.criteria) {
      message += `Criteria: ${item.criteria}\n`;
    }
    message += `Conformance Level: ${item.conformanceLevel}\n`;
    message += `Remarks: ${item.remarks}\n\n`;
  });

  message += `Return a JSON array (one object per entry, in the same order) with this structure for each:
{
  "conformanceLevel": "Supports|Partially Supports|Does Not Support|Not Applicable|Not Evaluated",
  "web": "...",
  "electronicDocs": "...",
  "software": "...",
  "closed": "...",
  "authoring": "...",
  "comment": "Brief explanation of interpretation",
  "confidence": 85
}`;

  return message;
}

/**
 * Writes interpreted values to the sheet
 */
function writeInterpretedValues(
  sheet,
  columnMap,
  rowNum,
  interpretation,
  confidenceThreshold,
) {
  // Validate and write each field
  const fields = [
    { key: "conformanceLevel", col: columnMap.CONFORMANCE_INTERPRETED },
    { key: "web", col: columnMap.WEB_INTERPRETED },
    { key: "electronicDocs", col: columnMap.ELECTRONIC_DOCS_INTERPRETED },
    { key: "software", col: columnMap.SOFTWARE_INTERPRETED },
    { key: "closed", col: columnMap.CLOSED_INTERPRETED },
    { key: "authoring", col: columnMap.AUTHORING_INTERPRETED },
  ];

  for (const field of fields) {
    let value = interpretation[field.key] || "";

    // Log the raw value received from AI
    if (value) {
      Logger.log(`Row ${rowNum}, ${field.key}: Raw AI value = "${value}"`);
    }

    // Normalize the value to handle AI variations
    const normalizedValue = normalizeConformanceValue(value);

    // Validate against allowed values
    if (
      normalizedValue &&
      !CONFIG.VALID_CONFORMANCE_VALUES.includes(normalizedValue)
    ) {
      Logger.log(
        `Row ${rowNum}: INVALID value "${value}" (normalized: "${normalizedValue}") for ${field.key}. Defaulting to "Not Evaluated".`,
      );
      // Instead of leaving empty, use "Not Evaluated" as fallback
      sheet.getRange(rowNum, field.col).setValue("Not Evaluated");
    } else if (normalizedValue) {
      // Log if normalization changed the value
      if (normalizedValue !== value) {
        Logger.log(
          `Row ${rowNum}: Normalized "${value}" -> "${normalizedValue}" for ${field.key}`,
        );
      }
      sheet.getRange(rowNum, field.col).setValue(normalizedValue);
    } else {
      // Empty value - use "Not Evaluated" instead of leaving blank
      Logger.log(
        `Row ${rowNum}: Empty value for ${field.key}, using "Not Evaluated"`,
      );
      sheet.getRange(rowNum, field.col).setValue("Not Evaluated");
    }
  }

  // Check confidence and set Needs Review checkbox
  const confidence = parseInt(interpretation.confidence, 10) || 0;
  const needsReview = confidence < confidenceThreshold;

  Logger.log(
    `Row ${rowNum}: Confidence=${confidence}, Threshold=${confidenceThreshold}, NeedsReview=${needsReview}`,
  );

  // Write AI Comment ONLY when needs review is flagged
  const comment = needsReview
    ? interpretation.comment || "Low confidence - please review"
    : "";
  sheet.getRange(rowNum, columnMap.AI_COMMENT).setValue(comment);

  sheet.getRange(rowNum, columnMap.NEEDS_REVIEW).setValue(needsReview);
}

/*******************************************************
 * CHATGPT API INTEGRATION
 *******************************************************/

/**
 * Calls ChatGPT API for interpretation
 */
function callChatGPTForInterpretation(systemPrompt, userMessage) {
  // Get provider from sheet dynamically
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const provider = getAIProvider(ss);

  if (provider === "OPENAI") {
    return callOpenAIForInterpretation(systemPrompt, userMessage);
  } else if (provider === "GEMINI") {
    return callGeminiForInterpretation(systemPrompt, userMessage);
  } else if (provider === "CHATGPT") {
    return callPortkeyForInterpretation(systemPrompt, userMessage);
  } else {
    throw new Error(`Unsupported provider: ${provider}`);
  }
}

/**
 * Calls OpenAI API directly
 */
function callOpenAIForInterpretation(systemPrompt, userMessage) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const apiKey = getAPIKeyFromSheet(ss);

  const url = `${CONFIG.AI_MODEL.OPENAI_BASE_URL}/chat/completions`;

  const payload = {
    model: CONFIG.AI_MODEL.OPENAI_MODEL,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userMessage },
    ],
    max_tokens: 4096,
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${apiKey}`,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  Logger.log(`Calling OpenAI API: ${url}`);
  Logger.log(`Model: ${CONFIG.AI_MODEL.OPENAI_MODEL}`);

  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  Logger.log(`Response Status: ${statusCode}`);

  if (statusCode !== 200) {
    let errorDetails = responseText;
    try {
      const errorData = JSON.parse(responseText);
      errorDetails = JSON.stringify(errorData, null, 2);
    } catch (e) {
      // Response is not JSON
    }
    throw new Error(
      `OpenAI API returned ${statusCode}. Details: ${errorDetails}`,
    );
  }

  const data = JSON.parse(responseText);

  if (!data.choices || !data.choices[0] || !data.choices[0].message) {
    throw new Error(`Invalid response structure: ${responseText}`);
  }

  const content = data.choices[0].message.content;
  Logger.log(`AI Response: ${content}`);

  // Strip markdown code fences if present
  let cleanContent = content.trim();
  if (cleanContent.startsWith("```")) {
    cleanContent = cleanContent.replace(/^```(?:json)?\s*\n?/, "");
    cleanContent = cleanContent.replace(/\n?```\s*$/, "");
  }

  // Parse JSON response (handle both single object and array)
  try {
    const parsed = JSON.parse(cleanContent);
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch (e) {
    // If not JSON, try to extract JSON from the response
    const jsonMatch = cleanContent.match(/\[[\s\S]*\]|\{[\s\S]*\}/);
    if (jsonMatch) {
      const parsed = JSON.parse(jsonMatch[0]);
      return Array.isArray(parsed) ? parsed : [parsed];
    }
    throw new Error(`Failed to parse AI response as JSON: ${content}`);
  }
}

/**
 * Calls Gemini API for interpretation
 */
function callGeminiForInterpretation(systemPrompt, userMessage) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const apiKey = getAPIKeyFromSheet(ss);

  const url = `${CONFIG.AI_MODEL.GEMINI_BASE_URL}/models/${CONFIG.AI_MODEL.GEMINI_MODEL}:generateContent?key=${apiKey}`;

  // Gemini uses a different format: contents array with parts
  const payload = {
    contents: [
      {
        parts: [
          {
            text: `${systemPrompt}\n\n${userMessage}`,
          },
        ],
      },
    ],
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: 4096, // Sufficient for batches of 5-10 questions
    },
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  Logger.log(`Calling Gemini API: ${CONFIG.AI_MODEL.GEMINI_MODEL}`);
  Logger.log(`Payload size: ${JSON.stringify(payload).length} bytes`);

  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  Logger.log(`Response Status: ${statusCode}`);

  if (statusCode !== 200) {
    let errorDetails = responseText;
    try {
      const errorData = JSON.parse(responseText);
      errorDetails = JSON.stringify(errorData, null, 2);
    } catch (e) {
      // Response is not JSON
    }
    throw new Error(
      `Gemini API returned ${statusCode}. Details: ${errorDetails}`,
    );
  }

  const data = JSON.parse(responseText);

  // Gemini response structure: { candidates: [{ content: { parts: [{ text: "..." }] } }] }
  if (
    !data.candidates ||
    !data.candidates[0] ||
    !data.candidates[0].content ||
    !data.candidates[0].content.parts ||
    !data.candidates[0].content.parts[0]
  ) {
    throw new Error(`Invalid Gemini response structure: ${responseText}`);
  }

  const content = data.candidates[0].content.parts[0].text;
  Logger.log(`Gemini Response length: ${content.length} characters`);
  Logger.log(`Gemini Response (first 500 chars): ${content.substring(0, 500)}`);
  Logger.log(
    `Gemini Response (last 500 chars): ${content.substring(Math.max(0, content.length - 500))}`,
  );

  // Strip markdown code fences if present
  let cleanContent = content.trim();
  if (cleanContent.startsWith("```")) {
    cleanContent = cleanContent.replace(/^```(?:json)?\s*\n?/, "");
    cleanContent = cleanContent.replace(/\n?```\s*$/, "");
  }

  // Parse JSON response (handle both single object and array)
  try {
    const parsed = JSON.parse(cleanContent);
    const resultArray = Array.isArray(parsed) ? parsed : [parsed];
    Logger.log(
      `Successfully parsed ${resultArray.length} items from Gemini response`,
    );
    return resultArray;
  } catch (e) {
    Logger.log(`JSON parse error: ${e.message}`);
    Logger.log(`Attempting to extract valid JSON from response...`);

    // If not valid JSON, try to extract JSON array from the response
    const jsonMatch = cleanContent.match(/\[[\s\S]*\]/);
    if (jsonMatch) {
      try {
        const parsed = JSON.parse(jsonMatch[0]);
        const resultArray = Array.isArray(parsed) ? parsed : [parsed];
        Logger.log(`Extracted ${resultArray.length} items from partial JSON`);
        return resultArray;
      } catch (e2) {
        Logger.log(`Failed to parse extracted JSON: ${e2.message}`);

        // Try to fix incomplete JSON array by closing it
        let fixedJson = jsonMatch[0].trim();
        if (!fixedJson.endsWith("]")) {
          Logger.log(`JSON array not properly closed, attempting to fix...`);
          // Remove trailing comma and incomplete object
          fixedJson = fixedJson.replace(/,\s*\{[^}]*$/, "");
          if (!fixedJson.endsWith("]")) {
            fixedJson += "]";
          }
          try {
            const parsed = JSON.parse(fixedJson);
            const resultArray = Array.isArray(parsed) ? parsed : [parsed];
            Logger.log(
              `Fixed incomplete JSON, extracted ${resultArray.length} items`,
            );
            return resultArray;
          } catch (e3) {
            Logger.log(`Failed to fix incomplete JSON: ${e3.message}`);
          }
        }
      }
    }

    Logger.log(`Full response for debugging: ${cleanContent}`);
    throw new Error(
      `Failed to parse Gemini response as JSON. Response length: ${content.length} chars. Parse error: ${e.message}`,
    );
  }
}

/**
 * Calls Portkey Gateway (NYU)
 */
function callPortkeyForInterpretation(systemPrompt, userMessage) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const apiKey = getAPIKeyFromSheet(ss);

  const url = `${CONFIG.AI_MODEL.CHATGPT_BASE_URL}/chat/completions`;

  const payload = {
    model: CONFIG.AI_MODEL.CHATGPT_MODEL, // Use model string exactly as configured (with @ prefix)
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userMessage },
    ],
    max_tokens: 4096,
  };

  // Portkey only requires x-portkey-api-key header
  // The provider info is embedded in the model string (e.g., @openai-nyu-it-d-5b382a/gpt-4o-mini)
  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-portkey-api-key": apiKey,
      "User-Agent": "Google-Apps-Script",
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    validateHttpsCertificates: false, // Try disabling SSL validation
  };

  Logger.log(`Calling API: ${url}`);
  Logger.log(`Model: ${CONFIG.AI_MODEL.CHATGPT_MODEL}`);
  Logger.log(`Payload size: ${JSON.stringify(payload).length} bytes`);

  let response;
  try {
    response = UrlFetchApp.fetch(url, options);
  } catch (fetchError) {
    Logger.log(`Fetch error: ${fetchError.toString()}`);
    Logger.log(`Fetch error message: ${fetchError.message}`);
    Logger.log(`Fetch error name: ${fetchError.name}`);

    // Try to get more details
    if (fetchError.message.includes("Bad request")) {
      throw new Error(
        `Gateway rejected request. Check: 1) API key is valid, 2) Model name is correct, 3) Gateway URL is accessible from Apps Script`,
      );
    }
    throw new Error(`Failed to connect to API: ${fetchError.message}`);
  }

  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  Logger.log(`Response Status: ${statusCode}`);
  Logger.log(`Response Body: ${responseText}`);

  if (statusCode !== 200) {
    let errorDetails = responseText;
    try {
      const errorData = JSON.parse(responseText);
      errorDetails = JSON.stringify(errorData, null, 2);
    } catch (e) {
      // Response is not JSON
    }
    throw new Error(`API returned ${statusCode}. Details: ${errorDetails}`);
  }

  const data = JSON.parse(responseText);

  if (!data.choices || !data.choices[0] || !data.choices[0].message) {
    throw new Error(`Invalid response structure: ${responseText}`);
  }

  const content = data.choices[0].message.content;
  Logger.log(`AI Response: ${content}`);

  // Strip markdown code fences if present
  let cleanContent = content.trim();
  if (cleanContent.startsWith("```")) {
    cleanContent = cleanContent.replace(/^```(?:json)?\s*\n?/, "");
    cleanContent = cleanContent.replace(/\n?```\s*$/, "");
  }

  // Parse JSON response (handle both single object and array)
  try {
    const parsed = JSON.parse(cleanContent);
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch (e) {
    // If not JSON, try to extract JSON from the response
    const jsonMatch = cleanContent.match(/\[[\s\S]*\]|\{[\s\S]*\}/);
    if (jsonMatch) {
      const parsed = JSON.parse(jsonMatch[0]);
      return Array.isArray(parsed) ? parsed : [parsed];
    }
    throw new Error(`Failed to parse AI response as JSON: ${content}`);
  }
}

/*******************************************************
 * QUALITY CHECKLIST FUNCTIONALITY
 *******************************************************/

/**
 * Main function to analyze VPAT quality checklist
 * Reads questions from "Quality Requirements" sheet
 * Groups by criteria number and processes in batches
 * Writes AI responses back to the AI Response column
 */
function analyzeQualityChecklist() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Step 1: Get file ID from user
    const fileId = getVPATFileIdForAnalysis(ui);
    if (!fileId) return;

    // Step 2: Load questions from Quality Requirements sheet
    showProgress("Loading quality requirements...");
    const requirements = loadQualityRequirements(ss);

    if (requirements.length === 0) {
      throw new Error(
        `No requirements found in "${CONFIG.QUALITY_CHECKLIST.SHEET_NAME}" sheet`,
      );
    }

    // Step 3: Group requirements by criteria number
    const criteriaGroups = groupRequirementsByCriteria(requirements);
    Logger.log(
      `Grouped ${requirements.length} requirements into ${criteriaGroups.length} criteria groups`,
    );

    // Step 4: Load VPAT document text
    showProgress("Loading VPAT document...");
    const documentText = extractFullDocumentText(fileId);
    Logger.log(`Extracted ${documentText.length} characters from document`);

    // Step 5: Get system prompt
    const systemPrompt = getQualityChecklistPrompt(ss);

    // Step 6: Process criteria groups in batches
    showProgress(`Analyzing ${criteriaGroups.length} criteria groups...`);
    const allResponses = processCriteriaGroups(
      criteriaGroups,
      documentText,
      systemPrompt,
    );

    // Step 7: Write responses back to the AI Response column
    showProgress("Writing AI responses to sheet...");
    writeQualityResponses(ss, allResponses);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `✓ Quality analysis complete! ${allResponses.length} requirements evaluated`,
      "Success",
      10,
    );
  } catch (error) {
    Logger.log(`Error in analyzeQualityChecklist: ${error.message}`);
    Logger.log(error.stack);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Error: ${error.message}`,
      "Quality Analysis Failed",
      10,
    );
  } finally {
    hideProgress();
  }
}

/**
 * Prompts user for VPAT document file ID
 */
function getVPATFileIdForAnalysis(ui) {
  const response = ui.prompt(
    "VPAT Quality Analysis",
    "Enter the Google Drive File ID of the VPAT document to analyze:\n\n" +
      "(This should be the same VPAT document you want to evaluate)",
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return null;
  }

  const fileId = response.getResponseText().trim();

  if (!fileId) {
    throw new Error("File ID cannot be empty.");
  }

  return fileId;
}

/**
 * Loads quality requirements from the Quality Requirements sheet
 * Returns array of requirement objects with all columns
 */
function loadQualityRequirements(spreadsheet) {
  const sheetName = CONFIG.QUALITY_CHECKLIST.SHEET_NAME;
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(
      `Sheet "${sheetName}" not found. Please create it with columns: ` +
        `Req ID, Simplified Questions, AI Guidelines (optional), Response Type, Criteria, Criteria Name, Impact, Type, AI Response, Original from VPAT, AI Explanation`,
    );
  }

  const data = sheet.getDataRange().getValues();
  const requirements = [];

  // Skip header row (index 0)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const reqId = row[CONFIG.QUALITY_CHECKLIST.COLUMNS.REQ_ID - 1];
    const question = row[CONFIG.QUALITY_CHECKLIST.COLUMNS.QUESTION - 1];

    // Only include rows with Req ID and Question
    if (reqId && question && String(question).trim() !== "") {
      requirements.push({
        rowIndex: i + 1, // 1-based row number for writing back
        reqId: String(reqId).trim(),
        question: String(question).trim(),
        aiGuidelines: String(
          row[CONFIG.QUALITY_CHECKLIST.COLUMNS.AI_GUIDELINES - 1] || "",
        ).trim(),
        responseType: String(
          row[CONFIG.QUALITY_CHECKLIST.COLUMNS.RESPONSE_TYPE - 1] || "",
        ).trim(),
        criteria:
          parseInt(row[CONFIG.QUALITY_CHECKLIST.COLUMNS.CRITERIA - 1], 10) || 0,
        criteriaName: String(
          row[CONFIG.QUALITY_CHECKLIST.COLUMNS.CRITERIA_NAME - 1] || "",
        ).trim(),
        impact: String(
          row[CONFIG.QUALITY_CHECKLIST.COLUMNS.IMPACT - 1] || "",
        ).trim(),
        type: String(
          row[CONFIG.QUALITY_CHECKLIST.COLUMNS.TYPE - 1] || "",
        ).trim(),
      });
    }
  }

  Logger.log(`Loaded ${requirements.length} quality requirements`);
  return requirements;
}

/**
 * Groups requirements by criteria number
 * Returns array of criteria groups, each containing multiple requirements
 */
function groupRequirementsByCriteria(requirements) {
  const criteriaMap = new Map();

  // Group by criteria number
  for (const req of requirements) {
    const criteriaNum = req.criteria || 0;

    if (!criteriaMap.has(criteriaNum)) {
      criteriaMap.set(criteriaNum, []);
    }

    criteriaMap.get(criteriaNum).push(req);
  }

  // Convert map to sorted array of groups
  const groups = Array.from(criteriaMap.entries())
    .map(([criteriaNum, reqs]) => ({
      criteriaNum: criteriaNum,
      requirements: reqs,
    }))
    .sort((a, b) => a.criteriaNum - b.criteriaNum);

  Logger.log(`Created ${groups.length} criteria groups`);
  groups.forEach((g) => {
    Logger.log(
      `  Criteria ${g.criteriaNum}: ${g.requirements.length} requirements`,
    );
  });

  return groups;
}

/**
 * Extracts full text from VPAT document for analysis
 */
function extractFullDocumentText(fileId) {
  const documentData = loadDocument(fileId);

  try {
    const doc = DocumentApp.openById(documentData.tempDocId || fileId);
    const body = doc.getBody();
    let fullText = body.getText();

    // Truncate if too long
    const maxLength = CONFIG.QUALITY_CHECKLIST.MAX_DOC_LENGTH;
    if (fullText.length > maxLength) {
      Logger.log(
        `Document truncated from ${fullText.length} to ${maxLength} characters`,
      );
      fullText =
        fullText.substring(0, maxLength) +
        "\n\n[Document truncated for analysis...]";
    }

    // Cleanup temp doc if created
    if (documentData.tempDocId) {
      cleanup(documentData.tempDocId);
    }

    return fullText;
  } catch (error) {
    if (documentData.tempDocId) {
      cleanup(documentData.tempDocId);
    }
    throw error;
  }
}

/**
 * Gets quality checklist system prompt from Prompts sheet
 */
function getQualityChecklistPrompt(spreadsheet) {
  try {
    const promptsSheet = spreadsheet.getSheetByName(CONFIG.PROMPTS_SHEET_NAME);

    if (!promptsSheet) {
      Logger.log(
        "Prompts sheet not found, using default quality checklist prompt",
      );
      return getDefaultQualityChecklistPrompt();
    }

    const data = promptsSheet.getDataRange().getValues();

    // Find the prompt by name (skip header row)
    for (let i = 1; i < data.length; i++) {
      const name = data[i][CONFIG.PROMPT_NAME_COLUMN - 1];
      if (name === CONFIG.QUALITY_CHECKLIST.SYSTEM_PROMPT_NAME) {
        const promptText = data[i][CONFIG.PROMPT_TEXT_COLUMN - 1];
        if (promptText && String(promptText).trim() !== "") {
          Logger.log("Using quality checklist prompt from Prompts sheet");
          return String(promptText).trim();
        }
      }
    }

    Logger.log("Quality checklist prompt not found, using default");
    return getDefaultQualityChecklistPrompt();
  } catch (error) {
    Logger.log(`Error loading quality prompt: ${error.message}`);
    return getDefaultQualityChecklistPrompt();
  }
}

/**
 * Default system prompt for quality checklist analysis
 */
function getDefaultQualityChecklistPrompt() {
  return `You are a VPAT (Voluntary Product Accessibility Template) quality assurance expert.

Your task is to analyze VPAT documents and answer specific quality checklist questions.

For each question, you will receive:
1. The question text
2. Expected response type (e.g., Date, Yes/No, Short Text)
3. Optional AI guidelines with specific instructions or hints
4. Optional VPAT section name to focus on

Guidelines:
- "response": Answer ONLY according to the response type requested
  - For Yes/No questions: answer only "Yes" or "No"
  - For Date questions: provide just the date (e.g., "March 2024") or "Not Found"
  - For Short Text: provide brief text answer
- "originalFromVpat": Provide the EXACT quote from the VPAT document that supports your answer
  - Copy the text verbatim
  - If not found, use empty string
- "explanation": Explain your reasoning and interpretation
  - Why did you answer this way?
  - What made you confident or uncertain?
  - Any relevant context or nuances

Return your responses in a structured JSON format with these three separate fields.`;
}

/**
 * Processes criteria groups in batches using AI
 * Batches by number of requirements (not criteria groups) for better reliability
 */
function processCriteriaGroups(criteriaGroups, documentText, systemPrompt) {
  const allResponses = [];
  const batchSize = CONFIG.QUALITY_CHECKLIST.REQUIREMENTS_BATCH_SIZE;

  // Flatten all requirements from all criteria groups
  const allRequirements = [];
  for (const group of criteriaGroups) {
    for (const req of group.requirements) {
      allRequirements.push(req);
    }
  }

  Logger.log(
    `Total requirements to process: ${allRequirements.length}, batch size: ${batchSize}`,
  );

  // Process requirements in batches
  for (let i = 0; i < allRequirements.length; i += batchSize) {
    const batchRequirements = allRequirements.slice(i, i + batchSize);

    try {
      const batchNum = Math.floor(i / batchSize) + 1;
      const totalBatches = Math.ceil(allRequirements.length / batchSize);
      showProgress(
        `Processing batch ${batchNum} of ${totalBatches} (${batchRequirements.length} questions)...`,
      );

      // Build message for this batch of requirements
      const userMessage = buildQualityChecklistMessageFromRequirements(
        batchRequirements,
        documentText,
      );

      // Call AI
      Logger.log(
        `Sending batch ${batchNum}/${totalBatches} to AI: ${batchRequirements.length} requirements`,
      );
      const batchResponses = callChatGPTForInterpretation(
        systemPrompt,
        userMessage,
      );

      const expectedCount = batchRequirements.length;
      Logger.log(
        `AI returned ${batchResponses.length} responses (expected ${expectedCount})`,
      );

      if (batchResponses.length < expectedCount) {
        Logger.log(
          `⚠️ WARNING: AI returned incomplete response! Got ${batchResponses.length} out of ${expectedCount} expected responses.`,
        );
        Logger.log(
          `This may be due to token limits. Consider reducing batch size or document length.`,
        );
      }

      Logger.log(
        `AI response structure: ${JSON.stringify(batchResponses).substring(0, 500)}...`,
      );

      // Parse and match responses to requirements
      const parsedResponses = parseQualityChecklistResponsesFromRequirements(
        batchRequirements,
        batchResponses,
      );
      Logger.log(`Parsed ${parsedResponses.length} responses after matching`);
      allResponses.push(...parsedResponses);

      // Rate limiting between batches
      if (i + batchSize < allRequirements.length) {
        Utilities.sleep(CONFIG.API_DELAY_MS);
      }
    } catch (error) {
      Logger.log(`Error processing batch ${batchNum}: ${error.message}`);

      // Add error responses for failed batch
      for (const req of batchRequirements) {
        allResponses.push({
          rowIndex: req.rowIndex,
          reqId: req.reqId,
          response: `Error: ${error.message}`,
          originalFromVpat: "",
          explanation: `Failed to process: ${error.message}`,
        });
      }
    }
  }

  return allResponses;
}

/**
 * Builds AI message for a batch of requirements (flat list)
 */
function buildQualityChecklistMessageFromRequirements(
  requirements,
  documentText,
) {
  let message = `VPAT Document Content:\n${"=".repeat(80)}\n${documentText}\n${"=".repeat(80)}\n\n`;

  message += `Analyze the document above and answer the following quality checklist questions.\n\n`;

  for (let i = 0; i < requirements.length; i++) {
    const req = requirements[i];

    message += `Question ${i}:\n`;
    message += `Req ID: ${req.reqId}\n`;
    message += `Question: ${req.question}\n`;

    if (req.responseType) {
      message += `Expected Response Type: ${req.responseType}\n`;
    }

    if (req.criteriaName) {
      message += `VPAT Section to Check: ${req.criteriaName}\n`;
    }

    if (req.aiGuidelines) {
      message += `AI Guidelines: ${req.aiGuidelines}\n`;
    }

    message += `\n`;
  }

  message += `Return a JSON array with ${requirements.length} objects (one per question, in the same order):\n`;
  message += `[\n`;
  message += `  {\n`;
  message += `    "reqId": "E-07",\n`;
  message += `    "response": "Your direct answer ONLY based on Response Type (e.g., just 'Yes', 'March 2024', or brief text)",\n`;
  message += `    "originalFromVpat": "Exact quote from VPAT document if found, otherwise empty string",\n`;
  message += `    "explanation": "Your reasoning and interpretation that led to this answer",\n`;
  message += `    "confidence": 85\n`;
  message += `  },\n`;
  message += `  ...\n`;
  message += `]\n`;

  return message;
}

/**
 * Parses AI responses and matches them to requirements (flat list)
 */
function parseQualityChecklistResponsesFromRequirements(
  requirements,
  aiResponses,
) {
  const results = [];

  Logger.log(
    `Quality Checklist Parsing: ${requirements.length} requirements, ${aiResponses.length} AI responses`,
  );
  Logger.log(
    `Expected requirement IDs: ${requirements.map((r) => r.reqId).join(", ")}`,
  );

  if (aiResponses.length > 0) {
    Logger.log(
      `AI response IDs: ${aiResponses.map((r) => r.reqId || "NO_ID").join(", ")}`,
    );
  }

  // Create a map of reqId -> AI response for efficient lookup
  const aiResponseMap = new Map();
  aiResponses.forEach((resp, index) => {
    if (resp.reqId) {
      aiResponseMap.set(resp.reqId, resp);
    } else {
      Logger.log(`Warning: AI response at index ${index} missing reqId`);
    }
  });

  // Match AI responses to requirements
  for (let i = 0; i < requirements.length; i++) {
    const req = requirements[i];

    // Try to match by reqId first, fall back to index-based matching
    let aiResponse;
    if (aiResponseMap.has(req.reqId)) {
      aiResponse = aiResponseMap.get(req.reqId);
    } else if (aiResponses[i]) {
      aiResponse = aiResponses[i];
      Logger.log(`Matched by index for ${req.reqId} (AI didn't include reqId)`);
    } else {
      aiResponse = {};
      Logger.log(`No AI response found for ${req.reqId} at index ${i}`);
    }

    // Extract separate fields
    const response = aiResponse.response || "No response";
    const originalFromVpat = aiResponse.originalFromVpat || "";
    let explanation = aiResponse.explanation || "";

    // Add confidence to explanation if provided
    if (aiResponse.confidence !== undefined) {
      explanation +=
        (explanation ? "\n\n" : "") + `Confidence: ${aiResponse.confidence}%`;
    }

    results.push({
      rowIndex: req.rowIndex,
      reqId: req.reqId,
      response: response,
      originalFromVpat: originalFromVpat,
      explanation: explanation,
    });
  }

  return results;
}

/**
 * Builds AI message for a batch of criteria groups (legacy - kept for reference)
 */
function buildQualityChecklistMessage(criteriaGroups, documentText) {
  let message = `VPAT Document Content:\n${"=".repeat(80)}\n${documentText}\n${"=".repeat(80)}\n\n`;

  message += `Analyze the document above and answer the following quality checklist questions.\n`;
  message += `Questions are grouped by criteria. Answer each question based on the document content.\n\n`;

  let questionIndex = 0;

  for (const group of criteriaGroups) {
    message += `CRITERIA ${group.criteriaNum}:\n`;
    // Include criteria name if available for first requirement in group
    if (group.requirements.length > 0 && group.requirements[0].criteriaName) {
      message += `Section: ${group.requirements[0].criteriaName}\n`;
    }
    message += `-`.repeat(40) + `\n`;

    for (const req of group.requirements) {
      message += `\nQuestion ${questionIndex}:\n`;
      message += `Req ID: ${req.reqId}\n`;
      message += `Question: ${req.question}\n`;

      if (req.responseType) {
        message += `Expected Response Type: ${req.responseType}\n`;
      }

      if (req.criteriaName) {
        message += `VPAT Section to Check: ${req.criteriaName}\n`;
      }

      if (req.aiGuidelines) {
        message += `AI Guidelines: ${req.aiGuidelines}\n`;
      }

      questionIndex++;
    }

    message += `\n`;
  }

  message += `\nReturn a JSON array with ${questionIndex} objects (one per question, in the same order):\n`;
  message += `[\n`;
  message += `  {\n`;
  message += `    "reqId": "E-07",\n`;
  message += `    "response": "Your direct answer ONLY based on Response Type (e.g., just 'Yes', 'March 2024', or brief text)",\n`;
  message += `    "originalFromVpat": "Exact quote from VPAT document if found, otherwise empty string",\n`;
  message += `    "explanation": "Your reasoning and interpretation that led to this answer",\n`;
  message += `    "confidence": 85\n`;
  message += `  },\n`;
  message += `  ...\n`;
  message += `]\n`;

  return message;
}

/**
 * Parses AI responses and matches them to requirements
 */
function parseQualityChecklistResponses(criteriaGroups, aiResponses) {
  const results = [];

  // Flatten groups to get expected order
  const allRequirements = [];
  for (const group of criteriaGroups) {
    for (const req of group.requirements) {
      allRequirements.push(req);
    }
  }

  Logger.log(
    `Quality Checklist Parsing: ${allRequirements.length} requirements, ${aiResponses.length} AI responses`,
  );
  Logger.log(
    `Expected requirement IDs: ${allRequirements.map((r) => r.reqId).join(", ")}`,
  );

  if (aiResponses.length > 0) {
    Logger.log(
      `AI response IDs: ${aiResponses.map((r) => r.reqId || "NO_ID").join(", ")}`,
    );
  }

  // Create a map of reqId -> AI response for efficient lookup
  const aiResponseMap = new Map();
  aiResponses.forEach((resp, index) => {
    if (resp.reqId) {
      aiResponseMap.set(resp.reqId, resp);
      Logger.log(`Mapped AI response for reqId ${resp.reqId}`);
    } else {
      Logger.log(
        `Warning: AI response at index ${index} missing reqId, will use index-based matching`,
      );
    }
  });

  // Match AI responses to requirements
  for (let i = 0; i < allRequirements.length; i++) {
    const req = allRequirements[i];

    // Try to match by reqId first, fall back to index-based matching
    let aiResponse;
    if (aiResponseMap.has(req.reqId)) {
      aiResponse = aiResponseMap.get(req.reqId);
      Logger.log(`Matched by reqId: ${req.reqId}`);
    } else if (aiResponses[i]) {
      aiResponse = aiResponses[i];
      Logger.log(
        `Matched by index ${i}: ${req.reqId} (AI didn't include reqId)`,
      );
    } else {
      aiResponse = {};
      Logger.log(`No AI response found for ${req.reqId} at index ${i}`);
    }

    Logger.log(
      `Processing ${req.reqId}: response="${(aiResponse.response || "NO RESPONSE").substring(0, 50)}"`,
    );

    // Extract separate fields
    const response = aiResponse.response || "No response";
    const originalFromVpat = aiResponse.originalFromVpat || "";
    let explanation = aiResponse.explanation || "";

    // Add confidence to explanation if provided
    if (aiResponse.confidence !== undefined) {
      explanation +=
        (explanation ? "\n\n" : "") + `Confidence: ${aiResponse.confidence}%`;
    }

    results.push({
      rowIndex: req.rowIndex,
      reqId: req.reqId,
      response: response,
      originalFromVpat: originalFromVpat,
      explanation: explanation,
    });

    Logger.log(
      `\u2713 Prepared response for ${req.reqId}: ${response.substring(0, 30)}...`,
    );
  }

  return results;
}

/**
 * Writes AI responses back to the Quality Requirements sheet
 * Only fills the AI Response column (column H)
 */
function writeQualityResponses(spreadsheet, responses) {
  const sheetName = CONFIG.QUALITY_CHECKLIST.SHEET_NAME;
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }

  const responseCol = CONFIG.QUALITY_CHECKLIST.COLUMNS.AI_RESPONSE;
  const originalCol = CONFIG.QUALITY_CHECKLIST.COLUMNS.ORIGINAL_FROM_VPAT;
  const explanationCol = CONFIG.QUALITY_CHECKLIST.COLUMNS.AI_EXPLANATION;

  // Write each response to its row (3 columns)
  Logger.log(`Writing ${responses.length} responses to sheet...`);
  for (const resp of responses) {
    try {
      Logger.log(
        `Writing row ${resp.rowIndex} (${resp.reqId}): response="${resp.response.substring(0, 50)}...", original="${(resp.originalFromVpat || "").substring(0, 30)}..."`,
      );

      // Write AI Response (answer only)
      sheet.getRange(resp.rowIndex, responseCol).setValue(resp.response);

      // Write Original from VPAT (exact quote)
      sheet
        .getRange(resp.rowIndex, originalCol)
        .setValue(resp.originalFromVpat);

      // Write AI Explanation (reasoning)
      sheet.getRange(resp.rowIndex, explanationCol).setValue(resp.explanation);

      Logger.log(`✓ Wrote response to row ${resp.rowIndex}: ${resp.reqId}`);
    } catch (error) {
      Logger.log(
        `✗ Error writing response to row ${resp.rowIndex}: ${error.message}`,
      );
    }
  }

  Logger.log(
    `Wrote ${responses.length} responses to 3 columns (AI Response, Original from VPAT, AI Explanation)`,
  );
}

/*******************************************************
 * AI MODEL INTEGRATION (LEGACY - KEPT FOR COMPATIBILITY)
 *******************************************************/

/**
 * Calls AI model to process text
 * @param {string} text - Input text
 * @returns {Object} Parsed response
 */
function callAIModel(text) {
  const provider = CONFIG.AI_MODEL.PROVIDER;

  if (provider === "GEMINI") {
    return callGeminiAPI(text);
  } else if (provider === "OPENAI") {
    return callOpenAIAPI(text);
  } else {
    throw new Error(`Unsupported AI provider: ${provider}`);
  }
}

/**
 * Calls Gemini API (legacy function for compatibility)
 * @param {string} text - Input text
 * @returns {Object} Parsed response
 */
function callGeminiAPI(text) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const apiKey = getAPIKeyFromSheet(ss);

  const url = `${CONFIG.AI_MODEL.GEMINI_BASE_URL}/models/${CONFIG.AI_MODEL.GEMINI_MODEL}:generateContent?key=${apiKey}`;

  const payload = {
    contents: [
      {
        parts: [
          {
            text: text,
          },
        ],
      },
    ],
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: 4096,
    },
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (statusCode !== 200) {
    let errorDetails = responseText;
    try {
      const errorData = JSON.parse(responseText);
      errorDetails = JSON.stringify(errorData, null, 2);
    } catch (e) {
      // Response is not JSON
    }
    throw new Error(
      `Gemini API returned ${statusCode}. Details: ${errorDetails}`,
    );
  }

  const data = JSON.parse(responseText);

  if (
    !data.candidates ||
    !data.candidates[0] ||
    !data.candidates[0].content ||
    !data.candidates[0].content.parts ||
    !data.candidates[0].content.parts[0]
  ) {
    throw new Error(`Invalid Gemini response structure: ${responseText}`);
  }

  const content = data.candidates[0].content.parts[0].text;

  try {
    return JSON.parse(content);
  } catch (e) {
    return { response: content };
  }
}

/**
 * Calls OpenAI API
 * @param {string} text - Input text
 * @returns {Object} Parsed response
 */
function callOpenAIAPI(text) {
  const apiKey =
    PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  if (!apiKey) {
    throw new Error("OPENAI_API_KEY not set in Script Properties.");
  }

  // TODO: Implement OpenAI API call when needed
  throw new Error("OpenAI API integration not yet implemented.");
}
