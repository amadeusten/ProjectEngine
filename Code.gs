/**
 * @OnlyCurrentDoc
 * This script manages multiple production applications with a shared project management system.
 */

// An object to namespace all functions related to the "Materials" sheet.
const materialsSheet = {
  NAME: 'Materials', // The name of the sheet.

  /**
   * Gets the 'Materials' sheet object from the active spreadsheet.
   * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} The sheet object or null if not found.
   */
  getSheet: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(this.NAME);
    if (!sheet) {
      console.error(`Sheet '${this.NAME}' not found. Please ensure a sheet with this exact name exists in the current spreadsheet.`);
    }
    return sheet;
  },

  /**
   * Fetches and processes all material data from the sheet for the Fabrication App.
   * @returns {Array<Object>} An array of objects, each with 'name' and 'unitCost' properties.
   */
  getData: function() {
    const sheet = this.getSheet();
    if (!sheet) {
      return []; // Return empty array if the sheet doesn't exist.
    }
    // Get data from Column A to Column P to include all relevant fields from the new sheet.
    const range = sheet.getRange('A2:P' + sheet.getLastRow());
    const values = range.getValues();
    
    // Process the data into an array of objects.
    return values
      .filter(row => {
        const name = row[1]; // Column B: Material Name
        const primaryCategory = row[4]; // Column E: Primary Category
        // Ensure the row has a name and the category includes 'FABRICATION'.
        return name && name.toString().trim() !== "" && primaryCategory && primaryCategory.toString().includes('FABRICATION');
      })
      .map(row => {
        // Trim the name to prevent whitespace issues during lookup.
        const name = row[1].toString().trim();     // Column B: Material Name
        let unitCost = row[9]; // Column J: Unit Cost
        
        // Clean the unitCost value (e.g., from "$1,540.16" to 1540.16)
        if (unitCost && typeof unitCost === 'string') {
          // Remove currency symbols, spaces, and commas, then convert to a number.
          const cleanedCost = parseFloat(unitCost.replace(/[^0-9.-]+/g,""));
          unitCost = isNaN(cleanedCost) ? 0 : cleanedCost;
        } else if (typeof unitCost !== 'number') {
          unitCost = 0; // Default to 0 if it's not a number or a parsable string.
        }

        return {
          name: name,
          unitCost: unitCost
        };
      });
  }
};

// An object to namespace all functions related to the "Personnel" sheet.
const personnelSheet = {
  NAME: 'Personnel', // The name of the sheet.

  /**
   * Gets the 'Personnel' sheet object from the active spreadsheet.
   * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} The sheet object or null if not found.
   */
  getSheet: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(this.NAME);
    if (!sheet) {
      console.error(`Sheet '${this.NAME}' not found. Please ensure a sheet with this exact name exists in the current spreadsheet.`);
    }
    return sheet;
  },

  /**
   * Fetches and processes all personnel data from the sheet.
   * @returns {Array<Object>} An array of objects, each with 'name' and 'projectRate' properties.
   */
  getData: function() {
    const sheet = this.getSheet();
    if (!sheet) {
      return []; // Return empty array if the sheet doesn't exist.
    }
    // Get data from Column B (Name) and Column C (ProjectRate).
    const range = sheet.getRange('B2:C' + sheet.getLastRow());
    const values = range.getValues();
    
    // Process the data into an array of objects.
    return values
      .filter(row => row[0] && row[0].toString().trim() !== "") // Filter out rows with no name.
      .map(row => {
        const name = row[0].toString().trim(); // Column B: Name
        let projectRate = row[1]; // Column C: ProjectRate

        if (typeof projectRate !== 'number') {
          projectRate = parseFloat(projectRate) || 0;
        }

        return {
          name: name,
          projectRate: projectRate
        };
      });
  }
};

// An object to namespace all functions related to the fabrication application.
const fabricationApp = {
  /**
   * Opens the modal dialog for the fabrication application.
   */
  showDialog: function() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('FabricationIndex')
        .setWidth(750)
        .setHeight(850);
    const ui = SpreadsheetApp.getUi();
    ui.showModalDialog(htmlOutput, 'Fabrication Details');
  },

  /**
   * Gets materials data for the app.
   * @returns {Array<Object>} An array of objects with material data.
   */
  getMaterials: function() {
    try {
      return materialsSheet.getData();
    } catch (e) {
      console.error("Error in fabricationApp.getMaterials: " + e.toString());
      return [];
    }
  },

  /**
   * Gets personnel data for the app.
   * @returns {Array<Object>} An array of objects with personnel data.
   */
  getPersonnel: function() {
    try {
      return personnelSheet.getData();
    } catch (e) {
      console.error("Error in fabricationApp.getPersonnel: " + e.toString());
      return [];
    }
  },

  /**
   * Opens the fabrication app with pre-populated data for editing.
   * @param {string} logId - The unique log ID to restore data from
   */
  openForEdit: function(logId) {
    const formData = projectSheet.getLoggedFormData(logId, 'FabricationLog');
    if (formData) {
      const htmlOutput = HtmlService.createHtmlOutputFromFile('FabricationIndex')
          .setWidth(750)
          .setHeight(850);
      
      // Pass the form data to the client
      const htmlContent = htmlOutput.getContent();
      const modifiedContent = htmlContent.replace(
        '<script>',
        `<script>window.editFormData = ${JSON.stringify(formData)};`
      );
      
      const modifiedOutput = HtmlService.createHtmlOutput(modifiedContent)
          .setWidth(750)
          .setHeight(850);
      
      const ui = SpreadsheetApp.getUi();
      ui.showModalDialog(modifiedOutput, 'Edit Fabrication Details');
    } else {
      this.showDialog();
    }
  },

  /**
   * Adds a fabrication item to the project sheet.
   * @param {Object} fabricationData - Object containing description, dimensions, totalPrice, and formData
   * @returns {Object} Result object with success status and message
   */
  addToProject: function(fabricationData) {
    try {
      return projectSheet.addProjectItem(fabricationData, 'FAB', 'FabricationLog');
    } catch (e) {
      console.error("Error in fabricationApp.addToProject: " + e.toString());
      return {
        success: false,
        message: `Error adding to project: ${e.toString()}`,
        rowNumber: null,
        logId: null
      };
    }
  }
};

// An object to namespace all functions related to the apparel estimate application.
const apparelApp = {
  /**
   * Opens the modal dialog for the apparel estimate application.
   */
  showDialog: function() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('ApparelIndex')
        .setWidth(750)
        .setHeight(850);
    const ui = SpreadsheetApp.getUi();
    ui.showModalDialog(htmlOutput, 'Apparel / Screen Printing');
  },

  /**
   * Opens the apparel app with pre-populated data for editing.
   * @param {string} logId - The unique log ID to restore data from
   */
  openForEdit: function(logId) {
    const formData = projectSheet.getLoggedFormData(logId, 'ApparelLog');
    if (formData) {
      const htmlOutput = HtmlService.createHtmlOutputFromFile('ApparelIndex')
          .setWidth(750)
          .setHeight(850);
      
      // Pass the form data to the client
      const htmlContent = htmlOutput.getContent();
      const modifiedContent = htmlContent.replace(
        '<script>',
        `<script>window.editFormData = ${JSON.stringify(formData)};`
      );
      
      const modifiedOutput = HtmlService.createHtmlOutput(modifiedContent)
          .setWidth(750)
          .setHeight(850);
      
      const ui = SpreadsheetApp.getUi();
      ui.showModalDialog(modifiedOutput, 'Edit Apparel Estimate');
    } else {
      this.showDialog();
    }
  },

  /**
   * Adds an apparel estimate to the project sheet.
   * @param {Object} apparelData - Object containing description, quantity, totalPrice, and formData
   * @returns {Object} Result object with success status and message
   */
  addToProject: function(apparelData) {
    try {
      return projectSheet.addProjectItem(apparelData, 'APP', 'ApparelLog');
    } catch (e) {
      console.error("Error in apparelApp.addToProject: " + e.toString());
      return {
        success: false,
        message: `Error adding to project: ${e.toString()}`,
        rowNumber: null,
        logId: null
      };
    }
  }
};

// An object to namespace all functions related to project data management.
const projectSheet = {
  /**
   * Gets the active sheet (assumes it's the project sheet where data should be written).
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} The active sheet object.
   */
  getActiveSheet: function() {
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  },

  /**
   * Logs form data to a hidden sheet for edit functionality.
   * @param {Object} formData - Complete form data
   * @param {number} projectRowNumber - The row number in the project sheet
   * @param {string} logIdPrefix - Prefix for the log ID (e.g., 'FAB', 'APP')
   * @param {string} logSheetName - Name of the log sheet (e.g., 'FabricationLog', 'ApparelLog')
   * @returns {string} Unique log ID
   */
  logFormData: function(formData, projectRowNumber, logIdPrefix, logSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let logSheet = spreadsheet.getSheetByName(logSheetName);
      
      // Create log sheet if it doesn't exist
      if (!logSheet) {
        logSheet = spreadsheet.insertSheet(logSheetName);
        logSheet.hideSheet();
        logSheet.getRange(1, 1, 1, 4).setValues([
          ['LogID', 'ProjectRow', 'Timestamp', 'FormData']
        ]);
      }
      
      // Generate unique log ID
      const logId = `${logIdPrefix}_${Date.now()}_${projectRowNumber}`;
      const timestamp = new Date();
      
      // Add original row number to form data for edit tracking
      const formDataWithRow = {
        ...formData,
        originalRowNumber: projectRowNumber
      };
      const formDataJson = JSON.stringify(formDataWithRow);
      
      // Find next empty row in log sheet
      const lastLogRow = logSheet.getLastRow();
      const nextLogRow = lastLogRow + 1;
      
      // Write log data
      logSheet.getRange(nextLogRow, 1, 1, 4).setValues([
        [logId, projectRowNumber, timestamp, formDataJson]
      ]);
      
      return logId;
      
    } catch (error) {
      console.error('Error logging form data:', error);
      return null;
    }
  },

  /**
   * Updates existing log data for an edited item.
   * @param {Object} formData - Complete form data
   * @param {number} projectRowNumber - The row number in the project sheet
   * @param {string} logIdPrefix - Prefix for the log ID (e.g., 'FAB', 'APP')
   * @param {string} logSheetName - Name of the log sheet
   * @returns {string} Updated log ID
   */
  updateLogFormData: function(formData, projectRowNumber, logIdPrefix, logSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = spreadsheet.getSheetByName(logSheetName);
      
      if (!logSheet) {
        return this.logFormData(formData, projectRowNumber, logIdPrefix, logSheetName);
      }
      
      const dataRange = logSheet.getDataRange();
      const values = dataRange.getValues();
      
      // Find existing log entry for this row
      let existingRowIndex = -1;
      for (let i = 1; i < values.length; i++) {
        if (values[i][1] === projectRowNumber) {
          existingRowIndex = i + 1;
          break;
        }
      }
      
      // Generate new log ID with current timestamp
      const logId = `${logIdPrefix}_${Date.now()}_${projectRowNumber}`;
      const timestamp = new Date();
      
      const formDataWithRow = {
        ...formData,
        originalRowNumber: projectRowNumber
      };
      const formDataJson = JSON.stringify(formDataWithRow);
      
      if (existingRowIndex > 0) {
        logSheet.getRange(existingRowIndex, 1, 1, 4).setValues([
          [logId, projectRowNumber, timestamp, formDataJson]
        ]);
      } else {
        const lastLogRow = logSheet.getLastRow();
        const nextLogRow = lastLogRow + 1;
        logSheet.getRange(nextLogRow, 1, 1, 4).setValues([
          [logId, projectRowNumber, timestamp, formDataJson]
        ]);
      }
      
      return logId;
      
    } catch (error) {
      console.error('Error updating log data:', error);
      return null;
    }
  },

  /**
   * Retrieves logged form data by log ID.
   * @param {string} logId - The unique log ID
   * @param {string} logSheetName - Name of the log sheet to search
   * @returns {Object|null} Form data object or null if not found
   */
  getLoggedFormData: function(logId, logSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = spreadsheet.getSheetByName(logSheetName);
      
      if (!logSheet) {
        return null;
      }
      
      const dataRange = logSheet.getDataRange();
      const values = dataRange.getValues();
      
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === logId) {
          const formDataJson = values[i][3];
          return JSON.parse(formDataJson);
        }
      }
      
      return null;
      
    } catch (error) {
      console.error('Error retrieving form data:', error);
      return null;
    }
  },

  /**
   * Creates an edit instruction for the spreadsheet.
   * @param {string} logId - The unique log ID
   * @returns {string} Edit instruction text
   */
  createEditInstruction: function(logId) {
    return "Edit";
  },

  /**
   * Generates the next sequential ID for fabrication items
   * @returns {string} Next fabrication ID (e.g., F01, F02, F03)
   */
  getNextFabricationId: function() {
    const sheet = this.getActiveSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let maxNumber = 0;
    
    // Look through column B for existing fabrication IDs
    for (let i = 1; i < values.length; i++) { // Start from row 2 (index 1)
      const cellValue = values[i][1]; // Column B (index 1)
      if (cellValue && typeof cellValue === 'string' && cellValue.startsWith('F')) {
        const numberPart = cellValue.substring(1);
        const number = parseInt(numberPart);
        if (!isNaN(number) && number > maxNumber) {
          maxNumber = number;
        }
      }
    }
    
    // Generate next ID with zero padding
    const nextNumber = maxNumber + 1;
    return `F${nextNumber.toString().padStart(2, '0')}`;
  },

  /**
   * Updates an existing project item in the sheet.
   */
  updateProjectItem: function(itemData, logIdPrefix, logSheetName) {
    try {
      const sheet = this.getActiveSheet();
      
      if (!itemData || typeof itemData !== 'object') {
        throw new Error('Invalid item data provided');
      }

      const { description, quantity, dimensions, totalPrice, formData, originalRowNumber } = itemData;
      
      const rowNum = parseInt(originalRowNumber);
      if (!rowNum || isNaN(rowNum) || rowNum < 1) {
        throw new Error(`Invalid original row number for update: ${originalRowNumber}`);
      }
      
      const maxRows = sheet.getMaxRows();
      if (rowNum > maxRows) {
        throw new Error(`Row number ${rowNum} exceeds sheet maximum rows ${maxRows}`);
      }
      
      const logId = this.updateLogFormData(formData, rowNum, logIdPrefix, logSheetName);
      
      let rowData;
      
      // Different column layouts for different app types
      if (logIdPrefix === 'FAB') {
        // Fabrication: A=empty, B=ID (preserve existing), C=Description, D=Dimensions, E=empty, F=Total Price, G=Edit
        const existingId = sheet.getRange(rowNum, 2).getValue() || this.getNextFabricationId();
        rowData = [
          '',                           // Column A (empty)
          existingId,                   // Column B (preserve existing ID)
          description || '',            // Column C (Description)
          dimensions || '',             // Column D (Dimensions)  
          '',                           // Column E (empty)
          totalPrice || 0,              // Column F (Total Price)
          'Edit'                        // Column G (Edit - preserve)
        ];
        
        // Write to columns A-F only, preserve G
        const range = sheet.getRange(rowNum, 1, 1, 6);
        range.setValues([rowData.slice(0, 6)]);
        
        // Format the total price as currency
        const priceCell = sheet.getRange(rowNum, 6); // Column F
        priceCell.setNumberFormat('$#,##0.00');
        
        // Update the edit cell note
        if (logId) {
          const editCell = sheet.getRange(rowNum, 7); // Column G
          editCell.setNote(`LogID: ${logId}\n\nTo edit this item:\n1. Select this cell\n2. Go to Production > Edit Selected Item\n\nLast updated: ${new Date().toLocaleString()}`);
        }
        
      } else {
        // Apparel and other apps: A=Description, B=Quantity, C=empty, D=Total Price, E=Edit
        rowData = [
          description || '',
          quantity || '',
          '',
          totalPrice || 0
        ];
        
        const range = sheet.getRange(rowNum, 1, 1, 4);
        range.setValues([rowData]);
        
        const priceCell = sheet.getRange(rowNum, 4);
        priceCell.setNumberFormat('$#,##0.00');
        
        if (logId) {
          const editCell = sheet.getRange(rowNum, 5);
          editCell.setNote(`LogID: ${logId}\n\nTo edit this item:\n1. Select this cell\n2. Go to Production > Edit Selected Item\n\nLast updated: ${new Date().toLocaleString()}`);
        }
      }
      
      return {
        success: true,
        message: `Item updated in row ${rowNum}`,
        rowNumber: rowNum,
        logId: logId,
        isUpdate: true
      };
      
    } catch (error) {
      console.error('Error updating project item:', error);
      return {
        success: false,
        message: `Error updating item: ${error.message}`,
        rowNumber: null,
        logId: null,
        isUpdate: false
      };
    }
  },

  /**
   * Adds a new project item to the sheet or updates existing one.
   */
  addProjectItem: function(itemData, logIdPrefix, logSheetName) {
    try {
      let originalRowNumber = null;
      
      if (itemData.originalRowNumber) {
        originalRowNumber = itemData.originalRowNumber;
      } else if (itemData.formData && itemData.formData.originalRowNumber) {
        originalRowNumber = itemData.formData.originalRowNumber;
      }
      
      if (originalRowNumber && originalRowNumber > 0) {
        return this.updateProjectItem({
          ...itemData,
          originalRowNumber: originalRowNumber
        }, logIdPrefix, logSheetName);
      }
      
      const sheet = this.getActiveSheet();
      
      if (!itemData || typeof itemData !== 'object') {
        throw new Error('Invalid item data provided');
      }

      const { description, quantity, dimensions, totalPrice, formData } = itemData;
      
      const lastRow = sheet.getLastRow();
      const nextRow = lastRow + 1;
      
      const logId = this.logFormData(formData, nextRow, logIdPrefix, logSheetName);
      
      const editInstruction = logId ? this.createEditInstruction(logId) : 'Edit';
      
      let rowData;
      let editColumnIndex;
      
      // Different column layouts for different app types
      if (logIdPrefix === 'FAB') {
        // Fabrication: A=empty, B=Auto ID, C=Description, D=Dimensions, E=empty, F=Total Price, G=Edit
        const fabricationId = this.getNextFabricationId();
        rowData = [
          '',                           // Column A (empty)
          fabricationId,                // Column B (Auto-generated ID)
          description || '',            // Column C (Description)
          dimensions || '',             // Column D (Dimensions)
          '',                           // Column E (empty)
          totalPrice || 0,              // Column F (Total Price)
          editInstruction               // Column G (Edit)
        ];
        editColumnIndex = 7; // Column G
      } else {
        // Apparel and other apps: A=Description, B=Quantity, C=empty, D=Total Price, E=Edit
        rowData = [
          description || '',
          quantity || '',
          '',
          totalPrice || 0,
          editInstruction
        ];
        editColumnIndex = 5; // Column E
      }
      
      const range = sheet.getRange(nextRow, 1, 1, rowData.length);
      range.setValues([rowData]);
      
      // Format the total price as currency
      if (logIdPrefix === 'FAB') {
        const priceCell = sheet.getRange(nextRow, 6); // Column F for fabrication
        priceCell.setNumberFormat('$#,##0.00');
      } else {
        const priceCell = sheet.getRange(nextRow, 4); // Column D for apparel
        priceCell.setNumberFormat('$#,##0.00');
      }
      
      if (logId) {
        const editCell = sheet.getRange(nextRow, editColumnIndex);
        editCell.setNote(`LogID: ${logId}\n\nTo edit this item:\n1. Select this cell\n2. Go to Production > Edit Selected Item`);
        editCell.setBackground('#e3f2fd');
        editCell.setFontColor('#1976d2');
        editCell.setFontWeight('bold');
      }
      
      return {
        success: true,
        message: `Item added to row ${nextRow}`,
        rowNumber: nextRow,
        logId: logId,
        isUpdate: false
      };
      
    } catch (error) {
      console.error('Error adding project item:', error);
      return {
        success: false,
        message: `Error adding item: ${error.message}`,
        rowNumber: null,
        logId: null,
        isUpdate: false
      };
    }
  }
};

/**
 * Creates a custom menu in the UI when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Production')
      .addItem('Fabrication', 'openFabricationApp')
      .addItem('Apparel Estimate', 'openApparelApp')
      .addSeparator()
      .addItem('Edit Selected Item', 'editSelectedItem')
      .addToUi();
}

/**
 * Global function to open the fabrication app.
 */
function openFabricationApp() {
  fabricationApp.showDialog();
}

/**
 * Global function to open the apparel app.
 */
function openApparelApp() {
  apparelApp.showDialog();
}

/**
 * Global function for fabrication app to get materials data.
 */
function getMaterials() {
  return fabricationApp.getMaterials();
}

/**
 * Global function for fabrication app to get personnel data.
 */
function getPersonnel() {
  return fabricationApp.getPersonnel();
}

/**
 * Global function for fabrication app to add data to project.
 */
function addFabricationToProject(fabricationData) {
  return fabricationApp.addToProject(fabricationData);
}

/**
 * Global function for apparel app to add data to project.
 */
function addApparelToProject(apparelData) {
  return apparelApp.addToProject(apparelData);
}

/**
 * Global function to open fabrication app for editing with specific log ID.
 */
function openFabricationAppForEdit(logId) {
  return fabricationApp.openForEdit(logId);
}

/**
 * Global function to open apparel app for editing with specific log ID.
 */
function openApparelAppForEdit(logId) {
  return apparelApp.openForEdit(logId);
}

/**
 * Global function to get logged form data by ID and sheet name.
 */
function getLoggedFormData(logId, logSheetName) {
  return projectSheet.getLoggedFormData(logId, logSheetName);
}

/**
 * Function to edit the selected item from the menu.
 */
function editSelectedItem() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const activeCell = sheet.getActiveCell();
    
    // Check if the selected cell is in the edit column and has a note with LogID
    const column = activeCell.getColumn();
    if (column !== 5 && column !== 7) { // Column E (5) for apparel, Column G (7) for fabrication
      SpreadsheetApp.getUi().alert(
        'Edit Item', 
        'Please select an "Edit" cell first, then try again.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const cellNote = activeCell.getNote();
    if (!cellNote || !cellNote.includes('LogID:')) {
      SpreadsheetApp.getUi().alert(
        'Edit Item', 
        'No edit data found for this item.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const logIdMatch = cellNote.match(/LogID:\s*([^\n\r]+)/);
    if (!logIdMatch) {
      SpreadsheetApp.getUi().alert(
        'Edit Item', 
        'Could not find LogID in the selected cell.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const logId = logIdMatch[1].trim();
    
    if (logId.startsWith('FAB_')) {
      fabricationApp.openForEdit(logId);
    } else if (logId.startsWith('APP_')) {
      apparelApp.openForEdit(logId);
    } else {
      SpreadsheetApp.getUi().alert(
        'Edit Item', 
        'Unknown item type. Cannot determine which editor to open.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
  } catch (error) {
    console.error('Error in editSelectedItem:', error);
    SpreadsheetApp.getUi().alert(
      'Error', 
      'An error occurred while trying to edit the item: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
