/**
 * @OnlyCurrentDoc
 * This script manages multiple production applications with a shared project management system.
 */

// An object to namespace all functions related to the "Materials" sheet.
const materialsSheet = {
  NAME: 'Materials',

  getSheet: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(this.NAME);
    if (!sheet) {
      console.error(`Sheet '${this.NAME}' not found. Please ensure a sheet with this exact name exists in the current spreadsheet.`);
    }
    return sheet;
  },

  getData: function(category = 'FABRICATION') {
    const sheet = this.getSheet();
    if (!sheet) {
      return [];
    }
    const range = sheet.getRange('A2:P' + sheet.getLastRow());
    const values = range.getValues();

    const filteredValues = values
      .filter(row => {
        const name = row[1];
        const primaryCategory = row[4];
        return name && name.toString().trim() !== "" && primaryCategory && primaryCategory.toString().toUpperCase().includes(category);
      });

    if (category === 'PRINT') {
        return filteredValues.map(row => {
            const name = row[1].toString().trim(); // Column B
            const type = row[6] ? row[6].toString().trim().toUpperCase() : 'SHEET'; // Column G (Material Type)
            const width = parseFloat(row[7]) || 0; // Column H (inches)
            const length = parseFloat(row[8]) || 0; // Column I (feet for ROLL, inches for SHEET)
            let sheetCost = row[9]; // Column J

            if (sheetCost && typeof sheetCost === 'string') {
                const cleanedCost = parseFloat(sheetCost.replace(/[^0-9.-]+/g,""));
                sheetCost = isNaN(cleanedCost) ? 0 : cleanedCost;
            } else if (typeof sheetCost !== 'number') {
                sheetCost = 0;
            }
            
            // Calculate linear foot cost for ROLL materials
            // For ROLL: length is already in feet, so cost per linear foot = unit cost / length
            let costLinFt = 0;
            if (type === 'ROLL' && length > 0) {
                costLinFt = sheetCost / length;
            }
            
            // Return: [name, type, width, height, costSheet, costLinFt]
            return [name, type, width, length, sheetCost, costLinFt];
        });
    } else { // 'FABRICATION' and default
        return filteredValues.map(row => {
            const name = row[1].toString().trim();
            let unitCost = row[9];
            if (unitCost && typeof unitCost === 'string') {
                const cleanedCost = parseFloat(unitCost.replace(/[^0-9.-]+/g,""));
                unitCost = isNaN(cleanedCost) ? 0 : cleanedCost;
            } else if (typeof unitCost !== 'number') {
                unitCost = 0;
            }
            return {
              name: name,
              unitCost: unitCost
            };
        });
    }
  }
};

// An object to namespace all functions related to the "Personnel" sheet.
const personnelSheet = {
  NAME: 'Personnel',

  getSheet: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(this.NAME);
    if (!sheet) {
      console.error(`Sheet '${this.NAME}' not found. Please ensure a sheet with this exact name exists in the current spreadsheet.`);
    }
    return sheet;
  },

  getData: function() {
    const sheet = this.getSheet();
    if (!sheet) {
      return [];
    }
    const range = sheet.getRange('B2:C' + sheet.getLastRow());
    const values = range.getValues();
    
    return values
      .filter(row => row[0] && row[0].toString().trim() !== "")
      .map(row => {
        const name = row[0].toString().trim();
        let projectRate = row[1];

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
  showDialog: function() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('FabricationIndex')
        .setWidth(750)
        .setHeight(850);
    const ui = SpreadsheetApp.getUi();
    ui.showModalDialog(htmlOutput, 'Fabrication Details');
  },

  getMaterials: function() {
    try {
      return materialsSheet.getData('FABRICATION');
    } catch (e) {
      console.error("Error in fabricationApp.getMaterials: " + e.toString());
      return [];
    }
  },

  getPersonnel: function() {
    try {
      return personnelSheet.getData();
    } catch (e) {
      console.error("Error in fabricationApp.getPersonnel: " + e.toString());
      return [];
    }
  },

  openForEdit: function(logId) {
    const formData = projectSheet.getLoggedFormData(logId, 'FabricationLog');
    if (formData) {
      const htmlOutput = HtmlService.createHtmlOutputFromFile('FabricationIndex')
          .setWidth(750)
          .setHeight(850);
      
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

// An object to namespace all functions related to the apparel application.
const apparelApp = {
  showDialog: function() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('ApparelIndex')
        .setWidth(750)
        .setHeight(850);
    const ui = SpreadsheetApp.getUi();
    ui.showModalDialog(htmlOutput, 'Apparel / Screen Printing');
  },

  openForEdit: function(logId) {
    const formData = projectSheet.getLoggedFormData(logId, 'ApparelLog');
    if (formData) {
      const htmlOutput = HtmlService.createHtmlOutputFromFile('ApparelIndex')
          .setWidth(750)
          .setHeight(850);
      
      const htmlContent = htmlOutput.getContent();
      const modifiedContent = htmlContent.replace(
        '<script>',
        `<script>window.editFormData = ${JSON.stringify(formData)};`
      );
      
      const modifiedOutput = HtmlService.createHtmlOutput(modifiedContent)
          .setWidth(750)
          .setHeight(850);
      
      const ui = SpreadsheetApp.getUi();
      ui.showModalDialog(modifiedOutput, 'Edit Apparel');
    } else {
      this.showDialog();
    }
  },

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

// An object to namespace all functions related to the printing estimate application.
const printingApp = {
  showDialog: function() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('PrintingIndex')
        .setWidth(750)
        .setHeight(850);
    const ui = SpreadsheetApp.getUi();
    ui.showModalDialog(htmlOutput, 'PrintCut Estimate');
  },

  getMaterials: function() {
    try {
      return materialsSheet.getData('PRINT');
    } catch (e) {
      console.error("Error in printingApp.getMaterials: " + e.toString());
      return [];
    }
  },

  openForEdit: function(logId) {
    const formData = projectSheet.getLoggedFormData(logId, 'PrintingLog');
    if (formData) {
      const htmlOutput = HtmlService.createHtmlOutputFromFile('PrintingIndex')
          .setWidth(750)
          .setHeight(850);
      
      const htmlContent = htmlOutput.getContent();
      const modifiedContent = htmlContent.replace(
        '<script>',
        `<script>window.editFormData = ${JSON.stringify(formData)};`
      );
      
      const modifiedOutput = HtmlService.createHtmlOutput(modifiedContent)
          .setWidth(750)
          .setHeight(850);
      
      const ui = SpreadsheetApp.getUi();
      ui.showModalDialog(modifiedOutput, 'Edit PrintCut Estimate');
    } else {
      this.showDialog();
    }
  },

  addToProject: function(printingData) {
    try {
      return projectSheet.addProjectItem(printingData, 'PRT', 'PrintingLog');
    } catch (e) {
      console.error("Error in printingApp.addToProject: " + e.toString());
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
  getActiveSheet: function() {
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  },

  logFormData: function(formData, projectRowNumber, logIdPrefix, logSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let logSheet = spreadsheet.getSheetByName(logSheetName);
      
      if (!logSheet) {
        logSheet = spreadsheet.insertSheet(logSheetName);
        logSheet.hideSheet();
        logSheet.getRange(1, 1, 1, 4).setValues([
          ['LogID', 'ProjectRow', 'Timestamp', 'FormData']
        ]);
      }
      
      const logId = `${logIdPrefix}_${Date.now()}_${projectRowNumber}`;
      const timestamp = new Date();
      
      const formDataWithRow = {
        ...formData,
        originalRowNumber: projectRowNumber
      };
      const formDataJson = JSON.stringify(formDataWithRow);
      
      const lastLogRow = logSheet.getLastRow();
      const nextLogRow = lastLogRow + 1;
      
      logSheet.getRange(nextLogRow, 1, 1, 4).setValues([
        [logId, projectRowNumber, timestamp, formDataJson]
      ]);
      
      return logId;
      
    } catch (error) {
      console.error('Error logging form data:', error);
      return null;
    }
  },

  updateLogFormData: function(formData, projectRowNumber, logIdPrefix, logSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = spreadsheet.getSheetByName(logSheetName);
      
      if (!logSheet) {
        return this.logFormData(formData, projectRowNumber, logIdPrefix, logSheetName);
      }
      
      const dataRange = logSheet.getDataRange();
      const values = dataRange.getValues();
      
      let existingRowIndex = -1;
      for (let i = 1; i < values.length; i++) {
        if (values[i][1] === projectRowNumber) {
          existingRowIndex = i + 1;
          break;
        }
      }
      
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

  createEditInstruction: function(logId) {
    return "Edit";
  },

  getNextFabricationId: function() {
    const sheet = this.getActiveSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let maxNumber = 0;
    
    for (let i = 1; i < values.length; i++) {
      const cellValue = values[i][1];
      if (cellValue && typeof cellValue === 'string' && cellValue.startsWith('F')) {
        const numberPart = cellValue.substring(1);
        const number = parseInt(numberPart);
        if (!isNaN(number) && number > maxNumber) {
          maxNumber = number;
        }
      }
    }
    
    const nextNumber = maxNumber + 1;
    return `F${nextNumber.toString().padStart(2, '0')}`;
  },

  getNextApparelId: function() {
    const sheet = this.getActiveSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let maxNumber = 0;
    
    for (let i = 1; i < values.length; i++) {
      const cellValue = values[i][1];
      if (cellValue && typeof cellValue === 'string' && cellValue.startsWith('AP')) {
        const numberPart = cellValue.substring(2);
        const number = parseInt(numberPart);
        if (!isNaN(number) && number > maxNumber) {
          maxNumber = number;
        }
      }
    }
    
    const nextNumber = maxNumber + 1;
    return `AP${nextNumber.toString().padStart(2, '0')}`;
  },

  getNextPrintingId: function() {
    const sheet = this.getActiveSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let maxNumber = 0;
    
    for (let i = 1; i < values.length; i++) {
      const cellValue = values[i][1];
      if (cellValue && typeof cellValue === 'string' && cellValue.startsWith('PR')) {
        const numberPart = cellValue.substring(2);
        const number = parseInt(numberPart);
        if (!isNaN(number) && number > maxNumber) {
          maxNumber = number;
        }
      }
    }
    
    const nextNumber = maxNumber + 1;
    return `PR${nextNumber.toString().padStart(2, '0')}`;
  },

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
      let editColumnIndex;
      
      if (logIdPrefix === 'FAB') {
        const existingId = sheet.getRange(rowNum, 2).getValue() || this.getNextFabricationId();
        rowData = [
          '',
          existingId,
          description || '',
          dimensions || '',
          '',
          totalPrice || 0,
          'Edit'
        ];
        editColumnIndex = 7;
        
        const range = sheet.getRange(rowNum, 1, 1, 6);
        range.setValues([rowData.slice(0, 6)]);
        
        const priceCell = sheet.getRange(rowNum, 6);
        priceCell.setNumberFormat('$#,##0.00');
        
      } else if (logIdPrefix === 'APP') {
        const existingId = sheet.getRange(rowNum, 2).getValue() || this.getNextApparelId();
        const unitPrice = quantity && quantity > 0 ? (totalPrice / quantity) : 0;
        
        rowData = [
          '',
          existingId,
          description || '',
          quantity || '',
          unitPrice,
          totalPrice || 0,
          'Edit'
        ];
        editColumnIndex = 7;
        
        const range = sheet.getRange(rowNum, 1, 1, 6);
        range.setValues([rowData.slice(0, 6)]);
        
        const unitPriceCell = sheet.getRange(rowNum, 5);
        const totalPriceCell = sheet.getRange(rowNum, 6);
        unitPriceCell.setNumberFormat('$#,##0.00');
        totalPriceCell.setNumberFormat('$#,##0.00');
        
      } else if (logIdPrefix === 'PRT') {
        const existingId = sheet.getRange(rowNum, 2).getValue() || this.getNextPrintingId();
        const unitPrice = quantity && quantity > 0 ? (totalPrice / quantity) : 0;
        
        rowData = [
          '',
          existingId,
          description || '',
          quantity || '',
          unitPrice,
          totalPrice || 0,
          'Edit'
        ];
        editColumnIndex = 7;
        
        const range = sheet.getRange(rowNum, 1, 1, 6);
        range.setValues([rowData.slice(0, 6)]);
        
        const unitPriceCell = sheet.getRange(rowNum, 5);
        const totalPriceCell = sheet.getRange(rowNum, 6);
        unitPriceCell.setNumberFormat('$#,##0.00');
        totalPriceCell.setNumberFormat('$#,##0.00');
      }
      
      if (logId) {
        const editCell = sheet.getRange(rowNum, editColumnIndex);
        editCell.setNote(`LogID: ${logId}\n\nTo edit this item:\n1. Select this cell\n2. Go to Production > Edit Selected Item\n\nLast updated: ${new Date().toLocaleString()}`);
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
      
      if (logIdPrefix === 'FAB') {
        const fabricationId = this.getNextFabricationId();
        rowData = [
          '',
          fabricationId,
          description || '',
          dimensions || '',
          '',
          totalPrice || 0,
          editInstruction
        ];
        editColumnIndex = 7;
        
        const range = sheet.getRange(nextRow, 1, 1, rowData.length);
        range.setValues([rowData]);
        
        const priceCell = sheet.getRange(nextRow, 6);
        priceCell.setNumberFormat('$#,##0.00');
        
      } else if (logIdPrefix === 'APP') {
        const apparelId = this.getNextApparelId();
        const unitPrice = quantity && quantity > 0 ? (totalPrice / quantity) : 0;
        
        rowData = [
          '',
          apparelId,
          description || '',
          quantity || '',
          unitPrice,
          totalPrice || 0,
          editInstruction
        ];
        editColumnIndex = 7;
        
        const range = sheet.getRange(nextRow, 1, 1, rowData.length);
        range.setValues([rowData]);
        
        const unitPriceCell = sheet.getRange(nextRow, 5);
        const totalPriceCell = sheet.getRange(nextRow, 6);
        unitPriceCell.setNumberFormat('$#,##0.00');
        totalPriceCell.setNumberFormat('$#,##0.00');
        
      } else if (logIdPrefix === 'PRT') {
        const printingId = this.getNextPrintingId();
        const unitPrice = quantity && quantity > 0 ? (totalPrice / quantity) : 0;
        
        rowData = [
          '',
          printingId,
          description || '',
          quantity || '',
          unitPrice,
          totalPrice || 0,
          editInstruction
        ];
        editColumnIndex = 7;
        
        const range = sheet.getRange(nextRow, 1, 1, rowData.length);
        range.setValues([rowData]);
        
        const unitPriceCell = sheet.getRange(nextRow, 5);
        const totalPriceCell = sheet.getRange(nextRow, 6);
        unitPriceCell.setNumberFormat('$#,##0.00');
        totalPriceCell.setNumberFormat('$#,##0.00');
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

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Production')
      .addItem('Fabrication', 'openFabricationApp')
      .addItem('Apparel', 'openApparelApp')
      .addItem('Printing', 'openPrintingApp')
      .addSeparator()
      .addItem('Edit Selected Item', 'editSelectedItem')
      .addToUi();
}

function openFabricationApp() {
  fabricationApp.showDialog();
}

function openApparelApp() {
  apparelApp.showDialog();
}

function openPrintingApp() {
  printingApp.showDialog();
}

function getMaterials() {
  return fabricationApp.getMaterials();
}

function getPersonnel() {
  return fabricationApp.getPersonnel();
}

function getPrintingMaterials() {
  return printingApp.getMaterials();
}

function addFabricationToProject(fabricationData) {
  return fabricationApp.addToProject(fabricationData);
}

function addApparelToProject(apparelData) {
  return apparelApp.addToProject(apparelData);
}

function addPrintingToProject(printingData) {
  return printingApp.addToProject(printingData);
}

function openFabricationAppForEdit(logId) {
  return fabricationApp.openForEdit(logId);
}

function openApparelAppForEdit(logId) {
  return apparelApp.openForEdit(logId);
}

function openPrintingAppForEdit(logId) {
  return printingApp.openForEdit(logId);
}

function getLoggedFormData(logId, logSheetName) {
  return projectSheet.getLoggedFormData(logId, logSheetName);
}

function editSelectedItem() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const activeCell = sheet.getActiveCell();
    
    const column = activeCell.getColumn();
    if (column !== 7) {
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
    } else if (logId.startsWith('PRT_')) {
      printingApp.openForEdit(logId);
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
