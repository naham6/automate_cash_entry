
const CONFIG = {
  SPREADSHEET_ID:
    '',

  SHEET_NAME: 'Sheet1',

  TIMEZONE: 'Asia/Dhaka',

  APP_TITLE: 'Collection',

  ACCESS_CODE: '1234',

  SECTIONS: [
    'Brand1',
    'Brand2'
  ],

  // Salesperson names are in column A.
  NAME_COLUMN: 1
};




const AUDIT_CONFIG = {
  SPREADSHEET_ID:
    '',

  SHEET_NAME: 'Log',

  TIMEZONE: 'Asia/Dhaka'
};

const AUDIT_HEADERS = [
  'Timestamp',
  'Collection Type',
  'Salesperson',
  'Sales Date',
  'Action',
  'Amount Submitted',
  'Previous Value',
  'New Value',
  'Submission Status'
];



function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle(CONFIG.APP_TITLE)
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}



function testSpreadsheetConnection() {
  const spreadsheet =
    SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  Logger.log('Spreadsheet name: ' + spreadsheet.getName());
  Logger.log('Spreadsheet ID: ' + spreadsheet.getId());

  spreadsheet.getSheets().forEach(function(sheet) {
    Logger.log('Sheet: ' + sheet.getName());
  });
}



function setupCashCollectionLog() {
  const spreadsheet = SpreadsheetApp.openById(
    AUDIT_CONFIG.SPREADSHEET_ID
  );

  let sheet = spreadsheet.getSheetByName(
    AUDIT_CONFIG.SHEET_NAME
  );

  if (!sheet) {
    sheet = spreadsheet.insertSheet(
      AUDIT_CONFIG.SHEET_NAME
    );
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(
      1,
      1,
      1,
      AUDIT_HEADERS.length
    ).setValues([AUDIT_HEADERS]);

    sheet.setFrozenRows(1);
  } else {
    
    const existingHeaders = sheet.getRange(
      1,
      1,
      1,
      AUDIT_HEADERS.length
    ).getDisplayValues()[0];

    const headersMatch = AUDIT_HEADERS.every(
      function(header, index) {
        return String(existingHeaders[index] || '').trim() === header;
      }
    );

    if (!headersMatch) {
      throw new Error(
        'The existing Log tab has different headers. ' +
        'Please check the tab before running setup again.'
      );
    }
  }

  Logger.log(
    'Cash log is ready: ' +
    spreadsheet.getUrl() +
    ' / ' +
    AUDIT_CONFIG.SHEET_NAME
  );

  return {
    success: true,
    message: 'Cash audit log is ready.'
  };
}




function logCollectionSubmission(
  sectionName,
  salesperson,
  selectedDate,
  action,
  amountEntered,
  previousValue,
  newValue
) {
  try {
    const spreadsheet = SpreadsheetApp.openById(
      AUDIT_CONFIG.SPREADSHEET_ID
    );

    let sheet = spreadsheet.getSheetByName(
      AUDIT_CONFIG.SHEET_NAME
    );

    // Automatically create and initialize the log tab if it has not yet been created.

    if (!sheet) {
      sheet = spreadsheet.insertSheet(
        AUDIT_CONFIG.SHEET_NAME
      );

      sheet.getRange(
        1,
        1,
        1,
        AUDIT_HEADERS.length
      ).setValues([AUDIT_HEADERS]);

      sheet.setFrozenRows(1);
    } else if (sheet.getLastRow() === 0) {
      sheet.getRange(
        1,
        1,
        1,
        AUDIT_HEADERS.length
      ).setValues([AUDIT_HEADERS]);

      sheet.setFrozenRows(1);
    }

    const timestamp = Utilities.formatDate(
      new Date(),
      AUDIT_CONFIG.TIMEZONE,
      'yyyy-MM-dd HH:mm:ss'
    );

    const displayDate = formatIsoDateForDisplay(
      selectedDate
    );

    sheet.appendRow([
      timestamp,
      sectionName,
      salesperson,
      displayDate,
      action,
      amountEntered,
      previousValue,
      newValue,
      'SUCCESS'
    ]);

    SpreadsheetApp.flush();

  } catch (error) {

    console.error(
      'Cash collection audit logging failed: ' +
      (error && error.message ? error.message : error)
    );
  }
}



function verifyAccessCode(code) {
  code = String(code || '').trim();

  if (!/^\d{4}$/.test(code)) {
    return {
      success: false,
      message: 'Please enter the 4-digit access code.'
    };
  }

  if (code !== CONFIG.ACCESS_CODE) {
    return {
      success: false,
      message: 'Incorrect access code.'
    };
  }

  return { success: true };
}




function getAppData() {
  try {
    validateConfiguration();

    const spreadsheet =
      SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    const sheet = getSalesSheet(spreadsheet);
    const today = new Date();

    const selectedDate = Utilities.formatDate(
      today,
      CONFIG.TIMEZONE,
      'yyyy-MM-dd'
    );

    const formattedDate = Utilities.formatDate(
      today,
      CONFIG.TIMEZONE,
      'MMMM d, yyyy'
    );

    const sectionDefinitions =
      getSectionDefinitions(sheet);

    const sections = sectionDefinitions.map(function(section) {
      return {
        name: section.name,

        salespeople: getSalespeopleForSection(
          sheet,
          section
        ),

        availableDates: getAvailableDateColumnsForSection(
          sheet,
          section
        )
      };
    });

    return {
      success: true,
      date: formattedDate,
      selectedDate: selectedDate,
      sections: sections
    };

  } catch (error) {
    console.error(error);

    return {
      success: false,
      message: friendlyErrorMessage(error)
    };
  }
}




function getSalesSheet(spreadsheet) {
  const sheet =
    spreadsheet.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error('INVALID_SHEET_NAME');
  }

  return sheet;
}


function getSectionDefinitions(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow === 0 || lastColumn === 0) {
    throw new Error('NO_SHEETS_FOUND');
  }

  const values = sheet
    .getRange(1, 1, lastRow, lastColumn)
    .getDisplayValues();

  const sectionDefinitions = [];

  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    let foundSection = null;

    for (
      let columnIndex = 0;
      columnIndex < values[rowIndex].length;
      columnIndex++
    ) {
      const cellText = cleanCellText(
        values[rowIndex][columnIndex]
      );

      if (isConfiguredSection(cellText)) {
        foundSection = cellText;
        break;
      }
    }

    if (!foundSection) {
      continue;
    }

    const titleRow = rowIndex + 1;

    const headerRow = findSectionHeaderRow(
      sheet,
      titleRow,
      lastRow
    );

    if (headerRow === -1) {
      throw new Error(
        'SECTION_HEADER_NOT_FOUND:' +
        normalizeName(foundSection).toUpperCase()
      );
    }

    sectionDefinitions.push({
      name: foundSection,
      titleRow: titleRow,
      headerRow: headerRow,
      endRow: lastRow
    });
  }

  if (sectionDefinitions.length === 0) {
    throw new Error('NO_COLLECTION_SECTIONS');
  }

  sectionDefinitions.sort(function(a, b) {
    return a.titleRow - b.titleRow;
  });

  for (let i = 0; i < sectionDefinitions.length; i++) {
    if (i < sectionDefinitions.length - 1) {
      sectionDefinitions[i].endRow =
        sectionDefinitions[i + 1].titleRow - 1;
    }
  }

  return sectionDefinitions;
}



function findSectionHeaderRow(sheet, titleRow, lastRow) {
  const searchStart = titleRow + 1;
  const searchEnd = Math.min(titleRow + 3, lastRow);

  if (searchStart > searchEnd) {
    return -1;
  }

  const values = sheet
    .getRange(
      searchStart,
      CONFIG.NAME_COLUMN,
      searchEnd - searchStart + 1,
      1
    )
    .getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    if (normalizeName(values[i][0]) === 'name') {
      return searchStart + i;
    }
  }

  return -1;
}


function isConfiguredSection(name) {
  const target = normalizeName(name);

  return CONFIG.SECTIONS.some(function(sectionName) {
    return normalizeName(sectionName) === target;
  });
}



function getSectionDefinition(sheet, sectionName) {
  sectionName = String(sectionName || '').trim();

  if (!sectionName) {
    throw new Error('SECTION_REQUIRED');
  }

  const sections = getSectionDefinitions(sheet);
  const target = normalizeName(sectionName);

  for (let i = 0; i < sections.length; i++) {
    if (normalizeName(sections[i].name) === target) {
      return sections[i];
    }
  }

  throw new Error('SECTION_NOT_FOUND');
}



function getSalespeopleForSection(sheet, section) {
  const firstRow = section.headerRow + 1;

  const lastSalespersonRow =
    findLastSalespersonRow(sheet, section);

  if (lastSalespersonRow < firstRow) {
    throw new Error('NO_SALESPERSONS:' + section.name);
  }

  const values = sheet
    .getRange(
      firstRow,
      CONFIG.NAME_COLUMN,
      lastSalespersonRow - firstRow + 1,
      1
    )
    .getDisplayValues();

  const names = [];

  values.forEach(function(row) {
    const name = cleanCellText(row[0]);
    const normalized = normalizeName(name);

    if (
      name !== '' &&
      normalized !== 'day wise collection' &&
      normalized !== 'name' &&
      !isConfiguredSection(name)
    ) {
      names.push(name);
    }
  });

  return names;
}



function findLastSalespersonRow(sheet, section) {
  const firstRow = section.headerRow + 1;
  const lastRow = section.endRow;

  if (lastRow < firstRow) {
    return firstRow - 1;
  }

  const values = sheet
    .getRange(
      firstRow,
      CONFIG.NAME_COLUMN,
      lastRow - firstRow + 1,
      1
    )
    .getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const name = normalizeName(values[i][0]);

    if (
      name === 'day wise collection' ||
      name === 'total collection by person'
    ) {
      return firstRow + i - 1;
    }

    if (
      i > 0 &&
      isConfiguredSection(values[i][0])
    ) {
      return firstRow + i - 1;
    }
  }

  return lastRow;
}



function getAvailableDateColumnsForSection(sheet, section) {
  const lastColumn = sheet.getLastColumn();

  if (lastColumn === 0) {
    return [];
  }

  const range = sheet.getRange(
    section.headerRow,
    1,
    1,
    lastColumn
  );

  const displayValues = range.getDisplayValues()[0];
  const rawValues = range.getValues()[0];

  const dates = [];

  for (let index = 0; index < displayValues.length; index++) {
    const parsed = parseDateHeader(
      displayValues[index],
      rawValues[index]
    );

    if (!parsed) {
      continue;
    }

    dates.push({
      column: index + 1,
      date: parsed.isoDate,
      label: parsed.label
    });
  }

  dates.sort(function(a, b) {
    return a.date.localeCompare(b.date);
  });

  return dates;
}



function parseDateHeader(displayValue, rawValue) {
  const text = cleanCellText(displayValue);

  const match = text.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/
  );

  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]);
    let year = Number(match[3]);

    if (year < 100) {
      year += 2000;
    }

    return buildDateInfo(year, month, day);
  }

  if (
    Object.prototype.toString.call(rawValue) === '[object Date]' &&
    !isNaN(rawValue.getTime())
  ) {
    const year = Number(
      Utilities.formatDate(rawValue, CONFIG.TIMEZONE, 'yyyy')
    );

    const month = Number(
      Utilities.formatDate(rawValue, CONFIG.TIMEZONE, 'M')
    );

    const day = Number(
      Utilities.formatDate(rawValue, CONFIG.TIMEZONE, 'd')
    );

    return buildDateInfo(year, month, day);
  }

  return null;
}


function buildDateInfo(year, month, day) {
  const date = new Date(year, month - 1, day);

  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }

  const isoDate =
    year + '-' +
    String(month).padStart(2, '0') + '-' +
    String(day).padStart(2, '0');

  return {
    isoDate: isoDate,
    label: Utilities.formatDate(
      date,
      CONFIG.TIMEZONE,
      'MMMM d, yyyy'
    )
  };
}


function checkSale(salesperson, amount, selectedDate, sectionName) {
  try {
    validateConfiguration();

    salesperson = String(salesperson || '').trim();

    if (!salesperson) {
      throw new Error('SALESPERSON_REQUIRED');
    }

    const numericAmount = Number(amount);

    if (
      amount === '' ||
      amount === null ||
      amount === undefined ||
      !Number.isFinite(numericAmount) ||
      numericAmount < 0
    ) {
      throw new Error('INVALID_AMOUNT');
    }

    selectedDate = validateSelectedDate(selectedDate);

    const spreadsheet =
      SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    const sheet = getSalesSheet(spreadsheet);

    const section =
      getSectionDefinition(sheet, sectionName);

    const dayColumn = findDayColumnInSection(
      sheet,
      section,
      selectedDate
    );

    if (dayColumn === -1) {
      throw new Error('DAY_COLUMN_NOT_FOUND');
    }

    const salespersonRow = findSalespersonRowInSection(
      sheet,
      section,
      salesperson
    );

    if (salespersonRow === -1) {
      throw new Error('SALESPERSON_NOT_FOUND');
    }

    if (
      !isEditableDateCell(
        sheet,
        section,
        salespersonRow,
        dayColumn
      )
    ) {
      throw new Error('INVALID_EDIT_CELL');
    }

    const cell = sheet.getRange(salespersonRow, dayColumn);
    const existingValue = readNumericCellValue(cell);

    return {
      success: true,
      section: section.name,
      salesperson: salesperson,
      amountEntered: numericAmount,
      previousValue: existingValue,
      date: formatIsoDateForDisplay(selectedDate),
      selectedDate: selectedDate,
      dayColumn: dayColumn
    };

  } catch (error) {
    console.error(error);

    return {
      success: false,
      message: friendlyErrorMessage(error)
    };
  }
}


function saveSale(
  salesperson,
  amount,
  action,
  selectedDate,
  sectionName
) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;

  try {
    lock.waitLock(15000);
    lockAcquired = true;

    validateConfiguration();

    salesperson = String(salesperson || '').trim();

    if (!salesperson) {
      throw new Error('SALESPERSON_REQUIRED');
    }

    const numericAmount = Number(amount);

    if (
      amount === '' ||
      amount === null ||
      amount === undefined ||
      !Number.isFinite(numericAmount) ||
      numericAmount < 0
    ) {
      throw new Error('INVALID_AMOUNT');
    }

    action = String(action || 'ADD').toUpperCase();

    if (action !== 'ADD' && action !== 'REPLACE') {
      throw new Error('INVALID_ACTION');
    }

    selectedDate = validateSelectedDate(selectedDate);

    const spreadsheet =
      SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    const sheet = getSalesSheet(spreadsheet);

    const section =
      getSectionDefinition(sheet, sectionName);

    const dayColumn = findDayColumnInSection(
      sheet,
      section,
      selectedDate
    );

    if (dayColumn === -1) {
      throw new Error('DAY_COLUMN_NOT_FOUND');
    }

    const salespersonRow = findSalespersonRowInSection(
      sheet,
      section,
      salesperson
    );

    if (salespersonRow === -1) {
      throw new Error('SALESPERSON_NOT_FOUND');
    }

    if (
      !isEditableDateCell(
        sheet,
        section,
        salespersonRow,
        dayColumn
      )
    ) {
      throw new Error('INVALID_EDIT_CELL');
    }

    const cell = sheet.getRange(salespersonRow, dayColumn);

    // Read the latest value inside the lock.
    const existingValue = readNumericCellValue(cell);

    let newValue;

    if (action === 'ADD') {
      newValue = existingValue + numericAmount;
    } else {
      newValue = numericAmount;
    }

    cell.setValue(newValue);
    cell.setNumberFormat('#,##0.##');

    SpreadsheetApp.flush();

    logCollectionSubmission(
      section.name,
      salesperson,
      selectedDate,
      action,
      numericAmount,
      existingValue,
      newValue
    );

    return {
      success: true,
      section: section.name,
      salesperson: salesperson,
      amountEntered: numericAmount,
      previousValue: existingValue,
      newValue: newValue,
      date: formatIsoDateForDisplay(selectedDate),
      selectedDate: selectedDate,
      action: action
    };

  } catch (error) {
    console.error(error);

    return {
      success: false,
      message: friendlyErrorMessage(error)
    };

  } finally {
    if (lockAcquired) {
      try {
        lock.releaseLock();
      } catch (ignore) {
        // Nothing to do.
      }
    }
  }
}



function readNumericCellValue(cell) {
  const rawValue = cell.getValue();

  if (
    rawValue === '' ||
    rawValue === null ||
    rawValue === undefined
  ) {
    return 0;
  }

  if (
    typeof rawValue === 'number' &&
    Number.isFinite(rawValue) &&
    rawValue >= 0
  ) {
    return rawValue;
  }

  const existingText = cleanCellText(
    cell.getDisplayValue()
  );

  if (existingText === '') {
    return 0;
  }

  const parsedValue = Number(
    existingText.replace(/,/g, '')
  );

  if (
    !Number.isFinite(parsedValue) ||
    parsedValue < 0
  ) {
    throw new Error('EXISTING_CELL_NOT_NUMERIC');
  }

  return parsedValue;
}



function findDayColumnInSection(sheet, section, selectedDate) {
  const dateColumns =
    getAvailableDateColumnsForSection(sheet, section);

  for (let i = 0; i < dateColumns.length; i++) {
    if (dateColumns[i].date === selectedDate) {
      return dateColumns[i].column;
    }
  }

  return -1;
}

function findSalespersonRowInSection(
  sheet,
  section,
  salesperson
) {
  const target = normalizeName(salesperson);

  const firstRow = section.headerRow + 1;
  const lastRow = findLastSalespersonRow(sheet, section);

  if (lastRow < firstRow) {
    return -1;
  }

  const names = sheet
    .getRange(
      firstRow,
      CONFIG.NAME_COLUMN,
      lastRow - firstRow + 1,
      1
    )
    .getDisplayValues();

  for (let index = 0; index < names.length; index++) {
    const cellName = cleanCellText(names[index][0]);

    if (
      cellName !== '' &&
      normalizeName(cellName) === target
    ) {
      return firstRow + index;
    }
  }

  return -1;
}



function isEditableDateCell(sheet, section, row, column) {
  if (!sheet || !section) {
    return false;
  }

  //Never allow writing to the Name column.
  if (
    !Number.isInteger(column) ||
    column <= CONFIG.NAME_COLUMN ||
    column > sheet.getLastColumn()
  ) {
    return false;
  }

  if (
    !Number.isInteger(section.headerRow) ||
    !Number.isInteger(section.endRow) ||
    section.headerRow < 1 ||
    section.endRow < section.headerRow
  ) {
    return false;
  }

  const firstSalespersonRow = section.headerRow + 1;

  const lastSalespersonRow =
    findLastSalespersonRow(sheet, section);

  if (
    !Number.isInteger(row) ||
    row < firstSalespersonRow ||
    row > lastSalespersonRow ||
    row > section.endRow
  ) {
    return false;
  }

  const salespersonName = cleanCellText(
    sheet.getRange(row, CONFIG.NAME_COLUMN).getDisplayValue()
  );

  const normalizedSalesperson = normalizeName(salespersonName);

  if (
    !salespersonName ||
    normalizedSalesperson === 'name' ||
    normalizedSalesperson === 'day wise collection' ||
    normalizedSalesperson === 'total collection by person' ||
    isConfiguredSection(salespersonName)
  ) {
    return false;
  }


  const dateColumns =
    getAvailableDateColumnsForSection(sheet, section);

  return dateColumns.some(function(item) {
    return item.column === column;
  });
}



function validateSelectedDate(selectedDate) {
  selectedDate = String(selectedDate || '').trim();

  if (!/^\d{4}-\d{2}-\d{2}$/.test(selectedDate)) {
    throw new Error('INVALID_DATE');
  }

  const parts = selectedDate.split('-');

  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const day = Number(parts[2]);

  const date = new Date(year, month - 1, day);

  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    throw new Error('INVALID_DATE');
  }

  const now = new Date();

  const currentYear = Number(
    Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy')
  );

  const currentMonth = Number(
    Utilities.formatDate(now, CONFIG.TIMEZONE, 'M')
  );

  if (year !== currentYear || month !== currentMonth) {
    throw new Error('INVALID_DATE');
  }

  return selectedDate;
}



function formatIsoDateForDisplay(selectedDate) {
  const parts = selectedDate.split('-');


  const date = new Date(Date.UTC(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2]),
    12,
    0,
    0
  ));

  return Utilities.formatDate(
    date,
    CONFIG.TIMEZONE,
    'MMMM d, yyyy'
  );
}




function cleanCellText(text) {
  return String(text === null || text === undefined ? '' : text)
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}



function normalizeName(name) {
  return cleanCellText(name).toLowerCase();
}



function validateConfiguration() {
  if (!CONFIG.SPREADSHEET_ID) {
    throw new Error('INVALID_SPREADSHEET_ID');
  }

  if (!CONFIG.SHEET_NAME) {
    throw new Error('INVALID_SHEET_NAME');
  }

  if (!CONFIG.ACCESS_CODE) {
    throw new Error('INVALID_ACCESS_CODE');
  }

  if (
    !Array.isArray(CONFIG.SECTIONS) ||
    CONFIG.SECTIONS.length === 0
  ) {
    throw new Error('INVALID_SECTIONS');
  }
}



function friendlyErrorMessage(error) {
  const message =
    error && error.message
      ? String(error.message)
      : '';

  switch (message) {
    case 'INVALID_SPREADSHEET_ID':
      return 'Spreadsheet ID is not configured.';

    case 'INVALID_SHEET_NAME':
      return 'The configured collection sheet was not found.';

    case 'INVALID_ACCESS_CODE':
      return 'Access code is not configured.';

    case 'INVALID_SECTIONS':
      return 'Collection sections are not configured.';

    case 'NO_SHEETS_FOUND':
      return 'No spreadsheet sheet or data was found.';

    case 'NO_COLLECTION_SECTIONS':
      return 'Brand sections were not found in the sheet.';

    case 'SECTION_REQUIRED':
      return 'Please select a collection type.';

    case 'SECTION_NOT_FOUND':
      return 'Selected collection type was not found in the sheet.';

    case 'SECTION_HEADER_NOT_FOUND:BRAND1':
      return 'Brand1 header row was not found.';

    case 'SECTION_HEADER_NOT_FOUND:BRAND2':
      return 'Brand2 header row was not found.';

    case 'SALESPERSON_REQUIRED':
      return 'Please select a salesperson.';

    case 'SALESPERSON_NOT_FOUND':
      return 'The selected salesperson was not found in this collection section.';

    case 'INVALID_AMOUNT':
      return 'Please enter a valid collection amount greater than or equal to 0.';

    case 'INVALID_DATE':
      return 'Please select a valid date from the current month.';

    case 'DAY_COLUMN_NOT_FOUND':
      return 'The selected date column was not found in this section.';

    case 'EXISTING_CELL_NOT_NUMERIC':
      return 'The existing collection cell does not contain a valid number.';

    case 'INVALID_ACTION':
      return 'Invalid save action.';

    case 'INVALID_EDIT_CELL':
      return 'For safety, this cell is not allowed to be changed by the website.';

    case 'NO_SALESPERSONS:Brand1':
      return 'No salesperson rows were found in Brand1.';

    case 'NO_SALESPERSONS:Brand2':
      return 'No salesperson rows were found in Brand2.';

    default:
      if (message.indexOf('Lock timeout') !== -1) {
        return 'Another collection is being saved. Please try again.';
      }

      return message || 'An unexpected server error occurred.';
  }
}



function testExecutionIdentity() {
  Logger.log(
    'Effective user: ' +
    Session.getEffectiveUser().getEmail()
  );
}



function testWebAppIdentity() {
  const user = Session.getActiveUser().getEmail();

  Logger.log('Active user: ' + user);
}
