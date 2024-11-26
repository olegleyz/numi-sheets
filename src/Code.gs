// Configuration
const CONFIG = {
  SHEETS: {
    SETTINGS: 'Settings',
    ADD_CUSTOMER: 'Add Customer',
    TEMPLATE: 'Template'
  },
  CELLS: {
    API_KEY: 'B2',
    API_URL: 'B3'
  }
};

// Add custom menu when the spreadsheet opens
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Numi Foreseeing')
    .addItem('Add New Customer', 'addCustomerFromSheet')
    .addToUi();
    
  // Setup sheets when spreadsheet opens
  setupSheets();
}

// Setup sheets with correct structure
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create Add Customer sheet if it doesn't exist
  let addCustomerSheet = ss.getSheetByName(CONFIG.SHEETS.ADD_CUSTOMER);
  if (!addCustomerSheet) {
    addCustomerSheet = ss.insertSheet(CONFIG.SHEETS.ADD_CUSTOMER);
  }
  setupAddCustomerSheet(addCustomerSheet);
  
  // Create Template sheet if it doesn't exist
  let templateSheet = ss.getSheetByName(CONFIG.SHEETS.TEMPLATE);
  if (!templateSheet) {
    templateSheet = ss.insertSheet(CONFIG.SHEETS.TEMPLATE);
    setupTemplateSheet(templateSheet);
  }
  
  // Create Settings sheet if it doesn't exist
  let settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(CONFIG.SHEETS.SETTINGS);
    const settingsData = [
      ['Setting', 'Value'],
      ['API Key', ''],
      ['API URL', '']
    ];
    settingsSheet.getRange(1, 1, settingsData.length, 2).setValues(settingsData);
    settingsSheet.setFrozenRows(1);
  }
  
  // Delete the default "Sheet1" if it exists
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1) {
    ss.deleteSheet(sheet1);
  }
}

// Setup the Add Customer sheet with form-like interface
function setupAddCustomerSheet(sheet) {
  // Clear existing content and formatting
  sheet.clear();
  sheet.clearFormats();
  
  // Set column widths
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 200);
  
  // Add labels and input cells
  const labels = [
    ['First Name:', 'Enter first name'],
    ['Last Name:', 'Enter last name'],
    ['Year of Birth:', 'YYYY'],
    ['Month of Birth:', 'MM (1-12)'],
    ['Day of Birth:', 'DD (1-31)']
  ];
  
  // Add labels and placeholder text
  for (let i = 0; i < labels.length; i++) {
    const row = i + 2; // Start from row 2
    sheet.getRange(row, 1).setValue(labels[i][0]);
    
    const inputCell = sheet.getRange(row, 2);
    inputCell.setValue(labels[i][1])
             .setBackground('#E2EFDA')  // Light green background
             .setFontColor('#666666')   // Gray text for placeholder
             .setHorizontalAlignment('center'); // Center-align input cells
  }
  
  // Add title
  sheet.getRange(1, 1, 1, 2)
       .merge()
       .setValue('Add New Customer')
       .setHorizontalAlignment('center')
       .setFontWeight('bold')
       .setFontSize(14);
  
  // Add instructions with text wrapping
  const instructionCell = sheet.getRange(7, 1, 1, 2);
  
  // Set cell properties before merging
  instructionCell.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  instructionCell.setRowHeight(45);
  
  // Now merge and set other properties
  instructionCell.merge()
                .setValue('Fill in the details above and check the box below to add the customer')
                .setHorizontalAlignment('left')
                .setVerticalAlignment('middle')
                .setFontStyle('italic')
                .setBackground('#f3f3f3'); // Light grey background to match other formatting

  // Add checkbox label
  sheet.getRange(8, 1).setValue('Add Customer:')
       .setHorizontalAlignment('right')
       .setFontWeight('bold');
  
  // Add checkbox using data validation
  const checkboxCell = sheet.getRange(8, 2);
  const rule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();
  
  checkboxCell.setDataValidation(rule)
              .setValue(false)
              .setBackground('#E2EFDA')  // Light green background to match input fields
              .setHorizontalAlignment('center');
              
  // Create the trigger for the checkbox
  createTrigger();
}

// Setup the template sheet
function setupTemplateSheet(sheet) {
  // Clear existing content and formatting
  sheet.clear();
  sheet.clearFormats();
  
  // Set column widths
  sheet.setColumnWidths(1, 6, 100);
  
  // Add basic info section
  sheet.getRange('A1').setValue('Name');
  sheet.getRange('A2').setValue('Date of Birth');
  
  // Add forecast info section
  sheet.getRange('A3').setValue('Current Solar');
  sheet.getRange('A4').setValue('Next Solar');
  sheet.getRange('A5').setValue('Current Personal Life');
  sheet.getRange('A6').setValue('Next Personal Life');
  
  // Add table headers for months (will be filled dynamically)
  sheet.getRange('A7:F7').setBackground('#f3f3f3');
  sheet.getRange('A8:F8');  // Energy values row
  sheet.getRange('A9:F9').setBackground('#f3f3f3');
  sheet.getRange('A10:F10');  // Energy values row
  
  // Format headers
  sheet.getRange('A1:A6').setFontWeight('bold');
}

// Create an installable trigger for the onEdit function
function createTrigger() {
  // Delete any existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Create a new trigger
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

// Handle Add Customer menu action
function addCustomerFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ADD_CUSTOMER);
  const settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  
  // Get input values
  const firstName = sheet.getRange(2, 2).getValue();
  const lastName = sheet.getRange(3, 2).getValue();
  const yearOfBirth = sheet.getRange(4, 2).getValue();
  const monthOfBirth = sheet.getRange(5, 2).getValue();
  const dayOfBirth = sheet.getRange(6, 2).getValue();
  
  // Check if values are placeholders
  if (firstName === 'Enter first name' || 
      lastName === 'Enter last name' || 
      yearOfBirth === 'YYYY' ||
      monthOfBirth === 'MM (1-12)' ||
      dayOfBirth === 'DD (1-31)') {
    SpreadsheetApp.getUi().alert('Please fill in all fields with actual values');
    return;
  }
  
  // Validate input
  if (!firstName || !lastName || !yearOfBirth || !monthOfBirth || !dayOfBirth) {
    SpreadsheetApp.getUi().alert('Please fill in all fields');
    return;
  }
  
  // Validate date components
  const year = parseInt(yearOfBirth);
  const month = parseInt(monthOfBirth);
  const day = parseInt(dayOfBirth);
  
  if (isNaN(year) || year < 1900 || year > new Date().getFullYear()) {
    SpreadsheetApp.getUi().alert('Please enter a valid year (1900-present)');
    return;
  }
  
  if (isNaN(month) || month < 1 || month > 12) {
    SpreadsheetApp.getUi().alert('Please enter a valid month (1-12)');
    return;
  }
  
  if (isNaN(day) || day < 1 || day > 31) {
    SpreadsheetApp.getUi().alert('Please enter a valid day (1-31)');
    return;
  }
  
  // Format date as YYYY-MM-DD
  const dateOfBirth = `${year}-${month.toString().padStart(2, '0')}-${day.toString().padStart(2, '0')}`;
  
  // Get API configuration
  const apiKey = settingsSheet.getRange(CONFIG.CELLS.API_KEY).getValue();
  const apiUrl = settingsSheet.getRange(CONFIG.CELLS.API_URL).getValue();
  
  if (!apiKey || !apiUrl) {
    SpreadsheetApp.getUi().alert('Please configure API key and URL in the Settings sheet.');
    return;
  }
  
  try {
    // Call add_customer API
    const response = UrlFetchApp.fetch(`${apiUrl}/add_customer`, {
      method: 'POST',
      headers: {
        'x-api-key': apiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        first_name: firstName,
        last_name: lastName,
        date_of_birth: dateOfBirth,
        parameters: {}  // Adding empty parameters dictionary
      }),
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    console.log('API Response:', response.getContentText());
    
    if (responseCode === 200 || responseCode === 201) {
      // Create customer tab after successful API call
      createCustomerTab(firstName, lastName, dateOfBirth);
      

      // Clear input fields
      sheet.getRange(2, 2).setValue('Enter first name');
      sheet.getRange(3, 2).setValue('Enter last name');
      sheet.getRange(4, 2).setValue('YYYY');
      sheet.getRange(5, 2).setValue('MM (1-12)');
      sheet.getRange(6, 2).setValue('DD (1-31)');
      return true;
    } else {
      SpreadsheetApp.getUi().alert('Failed to add customer. Please try again.');
      return false;
    }
  } catch (error) {
    console.error('Error adding customer:', error);
    SpreadsheetApp.getUi().alert('Error adding customer. Please check the logs for details.');
    return false;
  }
}

// Create a new tab for a customer based on the template
function createCustomerTab(firstName, lastName, dateOfBirth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName(CONFIG.SHEETS.TEMPLATE);
  const settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  
  // Create tab name
  const tabName = `${firstName} ${lastName}`;
  
  // Check if tab already exists
  let customerSheet = ss.getSheetByName(tabName);
  if (!customerSheet) {
    // Copy template
    customerSheet = templateSheet.copyTo(ss);
    customerSheet.setName(tabName);
  }
  
  // Get forecast data
  const apiKey = settingsSheet.getRange(CONFIG.CELLS.API_KEY).getValue();
  const apiUrl = settingsSheet.getRange(CONFIG.CELLS.API_URL).getValue();
  
  try {
    const forecast = getForecast({
      firstName: firstName,
      lastName: lastName,
      dateOfBirth: dateOfBirth
    }, apiUrl, apiKey);
    
    if (!forecast) return;
    
    // Log the forecast object for debugging
    console.log('Forecast object:', JSON.stringify(forecast, null, 2));
    
    // Fill in customer info
    customerSheet.getRange('B1').setValue(`${firstName} ${lastName}`);
    customerSheet.getRange('B2').setValue(dateOfBirth);
    
    let msg = 'Current Solar: ' + forecast.current_solar + '\n';
    msg += 'Next Solar: ' + forecast.next_solar + '\n';
    msg += 'Current Personal Life: ' + forecast.current_personal_life + '\n';
    msg += 'Next Personal Life: ' + forecast.next_personal_life + '\n';
    
    console.log('Debug message:', msg);
    
    
    // Fill in forecast info
    customerSheet.getRange('B3').setValue(forecast.current_solar);
    customerSheet.getRange('B4').setValue(forecast.next_solar);
    customerSheet.getRange('B5').setValue(forecast.current_personal_life);
    customerSheet.getRange('B6').setValue(forecast.next_personal_life);
    
    // Fill in monthly energy table
    const months = forecast.energy_of_months;
    if (months && Array.isArray(months)) {
      for (let i = 0; i < months.length; i++) {
        const row = Math.floor(i / 6) * 2 + 7;
        const col = (i % 6) + 1;
        
        // Set month/year
        customerSheet.getRange(row, col).setValue(`${months[i].month}/${months[i].year}`);
        
        // Set energy value with conditional formatting
        const energyCell = customerSheet.getRange(row + 1, col);
        energyCell.setValue(months[i].energy);
        
        if (months[i].is_karmic) {
          energyCell.setFontColor('red');
        }
      }
    } else {
      console.error('No months data or invalid format:', months);
    }
    
    // Activate the new sheet
    customerSheet.activate();
    
  } catch (error) {
    console.error('Error creating customer tab:', error);
    SpreadsheetApp.getUi().alert('Error creating customer tab. Please check the logs for details.');
  }
}

// Handle checkbox change
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  
  // Check if this is our Add Customer checkbox
  if (sheet.getName() === CONFIG.SHEETS.ADD_CUSTOMER && 
      range.getRow() === 8 && 
      range.getColumn() === 2) {
    
    // If checkbox is checked
    if (range.getValue() === true) {
      // Add the customer
      const success = addCustomerFromSheet();
      
      // Uncheck the box
      SpreadsheetApp.flush(); // Make sure all changes are applied
      range.setValue(false);
      
      // Show status message
      if (success) {
        sheet.getRange(9, 1, 1, 2).merge()
             .setValue('✓ Customer added successfully!')
             .setBackground('#E2EFDA')
             .setFontColor('#38761D')
             .setHorizontalAlignment('center');
        
        // Clear the success message after 3 seconds
        Utilities.sleep(3000);
        sheet.getRange(9, 1, 1, 2).clear();
      }
    }
  }
}

// Function to load customers from API
function loadCustomers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(CONFIG.SHEETS.INPUT);
  const settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  
  // Get API configuration
  const apiKey = settingsSheet.getRange(CONFIG.CELLS.API_KEY).getValue();
  const apiUrl = settingsSheet.getRange(CONFIG.CELLS.API_URL).getValue();
  
  console.log('API Config - URL:', apiUrl, 'Key:', apiKey ? 'Present' : 'Missing');
  
  if (!apiKey || !apiUrl) {
    SpreadsheetApp.getUi().alert('Please configure API key and URL in the Settings sheet.');
    return;
  }

  try {
    // Call the get_customers API
    const fullUrl = `${apiUrl}/get_customers`;
    console.log('Calling API:', fullUrl);
    
    const response = UrlFetchApp.fetch(fullUrl, {
      method: 'GET',
      headers: {
        'x-api-key': apiKey
      },
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    console.log('API Response - Status:', responseCode);
    console.log('API Response - Body:', responseText);
    
    if (responseCode !== 200) {
      SpreadsheetApp.getUi().alert('Failed to fetch customers. Please check your API configuration.');
      return;
    }
    
    const responseData = JSON.parse(responseText);
    console.log('Response Data Type:', typeof responseData);
    console.log('Response Data Keys:', Object.keys(responseData));
    
    // Try to find customers array in the response
    let customers = responseData;
    if (responseData.data) {
      customers = responseData.data;
    } else if (responseData.customers) {
      customers = responseData.customers;
    } else if (responseData.results) {
      customers = responseData.results;
    }
    
    console.log('Extracted Customers:', customers);
    console.log('Customers is array?', Array.isArray(customers));
    
    if (!Array.isArray(customers)) {
      console.error('Customers is not an array:', customers);
      SpreadsheetApp.getUi().alert('Invalid data format received from API.');
      return;
    }
    
    // Clear existing data (except headers)
    const lastRow = inputSheet.getLastRow();
    if (lastRow > 1) {
      inputSheet.getRange(2, 1, lastRow - 1, 5).clearContent();
    }
    
    // Populate the sheet with customer data
    if (customers.length > 0) {
      console.log('Sample customer:', customers[0]);
      
      const data = customers.map(customer => {
        const row = [
          customer.firstName || customer.first_name || '',
          customer.lastName || customer.last_name || '',
          customer.dateOfBirth || customer.date_of_birth || '',
          customer.year || '',
          customer.month || ''
        ];
        console.log('Processed row:', row);
        return row;
      });
      
      console.log('Final data to write:', data);
      
      inputSheet.getRange(2, 1, data.length, 5).setValues(data);
      
      // Format date column
      const dateRange = inputSheet.getRange(2, 3, data.length, 1);
      dateRange.setNumberFormat('yyyy-mm-dd');
      
      SpreadsheetApp.getUi().alert(`Successfully loaded ${data.length} customers.`);
    } else {
      console.log('No customers found in response');
      SpreadsheetApp.getUi().alert('No customers found.');
    }
  } catch (error) {
    console.error('Error loading customers:', error);
    console.error('Error details:', error.stack);
    SpreadsheetApp.getUi().alert('Error loading customers. Please check the logs for details.');
  }
}

// Main function to calculate forecast
function calculateForecast() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(CONFIG.SHEETS.INPUT);
  const forecastSheet = ss.getSheetByName(CONFIG.SHEETS.FORECAST);
  const settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  
  // Get API configuration
  const apiKey = settingsSheet.getRange(CONFIG.CELLS.API_KEY).getValue();
  const apiUrl = settingsSheet.getRange(CONFIG.CELLS.API_URL).getValue();
  
  if (!apiKey || !apiUrl) {
    SpreadsheetApp.getUi().alert('Please configure API key and URL in the Settings sheet.');
    return;
  }
  
  // Get customer data
  const customers = getCustomerData(inputSheet);
  
  // Clear previous forecasts
  clearForecast();
  
  // Calculate and display forecasts
  displayForecasts(customers, forecastSheet, apiUrl, apiKey);
}

// Get customer data from input sheet
function getCustomerData(sheet) {
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];
  
  const customers = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[0]) continue; // Skip empty rows
    
    customers.push({
      firstName: row[0],
      lastName: row[1],
      dateOfBirth: formatDate(row[2]),
      year: row[3] || new Date().getFullYear(),
      month: row[4] || new Date().getMonth() + 1
    });
  }
  
  return customers;
}

// Format date to YYYY-MM-DD
function formatDate(date) {
  if (!date) return '';
  const d = new Date(date);
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

// Call API and display forecasts
function displayForecasts(customers, sheet, apiUrl, apiKey) {
  // Set headers
  const headers = [
    'First Name', 'Last Name', 'Date of Birth',
    'Current Solar', 'Current Personal Life',
    'Next Solar', 'Next Personal Life',
    'Month', 'Year', 'Energy', 'Is Karmic'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  let row = 2;
  customers.forEach(customer => {
    try {
      const forecast = getForecast(customer, apiUrl, apiKey);
      if (!forecast) return;
      
      // Display customer info and main forecast
      sheet.getRange(row, 1, 1, 3).setValues([[
        customer.firstName,
        customer.lastName,
        customer.dateOfBirth
      ]]);
      
      sheet.getRange(row, 4, 1, 4).setValues([[
        forecast.current_solar,
        forecast.current_personal_life,
        forecast.next_solar,
        forecast.next_personal_life
      ]]);
      
      // Display monthly energies
      forecast.energy_of_months.forEach(month => {
        sheet.getRange(row, 8, 1, 4).setValues([[
          month.month,
          month.year,
          month.energy,
          month.is_karmic
        ]]);
        row++;
      });
      
    } catch (error) {
      Logger.log(`Error processing forecast for ${customer.firstName} ${customer.lastName}: ${error}`);
    }
  });
}

// Get forecast for a customer
function getForecast(customer, apiUrl, apiKey) {
  try {
    // Build query parameters
    let queryParams = [
      `first_name=${encodeURIComponent(customer.firstName)}`,
      `last_name=${encodeURIComponent(customer.lastName)}`,
      `date_of_birth=${encodeURIComponent(customer.dateOfBirth)}`
    ];
    
    // Only add year and month if they are defined and not "undefined"
    if (customer.year && customer.year !== "undefined") {
      queryParams.push(`year=${encodeURIComponent(customer.year)}`);
    }
    
    if (customer.month && customer.month !== "undefined") {
      queryParams.push(`month=${encodeURIComponent(customer.month)}`);
    }
    
    // Construct URL with query parameters
    const url = `${apiUrl}/get_forecast?${queryParams.join('&')}`;
    
    // Call forecast API
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'x-api-key': apiKey
      },
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    let responseText = response.getContentText();
    console.log('Raw API Response:', responseText);
    
    if (responseCode === 200) {
      console.log('Cleaned Response:', responseText);
      
      try {
        // Parse the response text into a JavaScript object
        const parsedResponse = JSON.parse(responseText);
        console.log('Parsed response type:', typeof parsedResponse);
        console.log('Parsed response keys:', Object.keys(parsedResponse));
        console.log('Full parsed response:', JSON.stringify(parsedResponse, null, 2));
        
        // Verify the structure
        if (typeof parsedResponse === 'object' && parsedResponse !== null) {
          // Extract the forecast data
          const forecast = {
            current_solar: parsedResponse.forecast.current_solar,
            next_solar: parsedResponse.forecast.next_solar,
            current_personal_life: parsedResponse.forecast.current_personal_life,
            next_personal_life: parsedResponse.forecast.next_personal_life,
            energy_of_months: parsedResponse.forecast.energy_of_months
          };
          
          console.log('Extracted forecast:', JSON.stringify(forecast, null, 2));
          return forecast;
        } else {
          console.error('Invalid response structure:', parsedResponse);
          SpreadsheetApp.getUi().alert('Invalid forecast data structure received from API');
          return null;
        }
      } catch (parseError) {
        console.error('JSON Parse Error:', parseError);
        console.error('Response that failed to parse:', responseText);
        SpreadsheetApp.getUi().alert('Failed to parse forecast data');
        return null;
      }
    } else {
      console.error('Error getting forecast. Response:', responseText);
      SpreadsheetApp.getUi().alert('Failed to get forecast. Please try again.');
      return null;
    }
  } catch (error) {
    console.error('Error in getForecast:', error);
    SpreadsheetApp.getUi().alert('Error getting forecast. Please check the logs for details.');
    return null;
  }
}

// Clear forecast sheet
function clearForecast() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEETS.FORECAST);
  sheet.clear();
}
