# Numi Sheets Integration

This project provides Google Sheets integration for the Numi Foreseeing application. It includes:

1. Google Sheets Template
2. Apps Script Extension
3. API Integration Code

## Setup Instructions

1. Create a copy of the template spreadsheet
2. Open Script Editor in Google Sheets (Extensions > Apps Script)
3. Copy the contents of `src/Code.gs` into the Script Editor
4. Save and authorize the script
5. Refresh the spreadsheet to see the new menu items

## Template Structure

The template includes the following sheets:
1. Input Sheet: Customer information entry
2. Forecast Sheet: Displays calculated forecasts
3. Settings Sheet: API configuration

## Usage

1. Enter customer information in the Input sheet
2. Use the custom menu to trigger forecast calculations
3. View results in the Forecast sheet

## Configuration

1. Get your API key from the Numi Foreseeing application
2. Enter the API key in the Settings sheet
3. Configure the API endpoint URL in the Settings sheet
