# Google Sheets Template Structure

## Input Sheet
Column headers (Row 1):
- A: First Name
- B: Last Name
- C: Date of Birth
- D: Year (optional)
- E: Month (optional)

## Forecast Sheet
(Automatically populated by the script)
Column headers:
- A: First Name
- B: Last Name
- C: Date of Birth
- D: Current Solar
- E: Current Personal Life
- F: Next Solar
- G: Next Personal Life
- H: Month
- I: Year
- J: Energy
- K: Is Karmic

## Settings Sheet
Configuration cells:
- A2: API Key Label
- B2: [Your API Key]
- A3: API URL Label
- B3: [Your API URL]

### Cell Formatting
- Date of Birth: Date format (YYYY-MM-DD)
- Energy values: Number format with 2 decimal places
- Is Karmic: Boolean (TRUE/FALSE)

### Protected Ranges
- Settings sheet: B2 (API Key)
- All headers (Row 1) in each sheet

### Data Validation
- Month: Numbers 1-12
- Year: 4-digit year (YYYY)
- First Name, Last Name: Text only
- Date of Birth: Date format only
