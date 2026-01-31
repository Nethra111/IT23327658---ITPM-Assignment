# IT23327658 – Playwright Automation (SwiftTranslator)

This repository contains a complete Playwright test project that automates identified scenarios for swifttranslator.com and records execution results back into the provided Excel test case template.

## What This Project Does
- Loads test cases from the Excel file:
  - Functional: Pos_Fun_####, Neg_Fun_####
  - UI: UI_Fun_####
- Executes all identified scenarios against https://www.swifttranslator.com/
- Writes execution results into a new Excel output file:
  - Actual output
  - Status (Pass/Fail)
- Generates an HTML report after every run

## Project Structure
IT23327658/
- tests/
  - IT2332765_translator.spec.js
  - IT2332765_ui.spec.js
- playwright.config.cjs
- IT23327658 - ITPM Excel.xlsx
- package.json

## Prerequisites
- Node.js (LTS recommended)
- npm
- Internet connection

## Install
cd /Users/tharushinethra/Desktop/IT23327658
npm i -D @playwright/test exceljs xlsx
npx playwright install

## Run Tests
npx playwright test --config=playwright.config.cjs --workers=1 --reporter=html

## View Report
npx playwright show-report

## Debug a Single Test
PWDEBUG=1 npx playwright test --config=playwright.config.cjs --workers=1 --headed -g Pos_Fun_0001
PWDEBUG=1 npx playwright test --config=playwright.config.cjs --workers=1 --headed -g UI_Fun_0001

## Output Files
After running tests, the executed Excel will be generated in the project root:
- IT23327658 - ITPM Excel_EXECUTED.xlsx

This file includes updated:
- Actual output
- Status

## Notes
- workers=1 is used to avoid parallel write conflicts when updating the Excel output file.
- Test cases are read from the " Test cases" sheet inside the Excel file.
