# SAP Invoice ZBEX Processor

VBA automation that inserts the **ZBEX** output message into SAP invoices using transaction **VF02**.

## Overview

This project automates a repetitive SAP task: opening invoices in VF02, inserting the ZBEX output type in the first available line, configuring spool parameters and saving the document.

## Problem

Manually processing multiple invoices in VF02 to add the ZBEX message is time-consuming and error-prone.

## Solution

The macro:

- Reads invoice numbers from an Excel sheet
- Opens each invoice in transaction VF02
- Inserts **ZBEX** in the first empty output line
- Configures the message (SPOOLONLY, Print Immediately, etc.)
- Saves the invoice
- Removes the processed row from Excel

## Workflow
Excel Sheet (invoice)
↓
Read Invoice Number
↓
SAP Transaction VF02
↓
Insert ZBEX output
↓
Configure Message
↓
Save
↓
Delete row from Excel

## Requirements

- SAP GUI with Scripting enabled
- User logged into SAP
- Excel file with a sheet named `invoice`
- Column A containing the invoice numbers (starting from A2)

## How to Use

1. Open the Excel file
2. Make sure the sheet is named exactly `invoice`
3. Place the invoice numbers in column A (starting at A2)
4. Press `Alt + F11` and import `src/Invoice_ZBEX.bas`
5. Run the macro `Invoice_ZBEX`
6. Keep the SAP window visible (do not minimize it completely)

## Important Notes

- SAP GUI Scripting must be enabled
- The user must be already logged into SAP
- The macro waits while `session.Busy = True`
- If there is no empty line to insert ZBEX, a message is shown and the process stops for that invoice
