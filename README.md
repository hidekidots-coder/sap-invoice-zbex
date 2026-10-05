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
