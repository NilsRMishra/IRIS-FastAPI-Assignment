# FastAPI Excel Processor Assignment - IRIS Assignment

## Overview

This project is a FastAPI application designed to process and expose data from an Excel file (`capbudg.xls`) through a RESTful API.

Unlike typical structured Excel files, the given file contains a **single worksheet (`CapBudgWS`)** where **multiple logical tables are embedded horizontally as well as vertically**. Each "table" begins with a capitalized section header like `"INITIAL INVESTMENT"`, `"SALVAGE VALUE"`, etc.

The app parses these logical sections and provides endpoints to:
- List available tables
- Fetch row labels from a table
- Compute the sum of numeric values in any specific row

---

## File Structure

```bash
|--Data
  |--capbudg.xls # Provided Excel file
|-- main.py # FastAPI app
|-- notebook.ipynb # Trial codes for preprocessing
|-- utils.py # Utility functions
|-- README.md # This file
```

---

## Features

### 1. Load Excel Data

Automatically reads and splits `capbudg.xls` into multiple logical tables based on capitalized section headers.

---

### 2. API Endpoints

> **Base URL:** `http://localhost:9090`

#### `GET /list_tables`
Returns all logical tables parsed from the Excel sheet.

**Example Response:**
```json
{
  "tables": [
    "INITIAL INVESTMENT",
    "CASHFLOW DETAILS",
    "DISCOUNT RATE",
    "WORKING CAPITAL",
    "GROWTH RATES",
    "SALVAGE VALUE",
    "OPERATING CASHFLOWS",
    "Investment Measures",
    "BOOK VALUE & DEPRECIATION"
  ]
}
```

#### `GET /get_table_details?table_name=SALVAGE VALUE`
Returns the row labels (table keys) of the selected table.

**Example Response:**
```json
{
  "table_name": "SALVAGE VALUE",
  "row_names": [
    "Equipment",
    "Working Capital"
  ]
}
```

#### `GET /row_sum?table_name=OPERATING CASHFLOWS&row_name=Revenues`
Calculates and returns the numeric sum of all values in the specified row.

**Example Response:**
```json
{
  "table_name": "OPERATING CASHFLOWS",
  "row_name": "Revenues",
  "sum": "537024.00"
}
```

---

## Steps for Execution

### Prerequisites

#### `1. Install dependencies:`
pip install fastapi uvicorn pandas xlrd

#### `2. Run the App:`
uvicorn main:app --reload --port 9090

#### `3. Access in Browser using:`
http://localhost:9090/{API Endpoints}


## Potential Improvements

#### 1. Handle Excel uploads dynamically via a POST /upload_excel endpoint.

#### 2. Add API Endpoints to fetch different statistical summaries (mean, std, etc.), `docs` with example responses and error models and more.

## Known Edge Cases

#### 1. Table Names and Row labels must exactly match what's present in the data.

#### 2. Table name label duplicated multiple times — currently returns only the second match.

<!-- #### 3. Formulas or merged cells in Excel might not behave as expected unless handled via a better file processing engine. -->
