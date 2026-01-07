# Amista DB Tool

A Windows desktop application for performing batch operations on SAP Business One databases using data from Excel files.

## Overview

Amista DB Tool connects to SAP Business One via the DI API (SAPbobsCOM) and executes batch data operations based on Excel file input. It provides a simple GUI with job selection, execution logging, and progress tracking.

## Features

- **SAP Business One Integration** - Connects via DI API with support for HANA and SQL Server databases
- **Excel File Processing** - Reads `.xls` and `.xlsx` files using ExcelDataReader
- **Real-time Logging** - Displays operation progress with timestamps
- **Progress Tracking** - Visual progress bar for batch operations

## Available Jobs

### Job 1 - Delete BP Catalogs from XLS

Deletes Business Partner Catalog entries (Alternate Catalog Numbers) from SAP based on an Excel file containing:

| Column A | Column B |
|----------|----------|
| CardCode | ItemCode |

The job queries the `OSCN` table to find matching catalog entries and removes them.

**Sample file:** `Job1Sample.xlsx`

## Requirements

- Windows OS
- .NET 6.0 Runtime (x64)
- SAP Business One DI API (SAPBusinessOneSDK.dll)
- SAP Business One client installed and configured

## Configuration

Edit `appsettings.json` with your SAP connection details:

```json
{
  "SapConnection": {
    "Server": "YOUR_DB_SERVER",
    "DbServerType": "dst_HANADB",
    "DBUserName": "SYSTEM",
    "DBPassword": "YOUR_DB_PASSWORD",
    "CompanyDB": "YOUR_COMPANY_DB",
    "UserName": "YOUR_SAP_USER",
    "Password": "YOUR_SAP_PASSWORD",
    "LicenseServer": "YOUR_LICENSE_SERVER",
    "SLDServer": "YOUR_SLD_SERVER"
  }
}
```

**DbServerType options:**
- `dst_HANADB` - SAP HANA
- `dst_MSSQL2019` - SQL Server 2019
- `dst_MSSQL2017` - SQL Server 2017
- `dst_MSSQL2016` - SQL Server 2016

## Building

```bash
dotnet restore
dotnet build
```

## Usage

1. Launch the application
2. Select the job to execute (e.g., "Job 1 - Delete BP Catalogs from XLS")
3. Click **Execute**
4. Select the Excel file when prompted
5. Monitor progress in the log window
6. Use **Copy to Clipboard** to copy the log output

## Project Structure

```
AmistaDBTool/
├── AmistaDBTool.csproj    # Project file
├── Program.cs             # Application entry point
├── MainForm.cs            # Main UI form
├── JobManager.cs          # Job execution logic
├── SapConnector.cs        # SAP DI API connection handler
├── appsettings.json       # Configuration file
└── SAPBusinessOneSDK.dll  # SAP Business One SDK
```

## Logging

- **debug.log** - Application startup and connection debug info
- **crash_log.txt** - Unhandled exception details

## License

Proprietary - Amista USA
