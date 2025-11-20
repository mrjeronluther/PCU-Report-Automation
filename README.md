# Payables Custodianship Unit Data Consolidation and Reporting Scripts Automation

This repository contains two Google Apps Scripts designed to automate the process of consolidating data from multiple Google Sheets and then generating aggregated summary reports from that consolidated data.

## Table of Contents
- [Project Overview](#project-overview)
- [User Documentation](#user-documentation)
  - [Part 1: Data Consolidation](#part-1-data-consolidation)
    - [Initial Setup](#initial-setup)
    - [Running the Consolidation](#running-the-consolidation)
    - [Resetting the Process](#resetting-the-process)
  - [Part 2: Generating Reports](#part-2-generating-reports)
    - [Prerequisites](#prerequisites)
    - [Generating the Report](#generating-the-report)
- [Developer Documentation](#developer-documentation)
  - [Script 1: Conso Data (`Conso Data.gs`)](#script-1-conso-data-conso-datags)
    - [Configuration Constants](#configuration-constants)
    - [Core Functions](#core-functions)
  - [Script 2: Conso Report Data (`Conso Report Data.gs`)](#script-2-conso-report-data-conso-report-datags)
    - [Main Function](#main-function)
    - [ReportProcessor Class](#reportprocessor-class)

## Project Overview

This solution consists of two primary scripts:

1.  **Conso Data Script**: This script collects data from a user-defined list of source Google Sheets. It is designed to be robust and can handle very large datasets by batching writes, managing Google's execution time limits, and automatically creating new "Part" spreadsheets when the Google Sheets cell limit is reached.
2.  **Conso Report Data Script**: After consolidation, this script processes all the collected data. It deduplicates the records and generates several summary reports, organizing the financial data into different dashboard tabs based on the service category.

---

## User Documentation

This guide explains how to use the "Consolidator" and "Generate Report" scripts in your Google Sheet.

### Part 1: Data Consolidation

This script collects data from a list of Google Sheets and consolidates it into a central location.

#### Initial Setup

Before running the consolidation for the first time, you need to provide the list of source files.

1.  Open the Google Sheet that contains the script.
2.  Navigate to the **`SOURCEFILES`** tab.
3.  In column **B (`File Link`)**, paste the Google Sheet URLs of all the files you want to consolidate.
4.  In column **A (`Tab Names`)**, for each corresponding file link, enter the names of the tabs you want to pull data from.
    > **Note:** If you need to pull data from multiple tabs within the same file, separate the tab names with a comma (e.g., `Sheet1,Sheet2,Sheet3`).

#### Running the Consolidation

Once the `SOURCEFILES` sheet is set up, you can start the process.

1.  Click on the custom **`Consolidator`** menu in the spreadsheet toolbar.
2.  Select **`Start or Resume Consolidation`**.
3.  A dialog box will appear asking you to confirm starting a new consolidation. Click **`Yes`**.
4.  The script will begin processing the files. This can take a long time depending on the amount of data.
5.  If the script's execution time runs out, it will pause and save its progress. To continue, simply click **`Start or Resume Consolidation`** again.

#### Resetting the Process

If you need to start the entire consolidation over from the beginning:

1.  Click on the **`Consolidator`** menu.
2.  Select **`Reset Consolidation Process`**.
3.  Confirm that you want to reset.
    > **Warning:** This action will clear all previously consolidated data in the `Conso` sheet, the `Issue` log, and the central `Extension Tab` log.

### Part 2: Generating Reports

After the data has been consolidated, you can generate summary reports.

#### Prerequisites

*   The data consolidation process must have been completed successfully.
*   The `SOURCEFILES` tab must be correctly filled out, as the report script uses this list to map services to destination tabs.

#### Generating the Report

1.  In the same spreadsheet, click on the **`Consolidator`** menu.
2.  Select **`Generate Report`**.
3.  The script will begin analyzing the consolidated data and creating summary reports in new tabs.
4.  The script will create new dashboard tabs (e.g., `MSPdash`, `PROJdash`) based on the mappings defined in the script.
5.  Each new tab will contain several summary tables:
    *   Summary by Service and Year
    *   Summary by Service and Property
    *   Summary by Service and Supplier/Payee
    *   Summary by Service and Payor Company

> **Note:** If the report generation is interrupted, you can run it again by clicking **`Generate Report`**. It is designed to resume from where it left off.

---

## Developer Documentation

This section provides technical details about the scripts for developers who may need to modify, extend, or troubleshoot the code.

### Script 1: Conso Data (`Conso Data.gs`)

This script handles the aggregation of raw data from various sources.

#### Configuration Constants

These global variables at the top of the script control its core behavior.

*   `LOG_SPREADSHEET_ID`: **String** - The ID of the Google Spreadsheet where a log of newly created consolidated files ("Part" files) is maintained.
*   `LOG_SHEET_NAME`: **String** - The name of the sheet within the log spreadsheet where new file links are stored.
*   `ISSUE_LOG_SHEET_NAME`: **String** - The name of the sheet within the main spreadsheet where errors (e.g., inability to access a source file) are logged.
*   `DESTINATION_FOLDER_ID`: **String** - The ID of the Google Drive folder where new consolidated spreadsheets will be created when the previous one becomes full.
*   `MAX_EXECUTION_TIME_MINUTES`: **Number** - The maximum number of minutes the script will run before pausing and saving its state to prevent execution timeouts.
*   `BATCH_WRITE_SIZE`: **Number** - The number of rows to hold in memory before writing them to the spreadsheet. This improves performance.
*   `GOOGLE_SHEET_CELL_LIMIT`: **Number** - The cell limit for a Google Sheet, used to determine when to create a new consolidation file.

#### Core Functions

*   `onOpen()`: A special trigger that runs when the spreadsheet is opened. It creates the custom `Consolidator` menu in the UI.
*   `createNewConsolidationSheet(state)`: Called when the current `Conso` sheet is full. It creates a new spreadsheet, logs it, and updates the script's state object.
*   `consolidateToMaxPerformance()`: The main consolidation function. It uses `PropertiesService` to manage state, allowing it to be paused and resumed. It iterates through source files, extracts data, handles batch writing, and manages errors and timeouts.
*   `resetConsolidationProcess()`: Clears the script's saved state from `PropertiesService` and clears all relevant data and log sheets.
*   `extractIdFromUrl(url)`: A utility function that uses a regular expression to get a spreadsheet ID from its URL.

### Script 2: Conso Report Data (`Conso Report Data.gs`)

This script generates summary reports from the data collected by the Conso Data script. It is built using a class-based structure for better organization.

#### Main Function

*   `uisadaaa()`: The primary function executed by the user from the menu. It instantiates the `ReportProcessor` class, runs pre-flight checks to ensure the environment is ready, manages the list of processed files for resumability, and triggers the main processing logic.

#### ReportProcessor Class

This class encapsulates all the logic for generating reports.

*   `constructor(spreadsheet, properties)`: Initializes the class and sets key configuration properties, including:
    *   `MASTER_EXTENSION_URL`: URL of the spreadsheet containing the log of all "Part" files.
    *   `ARCHIVE_FOLDER_ID`: Drive folder ID where new archived report files will be created.
    *   `MAX_ROWS_PER_SHEET`: The row limit for a single report tab before a new archive file is created.
    *   `columnNames`: An object mapping friendly names to the exact column headers in the `Conso` data. **This is critical for the script's logic.**
    *   `tabNameMapping`: An object mapping service names from `SOURCEFILES` to the desired destination tab names.
    *   `statusReportMapping`: Defines how raw status values (e.g., "RT ADMIN | COMPLIANCE") are grouped into cleaned report categories (e.g., "WITH ADMIN").

*   `runPreFlightChecks()`: A vital method that verifies the existence of all required resources (folders, files, sheets) and validates that the `Conso` data contains all necessary headers before processing begins.

*   `getAllConsoData()`: A recursive function that gathers all consolidated data. It starts with the local `Conso` sheet and then follows the links in the `Extension Tabs` sheet to fetch data from all "Part" files.

*   `_deduplicateConsoData(...)`: Removes duplicate rows from the aggregated raw data. It creates a unique key for each row based on multiple columns to ensure that only the latest version of a record is used in the report.

*   `processFiles(...)`: The core method that orchestrates report generation. It iterates through the files in `SOURCEFILES`, filters the deduplicated `Conso` data for each, and passes the filtered data to the aggregation logic. It also manages state for pausing and resuming.

*   `aggregateDataFromFile(...)`: Performs the actual financial data aggregation for a single source file, grouping the amounts based on the `statusReportMapping`.

*   `writeAggregatedDataToSheets(...)`: Takes the final aggregated data and writes it to the appropriate destination sheets. It handles formatting, sorting, and the creation of report blocks with titles and headers.

*   `_getOrCreateDestinationSheet(...)`: A robust function that manages where the report data is written. If a local report tab is full, it will find the latest archive file in the `ARCHIVE_FOLDER_ID` or create a new one, ensuring the report can continue without size limitations.
