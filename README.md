# GPX File Transfer Automation

A Python desktop application for automating the selection and copying of
GPX files using verified parcel records from an Excel workbook.

The application is intended to replace the current Excel VBA workflow.
Instead of editing source and destination folder paths in a macro for
every run, users select the workbook and folders through a graphical
interface.

> **Project status:** Proposed / Initial development\
> The repository structure and commands below are the recommended
> starting point. Update this README as features are implemented and
> tested.

## Objectives

-   Read verified parcel IDs from an Excel workbook.
-   Let users select the worksheet and the ID column.
-   Generate GPX filenames by appending `.gpx` to each valid ID.
-   Find matching GPX files in a selected source folder.
-   Copy matching files to a selected destination folder.
-   Report missing files, duplicate IDs, existing destination files, and
    copy errors.
-   Export a detailed Excel report after each transfer.

## Proposed technology stack

  Technology     Purpose
  -------------- -------------------------------------------------
  Python 3.12+   Main programming language
  Tkinter        Desktop graphical user interface
  OpenPyXL       Read Excel workbooks and generate Excel reports
  `pathlib`      File and directory path handling
  `shutil`       Copy GPX files
  `logging`      Application activity and error logging
  PyInstaller    Package the application as a Windows executable

Tkinter, `pathlib`, `shutil`, and `logging` are included with standard
Python installations. OpenPyXL and PyInstaller are installed separately.

## How it works

1.  Launch the GPX File Transfer application.
2.  Select the Excel workbook containing verified parcel records.
3.  Select the worksheet and the column containing the RSBSA ID or
    GeoRef ID.
4.  The application reads the IDs and generates filenames such as
    `123456789.gpx`.
5.  Select the folder containing the existing GPX files.
6.  Select the destination folder for the copied files.
7.  Start the transfer and monitor its progress.
8.  Review the transfer summary and the generated Excel report.

## Proposed application workflow

``` text
Verified Parcels Excel Workbook
              |
              v
   Select Worksheet and ID Column
              |
              v
      Extract and Validate IDs
              |
              v
     Generate filenames (.gpx)
              |
              v
       Select Source Folder
              |
              v
     Select Destination Folder
              |
              v
       Find and Copy GPX Files
              |
              v
     Display Results and Summary
              |
              v
       Generate Excel Report
```

## Main features

  -----------------------------------------------------------------------
  Feature                 Description             Priority
  ----------------------- ----------------------- -----------------------
  Excel workbook          Browse for the verified High
  selection               parcels workbook        

  Worksheet selection     Choose the worksheet to High
                          process                 

  ID column selection     Choose RSBSA ID or      High
                          GeoRef ID               

  GPX filename generation Append `.gpx` to each   High
                          valid ID                

  Source folder selection Choose the folder       High
                          containing GPX files    

  Destination folder      Choose where matching   High
  selection               files will be copied    

  File matching and       Copy matching GPX files High
  copying                 to the destination      

  Transfer summary        Show copied, missing,   High
                          skipped, and failed     
                          counts                  

  Excel report            Export per-file         High
                          processing results      

  Progress indicator      Show transfer progress  Medium

  Duplicate detection     Identify repeated IDs   Medium
                          in the workbook         

  Activity logging        Record processing       Medium
                          events and errors       
  -----------------------------------------------------------------------

## Recommended repository structure

``` text
gpx-file-transfer-automation/
|
├── app/
│   ├── __init__.py
│   ├── gui.py
│   ├── excel_handler.py
│   ├── file_handler.py
│   └── report_generator.py
|
├── docs/
│   └── project.md
|
├── logs/
|
├── output/
|
├── main.py
├── requirements.txt
├── .gitignore
└── README.md
```

Module responsibilities:

-   `main.py` --- application entry point.
-   `app/gui.py` --- Tkinter window, controls, dialogs, and progress
    display.
-   `app/excel_handler.py` --- workbook, worksheet, and ID extraction.
-   `app/file_handler.py` --- filename matching, validation, and GPX
    copying.
-   `app/report_generator.py` --- Excel transfer report generation.
-   `docs/project.md` --- project scope, requirements, workflow, and
    development plan.

## Requirements

-   Windows 10 or Windows 11
-   Python 3.12 or newer
-   A verified parcel workbook in `.xlsx` format
-   Access to the source folder containing GPX files
-   Permission to read the source folder and write to the destination
    folder

### Supported Excel formats

The initial version should support `.xlsx`. Support for `.xlsm` can be
added if needed. OpenPyXL does not read legacy `.xls` files; convert
those workbooks to `.xlsx` before processing.

## Installation

### 1. Clone the repository

``` bash
git clone https://github.com/<your-username>/gpx-file-transfer-automation.git
cd gpx-file-transfer-automation
```

Replace `<your-username>` with the GitHub account or organization that
will own the repository.

### 2. Create a virtual environment

``` bash
python -m venv .venv
```

### 3. Activate the virtual environment

Windows CMD:

``` bat
.venv\Scripts\activate
```

PowerShell:

``` powershell
.\.venv\Scripts\Activate.ps1
```

### 4. Install dependencies

``` bash
pip install -r requirements.txt
```

The initial `requirements.txt` should contain:

``` text
openpyxl
```

For packaging the application, install PyInstaller when needed:

``` bash
pip install pyinstaller
```

## Running the application

Once the application entry point and GUI have been implemented, run:

``` bash
python main.py
```

The application should open a Tkinter window where users can select the
workbook, worksheet, ID column, source folder, and destination folder.

## Transfer behavior and file safety

The application should follow these rules:

1.  **Copy, do not move.** Use `shutil.copy2()` so original GPX files
    remain in the source folder.
2.  **Do not overwrite by default.** If a file already exists in the
    destination, record it as `Already Exists`.
3.  **Validate IDs.** Ignore blank values and trim surrounding
    whitespace.
4.  **Preserve identifiers.** Treat IDs as text where possible so
    leading zeros are not lost.
5.  **Handle duplicates.** Detect duplicate IDs and avoid copying the
    same requested file repeatedly.
6.  **Continue after errors.** Record a failed file and continue
    processing the remaining files.
7.  **Validate folders.** Confirm that the source exists and that the
    destination can be created or written to.

> **Important data note:** If Excel has already stored an ID as a number
> and removed its leading zeros, the application cannot know the
> original value unless the expected ID length or another formatting
> rule is defined. Confirm the ID format using a sample workbook before
> implementation.

## Transfer report

Generate an Excel report for each run. A suggested filename is:

``` text
GPX_Transfer_Report_YYYY-MM-DD_HHMMSS.xlsx
```

Recommended report columns:

  Column             Description
  ------------------ -------------------------------------
  ID                 Original RSBSA ID or GeoRef ID
  Filename           Generated GPX filename
  Source Path        Expected source file location
  Destination Path   Target file location
  Status             Result of the file operation
  Remarks            Additional details or error message

Recommended statuses:

-   `Copied` --- file was copied successfully.
-   `File Not Found` --- no matching file was found in the source
    folder.
-   `Already Exists` --- a file with the same name already exists in the
    destination.
-   `Skipped` --- record was blank, duplicated, or otherwise excluded.
-   `Error` --- file operation failed; see Remarks for details.

## Testing checklist

-   [ ] Select a valid Excel workbook.
-   [ ] Select the correct worksheet.
-   [ ] Select the correct ID column.
-   [ ] Confirm generated filenames are correct.
-   [ ] Test IDs with leading zeros.
-   [ ] Test blank IDs and IDs with surrounding spaces.
-   [ ] Test duplicate IDs.
-   [ ] Test a GPX file that exists in the source folder.
-   [ ] Test a GPX file that is missing from the source folder.
-   [ ] Test a file that already exists in the destination folder.
-   [ ] Test invalid or inaccessible folders.
-   [ ] Confirm original GPX files remain in the source folder.
-   [ ] Confirm report statuses and summary totals are accurate.
-   [ ] Confirm the GUI remains responsive during larger transfers.

## Packaging for Windows

After the application has been implemented and tested, package it with
PyInstaller:

``` bash
pyinstaller --onefile --windowed --name GPX_Transfer main.py
```

The executable will be created in:

``` text
dist/
└── GPX_Transfer.exe
```

Test the executable on the intended office computers before distributing
it.

## Development roadmap

1.  **Project setup** --- create the repository, virtual environment,
    and application structure.
2.  **Tkinter interface** --- build the main window and selection
    controls.
3.  **Excel processing** --- read the workbook and extract IDs.
4.  **Filename generation** --- validate IDs and generate `.gpx`
    filenames.
5.  **File transfer** --- locate and copy matching GPX files.
6.  **Reporting** --- generate the Excel report and transfer summary.
7.  **Testing** --- test normal, missing-file, duplicate, and error
    scenarios.
8.  **Packaging** --- create and validate the Windows executable.

## Definition of done

The first release is complete when users can select a workbook,
worksheet, ID column, source folder, and destination folder without
editing code; matching GPX files are copied without changing the
originals; exceptions are reported; and a reliable Excel report is
generated for each run.

## License

Choose and add a license before making the repository public or
distributing the application outside the intended organization.
