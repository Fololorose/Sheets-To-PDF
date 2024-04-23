# Excel Sheets Merger and Exporter to PDF

## Overview:
This VBA script facilitates merging multiple Excel files into a single workbook and exporting all sheets to a PDF document. Each sheet from the selected Excel files is copied into a new workbook, and then the entire workbook is exported as a PDF.

## Requirements:
- Microsoft Excel

## Usage:
1. Press `Alt + F11` to open the Visual Basic for Applications (VBA) editor.
2. Insert a new module by selecting `Insert > Module`.
3. Copy and paste the provided VBA script into the module.
4. Close the VBA editor.
5. Press `Alt + F8` to open the "Macro" dialog box.
6. Select the `MergeAndExportSheetsToPDF` macro.
7. Click `Run`.

## Instructions:
1. Upon running the macro, a file picker dialog will appear.
2. Navigate to and select the Excel files (*.xlsx) you want to merge and export.
3. The script will open each selected Excel file, copy each sheet into a new workbook, and rename the sheets accordingly.
4. After merging all sheets into a single workbook, the script will prompt you to save the resulting PDF file.
5. Specify the file path and name for the PDF file and click "Save".

## Note:
- Ensure that macros are enabled in Excel for the script to run successfully.
- The script will merge all sheets from the selected Excel files into a
