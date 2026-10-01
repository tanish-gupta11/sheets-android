# Sheets for Android

A small, offline spreadsheet editor for Android 8.0 and newer. The installable APK is [sheets.apk](releases/sheets.apk) (about 2.9 MB). It has no network, account, advertising, or analytics permission.

## What works

- Open and export Excel `.xlsx`, legacy Excel `.xls`, and CSV files through Android's file picker
- Edit cells in a touch grid with multiple worksheets
- Evaluate arithmetic, references, ranges, and common functions including `SUM`, `AVERAGE`, `MIN`, and `MAX`
- Undo/redo, find/replace, column sorting, insert/delete rows and columns
- Bold text and highlighted cell fills
- Keep a local autosave; use **Save** to export a file that can be opened in other spreadsheet apps

Worksheets are limited to 10,000 rows and 256 columns. This Android app was built independently; the Windows Sheets 0.4.0 installer was used as a feature reference and is not bundled here.

## Compatibility limits

This is a lightweight editor, not a full Excel or Windows Sheets replacement. Imported charts, pivot tables, merged cells, macros, conditional formatting, rich number/date formats, and formulas beyond the app's supported set are not preserved when exporting. Opening an Excel workbook shows this warning; exporting always creates a new file so the original is left untouched. Use a copy for valuable or complex workbooks.

## Build from source

Open the project in Android Studio, install Android SDK Platform 35, and run `:app:testDebugUnitTest` and `:app:assembleDebug`. The Gradle project downloads Apache POI HSSF for legacy `.xls` support. See [third-party components](THIRD_PARTY.md).

The APK in `releases/` is signed with a dedicated release certificate. Its private signing key is intentionally excluded from this public repository; an update to the same Android package must use that key. The key and its password backup should be kept safe by the owner.

## Verification

- Java workbook tests cover formulas, CSV quoting, and XLS/XLSX round trips.
- XLS and XLSX files created by independent libraries were imported in JVM tests.
- On a connected Android 16 device, the app launched, edited a cell, calculated a formula, imported XLS and XLSX, and exported both formats. The exported files opened with independent spreadsheet readers.
- The release APK was installed and launched on that device; Android package signatures were verified.

The app's own source is licensed under [MIT](LICENSE). Apache POI and its dependencies retain their respective Apache 2.0 licenses.
